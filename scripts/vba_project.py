# -*- coding: utf-8 -*-
"""
Génération d'un projet VBA binaire (« vbaProject.bin ») sans Excel.

Le fichier produit est un conteneur [MS-CFB] (Compound File Binary, version 3,
secteurs de 512 octets) qui contient un projet VBA conforme à [MS-OVBA] :

    PROJECT              flux texte (page de codes 1252) : ID, Document=, Module=, Name, ...
    PROJECTwm            correspondance noms MBCS / noms UTF-16 des modules
    VBA/                 storage
        _VBA_PROJECT     en-tête minimal, Version = 0xFFFF, sans cache de p-code
        dir              enregistrements d'information du projet, compressés
        <module>         un flux par module : source compressée (MODULEOFFSET = 0)

Comme le flux _VBA_PROJECT ne contient aucun cache (Version 0xFFFF), l'hôte
(Excel) doit recompiler le projet à partir du code source stocké dans les flux
de modules.  C'est la démarche décrite par [MS-OVBA] pour un écrivain tiers.

Le format de chaque élément reproduit celui observé dans un vbaProject.bin réel
produit par Excel (exemple fourni avec XlsxWriter 3.2.9), à l'exception des
données de cache (p-code, flux __SRP_*), volontairement absentes.

Dépendances : bibliothèque standard Python uniquement.

Point d'entrée :

    build_vba_project_bin(std_modules, doc_modules, project_name="VBAProject") -> bytes

    std_modules : {nom_module: source_vba (str)}            -> modules standard (.bas)
    doc_modules : [(nom_module, "workbook" | "worksheet"), ...]
                  ou [(nom_module, type, source_vba), ...]   -> modules de document
                  Les noms doivent être identiques aux codeName du classeur
                  (<workbookPr codeName="..."/> et <sheetPr codeName="..."/>).
"""

import hashlib
import re
import struct

__all__ = [
    "build_vba_project_bin",
    "compress",
    "decompress",
    "encrypt_data",
    "decrypt_data",
    "CODEPAGE",
]

# ---------------------------------------------------------------------------
# Constantes du projet (valeurs relevées dans le vbaProject.bin réel d'Excel)
# ---------------------------------------------------------------------------

CODEPAGE = 1252                 # PROJECTCODEPAGE : Windows-1252 (Europe occidentale)
_PY_CODEC = "cp1252"
_SYSKIND = 0x00000001           # PROJECTSYSKIND : 1 = Windows 32 bits (comme la référence)
_LCID = 0x00000409              # PROJECTLCID : MUST be 0x409
_LCID_INVOKE = 0x00000409       # PROJECTLCIDINVOKE : MUST be 0x409
_VERSION_MAJOR = 0x52671F51     # PROJECTVERSION (valeurs de la référence ; aucune
_VERSION_MINOR = 0x0020         #   signification tant qu'aucun cache n'est présent)
_COOKIE = 0xFFFF                # PROJECTCOOKIE / MODULECOOKIE : ignorés à la lecture

# Références de bibliothèques, octet pour octet identiques à la référence Excel.
# (nom, LibidRegistered) ; Libid au format « *\G{GUID}#version#lcid#chemin#description ».
_REFERENCES = (
    ("stdole",
     "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#"
     "C:\\WINDOWS\\system32\\stdole2.tlb#OLE Automation"),
    ("Office",
     "*\\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#"
     "C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE12\\MSO.DLL#"
     "Microsoft Office 12.0 Object Library"),
)

# Attributs des modules de document, dans l'ordre et avec les valeurs du fichier
# de référence.  VB_Base désigne la classe Excel dont le module dérive :
#   {00020819-...} = Workbook, {00020820-...} = Worksheet.
_DOC_BASE = {
    "workbook": "0{00020819-0000-0000-C000-000000000046}",
    "worksheet": "0{00020820-0000-0000-C000-000000000046}",
}
_DOC_ATTRIBUTES = (
    ("VB_GlobalNameSpace", "False"),
    ("VB_Creatable", "False"),
    ("VB_PredeclaredId", "True"),
    ("VB_Exposed", "True"),
    ("VB_TemplateDerived", "False"),
    ("VB_Customizable", "True"),
)

# Ligne [Host Extender Info] écrite par Excel (extension « VBE »).
_HOST_EXTENDER = "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000"

# Noms interdits pour un module : flux réservés du storage VBA et noms déjà
# utilisés par l'hôte ou les bibliothèques référencées (conflit de noms VBA).
_RESERVED_NAMES = {"dir", "vba", "excel", "stdole", "office", "msforms"}

# Identifiant VBA restreint à l'ASCII (évite toute ambiguïté de casse dans
# l'arbre rouge-noir du CFB) et limité à 31 caractères (limite des noms CFB).
_IDENT_RE = re.compile(r"[A-Za-z][A-Za-z0-9_]{0,30}\Z")
_VB_NAME_RE = re.compile(r'Attribute\s+VB_Name\s*=\s*"([^"]*)"\s*\Z', re.IGNORECASE)
_MAX_LINE = 1023                # longueur maximale d'une ligne dans l'éditeur VBA
_MODULE_ATTR_RE = re.compile(
    r"\s*Attribute\s+VB_(Base|GlobalNameSpace|Creatable|PredeclaredId|Exposed|"
    r"TemplateDerived|Customizable)\s*=", re.IGNORECASE)


# ===========================================================================
# 1. Compression [MS-OVBA] 2.4.1
# ===========================================================================
#
# Conteneur compressé = octet de signature 0x01 suivi de « chunks ».  Chaque
# chunk couvre au plus 4096 octets décompressés et commence par un en-tête de
# 16 bits : bits 0-11 = taille du chunk - 3, bits 12-14 = 0b011, bit 15 = 1 si
# compressé.  Les données compressées sont des séquences de 8 jetons précédées
# d'un octet de drapeaux (bit à 0 = littéral, bit à 1 = CopyToken de 2 octets).
#
# L'implémentation suit l'algorithme normatif de la spécification (2.4.1.3.7 à
# 2.4.1.3.10, 2.4.1.3.19) : même recherche de correspondance (« Matching »),
# donc même sortie que l'algorithme de référence.  La recherche est seulement
# accélérée par un index des préfixes de 3 octets, ce qui ne change pas le
# résultat : une correspondance retenue doit faire au moins 3 octets, et parmi
# les plus longues c'est la plus proche qui est gardée (comparaison stricte en
# parcourant les candidats du plus proche au plus lointain).

_CHUNK = 4096


def _copytoken_help(difference):
    """CopyToken Help (2.4.1.3.19.1) : renvoie (bit_count, length_mask, maximum_length)."""
    # BitCount = plafond(log2(difference)), au minimum 4.
    bit_count = max((difference - 1).bit_length(), 4)
    length_mask = 0xFFFF >> bit_count
    maximum_length = length_mask + 3
    return bit_count, length_mask, maximum_length


def _compress_chunk(data, start, end):
    """Compresse data[start:end] (au plus 4096 octets) ; renvoie le CompressedChunk complet."""
    out = bytearray(b"\x00\x00")       # en-tête réservé, rempli à la fin
    cend = 4098                        # CompressedEnd relatif au début du chunk
    cur = start                        # DecompressedCurrent
    index = {}                         # préfixe de 3 octets -> positions antérieures
    indexed = start                    # positions < indexed déjà indexées

    def _index_up_to(limit):
        nonlocal indexed
        while indexed < limit:
            if indexed + 3 <= end:
                index.setdefault(data[indexed:indexed + 3], []).append(indexed)
            indexed += 1

    while cur < end and len(out) < cend:
        flag_pos = len(out)            # « Compressing a TokenSequence »
        out.append(0)
        flags = 0
        for bit in range(8):
            if not (cur < end and len(out) < cend):
                break
            # --- Matching ---
            best_len, best_cand = 0, 0
            if cur + 3 <= end:
                _index_up_to(cur)
                limit = end - cur
                for cand in reversed(index.get(data[cur:cur + 3], ())):
                    n = 3
                    while n < limit and data[cand + n] == data[cur + n]:
                        n += 1
                    if n > best_len:
                        best_len, best_cand = n, cand
                        if n == limit:     # aucune correspondance plus longue possible
                            break
            if best_len >= 3:
                bit_count, _, max_len = _copytoken_help(cur - start)
                length = min(best_len, max_len)
                offset = cur - best_cand
                if len(out) + 1 < cend:
                    token = ((offset - 1) << (16 - bit_count)) | (length - 3)
                    out += struct.pack("<H", token)
                    flags |= 1 << bit
                    cur += length
                else:
                    out.extend(b"\x00" * (cend - len(out)))   # CompressedCurrent = CompressedEnd
                    break
            else:
                if len(out) < cend:
                    out.append(data[cur])
                    cur += 1
                else:
                    break
        out[flag_pos] = flags

    if cur < end:
        # « Compressing a RawChunk » : 4096 octets bruts (complétés par des 0x00).
        raw = bytes(data[start:end]).ljust(_CHUNK, b"\x00")
        return struct.pack("<H", 0x3000 | (_CHUNK + 2 - 3)) + raw
    size = len(out)
    out[0:2] = struct.pack("<H", 0x8000 | 0x3000 | ((size - 3) & 0x0FFF))
    return bytes(out)


def compress(data):
    """Compresse des octets selon [MS-OVBA] 2.4.1 (CompressedContainer)."""
    data = bytes(data)
    out = bytearray(b"\x01")
    pos = 0
    while pos < len(data):
        end = min(pos + _CHUNK, len(data))
        out += _compress_chunk(data, pos, end)
        pos = end
    return bytes(out)


def decompress(container):
    """Décompresse un CompressedContainer [MS-OVBA] 2.4.1.3.1 (contrôles stricts)."""
    container = bytes(container)
    if not container or container[0] != 0x01:
        raise ValueError("signature de conteneur compressé absente (0x01 attendu)")
    out = bytearray()
    pos = 1
    while pos < len(container):
        if pos + 2 > len(container):
            raise ValueError("en-tête de chunk tronqué à l'offset %d" % pos)
        header = struct.unpack_from("<H", container, pos)[0]
        size = (header & 0x0FFF) + 3
        if (header >> 12) & 0x7 != 0b011:
            raise ValueError("signature de chunk invalide à l'offset %d" % pos)
        compressed = header >> 15
        chunk_end = min(len(container), pos + size)
        chunk_start = len(out)
        cur = pos + 2
        if compressed:
            while cur < chunk_end:
                flags = container[cur]
                cur += 1
                for bit in range(8):
                    if cur >= chunk_end:
                        break
                    if not (flags >> bit) & 1:
                        out.append(container[cur])
                        cur += 1
                    else:
                        if cur + 2 > chunk_end:
                            raise ValueError("CopyToken tronqué")
                        token = struct.unpack_from("<H", container, cur)[0]
                        cur += 2
                        bit_count, length_mask, _ = _copytoken_help(len(out) - chunk_start)
                        length = (token & length_mask) + 3
                        offset = (token >> (16 - bit_count)) + 1
                        src = len(out) - offset
                        if src < chunk_start:
                            raise ValueError("CopyToken hors du chunk")
                        for k in range(length):
                            out.append(out[src + k])
        else:
            if size != 4098:
                raise ValueError("chunk non compressé de taille %d (4098 attendu)" % size)
            out += container[cur:cur + _CHUNK]
        pos += size
    return bytes(out)


# ===========================================================================
# 2. Chiffrement des propriétés CMG / DPB / GC  [MS-OVBA] 2.4.3
# ===========================================================================
#
# Ce « chiffrement » n'est qu'un brouillage : Seed aléatoire, Version = 2,
# ProjKey = somme modulo 256 des octets de l'identifiant ID="{...}" du flux
# PROJECT, puis octets ignorés, longueur (4 octets) et données, chaque octet
# étant combiné aux deux octets chiffrés précédents.

def _project_key(project_id):
    """ProjKey : somme (mod 256) des caractères de l'ID du projet, accolades comprises."""
    return sum(project_id.encode("ascii")) & 0xFF


def encrypt_data(project_id, data, random_bytes):
    """Chiffre `data` (Data Encryption, 2.4.3.2).

    random_bytes(n) doit fournir n octets « aléatoires » (graine et octets ignorés).
    """
    seed = random_bytes(1)[0]
    version_enc = seed ^ 2
    proj_key = _project_key(project_id)
    proj_key_enc = seed ^ proj_key
    out = bytearray([seed, version_enc, proj_key_enc])
    unencrypted_byte1 = proj_key
    encrypted_byte1 = proj_key_enc
    encrypted_byte2 = version_enc

    def _push(byte):
        nonlocal unencrypted_byte1, encrypted_byte1, encrypted_byte2
        byte_enc = byte ^ ((encrypted_byte2 + unencrypted_byte1) & 0xFF)
        out.append(byte_enc)
        encrypted_byte2 = encrypted_byte1
        encrypted_byte1 = byte_enc
        unencrypted_byte1 = byte

    ignored_length = (seed & 6) // 2
    for b in random_bytes(ignored_length):
        _push(b)
    for b in struct.pack("<I", len(data)):
        _push(b)
    for b in data:
        _push(b)
    return bytes(out)


def decrypt_data(encrypted):
    """Déchiffre (2.4.3.3) ; renvoie (proj_key, data).  Contrôle Version et longueur."""
    encrypted = bytes(encrypted)
    seed, version_enc, proj_key_enc = encrypted[0], encrypted[1], encrypted[2]
    if seed ^ version_enc != 2:
        raise ValueError("Version chiffrée différente de 2")
    proj_key = seed ^ proj_key_enc
    unencrypted_byte1, encrypted_byte1, encrypted_byte2 = proj_key, proj_key_enc, version_enc
    pos = 3
    decoded = bytearray()
    ignored_length = (seed & 6) // 2
    while pos < len(encrypted):
        byte_enc = encrypted[pos]
        byte = byte_enc ^ ((encrypted_byte2 + unencrypted_byte1) & 0xFF)
        decoded.append(byte)
        encrypted_byte2, encrypted_byte1, unencrypted_byte1 = encrypted_byte1, byte_enc, byte
        pos += 1
    length = struct.unpack_from("<I", decoded, ignored_length)[0]
    data = bytes(decoded[ignored_length + 4:])
    if len(data) != length:
        raise ValueError("longueur chiffrée %d, %d octets présents" % (length, len(data)))
    return proj_key, data


# ===========================================================================
# 3. Préparation du code source des modules
# ===========================================================================

def _check_name(name, what):
    if not isinstance(name, str) or not _IDENT_RE.match(name):
        raise ValueError("%s %r invalide : identifiant ASCII [A-Za-z][A-Za-z0-9_]*, "
                         "31 caractères au plus" % (what, name))


def _normalize_source(name, source):
    """Normalise une source VBA : fins de ligne CRLF, ligne Attribute VB_Name retirée.

    Renvoie la liste des lignes (sans fin de ligne).  Une ligne « Attribute VB_Name »
    en tête (comme dans un fichier .bas exporté) est retirée après contrôle du nom ;
    elle est régénérée par l'appelant.
    """
    if source is None:
        return []
    if not isinstance(source, str):
        raise TypeError("la source du module %s doit être une chaîne (str)" % name)
    source = source.lstrip("\ufeff")
    if "\x00" in source:
        raise ValueError("la source du module %s contient un caractère NUL" % name)
    lines = source.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    if lines and lines[-1] == "":
        lines.pop()                       # la dernière fin de ligne est rajoutée plus tard
    if lines:
        m = _VB_NAME_RE.match(lines[0].strip())
        if m:
            if m.group(1) != name:
                raise ValueError("le module %r déclare Attribute VB_Name = %r"
                                 % (name, m.group(1)))
            lines.pop(0)
    for number, line in enumerate(lines, 1):
        if len(line) > _MAX_LINE:
            raise ValueError("module %s, ligne %d : plus de %d caractères"
                             % (name, number, _MAX_LINE))
        if re.match(r"\s*Attribute\s+VB_Name\b", line, re.IGNORECASE):
            raise ValueError("module %s, ligne %d : Attribute VB_Name inattendu" % (name, number))
        # Les attributs de module de document sont générés ici : les accepter aussi dans
        # la source les dupliquerait (en-tête de module invalide pour l'éditeur VBA).
        if _MODULE_ATTR_RE.match(line):
            raise ValueError("module %s, ligne %d : attribut de module %r généré automatiquement, "
                             "à retirer de la source" % (name, number, line.strip()))
    return lines


def _encode(text, where):
    try:
        return text.encode(_PY_CODEC)
    except UnicodeEncodeError as exc:
        bad = text[exc.start:exc.end]
        line = text.count("\n", 0, exc.start) + 1
        content = text.splitlines()[line - 1].strip()
        raise ValueError("%s, ligne %d (texte stocké, en-tête Attribute compris) : caractère %r "
                         "non représentable en page de codes %d dans « %s »"
                         % (where, line, bad, CODEPAGE, content[:80])) from None


def _module_source(name, kind, source):
    """Code source complet stocké dans le flux du module (octets cp1252, CRLF)."""
    lines = ['Attribute VB_Name = "%s"' % name]
    if kind in _DOC_BASE:
        lines.append('Attribute VB_Base = "%s"' % _DOC_BASE[kind])
        lines += ["Attribute %s = %s" % kv for kv in _DOC_ATTRIBUTES]
    lines += _normalize_source(name, source)
    text = "\r\n".join(lines) + "\r\n"
    return _encode(text, "module %s" % name)


# ===========================================================================
# 4. Flux du projet VBA
# ===========================================================================

def _project_stream(project_id, project_name, modules, random_bytes):
    """Flux PROJECT [MS-OVBA] 2.3.1 (texte cp1252, lignes terminées par CRLF)."""
    lines = ['ID="%s"' % project_id]
    for name, kind, _ in modules:
        if kind == "standard":
            lines.append("Module=%s" % name)
        else:
            lines.append("Document=%s/&H00000000" % name)
    lines += [
        'Name="%s"' % project_name,
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        # Projet non protégé : ProjectProtectionState = 0x00000000 (4 octets),
        # pas de mot de passe (un octet 0x00), projet visible (0xFF).
        'CMG="%s"' % encrypt_data(project_id, b"\x00\x00\x00\x00", random_bytes).hex().upper(),
        'DPB="%s"' % encrypt_data(project_id, b"\x00", random_bytes).hex().upper(),
        'GC="%s"' % encrypt_data(project_id, b"\xFF", random_bytes).hex().upper(),
        "",
        "[Host Extender Info]",
        _HOST_EXTENDER,
        "",
        "[Workspace]",
    ]
    lines += ["%s=0, 0, 0, 0, C" % name for name, _, _ in modules]
    return _encode("\r\n".join(lines) + "\r\n", "flux PROJECT")


def _projectwm_stream(modules):
    """Flux PROJECTwm [MS-OVBA] 2.3.3 : (nom MBCS + NUL, nom UTF-16 + NUL) ..., puis 0x0000."""
    out = bytearray()
    for name, _, _ in modules:
        out += _encode(name, "nom de module") + b"\x00"
        out += name.encode("utf-16-le") + b"\x00\x00"
    out += b"\x00\x00"
    return bytes(out)


def _vba_project_stream():
    """Flux _VBA_PROJECT [MS-OVBA] 2.3.4.1 : Reserved1 0x61CC, Version 0xFFFF,
    Reserved2 0x00, Reserved3 0x0000, aucun PerformanceCache."""
    return struct.pack("<HHBH", 0x61CC, 0xFFFF, 0x00, 0x0000)


def _rec(rec_id, payload):
    """Enregistrement générique : Id (2 octets), Size (4 octets), données."""
    return struct.pack("<HI", rec_id, len(payload)) + payload


def _dir_stream(project_name, modules):
    """Flux dir décompressé [MS-OVBA] 2.3.4.2 (enregistrements dans l'ordre normatif)."""
    enc = lambda s: _encode(s, "flux dir")
    u16 = lambda s: s.encode("utf-16-le")
    d = bytearray()
    # --- PROJECTINFORMATION (2.3.4.2.1) ---
    d += _rec(0x0001, struct.pack("<I", _SYSKIND))            # PROJECTSYSKIND
    d += _rec(0x0002, struct.pack("<I", _LCID))               # PROJECTLCID
    d += _rec(0x0014, struct.pack("<I", _LCID_INVOKE))        # PROJECTLCIDINVOKE
    d += _rec(0x0003, struct.pack("<H", CODEPAGE))            # PROJECTCODEPAGE
    d += _rec(0x0004, enc(project_name))                      # PROJECTNAME
    d += _rec(0x0005, b"") + _rec(0x0040, b"")                # PROJECTDOCSTRING (+ Unicode)
    d += _rec(0x0006, b"") + _rec(0x003D, b"")                # PROJECTHELPFILEPATH (1 et 2)
    d += _rec(0x0007, struct.pack("<I", 0))                   # PROJECTHELPCONTEXT
    d += _rec(0x0008, struct.pack("<I", 0))                   # PROJECTLIBFLAGS
    # PROJECTVERSION : Id, Reserved = 4, VersionMajor (4), VersionMinor (2)
    d += struct.pack("<HIIH", 0x0009, 4, _VERSION_MAJOR, _VERSION_MINOR)
    d += _rec(0x000C, b"") + _rec(0x003C, b"")                # PROJECTCONSTANTS (+ Unicode)
    # --- PROJECTREFERENCES (2.3.4.2.2) ---
    for ref_name, libid in _REFERENCES:
        d += _rec(0x0016, enc(ref_name)) + _rec(0x003E, u16(ref_name))   # REFERENCENAME
        lib = enc(libid)
        # REFERENCEREGISTERED : Size, SizeOfLibid, Libid, Reserved1 (4) = 0, Reserved2 (2) = 0
        d += _rec(0x000D, struct.pack("<I", len(lib)) + lib + struct.pack("<IH", 0, 0))
    # --- PROJECTMODULES (2.3.4.2.3) ---
    d += _rec(0x000F, struct.pack("<H", len(modules)))        # PROJECTMODULES.Count
    d += _rec(0x0013, struct.pack("<H", _COOKIE))             # PROJECTCOOKIE
    for name, kind, _ in modules:
        d += _rec(0x0019, enc(name)) + _rec(0x0047, u16(name))   # MODULENAME (+ Unicode)
        d += _rec(0x001A, enc(name)) + _rec(0x0032, u16(name))   # MODULESTREAMNAME (+ Unicode)
        d += _rec(0x001C, b"") + _rec(0x0048, b"")               # MODULEDOCSTRING (+ Unicode)
        d += _rec(0x0031, struct.pack("<I", 0))                  # MODULEOFFSET : source en tête
        d += _rec(0x001E, struct.pack("<I", 0))                  # MODULEHELPCONTEXT
        d += _rec(0x002C, struct.pack("<H", _COOKIE))            # MODULECOOKIE
        # MODULETYPE : 0x0021 = module standard, 0x0022 = document / classe ; Reserved = 0
        d += struct.pack("<HI", 0x0021 if kind == "standard" else 0x0022, 0)
        d += struct.pack("<HI", 0x002B, 0)                       # Terminator du module
    d += struct.pack("<HI", 0x0010, 0)                         # Terminator du flux dir
    return bytes(d)


# ===========================================================================
# 5. Écriture du conteneur [MS-CFB] version 3
# ===========================================================================

_SECTOR = 512
_MINI_SECTOR = 64
_MINI_CUTOFF = 4096
_FREESECT, _ENDOFCHAIN, _FATSECT = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD
_NOSTREAM = 0xFFFFFFFF
_RED, _BLACK = 0, 1
_STORAGE, _STREAM, _ROOT = 1, 2, 5


def _cfb_key(name):
    """Ordre des frères dans l'arbre rouge-noir (MS-CFB 2.6.4) : longueur d'abord,
    puis comparaison des caractères mis en majuscules (noms ASCII ici)."""
    return (len(name), name.upper())


def _build_rb_tree(indices, entries):
    """Relie les entrées `indices` (déjà triées) en arbre binaire équilibré et colore
    l'arbre de façon à respecter les règles rouge-noir.  Renvoie l'index de la racine."""
    if not indices:
        return _NOSTREAM
    depth = {}

    def link(lo, hi, level):
        if lo > hi:
            return _NOSTREAM
        mid = (lo + hi + 1) // 2
        idx = indices[mid]
        depth[idx] = level
        entries[idx]["left"] = link(lo, mid - 1, level + 1)
        entries[idx]["right"] = link(mid + 1, hi, level + 1)
        return idx

    root = link(0, len(indices) - 1, 0)
    max_depth = max(depth.values())
    perfect = len(indices) == (1 << (max_depth + 1)) - 1
    for idx, level in depth.items():
        # Tous les niveaux sauf le dernier sont complets : colorer en rouge le dernier
        # niveau (s'il est incomplet) égalise le nombre de nœuds noirs sur tout chemin.
        entries[idx]["color"] = _RED if (not perfect and level == max_depth and level > 0) else _BLACK
    _check_rb_tree(root, entries)
    return root


def _check_rb_tree(root, entries):
    """Vérifie les invariants : racine noire, pas de rouge-rouge, hauteur noire constante,
    ordre gauche < nœud < droite."""
    if entries[root]["color"] != _BLACK:
        raise AssertionError("racine de l'arbre rouge-noir non noire")

    def walk(idx, parent_red):
        if idx == _NOSTREAM:
            return 1
        e = entries[idx]
        red = e["color"] == _RED
        if red and parent_red:
            raise AssertionError("deux nœuds rouges consécutifs")
        for side, sign in (("left", -1), ("right", 1)):
            child = e[side]
            if child != _NOSTREAM:
                if (_cfb_key(entries[child]["name"]) > _cfb_key(e["name"])) != (sign > 0):
                    raise AssertionError("ordre de l'arbre rouge-noir non respecté")
        hl, hr = walk(e["left"], red), walk(e["right"], red)
        if hl != hr:
            raise AssertionError("hauteur noire inégale")
        return hl + (0 if red else 1)

    walk(root, False)


def _build_cfb(tree):
    """Construit un fichier composé CFB v3 à partir de {nom: bytes | dict}."""
    # --- Entrées de répertoire (index 0 = Root Entry) ---
    entries = [{"name": "Root Entry", "type": _ROOT, "data": None, "kids": []}]

    def add(parent, content):
        for name in sorted(content, key=_cfb_key):
            if not (0 < len(name) <= 31) or any(c in name for c in "/\\:!"):
                raise ValueError("nom CFB invalide : %r" % name)
            value = content[name]
            idx = len(entries)
            if isinstance(value, dict):
                entries.append({"name": name, "type": _STORAGE, "data": None, "kids": []})
            else:
                entries.append({"name": name, "type": _STREAM, "data": bytes(value), "kids": []})
            entries[parent]["kids"].append(idx)
        for idx in list(entries[parent]["kids"]):
            if entries[idx]["type"] == _STORAGE:
                add(idx, content[entries[idx]["name"]])

    add(0, tree)
    lower = [e["name"].upper() for e in entries]
    for e in entries:
        e.update(left=_NOSTREAM, right=_NOSTREAM, child=_NOSTREAM, color=_BLACK,
                 start=0, size=0)
        names = [lower[k] for k in e["kids"]]
        if len(set(names)) != len(names):
            raise ValueError("noms en double dans le storage %r" % e["name"])
    for e in entries:
        if e["type"] in (_ROOT, _STORAGE):
            e["child"] = _build_rb_tree(e["kids"], entries)
    entries[0]["color"] = _BLACK

    # --- Répartition des flux : mini-flux (< 4096 octets) ou secteurs normaux ---
    mini = bytearray()
    minifat = []
    big = []
    for e in entries:
        if e["type"] != _STREAM:
            continue
        data = e["data"]
        e["size"] = len(data)
        if not data:
            e["start"] = _ENDOFCHAIN
        elif len(data) < _MINI_CUTOFF:
            first = len(mini) // _MINI_SECTOR
            count = -(-len(data) // _MINI_SECTOR)
            e["start"] = first
            minifat += [first + k + 1 for k in range(count - 1)] + [_ENDOFCHAIN]
            mini += data.ljust(count * _MINI_SECTOR, b"\x00")
        else:
            big.append(e)

    n_dir = -(-len(entries) // 4)
    n_minifat = -(-len(minifat) * 4 // _SECTOR)
    n_mini = -(-len(mini) // _SECTOR)
    big_counts = [-(-len(e["data"]) // _SECTOR) for e in big]
    rest = n_dir + n_minifat + n_mini + sum(big_counts)
    n_fat = 1
    while n_fat * (_SECTOR // 4) < n_fat + rest:
        n_fat += 1
    if n_fat > 109:
        raise ValueError("projet trop volumineux (DIFAT étendue non gérée)")

    # --- Plan des secteurs : FAT | répertoire | MiniFAT | mini-flux | gros flux ---
    fat = [_FATSECT] * n_fat
    sectors = []

    def chain(count):
        first = len(fat)
        for k in range(count):
            fat.append(first + k + 1 if k < count - 1 else _ENDOFCHAIN)
        return first

    dir_start = chain(n_dir)
    minifat_start = chain(n_minifat) if n_minifat else _ENDOFCHAIN
    if n_mini:
        entries[0]["start"] = chain(n_mini)
        entries[0]["size"] = len(mini)
    else:
        entries[0]["start"] = _ENDOFCHAIN
    for e, count in zip(big, big_counts):
        e["start"] = chain(count)
    assert len(fat) == n_fat + rest
    fat += [_FREESECT] * (n_fat * (_SECTOR // 4) - len(fat))

    # --- Contenu des secteurs ---
    fat_bytes = struct.pack("<%dI" % len(fat), *fat)
    dir_bytes = bytearray()
    for e in entries:
        name16 = e["name"].encode("utf-16-le") + b"\x00\x00"
        dir_bytes += name16.ljust(64, b"\x00")
        dir_bytes += struct.pack("<HBB", len(name16), e["type"], e["color"])
        dir_bytes += struct.pack("<III", e["left"], e["right"], e["child"])
        dir_bytes += b"\x00" * 16                       # CLSID nul
        dir_bytes += struct.pack("<I", 0)               # State Bits
        dir_bytes += struct.pack("<QQ", 0, 0)           # dates de création / modification nulles
        dir_bytes += struct.pack("<IQ", e["start"], e["size"])
    while len(dir_bytes) < n_dir * _SECTOR:             # entrées libres
        dir_bytes += b"\x00" * 64 + struct.pack("<HBB", 0, 0, 0)
        dir_bytes += struct.pack("<III", _NOSTREAM, _NOSTREAM, _NOSTREAM) + b"\x00" * 48
    minifat_bytes = struct.pack("<%dI" % len(minifat), *minifat)
    minifat_bytes += struct.pack("<I", _FREESECT) * (n_minifat * (_SECTOR // 4) - len(minifat))

    body = bytearray(fat_bytes)
    body += dir_bytes
    body += minifat_bytes
    body += bytes(mini).ljust(n_mini * _SECTOR, b"\x00")
    for e, count in zip(big, big_counts):
        body += e["data"].ljust(count * _SECTOR, b"\x00")
    assert len(body) == (n_fat + rest) * _SECTOR

    # --- En-tête (512 octets) ---
    difat = list(range(n_fat)) + [_FREESECT] * (109 - n_fat)
    header = struct.pack(
        "<8s16sHHHHH6sIIIIIIIII",
        b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1",    # signature
        b"\x00" * 16,                            # Header CLSID
        0x003E, 0x0003,                          # Minor / Major Version (v3)
        0xFFFE,                                  # Byte Order (petit-boutiste)
        9, 6,                                    # secteurs de 2^9 = 512 et mini-secteurs de 2^6 = 64
        b"\x00" * 6,                             # Reserved
        0,                                       # Number of Directory Sectors (0 en v3)
        n_fat,                                   # Number of FAT Sectors
        dir_start,                               # First Directory Sector Location
        0,                                       # Transaction Signature Number
        _MINI_CUTOFF,                            # Mini Stream Cutoff Size
        minifat_start,                           # First Mini FAT Sector Location
        n_minifat,                               # Number of Mini FAT Sectors
        _ENDOFCHAIN,                             # First DIFAT Sector Location (aucun)
        0,                                       # Number of DIFAT Sectors
    ) + struct.pack("<109I", *difat)
    assert len(header) == _SECTOR
    return header + bytes(body)


# ===========================================================================
# 6. Point d'entrée
# ===========================================================================

def _det_random(key):
    """Générateur d'octets déterministe (SHA-256 en mode compteur) : sortie reproductible."""
    state = {"counter": 0}

    def random_bytes(n):
        out = b""
        while len(out) < n:
            out += hashlib.sha256(key + struct.pack("<Q", state["counter"])).digest()
            state["counter"] += 1
        return out[:n]

    return random_bytes


def build_vba_project_bin(std_modules, doc_modules, project_name="VBAProject", *,
                          project_id=None, seed=None, module_order=None, strict=True):
    """Construit le contenu binaire d'un vbaProject.bin.

    std_modules  : dict {nom: source VBA (str)} des modules standard.  Une ligne
                   « Attribute VB_Name = "nom" » en tête est acceptée (fichier .bas).
    doc_modules  : liste ordonnée de (nom, "workbook" | "worksheet") ou
                   (nom, type, source) pour les modules de document.
    project_name : nom du projet VBA (« VBAProject » par défaut, comme Excel).
    project_id   : GUID « {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX} » de la ligne ID= ;
                   par défaut dérivé du contenu (sortie reproductible).
    seed         : octets/str servant de graine au brouillage CMG/DPB/GC (défaut : contenu).
    module_order : ordre facultatif des modules (tous les noms) ; défaut : modules de
                   document puis modules standard.
    strict       : exige exactement un module de type "workbook".
    """
    _check_name(project_name, "nom de projet")
    modules = []
    for item in doc_modules:
        if len(item) == 2:
            (name, kind), source = item, None
        elif len(item) == 3:
            name, kind, source = item
        else:
            raise ValueError("module de document attendu : (nom, type) ou (nom, type, source)")
        if kind not in _DOC_BASE:
            raise ValueError("type de module de document %r : 'workbook' ou 'worksheet' attendu" % kind)
        modules.append((name, kind, source))
    for name, source in std_modules.items():
        modules.append((name, "standard", source))
    if not modules:
        raise ValueError("le projet doit contenir au moins un module")
    if strict and sum(1 for _, k, _ in modules if k == "workbook") != 1:
        raise ValueError("un et un seul module de type 'workbook' (ThisWorkbook) est attendu")
    seen = set()
    for name, _, _ in modules:
        _check_name(name, "nom de module")
        if name.lower() in seen:
            raise ValueError("nom de module en double : %r" % name)
        if name.lower() in _RESERVED_NAMES or name.lower() == project_name.lower():
            raise ValueError("nom de module réservé ou en conflit : %r" % name)
        seen.add(name.lower())
    if module_order is not None:
        by_name = {m[0]: m for m in modules}
        if sorted(module_order) != sorted(by_name):
            raise ValueError("module_order doit contenir exactement les noms des modules")
        modules = [by_name[n] for n in module_order]

    # Sources complètes (avec attributs) : calculées avant tout pour l'empreinte.
    sources = [(name, kind, _module_source(name, kind, src)) for name, kind, src in modules]
    digest = hashlib.sha256(repr((project_name, [(n, k, s) for n, k, s in sources])).encode())
    if project_id is None:
        h = digest.hexdigest().upper()
        project_id = "{%s-%s-%s-%s-%s}" % (h[0:8], h[8:12], h[12:16], h[16:20], h[20:32])
    if not re.match(r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}\Z", project_id):
        raise ValueError("project_id doit être un GUID entre accolades, en majuscules")
    if seed is None:
        key = digest.digest()
    else:
        key = seed if isinstance(seed, bytes) else str(seed).encode("utf-8")
    random_bytes = _det_random(b"vba-project-seed:" + key)

    def compress_checked(data, what):
        # garde-fou : un dernier chunk incompressible de plus de 3640 octets serait complété
        # par des NUL (RawChunk de 4096 octets) ; impossible avec du texte VBA, mais vérifié.
        packed = compress(data)
        if decompress(packed) != data:
            raise ValueError("compression non réversible pour %s" % what)
        return packed

    vba_storage = {
        "_VBA_PROJECT": _vba_project_stream(),
        "dir": compress_checked(_dir_stream(project_name, modules), "le flux dir"),
    }
    for name, _, src in sources:
        vba_storage[name] = compress_checked(src, "le module %s" % name)   # MODULEOFFSET = 0
    tree = {
        "PROJECT": _project_stream(project_id, project_name, modules, random_bytes),
        "PROJECTwm": _projectwm_stream(modules),
        "VBA": vba_storage,
    }
    return _build_cfb(tree)


# ===========================================================================
# 7. Utilisation en ligne de commande
# ===========================================================================
#
#   python vba_project.py sortie.bin --bas PlanDeCharge.bas \
#          --workbook ThisWorkbook --sheet Feuil1 --sheet Feuil2 ...
#
# Le nom d'un module standard est lu dans sa ligne « Attribute VB_Name » (ou, à
# défaut, dans le nom du fichier).  Les fichiers .bas sont lus en UTF-8.

def _main(argv=None):
    import argparse
    import os

    parser = argparse.ArgumentParser(description="Génère un vbaProject.bin depuis des sources VBA.")
    parser.add_argument("sortie", help="fichier vbaProject.bin à écrire")
    parser.add_argument("--bas", action="append", default=[], help="module standard (.bas, UTF-8)")
    parser.add_argument("--workbook", default="ThisWorkbook", help="codeName du classeur")
    parser.add_argument("--sheet", action="append", default=[], help="codeName d'une feuille (ordre des onglets)")
    parser.add_argument("--project-name", default="VBAProject")
    args = parser.parse_args(argv)

    std = {}
    for path in args.bas:
        with open(path, encoding="utf-8-sig") as f:
            text = f.read()
        first = text.lstrip().split("\n", 1)[0].strip()
        m = _VB_NAME_RE.match(first)
        name = m.group(1) if m else os.path.splitext(os.path.basename(path))[0]
        std[name] = text
    docs = [(args.workbook, "workbook")] + [(s, "worksheet") for s in args.sheet]
    data = build_vba_project_bin(std, docs, args.project_name)
    with open(args.sortie, "wb") as f:
        f.write(data)
    print("%s : %d octets, modules : %s" % (args.sortie, len(data),
                                           ", ".join([d for d, _ in docs] + list(std))))


if __name__ == "__main__":
    _main()
