#!/usr/bin/env python3
"""
Lossless .docx & Text Chord Sheet Transposer (敬拜歌譜轉調工具 CLI)
Supports:
- Lossless .docx XML transposition (preserves all fonts, sizes, grey shading, margins, embedded fonts)
- Smart enharmonic spelling (# vs b based on target key, plus manual override)
- Nashville Number System (級數譜 1 - 6m - 4 - 5) & Dual Chord+Number mode
- Automatic metadata header key update (e.g., "Upper Room/活水/D" -> "Upper Room/活水/C")
- Capo recommendations for guitarists
"""

import argparse
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from typing import Optional, Tuple, List, Dict

NOTES_SHARP = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B']
NOTES_FLAT  = ['C', 'Db', 'D', 'Eb', 'E', 'F', 'Gb', 'G', 'Ab', 'A', 'Bb', 'B']

NOTE_TO_SEMITONE = {
    'C': 0, 'B#': 0,
    'C#': 1, 'Db': 1, 'C♯': 1, 'D♭': 1,
    'D': 2,
    'D#': 3, 'Eb': 3, 'D♯': 3, 'E♭': 3,
    'E': 4, 'Fb': 4,
    'F': 5, 'E#': 5,
    'F#': 6, 'Gb': 6, 'F♯': 6, 'G♭': 6,
    'G': 7,
    'G#': 8, 'Ab': 8, 'G♯': 8, 'A♭': 8,
    'A': 9,
    'A#': 10, 'Bb': 10, 'A♯': 10, 'B♭': 10,
    'B': 11, 'Cb': 11,
}

# Keys that naturally prefer flats (Circle of Fifths: F, Bb, Eb, Ab, Db, Gb and their relative minors)
FLAT_KEYS = {
    'F', 'Bb', 'Eb', 'Ab', 'Db', 'Gb',
    'Dm', 'Gm', 'Cm', 'Fm', 'Bbm', 'Ebm',
    'B♭', 'E♭', 'A♭', 'D♭', 'G♭'
}

# Major scale intervals in semitones from root: 1, 2, 3, 4, 5, 6, 7
SCALE_DEGREES_SHARP = ['1', '#1', '2', '#2', '3', '4', '#4', '5', '#5', '6', 'b7', '7']
SCALE_DEGREES_FLAT  = ['1', 'b2', '2', 'b3', '3', '4', 'b5', '5', 'b6', '6', 'b7', '7']

# Regex for matching a chord token:
# Root: [A-G][#b♯♭]?
# Quality/Extension: (?:maj|min|m|M|dim|aug|sus|add|no|[0-9]|[\+\-\^\(\)])*
# Optional Slash Bass: (?:/[A-G][#b♯♭]?)?
CHORD_TOKEN_RE = re.compile(
    r'(?<![A-Za-z0-9])'
    r'([A-G][#b♯♭]?)'
    r'((?:maj|min|m|M|dim|aug|sus|add|no|[0-9]|[\+\^\(\)])*)'
    r'(?:/([A-G][#b♯♭]?))?'
    r'(?![A-Za-z0-9])'
)

SECTION_HEADERS = {
    'intro', 'verse', 'verse1', 'verse2', 'verse3', 'verse4',
    'pre-chorus', 'prechorus', 'chorus', 'chorus1', 'chorus2',
    'bridge', 'bridge1', 'bridge2', 'tag', 'outro', 'interlude',
    'ending', 'solo', 'instrumental', 'vamp', 'turnaround'
}


def should_use_flats(target_key: str, accidental_pref: str = 'auto') -> bool:
    if accidental_pref == 'flat':
        return True
    if accidental_pref == 'sharp':
        return False
    return target_key in FLAT_KEYS


def semitone_to_note(semitone: int, use_flats: bool) -> str:
    idx = semitone % 12
    return NOTES_FLAT[idx] if use_flats else NOTES_SHARP[idx]


def semitone_to_number(semitone: int, key_semitone: int, use_flats: bool = True) -> str:
    diff = (semitone - key_semitone) % 12
    return SCALE_DEGREES_FLAT[diff] if use_flats else SCALE_DEGREES_SHARP[diff]


def normalize_fullwidth_for_match(s: str) -> str:
    """1-to-1 fullwidth ASCII -> halfwidth ASCII normalization preserving exact character offsets."""
    out = []
    for ch in s:
        code = ord(ch)
        if 0xFF01 <= code <= 0xFF5E:
            out.append(chr(code - 0xFEE0))
        elif code in (0x3000, 0x00A0):
            out.append(' ')
        else:
            out.append(ch)
    return ''.join(out)


def is_chord_line(line: str) -> bool:
    """
    Determine whether a line is a chord progression line (like 'G - Bm - ｜D - - -')
    rather than normal English/Chinese lyrics or section headers.
    """
    norm = normalize_fullwidth_for_match(line).strip()
    if not norm:
        return False
    low = norm.lower().rstrip(':')
    if low in SECTION_HEADERS or re.match(r'^(?:->)?\s*(intro|verse|chorus|bridge|tag|outro|interlude|pre-?chorus|v|pc|c|b)\s*\d*$', low):
        return False
    # If line contains Chinese characters and has no barlines or beat dashes, check carefully
    chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', norm))
    # Tokenize by whitespace and barlines/separators
    tokens = [t for t in re.split(r'[\s｜|\-–—―─_\/、,，\(\)\[\]\*]+', norm) if t]
    if not tokens:
        return False
    chord_matches = 0
    for tok in tokens:
        if CHORD_TOKEN_RE.fullmatch(tok):
            chord_matches += 1
    # Check if line has rhythmic markers like '- - -', '— —', '｜', or '|'
    has_grid_markers = bool(re.search(r'[｜|]', norm) or re.search(r'[-–—―─]\s*[-–—―─]', norm))
    if has_grid_markers and chord_matches >= 1:
        return True
    if chinese_chars > 0 and not has_grid_markers:
        return False
    # If majority of tokens are valid chords (and not a long English sentence)
    if len(tokens) <= 16 and chord_matches / len(tokens) >= 0.6:
        # Make sure it's not an English title like "To The One Who Is Seated On The Throne"
        english_words = [t for t in tokens if len(t) > 1 and t.isalpha() and not re.match(r'^[A-G](m|maj|min|dim|aug|sus|add)$', t)]
        if len(english_words) >= 2:
            return False
        return True
    return False


def match_metadata_key(line: str) -> Optional[Tuple[int, int, str, str]]:
    """
    Detect header metadata key (e.g. 'G key / 可萌 / 68拍', 'Key: D', 'JPCC/約書亞/G').
    Returns (start, end, key_str, root) aligned 1:1 with `line`, or None.
    """
    if not line or not line.strip() or is_chord_line(line):
        return None
    norm = normalize_fullwidth_for_match(line)

    m = re.search(r'(?:^|[\s/|(\[（【,;，；])((?:原調|調性|[Kk][Ee][Yy])\s*[:：]?\s*)([A-G][#b♯♭]?m?)(?=$|[\s/|)\]）】.,;，；])', norm)
    if m:
        key_str = m.group(2)
        root = key_str[0].upper() + re.sub(r'm$', '', key_str[1:])
        if root in NOTE_TO_SEMITONE:
            return (m.start(2), m.end(2), key_str, root)

    m = re.search(r'(?:^|[\s/|(\[（【,;，；])([A-G][#b♯♭]?m?)(\s*(?:[Kk][Ee][Yy]|調))(?=$|[\s/|)\]）】.,;，；])', norm)
    if m:
        key_str = m.group(1)
        root = key_str[0].upper() + re.sub(r'm$', '', key_str[1:])
        is_eng_sentence = (not re.search(r'[\u4e00-\u9fff/|：:]', norm)) and len(re.findall(r'[A-Za-z]{2,}', norm)) >= 3
        if root in NOTE_TO_SEMITONE and not is_eng_sentence:
            return (m.start(1), m.end(1), key_str, root)

    m = re.search(r'(/\s*)([A-G][#b♯♭]?m?)(\s*)$', norm)
    if m:
        key_str = m.group(2)
        root = key_str[0].upper() + re.sub(r'm$', '', key_str[1:])
        if root in NOTE_TO_SEMITONE:
            return (m.start(2), m.end(2), key_str, root)

    return None


def detect_key_from_lines(lines: List[str]) -> str:
    """
    Auto-detect original key from metadata line (e.g. 'G key / 可萌 / 68拍', 'Upper Room/活水/D', 'Key: D')
    or fallback to diatonic tonic scoring across all chords.
    """
    for line in lines[:12]:
        meta = match_metadata_key(line)
        if meta:
            return meta[3]
    # Fallback: collect all chords (including fullwidth)
    chords = []
    for line in lines:
        if is_chord_line(line):
            norm = normalize_fullwidth_for_match(line)
            for m in CHORD_TOKEN_RE.finditer(norm):
                semi = NOTE_TO_SEMITONE.get(m.group(1))
                if semi is not None:
                    q = m.group(2) or ''
                    is_minor = bool(re.match(r'^m(?!aj)|^min', q, re.IGNORECASE))
                    chords.append((semi, is_minor))
    if not chords:
        return 'D'
    best_k, best_score = chords[-1][0], -1e9
    for k in range(12):
        score = 0.0
        for semi, is_minor in chords:
            diff = (semi - k) % 12
            if not is_minor and diff == 0:
                score += 3
            elif not is_minor and diff in (5, 7):
                score += 2
            elif is_minor and diff == 9:
                score += 2
            elif is_minor and diff in (2, 4):
                score += 1
            else:
                score -= 2
        if not chords[-1][1] and chords[-1][0] == k:
            score += 3
        if not chords[0][1] and chords[0][0] == k:
            score += 1.5
        if score > best_score:
            best_score = score
            best_k = k
    preferred = ['C', 'Db', 'D', 'Eb', 'E', 'F', 'F#', 'G', 'Ab', 'A', 'Bb', 'B']
    return preferred[best_k]


def transpose_chord_token(
    root: str,
    quality: str,
    bass: Optional[str],
    src_semitone: int,
    dst_semitone: int,
    use_flats: bool,
    mode: str = 'chord'  # 'chord', 'number', 'dual'
) -> str:
    root_semi = NOTE_TO_SEMITONE.get(root)
    if root_semi is None:
        return f"{root}{quality}" + (f"/{bass}" if bass else "")

    new_root_semi = (root_semi - src_semitone + dst_semitone) % 12
    new_root_note = semitone_to_note(new_root_semi, use_flats)
    new_root_num = semitone_to_number(new_root_semi, dst_semitone, use_flats=True)

    new_bass_note = None
    new_bass_num = None
    if bass:
        bass_semi = NOTE_TO_SEMITONE.get(bass)
        if bass_semi is not None:
            new_bass_semi = (bass_semi - src_semitone + dst_semitone) % 12
            new_bass_note = semitone_to_note(new_bass_semi, use_flats)
            new_bass_num = semitone_to_number(new_bass_semi, dst_semitone, use_flats=True)
        else:
            new_bass_note = bass
            new_bass_num = bass

    chord_str = f"{new_root_note}{quality}" + (f"/{new_bass_note}" if new_bass_note else "")
    num_str = f"{new_root_num}{quality}" + (f"/{new_bass_num}" if new_bass_num else "")

    if mode == 'number':
        return num_str
    elif mode == 'dual':
        return f"{chord_str}({num_str})"
    return chord_str


INLINE_BRACKET_CHORD_RE = re.compile(
    r'\[([A-G][#b♯♭]?)((?:maj|min|m|M|dim|aug|sus|add|no|[0-9]|[\+\^\(\)])*)(?:/([A-G][#b♯♭]?))?\]'
)


def transpose_line(
    line: str,
    src_key: str,
    dst_key: str,
    accidental_pref: str = 'auto',
    mode: str = 'chord'
) -> str:
    src_semi = NOTE_TO_SEMITONE.get(src_key.rstrip('m'), 2)
    dst_semi = NOTE_TO_SEMITONE.get(dst_key.rstrip('m'), src_semi)
    use_flats = should_use_flats(dst_key, accidental_pref)

    norm_line = normalize_fullwidth_for_match(line)

    # Check if this is a metadata line (e.g. 'G key / 可萌 / 68拍', 'Upper Room/活水/D')
    meta_match = match_metadata_key(line)
    if meta_match:
        prefix_end, suffix_start, key_str, _ = meta_match
        label = f"{dst_key} (級數)" if mode == 'number' else dst_key
        if '(' in label and re.match(r'[A-Za-z]', line[suffix_start:]):
            label += ' '
        return (line[:prefix_end] + label + line[suffix_start:]) if label != key_str else line

    # 0-semitone identity pass in auto-accidental chord mode: preserve original borrowed chord spellings
    if src_semi == dst_semi and accidental_pref == 'auto' and mode == 'chord':
        return line

    # Check if line has inline bracketed chords like [G]我想認識[Bm]祢 (or fullwidth ［Ｇ］)
    bracket_matches = list(INLINE_BRACKET_CHORD_RE.finditer(norm_line))
    if bracket_matches:
        out_parts = []
        cur = 0
        for m in bracket_matches:
            out_parts.append(line[cur:m.start()])
            root, quality, bass = m.group(1), m.group(2) or '', m.group(3)
            transposed = transpose_chord_token(root, quality, bass, src_semi, dst_semi, use_flats, mode)
            out_parts.append(f"[{transposed}]")
            cur = m.end()
        out_parts.append(line[cur:])
        return ''.join(out_parts)

    if not is_chord_line(line):
        return line

    def repl(m: re.Match) -> str:
        root, quality, bass = m.group(1), m.group(2) or '', m.group(3)
        return transpose_chord_token(root, quality, bass, src_semi, dst_semi, use_flats, mode)

    matches = list(CHORD_TOKEN_RE.finditer(norm_line))
    if not matches:
        return line

    result = []
    cursor = 0
    debt = 0  # positive means previous replacement was longer by `debt` chars

    for m in matches:
        gap = line[cursor:m.start()]
        if debt > 0:
            m_spaces = re.search(r' {2,}', gap)
            if m_spaces:
                s_start, s_end = m_spaces.span()
                consume = min(debt, s_end - s_start - 1)
                gap = gap[:s_start] + (' ' * (s_end - s_start - consume)) + gap[s_end:]
                debt -= consume
            if debt > 0:
                trail = len(gap) - len(gap.rstrip(' '))
                consume = min(debt, max(0, trail - 1))
                if consume > 0:
                    gap = gap[:-consume]
                    debt -= consume
        elif debt < 0:
            pad = -debt
            pipe_pos = gap.find('｜') if '｜' in gap else gap.find('|')
            if pipe_pos >= 0:
                gap = gap[:pipe_pos] + (' ' * pad) + gap[pipe_pos:]
            else:
                gap = gap + (' ' * pad)
            debt = 0

        result.append(gap)
        old_tok = m.group(0)
        new_tok = repl(m)
        result.append(new_tok)
        debt += len(new_tok) - len(old_tok)
        cursor = m.end()

    tail = line[cursor:]
    if debt > 0:
        m_spaces = re.search(r' {2,}', tail)
        if m_spaces:
            s_start, s_end = m_spaces.span()
            consume = min(debt, s_end - s_start - 1)
            tail = tail[:s_start] + (' ' * (s_end - s_start - consume)) + tail[s_end:]
    elif debt < 0:
        tail = tail + (' ' * (-debt))
    result.append(tail)
    return ''.join(result)


def get_capo_suggestions(target_key: str) -> List[Dict[str, str]]:
    """Return practical acoustic guitar Capo positions for the target sounding key."""
    target_semi = NOTE_TO_SEMITONE.get(target_key.rstrip('m'), 0)
    easy_guitar_keys = [
        ('C', 0), ('G', 7), ('D', 2), ('A', 9), ('E', 4), ('F', 5)
    ]
    suggestions = []
    for play_key, play_semi in easy_guitar_keys:
        capo = (target_semi - play_semi) % 12
        if 0 <= capo <= 9:
            suggestions.append({
                'capo': capo,
                'play_key': play_key,
                'label': f"不夾 Capo (直接彈 {play_key} 調)" if capo == 0 else f"Capo {capo} (彈 {play_key} 調指法)"
            })
    suggestions.sort(key=lambda x: (0 if x['play_key'] in ('C', 'G', 'D') and x['capo'] <= 5 else 1, x['capo']))
    return suggestions


def compute_minimal_edits_for_line(
    line: str,
    src_key: str,
    dst_key: str,
    accidental_pref: str = 'auto',
    mode: str = 'chord'
) -> List[Tuple[int, int, str]]:
    """Compute minimal in-place offset edits (start, end, new_text) sorted descending by start."""
    src_semi = NOTE_TO_SEMITONE.get(src_key.rstrip('m'), 2)
    dst_semi = NOTE_TO_SEMITONE.get(dst_key.rstrip('m'), src_semi)
    use_flats = should_use_flats(dst_key, accidental_pref)
    norm_line = normalize_fullwidth_for_match(line)

    meta_match = match_metadata_key(line)
    if meta_match:
        start, end, key_str, _ = meta_match
        label = f"{dst_key} (級數)" if mode == 'number' else dst_key
        if '(' in label and re.match(r'[A-Za-z]', line[end:]):
            label += ' '
        return [(start, end, label)] if label != key_str else []

    if src_semi == dst_semi and accidental_pref == 'auto' and mode == 'chord':
        return []

    bracket_matches = list(INLINE_BRACKET_CHORD_RE.finditer(norm_line))
    if bracket_matches:
        edits = []
        for m in bracket_matches:
            new_tok = '[' + transpose_chord_token(m.group(1), m.group(2) or '', m.group(3), src_semi, dst_semi, use_flats, mode) + ']'
            if new_tok != m.group(0):
                edits.append((m.start(), m.end(), new_tok))
        edits.sort(key=lambda e: (-e[0], -e[1]))
        return edits

    if not is_chord_line(line):
        return []

    matches = list(CHORD_TOKEN_RE.finditer(norm_line))
    if not matches:
        return []

    edits: List[Tuple[int, int, str]] = []
    cursor = 0
    debt = 0

    for m in matches:
        gap_start, gap_end = cursor, m.start()
        gap = line[gap_start:gap_end]
        if '\t' in gap:
            debt = 0
        elif debt > 0:
            m_spaces = re.search(r' {2,}', gap)
            m_spaces_end = -1
            if m_spaces:
                consume = min(debt, (m_spaces.end() - m_spaces.start()) - 1)
                if consume > 0:
                    edits.append((gap_start + m_spaces.start(), gap_start + m_spaces.start() + consume, ''))
                    debt -= consume
                m_spaces_end = m_spaces.end()
            if debt > 0 and m_spaces_end < len(gap):
                m_trail = re.search(r' +$', gap)
                trail = len(m_trail.group(0)) if m_trail else 0
                consume = min(debt, max(0, trail - 1))
                if consume > 0:
                    edits.append((gap_end - consume, gap_end, ''))
                    debt -= consume
        elif debt < 0:
            pad = -debt
            pipe_pos = gap.find('｜') if '｜' in gap else gap.find('|')
            insert_at = (gap_start + pipe_pos) if pipe_pos >= 0 else gap_end
            edits.append((insert_at, insert_at, ' ' * pad))
            debt = 0

        old_tok = m.group(0)
        new_tok = transpose_chord_token(m.group(1), m.group(2) or '', m.group(3), src_semi, dst_semi, use_flats, mode)
        if new_tok != old_tok:
            edits.append((m.start(), m.end(), new_tok))
        debt += len(new_tok) - len(old_tok)
        cursor = m.end()

    edits.sort(key=lambda e: (-e[0], -e[1]))
    return edits


def compute_minimal_edits_for_paragraph(
    full_text: str,
    src_key: str,
    dst_key: str,
    accidental_pref: str = 'auto',
    mode: str = 'chord'
) -> List[Tuple[int, int, str]]:
    if not re.search(r'[\r\n\x0b]', full_text):
        return compute_minimal_edits_for_line(full_text, src_key, dst_key, accidental_pref, mode)
    all_edits: List[Tuple[int, int, str]] = []
    offset = 0
    parts = re.split(r'([\r\n\x0b])', full_text)
    for idx, part in enumerate(parts):
        if idx % 2 == 0 and part.strip():
            for s, e, txt in compute_minimal_edits_for_line(part, src_key, dst_key, accidental_pref, mode):
                all_edits.append((s + offset, e + offset, txt))
        offset += len(part)
    all_edits.sort(key=lambda e: (-e[0], -e[1]))
    return all_edits


def transpose_docx(
    input_path: str,
    output_path: str,
    target_key: str,
    source_key: Optional[str] = None,
    accidental_pref: str = 'auto',
    mode: str = 'chord'
) -> Dict[str, str]:
    """
    Losslessly transpose a .docx chord sheet while keeping all XML styling,
    embedded fonts, shading, alignments, headers/footers, and runs intact.
    """
    with zipfile.ZipFile(input_path, 'r') as zin:
        namelist = zin.namelist()
        all_files = {name: zin.read(name) for name in namelist}

    xml_bytes = all_files['word/document.xml']
    header_footer_names = [
        n for n in namelist
        if re.match(r'^word/(?:header|footer)\d*\.xml$', n, re.IGNORECASE)
    ]

    # Extract header + body paragraph texts to detect source key if not provided
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    para_texts: List[str] = []
    for part_name in header_footer_names + ['word/document.xml']:
        tree = ET.fromstring(all_files[part_name])
        for p in tree.findall('.//w:p', ns):
            t_list = [t.text for t in p.findall('.//w:t', ns) if t.text]
            para_texts.append(''.join(t_list))

    detected_src = source_key or detect_key_from_lines(para_texts)

    updated_files = dict(all_files)
    for part_name in ['word/document.xml'] + header_footer_names:
        updated_files[part_name] = transpose_docx_xml_bytes(
            all_files[part_name], detected_src, target_key, accidental_pref, mode
        )

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name in namelist:
            zout.writestr(name, updated_files[name])

    return {
        'source_key': detected_src,
        'target_key': target_key,
        'mode': mode,
        'output_path': output_path
    }


def transpose_docx_xml_bytes(
    xml_bytes: bytes,
    src_key: str,
    dst_key: str,
    accidental_pref: str = 'auto',
    mode: str = 'chord'
) -> bytes:
    """
    Operate directly on raw OOXML string via minimal per-run offset edits so 100% of
    XML namespaces, attributes, run shading, fonts, colors, tabs, and line breaks remain intact.
    """
    xml_str = xml_bytes.decode('utf-8')

    p_tag_re = re.compile(r'<w:p(?:\s[^>]*)?>|</w:p>')
    token_re = re.compile(r'(<w:t(?:\s+[^>]*)?>)([\s\S]*?)(</w:t>)|<w:tab\b[^>]*\/?>|<w:br\b[^>]*\/?>|<w:cr\b[^>]*\/?>')

    def unescape_xml(s: str) -> str:
        return (s.replace('&lt;', '<')
                 .replace('&gt;', '>')
                 .replace('&quot;', '"')
                 .replace('&apos;', "'")
                 .replace('&amp;', '&'))

    def escape_xml(s: str) -> str:
        return (s.replace('&', '&amp;')
                 .replace('<', '&lt;')
                 .replace('>', '&gt;'))

    def ensure_xml_space_preserve(tag: str) -> str:
        if 'xml:space="preserve"' in tag or "xml:space='preserve'" in tag:
            return tag
        return re.sub(r'<w:t\b([^>]*)>', r'<w:t xml:space="preserve"\1>', tag)

    leaf_ranges = []
    stack = []
    for m in p_tag_re.finditer(xml_str):
        if m.group(0).startswith('</'):
            if stack:
                open_start, body_start, has_child_p = stack.pop()
                if not has_child_p:
                    leaf_ranges.append((open_start, body_start, m.start(), m.end()))
        else:
            if stack:
                stack[-1][2] = True
            stack.append([m.start(), m.end(), False])

    if not leaf_ranges:
        return xml_bytes
    leaf_ranges.sort(key=lambda r: r[0])

    out_parts = []
    cursor = 0
    for open_start, body_start, body_end, close_end in leaf_ranges:
        out_parts.append(xml_str[cursor:body_start])
        p_body = xml_str[body_start:body_end]

        tokens = []
        full_text_parts = []
        offset = 0
        for tm in token_re.finditer(p_body):
            if tm.group(1) is not None:
                txt = unescape_xml(tm.group(2))
                tokens.append({
                    'type': 'text', 'text': txt,
                    'start': offset, 'end': offset + len(txt),
                    'match_start': tm.start(), 'match_end': tm.end(),
                    'open_tag': tm.group(1), 'close_tag': tm.group(3),
                    'modified': False
                })
                full_text_parts.append(txt)
                offset += len(txt)
            elif tm.group(0).startswith('<w:tab'):
                tokens.append({'type': 'tab', 'start': offset, 'end': offset + 1})
                full_text_parts.append('\t')
                offset += 1
            else:
                tokens.append({'type': 'br', 'start': offset, 'end': offset + 1})
                full_text_parts.append('\n')
                offset += 1

        full_text = ''.join(full_text_parts)
        edits = compute_minimal_edits_for_paragraph(full_text, src_key, dst_key, accidental_pref, mode) if full_text.strip() else []

        if not edits:
            out_parts.append(p_body)
        else:
            for ed_start, ed_end, ed_text in edits:
                if ed_end == ed_start:
                    target = next((t for t in tokens if t['type'] == 'text' and t['start'] <= ed_start < t['end']), None)
                    if target is None:
                        target = next((t for t in reversed(tokens) if t['type'] == 'text' and t['end'] == ed_start), None)
                    if target is None:
                        target = next((t for t in tokens if t['type'] == 'text' and t['start'] == ed_start), None)
                    if target is not None:
                        local_pos = len(target['text']) if ed_start >= target['end'] else max(0, min(len(target['text']), ed_start - target['start']))
                        target['text'] = target['text'][:local_pos] + ed_text + target['text'][local_pos:]
                        target['modified'] = True
                elif ed_end > ed_start:
                    overlapping = [t for t in tokens if t['end'] > ed_start and t['start'] < ed_end]
                    if not overlapping or any(t['type'] != 'text' for t in overlapping):
                        continue
                    is_first = True
                    for tok in overlapping:
                        orig_len = tok['end'] - tok['start']
                        local_s = max(0, ed_start - tok['start'])
                        local_e = min(orig_len, ed_end - tok['start'])
                        if is_first:
                            tok['text'] = tok['text'][:local_s] + ed_text + tok['text'][local_e:]
                            is_first = False
                        else:
                            tok['text'] = tok['text'][:local_s] + tok['text'][local_e:]
                        tok['modified'] = True

            rebuilt = []
            p_cur = 0
            for tok in tokens:
                if tok['type'] == 'text' and tok['modified']:
                    rebuilt.append(p_body[p_cur:tok['match_start']])
                    rebuilt.append(ensure_xml_space_preserve(tok['open_tag']) + escape_xml(tok['text']) + tok['close_tag'])
                    p_cur = tok['match_end']
            rebuilt.append(p_body[p_cur:])
            out_parts.append(''.join(rebuilt))

        out_parts.append(xml_str[body_end:close_end])
        cursor = close_end

    out_parts.append(xml_str[cursor:])
    return ''.join(out_parts).encode('utf-8')


def main():
    parser = argparse.ArgumentParser(description='Lossless .docx & Text Worship Chord Sheet Transposer')
    parser.add_argument('input', help='Input .docx or .txt file path')
    parser.add_argument('--to', '-t', required=True, help='Target key (e.g. C, Db, D, Eb, E, F, F#, G, Ab, A, Bb, B)')
    parser.add_argument('--from-key', '-f', default=None, help='Source key (auto-detected if omitted)')
    parser.add_argument('--accidental', '-a', choices=['auto', 'sharp', 'flat'], default='auto',
                        help='Accidental preference: auto (by key), sharp (#), or flat (b)')
    parser.add_argument('--mode', '-m', choices=['chord', 'number', 'dual'], default='chord',
                        help='Output mode: chord (standard), number (Nashville 1-7), or dual (Chord + Number)')
    parser.add_argument('--output', '-o', default=None, help='Output file path')
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Error: File not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    ext = os.path.splitext(args.input)[1].lower()
    if ext == '.docx':
        out_path = args.output
        if not out_path:
            base = os.path.splitext(os.path.basename(args.input))[0]
            # Replace trailing key in filename if present (e.g. "坐在寶座上的那位 D" -> "坐在寶座上的那位 C")
            suffix = f"{args.to}_級數" if args.mode == 'number' else args.to
            new_base = re.sub(r'\s+[A-G][#b♯♭]?m?$', f' {suffix}', base)
            if new_base == base:
                new_base = f"{base} ({suffix})"
            out_path = os.path.join(os.path.dirname(args.input) or '.', f"{new_base}.docx")

        res = transpose_docx(args.input, out_path, args.to, args.from_key, args.accidental, args.mode)
        print(f"✅ Transposed .docx: {res['source_key']} -> {res['target_key']} (mode={res['mode']})")
        print(f"📄 Saved to: {res['output_path']}")
        capos = get_capo_suggestions(args.to)
        if capos:
            print("🎸 Guitar Capo Suggestions:", " | ".join(c['label'] for c in capos[:3]))
    else:
        with open(args.input, 'r', encoding='utf-8') as f:
            lines = f.read().splitlines()
        src_key = args.from_key or detect_key_from_lines(lines)
        out_lines = [transpose_line(line, src_key, args.to, args.accidental, args.mode) for line in lines]
        out_text = '\n'.join(out_lines)
        if args.output:
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write(out_text + '\n')
            print(f"✅ Saved transposed text ({src_key} -> {args.to}) to {args.output}")
        else:
            print(out_text)


if __name__ == '__main__':
    main()
