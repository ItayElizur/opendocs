#!/usr/bin/env python
"""
Split one big C# static class into `partial class` files, by member name.

Written for Phases 1+3 of the officeoffice refactor
(docs/superpowers/plans/2026-08-27-phase1-3-file-split.md). The split is
mechanical and error-prone by hand across ~200 members, so it is scripted:
the script refuses to run unless EVERY member is assigned to exactly one
destination, which is the mistake that would otherwise be silent.

Usage:
    python tools/split-partial.py <config.json>            # do the split
    python tools/split-partial.py <config.json> --dry-run  # report only

Config shape:
{
  "source":      "PowerPointAiAddIn/PowerPointTools.cs",
  "encoding":    "utf-8",          // "utf-8-sig" for a file with a BOM
  "headerLines": 12,               // usings + namespace{ + class{  (inclusive)
  "tailLines":   2,                // the closing } } at the end
  "groups": {                      // suffix -> member names, in any order
    "Read":   ["ShapeText", "GetDeckContext"],
    "Charts": ["AddChartPpt"]
  }
}

Members not listed in any group STAY in the source file. That is deliberate:
list only what moves out, so the core file needs no enumeration.
Output files are named <source-stem>.<suffix>.cs next to the source.
"""
import io
import json
import os
import re
import sys

DECL = re.compile(r'^        (private|public|internal)\b')
LEADING = re.compile(r'^\s*(//|\[)')


def read_lines(path, encoding):
    # newline='' preserves CRLF vs LF exactly as found.
    with io.open(path, encoding=encoding, newline='') as fh:
        return fh.read().split('\n')


def find_blocks(lines):
    """Return [(name, start_idx, end_idx)] covering every member, where the
    block includes any contiguous comment/attribute lines directly above the
    declaration - comments must travel with the code they explain."""
    starts = []
    for i, line in enumerate(lines):
        if DECL.match(line):
            s = i
            while s - 1 >= 0 and LEADING.match(lines[s - 1]):
                s -= 1
            starts.append((s, i))
    blocks = []
    for n, (s, d) in enumerate(starts):
        end = starts[n + 1][0] - 1 if n + 1 < len(starts) else None
        blocks.append([member_name(lines[d]), s, end])
    return blocks


def member_name(line):
    """Extract the declared name from a member declaration line.

    Handles three shapes, in order - a nested TYPE has neither '(' nor '=',
    so the method/field pattern alone silently mis-names it (this bit us on
    WordTools' `private sealed class ParagraphIndexResolver`)."""
    m = re.search(r'\b(?:class|struct|interface|enum)\s+([A-Za-z_][A-Za-z0-9_]*)', line)
    if m:
        return m.group(1)
    # Method, field, const, or expression-bodied property (the '=' of '=>').
    m = re.search(r'([A-Za-z_][A-Za-z0-9_]*)\s*(?:\(|=)', line)
    if m:
        return m.group(1)
    raise SystemExit(
        'Could not extract a member name from this declaration - the script\n'
        'needs a new pattern for it rather than guessing:\n    %s' % line.strip())


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    cfg = json.load(io.open(sys.argv[1], encoding='utf-8'))
    dry = '--dry-run' in sys.argv

    src = cfg['source']
    enc = cfg.get('encoding', 'utf-8')
    head_n = cfg['headerLines']
    tail_n = cfg['tailLines']

    lines = read_lines(src, enc)
    header = lines[:head_n]
    # Everything after the members, minus the trailing empty string from the
    # final newline, is the closing braces.
    trailing_blank = 1 if lines and lines[-1] == '' else 0
    tail_start = len(lines) - trailing_blank - tail_n
    tail = lines[tail_start:]

    blocks = find_blocks(lines[:tail_start])
    for b in blocks:
        if b[2] is None:
            b[2] = tail_start - 1

    by_name = {}
    for name, s, e in blocks:
        if name in by_name:
            sys.exit('DUPLICATE member name in source: %s' % name)
        by_name[name] = (s, e)

    # --- safety: every assigned name must exist, exactly once ------------
    assigned, problems = {}, []
    for suffix, names in cfg['groups'].items():
        for n in names:
            if n not in by_name:
                problems.append('  group %-16s unknown member: %s' % (suffix, n))
            elif n in assigned:
                problems.append('  member %s assigned to BOTH %s and %s' % (n, assigned[n], suffix))
            else:
                assigned[n] = suffix
    if problems:
        sys.exit('REFUSING TO SPLIT - fix the config:\n' + '\n'.join(problems))

    stays = [n for n, _ in [(b[0], b[1]) for b in blocks] if n not in assigned]

    print('source        : %s (%d lines, %d members)' % (src, len(lines) - trailing_blank, len(blocks)))
    print('stays in core : %d members' % len(stays))
    for suffix in cfg['groups']:
        print('  .%-16s %d members' % (suffix + '.cs', len(cfg['groups'][suffix])))
    if dry:
        print('\n--dry-run: nothing written.')
        return

    # Make the class partial in the header we copy into every part.
    class_re = re.compile(r'^(\s*public static )(class\b)')
    header_partial = [class_re.sub(r'\1partial \2', h) for h in header]
    if header_partial == header:
        sys.exit('Could not find "public static class" in the header - check headerLines.')

    stem = os.path.splitext(src)[0]
    written = []
    for suffix, names in cfg['groups'].items():
        # Emit in original source order so diffs stay readable.
        ordered = sorted(names, key=lambda n: by_name[n][0])
        body = []
        for n in ordered:
            s, e = by_name[n]
            body.extend(lines[s:e + 1])
        out = header_partial + body + tail + ['']
        dest = '%s.%s.cs' % (stem, suffix)
        with io.open(dest, 'w', encoding=enc, newline='') as fh:
            fh.write('\n'.join(out))
        written.append(dest)
        print('wrote %-52s %5d lines' % (dest, len(out) - 1))

    # Rewrite the source with only the members that stay.
    keep = []
    for name, s, e in blocks:
        if name not in assigned:
            keep.extend(lines[s:e + 1])
    out = header_partial + keep + tail + ['']
    with io.open(src, 'w', encoding=enc, newline='') as fh:
        fh.write('\n'.join(out))
    print('wrote %-52s %5d lines  (source, rewritten)' % (src, len(out) - 1))

    print('\nNow: add these to the .csproj, then build + run the member-set diff.')
    for w in written:
        print('  <Compile Include="%s"><SubType>Code</SubType></Compile>' % os.path.basename(w))


if __name__ == '__main__':
    main()
