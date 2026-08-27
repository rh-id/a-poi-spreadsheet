#!/usr/bin/env python3
"""
a-poi-spreadsheet upstream source transform pipeline.

Transforms upstream Apache POI java sources into the Android-adapted,
re-packaged form used by this project:

  1. License banner insertion (Apache-2.0 §4 modification notice)
  2. Package re-name: org.apache.poi -> m.co.rh.id.apoi_spreadsheet.org.apache.poi
  3. Bridge imports (TempFile / Dimension / Dimension2D / Graphics2D -> base module)
  4. log4j2 -> android.util.Log conversion (fluent API, {} placeholders, throwables)
  5. Import block re-ordering to project convention
  6. Inline FQCN references in executable code converted to imports
     (collision-safe; strings/comments/javadoc untouched)
  7. Flags anything it cannot convert mechanically

Modes:
  validate   transform upstream @fork-point files and diff against the current
             repo to measure how faithfully the pipeline reproduces the
             original hand adaptation (residuals = manual adaptations).
  apply      transform upstream files at the new target ref and write them
             into the repo modules.
  report     list files that still need manual attention (leftover log4j/awt).

Usage:
  python transform_poi.py validate
  python transform_poi.py apply [--only-modified] [--files f1 f2 ...]
  python transform_poi.py report
"""
import argparse
import difflib
import os
import re
import shutil
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
UPD = Path(os.environ.get("POI_UPD_DIR", r"C:\Users\Ruby\AppData\Local\Temp\opencode\poi-upd"))
OLD_TREE = UPD / "old"
NEW_TREE = UPD / "new"
LISTS = UPD / "lists"
OUT_DIR = UPD / "out"

PREFIX = "m.co.rh.id.apoi_spreadsheet.org.apache.poi"
BASE = "m.co.rh.id.apoi_spreadsheet.base"
OLD_SHA = "6a8994ee0e6c59aa231570307a5dd213784993c3"
NEW_SHA = "094968cfc3d48224db08f0b7f0a6fc341b035114"  # REL_5_5_1

MODULES = {
    "poi": {"local": REPO / "poi" / "src" / "main" / "java" / "m" / "co" / "rh" / "id" / "apoi_spreadsheet",
            "up_old": "poi/src/main/java",
            "up_new": "poi/src/main/java"},
    "poi-ooxml": {"local": REPO / "poi-ooxml" / "src" / "main" / "java" / "m" / "co" / "rh" / "id" / "apoi_spreadsheet",
                  "up_old": "poi-ooxml/src/main/java",
                  "up_new": "poi-ooxml/src/main/java"},
}

BANNER_TPL = ("// Derived from Apache POI (https://github.com/apache/poi @ commit {sha}); "
              "this file has been modified for Android compatibility by the a-poi-spreadsheet project.")

# import-level bridges: upstream FQCN -> adapted FQCN
IMPORT_MAP = {
    "org.apache.poi.util.TempFile": BASE + ".util.TempFile",
    "java.awt.Dimension": BASE + ".model.Dimension",
    "java.awt.geom.Dimension2D": BASE + ".model.Dimension2D",
    "java.awt.Graphics2D": BASE + ".image.Graphics2D",
}
# awt imports that are known to survive on Android (kept as-is)
AWT_OK = {"java.awt.font.TextAttribute"}
LOG4J_PREFIX = "org.apache.logging.log4j"
LEVEL_MAP = {"debug": "d", "info": "i", "warn": "w", "error": "e", "trace": "v"}

HEADER_END = re.compile(r"^=+\s*\*/\s*$", re.M)
PACKAGE_RE = re.compile(r"^package\s+org\.apache\.poi(\.|;)", re.M)
IMPORT_RE = re.compile(r"^(import\s+(?:static\s+)?)([\w.]+(?:\.\*)?)(\s*;[^\n]*)$", re.M)
LOGGER_FIELD_RE = re.compile(
    r"^(\s*)((?:(?:private|protected|public|static|final|volatile)\s+)*)"
    r"Logger\s+(\w+)\s*=\s*(?:LogManager|PoiLogManager)\.getLogger\(\s*([^)]*?)\s*\)\s*;",
    re.M)
FLUENT_RE = re.compile(r"\b(\w+)\.at(Debug|Info|Warn|Error|Trace)\(\s*\)")
PLAIN_RE = re.compile(r"\b(\w+)\.(debug|info|warn|error|trace)\(")
ISENABLED_RE = re.compile(r"\b(\w+)\.is(Debug|Info|Warn|Error|Trace)Enabled\(\s*\)")
INT_HINT_RE = re.compile(
    r"^(?:\(int\)|\(long\)|\(short\))\s*"
    r"|^\d+[Ll]?$"
    r"|\.size\(\)$|\.length\(\)$|\.length$|\.intValue\(\)$|\.longValue\(\)$|\.ordinal\(\)$"
    r"|\.getRowCount\(\)$|\.getCellCount\(\)$|\.getColumn\(\)$|\.getRow\(\)$|\.numberOf\(\)$|\.count\(\)$")


class Flags:
    def __init__(self, name):
        self.name = name
        self.items = []

    def add(self, why, detail=""):
        self.items.append((why, detail.strip()))

    def __bool__(self):
        return bool(self.items)

    def __str__(self):
        out = []
        for why, d in self.items:
            out.append(f"    {why}" + (f": {d[:160]}" if d else ""))
        return "\n".join(out)


FQCN_CODE_RE = re.compile(r"(?<![\w.])org\.apache\.poi\.")


def repackage_code_refs(text, flags):
    """Repackage fully-qualified org.apache.poi references in code regions only.

    String literals and comments (incl. javadoc {@link ...}) keep the original
    names - this matches the original hand adaptation.
    """
    out = []
    i = 0
    n = len(text)
    in_str = in_chr = in_lc = in_bc = False
    code_start = 0

    def flush(end):
        seg = text[code_start:end]
        seg = FQCN_CODE_RE.sub(PREFIX + ".", seg)
        # org.apache.poi.schemas.* lives in the poi-ooxml-full jar - keep as-is
        # (restore must run AFTER the substitution, otherwise fresh schemas
        # refs get prefixed into the fork namespace)
        seg = seg.replace("m.co.rh.id.apoi_spreadsheet.org.apache.poi.schemas.", "org.apache.poi.schemas.")
        out.append(seg)

    while i < n:
        c = text[i]
        nxt = text[i + 1] if i + 1 < n else ""
        if in_str:
            out.append(c)
            if c == "\\" and nxt:
                out.append(nxt)
                i += 2
                continue
            if c == '"':
                in_str = False
                code_start = i + 1
        elif in_chr:
            out.append(c)
            if c == "\\" and nxt:
                out.append(nxt)
                i += 2
                continue
            if c == "'":
                in_chr = False
                code_start = i + 1
        elif in_lc:
            out.append(c)
            if c == "\n":
                in_lc = False
                code_start = i + 1
        elif in_bc:
            out.append(c)
            if c == "*" and nxt == "/":
                out.append(nxt)
                in_bc = False
                code_start = i + 2
                i += 2
                continue
        else:
            if c == '"':
                flush(i)
                out.append(c)
                in_str = True
            elif c == "'":
                flush(i)
                out.append(c)
                in_chr = True
            elif c == "/" and nxt == "/":
                flush(i)
                out.append("//")
                in_lc = True
                i += 2
                continue
            elif c == "/" and nxt == "*":
                flush(i)
                out.append("/*")
                in_bc = True
                i += 2
                continue
            # plain code chars are covered by the next flush() - do not append here
        i += 1
    if not (in_str or in_chr or in_lc or in_bc):
        flush(n)
    else:
        out.append(text[code_start:n])
        flags.add("unterminated string/comment while scanning")
    return "".join(out)


def reformat_repo_style(text, flags):
    """Pass the source through verbatim via a string/comment-aware scanner.

    No reformatting is performed: braces are not joined to the previous line
    and upstream layout is kept as-is (the hand adaptation relied on the IDE
    format that upstream files already match). The scanner is kept as the
    pipeline's format checkpoint - it walks code, literals and comments so an
    unterminated /* block comment is flagged and a future repo-style
    reformatter can slot in here without touching strings or javadoc.
    """
    out = []
    i = 0
    n = len(text)
    in_str = in_chr = in_lc = in_bc = False
    code_start = 0

    def flush(end):
        out.append(text[code_start:end])

    while i < n:
        c = text[i]
        nxt = text[i + 1] if i + 1 < n else ""
        if in_str:
            out.append(c)
            if c == "\\" and nxt:
                out.append(nxt)
                i += 2
                continue
            if c == '"':
                in_str = False
                code_start = i + 1
        elif in_chr:
            out.append(c)
            if c == "\\" and nxt:
                out.append(nxt)
                i += 2
                continue
            if c == "'":
                in_chr = False
                code_start = i + 1
        elif in_lc:
            out.append(c)
            if c == "\n":
                in_lc = False
                code_start = i + 1
        elif in_bc:
            out.append(c)
            if c == "*" and nxt == "/":
                out.append(nxt)
                in_bc = False
                code_start = i + 2
                i += 2
                continue
        else:
            if c == '"':
                flush(i)
                out.append(c)
                in_str = True
            elif c == "'":
                flush(i)
                out.append(c)
                in_chr = True
            elif c == "/" and nxt == "/":
                flush(i)
                out.append("//")
                in_lc = True
                i += 2
                continue
            elif c == "/" and nxt == "*":
                flush(i)
                out.append(text[i:])
                # comments are passed through unchanged (repo kept javadoc as-is
                # apart from the <p> tags already present upstream)
                e = text.find("*/", i + 2)
                if e < 0:
                    flags.add("unterminated comment in reformat")
                    return "".join(out)
                out[-1] = text[i:e + 2]
                i = e + 2
                code_start = i
                continue
        i += 1
    if not (in_str or in_chr or in_lc or in_bc):
        flush(n)
    return "".join(out)


# inline FQCN families the hand adaptation converts to imports when referenced
# from executable code: fork classes (incl. base module), android.graphics and
# the common JDK packages upstream qualifies in code. A .schemas. package
# segment is excluded: schemas classes stay in the poi-ooxml-full jar under
# their original package, so normalize must never import/repackage them.
FQCN_INLINE_RE = re.compile(
    r"(?<![\w.])"
    r"(?![\w.]*\.schemas\.)"
    r"(?:m\.co\.rh\.id\.apoi_spreadsheet(?:\.[a-z0-9_]+)+\.[A-Z]\w*"
    r"|android\.graphics(?:\.[a-z0-9_]+)*\.[A-Z]\w*"
    r"|javax?\.(?:util|io|time|nio|math|text|lang)(?:\.[a-z0-9_]+)*\.[A-Z]\w*"
    r"|javax\.xml(?:\.[a-z0-9_]+)*\.[A-Z]\w*)")
IMPORT_LINE_RE = re.compile(r"^import\s+(static\s+)?([\w.]+(?:\.\*)?)\s*;")
TYPE_DECL_RE = re.compile(r"\b(?:class|interface|enum|record)\s+([A-Z]\w*)")
IMPORT_GROUP_ORDER = ("static", "android", "third", "java", "poi")


def _import_group(fq, is_static=False):
    """Import group of an FQCN - same convention as the rebuild in transform()."""
    if is_static:
        return "static"
    if fq.startswith("android."):
        return "android"
    if fq.startswith("java.") or fq.startswith("javax."):
        return "java"
    if fq.startswith(PREFIX) or fq.startswith(BASE):
        return "poi"
    return "third"


def _code_regions(text):
    """(start, end) spans of executable code, i.e. outside string/char literals
    and //- and /* */ comments (same scanner approach as repackage_code_refs).
    """
    spans = []
    i = 0
    n = len(text)
    in_str = in_chr = in_lc = in_bc = False
    code_start = 0
    while i < n:
        c = text[i]
        nxt = text[i + 1] if i + 1 < n else ""
        if in_str:
            if c == "\\" and nxt:
                i += 2
                continue
            if c == '"':
                in_str = False
                code_start = i + 1
        elif in_chr:
            if c == "\\" and nxt:
                i += 2
                continue
            if c == "'":
                in_chr = False
                code_start = i + 1
        elif in_lc:
            if c == "\n":
                in_lc = False
                code_start = i + 1
        elif in_bc:
            if c == "*" and nxt == "/":
                in_bc = False
                code_start = i + 2
                i += 2
                continue
        else:
            if c == '"':
                spans.append((code_start, i))
                in_str = True
            elif c == "'":
                spans.append((code_start, i))
                in_chr = True
            elif c == "/" and nxt == "/":
                spans.append((code_start, i))
                in_lc = True
                i += 2
                continue
            elif c == "/" and nxt == "*":
                spans.append((code_start, i))
                in_bc = True
                i += 2
                continue
        i += 1
    if not (in_str or in_chr or in_lc or in_bc):
        spans.append((code_start, n))
    return spans


def normalize_fqcn_imports(text, current_package, flags):
    """Convert inline fully-qualified class names in executable code to imports.

    Reproduces the hand adaptation convention for the fork: an inline FQCN
    reference is shortened to its simple name and covered by a single-type
    import unless that would change meaning:

      a. FQCN lives in the file's own package      -> plain name, no import
         (shortened unless a real rebinding hazard exists: a different
         existing import, an own/nested type, or a second inline FQCN
         sharing the simple name keeps it qualified)
      b. simple name already imported to same FQCN -> shorten refs
      c. simple name collision (different existing import, own/nested type,
         bare use of the simple name already in code, or two inline FQCNs
         sharing one simple name)                  -> leave ALL refs qualified
      d. otherwise                                 -> add import + shorten

    Strings, chars, comments and javadoc are never touched, so deliberate
    android.graphics.Color FQCN APIs stay intact wherever a Color is already
    in scope, and Class.forName("java.nio...") literals survive. Idempotent:
    running the pass on its own output changes nothing.
    """
    if not FQCN_INLINE_RE.search(text):
        return text

    lines = text.split("\n")
    regions = _code_regions(text)

    # imports occupy the simple-name space; only code-region import lines count
    # (never a commented-out import inside a block comment)
    line_start = 0
    body_start = 0  # refs before this belong to import/package lines - skip
    imported = {}
    import_lines = []  # (fq, is_static, raw line)
    for ln in lines:
        m = IMPORT_LINE_RE.match(ln)
        if m and any(a <= line_start < b for a, b in regions):
            fq, is_static = m.group(2), bool(m.group(1))
            import_lines.append((fq, is_static, ln))
            body_start = line_start + len(ln) + 1
            simple = fq.rsplit(".", 1)[-1]
            # static imports of nested types share the simple-name space too
            if not fq.endswith(".*") and (not is_static or simple[:1].isupper()):
                imported.setdefault(simple, fq)
        line_start += len(ln) + 1

    # body code only (own-name / bare-use checks must not see import lines)
    code = "".join(text[max(a, body_start):b] for a, b in regions if b > body_start)
    refs = {}  # fqcn -> [absolute match offsets in the full text]
    for a, b in regions:
        if b <= body_start:
            continue
        for m in FQCN_INLINE_RE.finditer(text[a:b]):
            p = a + m.start()
            if p >= body_start:
                refs.setdefault(m.group(0), []).append(p)
    if not refs:
        return text

    own_names = set(TYPE_DECL_RE.findall(code))
    # same-package refs normally shorten without an import -> their simple
    # name will exist bare in code and must be treated as taken for inline
    # FQCNs from other packages that reuse it
    taken_bare = {fq.rsplit(".", 1)[-1] for fq in refs
                  if fq.rsplit(".", 1)[0] == current_package}

    # code copy without the FQCN matches, to detect pre-existing bare uses
    stripped = []
    last = 0
    for m in FQCN_INLINE_RE.finditer(code):
        stripped.append(code[last:m.start()])
        last = m.end()
    stripped.append(code[last:])
    stripped = "".join(stripped)

    def bare_used(simple):
        return re.search(r"(?<![\w.])" + re.escape(simple) + r"\b", stripped) is not None

    by_simple = {}  # simple name -> set of distinct inline FQCNs using it
    for fq in refs:
        by_simple.setdefault(fq.rsplit(".", 1)[-1], set()).add(fq)

    decisions = {}  # fq -> "shorten" | "keep"
    add_imports = []
    for fq in refs:
        simple = fq.rsplit(".", 1)[-1]
        same_package = fq.rsplit(".", 1)[0] == current_package
        # (a) folds into the fall-through below: a same-package ref stays
        # qualified only on a genuine rebinding hazard (JLS 7.5.1 - an
        # existing single-type import of the same simple name, an own/nested
        # type of that name, or a second inline FQCN sharing it). A bare use
        # of the name is NOT a hazard for a same-package ref: absent such a
        # collision the bare name resolves to the ref's own class (the
        # hssf/record files rely on this, e.g. bare Record vs inline
        # org.apache.poi.hssf.record.Record). Other-package refs keep the
        # full guard set incl. bare uses and names reserved by (a).
        if imported.get(simple) == fq:               # (b)
            decisions[fq] = "shorten"
        elif (simple in imported or simple in own_names
                or (not same_package and (bare_used(simple) or simple in taken_bare))
                or len(by_simple[simple]) > 1):      # (c)
            decisions[fq] = "keep"
        else:            # (a)/(d): shorten; a same-package ref needs no import
            decisions[fq] = "shorten"
            if not same_package:
                add_imports.append(fq)

    pos_list = []
    for fq, positions in refs.items():
        if decisions[fq] == "shorten":
            simple = fq.rsplit(".", 1)[-1]
            pos_list.extend((p, fq, simple) for p in positions)
    if not pos_list:
        return text

    # replace from the end so earlier offsets stay valid
    pos_list.sort(reverse=True)
    for p, fq, simple in pos_list:
        text = text[:p] + simple + text[p + len(fq):]

    if add_imports:
        flags.add("fqcn import added", ", ".join(sorted(set(add_imports))))
        entries = {}  # group -> list of (sort key, raw line)
        for fq, is_static, ln in import_lines:
            entries.setdefault(_import_group(fq, is_static), []).append((fq, ln))
        for fq in sorted(set(add_imports)):
            entries.setdefault(_import_group(fq), []).append((fq, f"import {fq};"))
        # group order as found; new groups slot in by project convention
        groups = []
        for fq, is_static, ln in import_lines:
            g = _import_group(fq, is_static)
            if g not in groups:
                groups.append(g)

        def rank(g):
            return IMPORT_GROUP_ORDER.index(g) if g in IMPORT_GROUP_ORDER else 99
        for fq in sorted(set(add_imports)):
            g = _import_group(fq)
            if g not in groups:
                at = len(groups)
                for i, eg in enumerate(groups):
                    if rank(eg) > rank(g):
                        at = i
                        break
                groups.insert(at, g)
        lines = text.split("\n")  # re-split: refs above may have been shortened
        # region guard: never let a commented-out import (col 0, inside /* */)
        # widen the replaced import region into the type body
        starts = []
        off = 0
        for ln in lines:
            starts.append(off)
            off += len(ln) + 1
        imp_idx = [i for i, ln in enumerate(lines)
                   if IMPORT_LINE_RE.match(ln)
                   and any(a <= starts[i] < b for a, b in regions)]
        block = []
        for gi, g in enumerate(groups):
            # IntelliJ-style order: a wildcard import sorts after the uppercase
            # single-type imports of its package but before lowercase ones
            # (matches the hand-cleaned files, e.g. Record; then record.*;)
            block.extend(ln for _, ln in
                         sorted(entries[g], key=lambda e: e[0].replace(".*", ".[")))
            if gi < len(groups) - 1:
                block.append("")
        if imp_idx:
            text = "\n".join(
                lines[:imp_idx[0]] + block + lines[imp_idx[-1] + 1:])
        else:
            # end the match before the newline: "[ \t]*$" keeps pm.end() on
            # the package line, so the "\n\n" below yields exactly ONE blank
            # line between package and the inserted import block (idempotent)
            pm = re.search(r"^package\s+[\w.]+;[ \t]*$", text, re.M)
            ins = pm.end() if pm else 0
            text = text[:ins] + "\n\n" + "\n".join(block) + "\n" + text[ins:]
    return text


def tag_name(var):
    if var == "LOG":
        return "TAG"
    if var.endswith("_LOG"):
        return var[:-4] + "_TAG"
    return var + "_TAG"


def match_paren(text, open_idx):
    """Return index of the ')' matching '(' at open_idx (respecting strings/comments)."""
    depth = 0
    i = open_idx
    n = len(text)
    in_str = in_chr = in_lc = in_bc = False
    while i < n:
        c = text[i]
        nxt = text[i + 1] if i + 1 < n else ""
        if in_str:
            if c == "\\":
                i += 2
                continue
            if c == '"':
                in_str = False
        elif in_chr:
            if c == "\\":
                i += 2
                continue
            if c == "'":
                in_chr = False
        elif in_lc:
            if c == "\n":
                in_lc = False
        elif in_bc:
            if c == "*" and nxt == "/":
                in_bc = False
                i += 1
        else:
            if c == '"':
                in_str = True
            elif c == "'":
                in_chr = True
            elif c == "/" and nxt == "/":
                in_lc = True
                i += 1
            elif c == "/" and nxt == "*":
                in_bc = True
                i += 1
            elif c == "(":
                depth += 1
            elif c == ")":
                depth -= 1
                if depth == 0:
                    return i
        i += 1
    return -1


def split_args(argstr):
    """Split top-level comma-separated args (strings/nesting aware)."""
    args = []
    depth = 0
    cur = []
    i = 0
    n = len(argstr)
    in_str = in_chr = False
    while i < n:
        c = argstr[i]
        if in_str:
            cur.append(c)
            if c == "\\":
                if i + 1 < n:
                    cur.append(argstr[i + 1])
                i += 2
                continue
            if c == '"':
                in_str = False
        elif in_chr:
            cur.append(c)
            if c == "\\":
                if i + 1 < n:
                    cur.append(argstr[i + 1])
                i += 2
                continue
            if c == "'":
                in_chr = False
        else:
            if c == '"':
                in_str = True
                cur.append(c)
            elif c == "'":
                in_chr = True
                cur.append(c)
            elif c in "([{":
                depth += 1
                cur.append(c)
            elif c in ")]}":
                depth -= 1
                cur.append(c)
            elif c == "," and depth == 0:
                args.append("".join(cur).strip())
                cur = []
                i += 1
                continue
            else:
                cur.append(c)
        i += 1
    tail = "".join(cur).strip()
    if tail:
        args.append(tail)
    return args


def unwrap_box(arg):
    m = re.match(r"^box\((.*)\)$", arg, re.S)
    return m.group(1).strip() if m else arg


def spec_for(arg):
    arg = unwrap_box(arg)
    if INT_HINT_RE.search(arg):
        return "%d"
    return "%s"


def is_string_literal(expr):
    expr = expr.strip()
    return expr.startswith('"') and expr.endswith('"') and not expr[:-1].count('"') > 1


def convert_message(msg, args, flags):
    """Build android Log message arg. msg is the raw first-arg expression."""
    if is_string_literal(msg) or (msg.startswith('"') and "+" in msg and re.match(r'^(".*"\s*(\+\s*"[^"]*")*)$', msg)):
        if not args:
            return msg
        # convert placeholders
        def repl(m):
            return spec_for(args[int(m.group(1))]) if m.group(1) is not None else "%s"
        idx = [0]

        def repl_seq(_m):
            s = spec_for(args[idx[0]]) if idx[0] < len(args) else "%s"
            idx[0] += 1
            return s
        conv = re.sub(r"\\\{\}|\{\}", lambda m: m.group(0)[1:] if m.group(0).startswith("\\") else None or repl_seq(m), msg)
        new_args = [unwrap_box(a) for a in args]
        return f"String.format({conv}, {', '.join(new_args)})"
    # non-literal message (variable, concat with identifiers, method call)
    if not args:
        return msg
    flags.add("log message not literal + extra args", msg)
    return msg


class Converter:
    """Converts log4j2 call sites to android.util.Log."""

    def __init__(self, flags, tagvars):
        self.flags = flags
        self.tagvars = tagvars  # var -> TAG name

    def convert_calls(self, text):
        # pass 1: fluent  VAR.atX().withThrowable(t)?.log( ... )
        out = []
        pos = 0
        for m in FLUENT_RE.finditer(text):
            var, level = m.group(1), m.group(2).lower()
            if var not in self.tagvars:
                continue
            tail = text[m.end():]
            wt = re.match(r"\s*\.withThrowable\(", tail)
            throwable = None
            end = m.end()
            if wt:
                p = m.end() + wt.end() - 1
                cp = match_paren(text, p)
                if cp < 0:
                    continue
                throwable = text[p + 1:cp].strip()
                end = cp + 1
            lm = re.match(r"\s*\.log\(", text[end:])
            if not lm:
                self.flags.add("fluent log not followed by .log(", text[m.start():m.start() + 80])
                continue
            p = end + lm.end() - 1
            cp = match_paren(text, p)
            if cp < 0:
                self.flags.add("unbalanced fluent .log(", text[m.start():m.start() + 80])
                continue
            argstr = text[p + 1:cp].strip()
            replacement = self.build_call(var, level, argstr, throwable, text[m.start():cp + 1])
            out.append((m.start(), cp + 1, replacement))
        for s, e, r in reversed(out):
            text = text[:s] + r + text[e:]
        # pass 2: plain  VAR.info( ... )
        out = []
        for m in PLAIN_RE.finditer(text):
            var, level = m.group(1), m.group(2)
            if var not in self.tagvars:
                continue
            p = m.end() - 1
            cp = match_paren(text, p)
            if cp < 0:
                self.flags.add("unbalanced plain log call", text[m.start():m.start() + 80])
                continue
            argstr = text[p + 1:cp].strip()
            replacement = self.build_call(var, level, argstr, None, text[m.start():cp + 1])
            out.append((m.start(), cp + 1, replacement))
        for s, e, r in reversed(out):
            text = text[:s] + r + text[e:]
        # pass 3: guards  VAR.isXEnabled()
        def guard(m):
            var, level = m.group(1), m.group(2)
            if var not in self.tagvars:
                return m.group(0)
            return f"Log.isLoggable({self.tagvars[var]}, Log.{level.upper()})"
        text = ISENABLED_RE.sub(guard, text)
        return text

    def build_call(self, var, level, argstr, throwable, orig):
        tag = self.tagvars[var]
        method = LEVEL_MAP[level]
        if argstr.startswith("() ->") or " -> " in argstr.split(",")[0]:
            self.flags.add("lazy lambda log message", orig)
            return orig
        args = split_args(argstr) if argstr else []
        if not args:
            self.flags.add("log call without args", orig)
            return orig
        msg = args[0]
        rest = [a for a in args[1:]]
        # unwrap fully-boxed args
        rest = [unwrap_box(a) for a in rest]
        # plain message without placeholders
        message = convert_message(msg, rest, self.flags)
        parts = [f"Log.{method}({tag}, {message}"]
        if throwable:
            parts.append(f", {throwable}")
        parts.append(")")
        return "".join(parts)


def transform(text, sha, flags):
    # ---------- imports ----------
    found_imports = []
    uses_log = False

    def import_repl(m):
        kw, fqcn, tail = m.group(1), m.group(2), m.group(3)
        if tail.strip() != ";":
            flags.add("import with trailing comment", fqcn)
        if fqcn.startswith(LOG4J_PREFIX):
            nonlocal uses_log
            uses_log = True
            return None  # drop
        if fqcn == "org.apache.logging.log4j.util.Unbox.box":
            return None
        # upstream 5.4+ wraps log4j acquisition in PoiLogManager - on Android this
        # vanishes together with the logger fields it initializes
        if fqcn == "org.apache.poi.logging.PoiLogManager":
            uses_log = True
            return None
        if fqcn in IMPORT_MAP:
            found_imports.append((kw, IMPORT_MAP[fqcn]))
            return "\x00MARKER\x00"
        if fqcn.startswith("java.awt.") or fqcn.startswith("javax.awt."):
            if fqcn not in AWT_OK:
                flags.add("unhandled java.awt import", fqcn)
            found_imports.append((kw, fqcn))
            return "\x00MARKER\x00"
        if fqcn == "org.apache.poi" or fqcn.startswith("org.apache.poi."):
            # org.apache.poi.schemas.* comes from the poi-ooxml-full jar - never repackage
            if not fqcn.startswith("org.apache.poi.schemas."):
                fqcn = PREFIX + fqcn[len("org.apache.poi"):]
        found_imports.append((kw, fqcn))
        return "\x00MARKER\x00"

    # note: IMPORT_RE group2 doesn't include 'static' - handled via kw
    marked = IMPORT_RE.sub(import_repl, text)

    # package rename
    marked = PACKAGE_RE.sub(lambda m: "package " + PREFIX + m.group(1), marked)

    # ---------- logger fields ----------
    tagvars = {}

    def field_repl(m):
        indent, mods, var, arg = m.groups()
        if arg.endswith(".class"):
            tag_val = arg[:-len(".class")].rsplit(".", 1)[-1]
        elif arg.startswith('"') and arg.endswith('"'):
            tag_val = arg[1:-1]
        else:
            flags.add("getLogger arg not class/string literal", arg)
            tag_val = arg
        tagvars[var] = tag_name(var)
        mods_str = mods.strip()
        return f"{indent}{mods_str} String {tag_name(var)} = \"{tag_val}\";"

    marked = LOGGER_FIELD_RE.sub(field_repl, marked)
    # any logger leftovers?
    for m in re.finditer(r"\bLogger\s+\w+\s*=|LogManager\.", marked):
        flags.add("leftover Logger/LogManager construct", m.group(0))

    # ---------- log calls ----------
    conv = Converter(flags, tagvars)
    marked = conv.convert_calls(marked)

    # ---------- inline FQCN references in code (not strings/comments) ----------
    marked = repackage_code_refs(marked, flags)

    # ---------- format to project style ----------
    marked = reformat_repo_style(marked, flags)

    # ---------- rebuild import block ----------
    statics = sorted({fq for kw, fq in found_imports if kw.startswith("import static")})
    normals = sorted({fq for kw, fq in found_imports if not kw.startswith("import static")})

    def group_of(fq):
        if fq.startswith("static "):
            return "static"
        if fq.startswith("android."):
            return "android"
        if fq.startswith("java.") or fq.startswith("javax."):
            return "java"
        if fq.startswith(PREFIX) or fq.startswith(BASE):
            return "poi"
        return "third"

    groups = {"static": statics and sorted(statics) or [],
              "android": [],
              "third": [], "java": [], "poi": []}
    for fq in normals:
        groups[group_of(fq)].append(fq)
    if uses_log or "Log." in marked:
        # only add android.util.Log when there are actual Log call sites / tagvars
        if tagvars or re.search(r"\bLog\.[diwev]\(TAG", marked) or re.search(r"Log\.isLoggable\(", marked):
            groups["android"].append("android.util.Log")

    lines = []
    for g in ("static", "android", "third", "java", "poi"):
        items = sorted(set(groups[g]))
        for fq in items:
            if g == "static":
                lines.append(f"import static {fq};")
            else:
                lines.append(f"import {fq};")
        if items:
            lines.append("")
    block = "\n".join(lines)

    # ---------- drop unused imports (matches the project's IDE optimize-imports) ----------
    import re as _re

    def code_only(s):
        out = []
        i = 0
        n = len(s)
        in_str = in_chr = False
        while i < n:
            c = s[i]
            nxt = s[i + 1] if i + 1 < n else ""
            if in_str:
                if c == "\\" and nxt:
                    i += 2
                    continue
                if c == '"':
                    in_str = False
            elif in_chr:
                if c == "\\" and nxt:
                    i += 2
                    continue
                if c == "'":
                    in_chr = False
            else:
                if c == '"':
                    in_str = True
                elif c == "'":
                    in_chr = True
                elif c == "/" and nxt == "/":
                    j = s.find("\n", i)
                    i = (j if j >= 0 else n)
                    continue
                elif c == "/" and nxt == "*":
                    j = s.find("*/", i + 2)
                    i = (j + 2 if j >= 0 else n)
                    continue
                else:
                    out.append(c)
            i += 1
        return "".join(out)

    body = code_only(marked)
    kept_lines = []
    for l in lines:
        m = _re.match(r"^import (?:static )?([\w.]+(?:\.\*)?);$", l)
        if m and not m.group(1).endswith(".*"):
            fq = m.group(1)
            simple = fq.rsplit(".", 1)[-1]
            if not _re.search(r"\b" + _re.escape(simple) + r"\b", body):
                # import unused in code; keep only if referenced in javadoc/comments
                if not _re.search(r"[{@,\s(]" + _re.escape(simple) + r"\b", marked):
                    continue
        kept_lines.append(l)
    # trim trailing empties
    while kept_lines and kept_lines[-1] == "":
        kept_lines.pop()
    block = "\n".join(kept_lines) + ("\n" if kept_lines else "")

    if "\x00MARKER\x00" in marked:
        # replace the whole original import section (first to last marker) with new block
        first = marked.index("\x00MARKER\x00")
        last = marked.rindex("\x00MARKER\x00")
        # extend to full lines
        ls = marked.rfind("\n", 0, first) + 1
        le = marked.find("\n", last)
        le = len(marked) if le < 0 else le
        # strip trailing blank lines inside old block region
        text = marked[:ls] + block + marked[le:]
    else:
        if re.search(r"^\s*import\s", marked, re.M):
            flags.add("no imports found")
        text = marked

    # ---------- banner ----------
    banner = BANNER_TPL.format(sha=sha)
    if "// Derived from Apache POI" not in text:
        m = HEADER_END.search(text)
        if m:
            ins = m.end()
            text = text[:ins] + "\n" + banner + text[ins:]
            # ensure exactly one blank line between banner and package
            text = re.sub(r"(//[^\n]*project\.)\s*\n(package )", r"\1\n\n\2", text, count=1)
        else:
            pm = re.search(r"^package\s+", text, re.M)
            if pm:
                ls = text.rfind("\n", 0, pm.start()) + 1
                text = banner + "\n\n" + text[ls:]
            else:
                flags.add("no license header / package found")

    # ---------- inline FQCN -> imports normalization (final pass) ----------
    pkg_m = re.search(r"^package\s+([\w.]+);", text, re.M)
    if pkg_m:
        text = normalize_fqcn_imports(text, pkg_m.group(1), flags)

    # ---------- final red-flag scan ----------
    for m in re.finditer(r"\bLOG\b|\bLogManager\b|org\.apache\.logging", text):
        flags.add("leftover log4j reference", m.group(0))
        break
    for m in re.finditer(r"java\.awt\.(?!font\.TextAttribute)", text):
        flags.add("leftover java.awt reference", text[max(0, m.start() - 40):m.end() + 40])
        break
    return text


def read(p):
    return p.read_text(encoding="utf-8")


def write(p, s):
    p.parent.mkdir(parents=True, exist_ok=True)
    # normalize line endings to \n like the repo files
    p.write_text(s.replace("\r\n", "\n").replace("\n\r", "\n"), encoding="utf-8", newline="\n")


def norm_for_diff(s):
    """Token-level normalization: whitespace-insensitive comparison.

    The original hand adaptation included an IDE reformat (collapsed javadocs,
    spacing changes), so byte-level equality is not the goal - we compare with
    all whitespace removed and banners stripped (only 505 of 1501 files carry
    the modification banner) to expose only semantic deltas.
    """
    s = re.sub(r"// Derived from Apache POI[^\n]*\n", "", s)
    return "".join(s.split())


def load_list(mod, name):
    p = LISTS / f"{mod}-{name}.txt"
    return [l for l in p.read_text().splitlines() if l.strip()] if p.exists() else []


def cmd_validate(args):
    report = []
    total = same = diff = 0
    for mod, cfg in MODULES.items():
        for rel in load_list(mod, "keep"):
            src = OLD_TREE / cfg["up_old"] / rel
            dst = cfg["local"] / rel
            if not src.exists() or not dst.exists():
                report.append(f"MISSING {mod} {rel}")
                continue
            flags = Flags(rel)
            try:
                out = transform(read(src), OLD_SHA, flags)
            except Exception as ex:
                report.append(f"ERROR {mod} {rel}: {ex}")
                continue
            total += 1
            if norm_for_diff(out) == norm_for_diff(read(dst)):
                same += 1
            else:
                diff += 1
                report.append(f"DIFF {mod} {rel}")
            if flags:
                report.append(f"FLAGS {mod} {rel}\n{flags}")
    (UPD / "validate-report.txt").write_text("\n".join(report), encoding="utf-8")
    print(f"validated {total}: identical={same} differing={diff}")
    print(f"full report: {UPD / 'validate-report.txt'}")


def cmd_apply(args):
    """Regenerate all upstream-changed files from REL_5_5_1 via the pipeline.

    Additionally produces a residual diff per file (original repo version vs
    pipeline output at the fork point) which documents the hand adaptations
    that must be re-applied on top of the regenerated files.
    """
    import difflib
    import subprocess

    changed = []
    for mod in MODULES:
        changed += [(mod, rel) for rel in load_list(mod, "changed-upstream")]

    manual = set()
    p = UPD / "apply-worklist.txt"
    if p.exists():
        for l in p.read_text().splitlines():
            if l.strip():
                mod, rel = l.split(" ", 1)
                manual.add((mod, rel))

    n_written = 0
    problems = []
    residuals = []
    for mod, rel in sorted(changed):
        cfg = MODULES[mod]
        src_new = NEW_TREE / cfg["up_new"] / rel
        dst = cfg["local"] / rel
        if not src_new.exists():
            problems.append(f"MISSING-UPSTREAM {mod} {rel}")
            continue
        flags = Flags(rel)
        new_text = transform(read(src_new), NEW_SHA, flags)
        write(dst, new_text)
        n_written += 1
        if flags:
            problems.append(f"FLAGS {mod} {rel}\n{flags}")
        if (mod, rel) in manual:
            # residual: original adapted repo file vs pipeline(base=fork point)
            git_path = str(dst.relative_to(REPO)).replace("\\", "/")
            r = subprocess.run(["git", "show", f"HEAD:{git_path}"], capture_output=True)
            repo_orig = r.stdout.decode("utf-8").replace("\r\n", "\n")
            src_old = OLD_TREE / cfg["up_old"] / rel
            base = transform(read(src_old), OLD_SHA, Flags(rel))
            d = list(difflib.unified_diff(
                base.replace("\r\n", "\n").splitlines(),
                repo_orig.splitlines(),
                "pipeline-base", "repo-adapted", lineterm="", n=3))
            if d:
                residuals.append(f"##### {mod} {rel}\n" + "\n".join(d))
    (UPD / "apply-report.txt").write_text("\n".join(problems), encoding="utf-8")
    (UPD / "residual-adaptations.diff").write_text("\n\n".join(residuals), encoding="utf-8")
    print(f"regenerated: {n_written}  with-flags: {len([p for p in problems if p.startswith('FLAGS')])}")
    print(f"residual diff for manual re-application: {UPD / 'residual-adaptations.diff'}")


def cmd_report(args):
    for mod in MODULES:
        print(f"== {mod} ==")
        for name in ("changed-upstream", "new-upstream-carried-pkg", "gone-upstream"):
            lst = load_list(mod, name)
            print(f"  {name}: {len(lst)}")


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("validate")
    a = sub.add_parser("apply")
    a.add_argument("--only-modified", action="store_true")
    a.add_argument("--files", nargs="*")
    sub.add_parser("report")
    args = ap.parse_args()
    {"validate": cmd_validate, "apply": cmd_apply, "report": cmd_report}[args.cmd](args)


if __name__ == "__main__":
    main()
