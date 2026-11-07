"""Text Extractor — pull the text out of files and archives into one file.

Run with no arguments to open the GUI:
    python text_extractor.py

Or use the command line:
    python text_extractor.py -o combined.txt notes.zip report.pdf src/ readme.md

Archives (.zip, .tar, .tar.gz/.tgz, .tar.bz2, .tar.xz, .gz, .bz2, .xz and,
with py7zr installed, .7z) are opened in memory, including archives nested
inside other archives. Plain-text files, PDFs (needs pypdf), Word (.docx),
PowerPoint (.pptx) and Excel (.xlsx) files have their text extracted.
Anything else is listed as skipped.
"""

from __future__ import annotations

import argparse
import bz2
import gzip
import io
import lzma
import os
import sys
import tarfile
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterator

# Guards against zip bombs and runaway nesting.
MAX_FILE_BYTES = 100 * 1024 * 1024   # largest single file read from an archive
MAX_DEPTH = 5                        # archives inside archives inside ...

ARCHIVE_EXTS = (".zip", ".tar", ".tar.gz", ".tgz", ".tar.bz2", ".tbz2", ".tar.xz", ".txz", ".7z")
SINGLE_COMPRESSED = {".gz": gzip.decompress, ".bz2": bz2.decompress, ".xz": lzma.decompress}
OFFICE_EXTS = (".docx", ".pptx", ".xlsx")
UNSUPPORTED_ARCHIVES = (".rar",)
JUNK = ("__MACOSX/", ".DS_Store", "Thumbs.db", "desktop.ini")


# ---------------------------------------------------------------- core logic

@dataclass
class Result:
    written: list[str] = field(default_factory=list)
    skipped: list[tuple[str, str]] = field(default_factory=list)   # (name, reason)
    chars: int = 0


def is_junk(name: str) -> bool:
    return any(part in name for part in JUNK if part.endswith("/")) or Path(name).name in JUNK


def lower_name(name: str) -> str:
    return name.lower().replace("\\", "/")


def decode_text(data: bytes) -> str | None:
    """Return text if the bytes look like text, else None."""
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        try:
            return data.decode("utf-16")
        except UnicodeDecodeError:
            return None
    if b"\x00" in data[:8192]:
        return None
    try:
        return data.decode("utf-8-sig")
    except UnicodeDecodeError:
        return data.decode("cp1252", errors="replace")


def read_limited(stream, name: str) -> bytes:
    data = stream.read(MAX_FILE_BYTES + 1)
    if len(data) > MAX_FILE_BYTES:
        raise ValueError(f"larger than {MAX_FILE_BYTES // (1024 * 1024)} MB")
    return data


# ---- document formats

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def docx_text(zf: zipfile.ZipFile) -> str:
    root = ET.fromstring(zf.read("word/document.xml"))
    lines = []
    for p in root.iter(W + "p"):
        parts = []
        for el in p.iter():
            if el.tag == W + "t" and el.text:
                parts.append(el.text)
            elif el.tag == W + "tab":
                parts.append("\t")
            elif el.tag in (W + "br", W + "cr"):
                parts.append("\n")
        lines.append("".join(parts))
    return "\n".join(lines)


def _numbered(names: list[str], prefix: str) -> list[str]:
    hits = [n for n in names if n.startswith(prefix) and n.endswith(".xml")]
    return sorted(hits, key=lambda n: int("".join(c for c in Path(n).stem if c.isdigit()) or 0))


def pptx_text(zf: zipfile.ZipFile) -> str:
    out = []
    for i, name in enumerate(_numbered(zf.namelist(), "ppt/slides/slide"), 1):
        root = ET.fromstring(zf.read(name))
        paras = ["".join(t.text or "" for t in p.iter(A + "t")) for p in root.iter(A + "p")]
        out.append(f"--- Slide {i} ---\n" + "\n".join(x for x in paras if x.strip()))
    return "\n\n".join(out)


def xlsx_text(zf: zipfile.ZipFile) -> str:
    names = zf.namelist()
    shared: list[str] = []
    if "xl/sharedStrings.xml" in names:
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
        shared = ["".join(t.text or "" for t in si.iter(S + "t")) for si in root.iter(S + "si")]
    out = []
    for i, name in enumerate(_numbered(names, "xl/worksheets/sheet"), 1):
        root = ET.fromstring(zf.read(name))
        rows = []
        for row in root.iter(S + "row"):
            cells = []
            for c in row.iter(S + "c"):
                kind, v = c.get("t"), c.find(S + "v")
                if kind == "s" and v is not None:
                    cells.append(shared[int(v.text)])
                elif kind == "inlineStr":
                    cells.append("".join(t.text or "" for t in c.iter(S + "t")))
                else:
                    cells.append(v.text if v is not None and v.text else "")
            rows.append("\t".join(cells).rstrip("\t"))
        out.append(f"--- Sheet {i} ---\n" + "\n".join(rows))
    return "\n\n".join(out)


def office_text(data: bytes, ext: str) -> str:
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        return {".docx": docx_text, ".pptx": pptx_text, ".xlsx": xlsx_text}[ext](zf)


def pdf_text(data: bytes) -> str:
    try:
        from pypdf import PdfReader
    except ImportError:
        raise ValueError("PDF support needs pypdf (python -m pip install pypdf)")
    reader = PdfReader(io.BytesIO(data))
    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except Exception:
            raise ValueError("password-protected PDF")
    pages = [(page.extract_text() or "").strip() for page in reader.pages]
    if not any(pages):
        raise ValueError("no extractable text (scanned PDF?)")
    return "\n\n".join(f"--- Page {i} ---\n{t}" for i, t in enumerate(pages, 1))


# ---- archives

def archive_members(data: bytes, ext: str) -> Iterator[tuple[str, bytes | Exception]]:
    """Yield (member name, bytes or the error that stopped it being read)."""
    if ext == ".zip":
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            for info in zf.infolist():
                if info.is_dir():
                    continue
                try:
                    if info.file_size > MAX_FILE_BYTES:
                        raise ValueError(f"larger than {MAX_FILE_BYTES // (1024 * 1024)} MB")
                    with zf.open(info) as f:
                        yield info.filename, read_limited(f, info.filename)
                except Exception as e:   # encrypted members, bad CRC, too big
                    yield info.filename, e
    elif ext == ".7z":
        try:
            import py7zr
        except ImportError:
            raise ValueError(".7z support needs py7zr (python -m pip install py7zr)")
        with py7zr.SevenZipFile(io.BytesIO(data)) as sz:
            for info in sz.list():
                if not info.is_directory and info.uncompressed > MAX_FILE_BYTES:
                    raise ValueError(f"'{info.filename}' is larger than {MAX_FILE_BYTES // (1024 * 1024)} MB")
            for name, bio in sorted(sz.readall().items()):
                yield name, bio.read()
    else:  # tar family
        with tarfile.open(fileobj=io.BytesIO(data), mode="r:*") as tf:
            for m in tf:
                if not m.isfile():
                    continue
                try:
                    if m.size > MAX_FILE_BYTES:
                        raise ValueError(f"larger than {MAX_FILE_BYTES // (1024 * 1024)} MB")
                    yield m.name, tf.extractfile(m).read()
                except Exception as e:
                    yield m.name, e


def archive_ext(name: str) -> str | None:
    low = lower_name(name)
    for ext in ARCHIVE_EXTS:
        if low.endswith(ext):
            return ".zip" if ext == ".zip" else (".7z" if ext == ".7z" else ".tar")
    return None


def walk(name: str, data: bytes, depth: int = 0) -> Iterator[tuple[str, str | None, str]]:
    """Yield (display name, text, skip reason) for a file, recursing into archives."""
    low = lower_name(name)
    ext = Path(low).suffix

    arc = archive_ext(low)
    if arc:
        if depth >= MAX_DEPTH:
            yield name, None, "archive nested too deeply"
            return
        try:
            for member, content in archive_members(data, arc):
                if is_junk(member):
                    continue
                full = f"{name}/{member}"
                if isinstance(content, Exception):
                    yield full, None, str(content) or type(content).__name__
                else:
                    yield from walk(full, content, depth + 1)
        except Exception as e:
            yield name, None, f"couldn't open archive: {e}"
        return

    if ext in SINGLE_COMPRESSED:
        if depth >= MAX_DEPTH:
            yield name, None, "archive nested too deeply"
            return
        try:
            inner = SINGLE_COMPRESSED[ext](data)
        except Exception as e:
            yield name, None, f"couldn't decompress: {e}"
            return
        if len(inner) > MAX_FILE_BYTES:
            yield name, None, f"larger than {MAX_FILE_BYTES // (1024 * 1024)} MB once decompressed"
            return
        yield from walk(name[: -len(ext)], inner, depth + 1)
        return

    if ext in UNSUPPORTED_ARCHIVES:
        yield name, None, "RAR archives aren't supported; convert to .zip or .7z"
        return

    try:
        if ext == ".pdf":
            yield name, pdf_text(data), ""
        elif ext in OFFICE_EXTS:
            yield name, office_text(data, ext), ""
        else:
            text = decode_text(data)
            if text is None:
                yield name, None, "binary file"
            else:
                yield name, text, ""
    except Exception as e:
        yield name, None, str(e) or type(e).__name__


def expand_inputs(paths: list[str]) -> Iterator[tuple[Path, str]]:
    """Yield (path, label) for each file; folders expand to their files, sorted."""
    for p in map(Path, paths):
        if p.is_dir():
            for f in sorted(p.rglob("*")):
                rel = f.relative_to(p).as_posix()
                if f.is_file() and not is_junk(rel):
                    yield f, f"{p.name}/{rel}"
        else:
            yield p, p.name


def extract_text(inputs: list[str], output: str, headers: bool = True) -> Result:
    """Extract text from every input (files, folders, archives) into one UTF-8 file."""
    if not inputs:
        raise ValueError("No input files given")
    out_path = Path(output).resolve()
    result = Result()
    with open(out_path, "w", encoding="utf-8", newline="\n") as out:
        for path, label in expand_inputs(inputs):
            if path.resolve() == out_path:
                continue
            try:
                data = path.read_bytes()
            except OSError as e:
                result.skipped.append((label, e.strerror or str(e)))
                continue
            for name, text, reason in walk(label, data):
                if text is None:
                    result.skipped.append((name, reason))
                    continue
                text = text.replace("\r\n", "\n").replace("\r", "\n").strip("\n")
                if not text.strip():
                    result.skipped.append((name, "empty"))
                    continue
                if result.written:
                    out.write("\n\n")
                if headers:
                    out.write(f"===== {name} =====\n\n")
                out.write(text + "\n")
                result.written.append(name)
                result.chars += len(text)
    return result


# ---------------------------------------------------------------------- CLI

def run_cli(argv: list[str]) -> int:
    p = argparse.ArgumentParser(description="Extract the text from files, folders and archives into one file.")
    p.add_argument("inputs", nargs="+", help="files, folders or archives (.zip, .tar.gz, .7z, ...)")
    p.add_argument("-o", "--output", default="combined.txt", help="output text file (default: combined.txt)")
    p.add_argument("--no-headers", action="store_true", help="don't write a '===== name =====' line before each file")
    p.add_argument("-q", "--quiet", action="store_true", help="don't list skipped files")
    args = p.parse_args(argv)
    try:
        r = extract_text(args.inputs, args.output, headers=not args.no_headers)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    print(f"Wrote text from {len(r.written)} file(s) ({r.chars:,} characters) to {args.output}")
    if r.skipped:
        print(f"Skipped {len(r.skipped)} file(s)" + ("" if args.quiet else ":"))
        if not args.quiet:
            for name, reason in r.skipped:
                print(f"  {name} - {reason}")
    return 0


# ---------------------------------------------------------------------- GUI

def run_gui() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("Text Extractor")
    root.geometry("760x440")
    root.minsize(560, 300)

    items: list[str] = []   # file or folder paths, in output order

    tree = ttk.Treeview(root, columns=("name", "kind", "location"), show="headings", selectmode="extended")
    tree.heading("name", text="Name")
    tree.heading("kind", text="Type")
    tree.heading("location", text="Location")
    tree.column("name", width=230)
    tree.column("kind", width=90, anchor="center")
    tree.column("location", width=400)

    def kind_of(path: str) -> str:
        p = Path(path)
        if p.is_dir():
            return "Folder"
        if archive_ext(p.name) or p.suffix.lower() in SINGLE_COMPRESSED:
            return "Archive"
        return p.suffix.lower().lstrip(".").upper() or "File"

    def refresh(select: list[int] | None = None) -> None:
        tree.delete(*tree.get_children())
        for i, path in enumerate(items):
            p = Path(path)
            tree.insert("", "end", iid=str(i), values=(p.name, kind_of(path), str(p.parent)))
        if select:
            tree.selection_set([str(i) for i in select])
        status.set(f"{len(items)} item(s)")

    def selected() -> list[int]:
        return sorted(int(i) for i in tree.selection())

    def add_files() -> None:
        paths = filedialog.askopenfilenames(filetypes=[("All files", "*.*"), ("Archives", "*.zip *.tar *.gz *.tgz *.bz2 *.xz *.7z"), ("Documents", "*.txt *.md *.pdf *.docx *.pptx *.xlsx *.csv"),])
        items.extend(p for p in paths if p not in items)
        refresh()

    def add_folder() -> None:
        path = filedialog.askdirectory()
        if path and path not in items:
            items.append(path)
            refresh()

    def remove() -> None:
        for i in reversed(selected()):
            del items[i]
        refresh()

    def move(delta: int) -> None:
        sel = selected()
        if not sel:
            return
        if (delta < 0 and sel[0] == 0) or (delta > 0 and sel[-1] == len(items) - 1):
            return
        for i in (sel if delta < 0 else list(reversed(sel))):
            items[i], items[i + delta] = items[i + delta], items[i]
        refresh([i + delta for i in sel])

    def extract() -> None:
        if not items:
            messagebox.showinfo("Nothing to extract", "Add some files, folders or archives first.")
            return
        out = filedialog.asksaveasfilename(defaultextension=".txt", initialfile="combined.txt", filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if not out:
            return
        if any(Path(out).resolve() == Path(p).resolve() for p in items):
            messagebox.showerror("Error", "Output can't be one of the input files.")
            return
        root.config(cursor="watch")
        status.set("Extracting…")
        root.update()
        try:
            r = extract_text(items, out, headers=headers.get())
        except Exception as e:
            messagebox.showerror("Extraction failed", str(e))
            return
        finally:
            root.config(cursor="")
            refresh()
        msg = f"Wrote text from {len(r.written)} file(s) ({r.chars:,} characters) to\n{out}"
        if r.skipped:
            shown = "\n".join(f"• {n} — {why}" for n, why in r.skipped[:15])
            more = f"\n…and {len(r.skipped) - 15} more" if len(r.skipped) > 15 else ""
            msg += f"\n\nSkipped {len(r.skipped)} file(s):\n{shown}{more}"
        messagebox.showinfo("Done", msg)

    btns = ttk.Frame(root, padding=(8, 8, 8, 0))
    for text, cmd in [("Add files…", add_files), ("Add folder…", add_folder), ("Remove", remove), ("Move up", lambda: move(-1)), ("Move down", lambda: move(1))]:
        ttk.Button(btns, text=text, command=cmd).pack(side="left", padx=(0, 6))
    ttk.Button(btns, text="Extract →", command=extract).pack(side="right")
    btns.pack(fill="x")

    tree.pack(fill="both", expand=True, padx=8, pady=8)
    tree.bind("<Delete>", lambda e: remove())

    bottom = ttk.Frame(root, padding=(8, 0, 8, 6))
    status = tk.StringVar(value="0 item(s) — add files, folders or archives")
    ttk.Label(bottom, textvariable=status).pack(side="left")
    headers = tk.BooleanVar(value=True)
    ttk.Checkbutton(bottom, text="Label each file with a header", variable=headers).pack(side="right")
    bottom.pack(fill="x")

    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) > 1:
        sys.exit(run_cli(sys.argv[1:]))
    run_gui()
