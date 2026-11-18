#!/usr/bin/env python3
"""
Google Drive / Google Docs Direct Chord Sheet Transposer
Integrates `gdrive` + `transpose_docx.py` so you can pass a Google Drive File ID
(or search by song name), transpose it losslessly to one or more keys (or Nashville Numbers),
and optionally upload the resulting .docx back to Google Drive.
"""

import argparse
import json
import os
import subprocess
import sys
import tempfile
from transpose_docx import transpose_docx, get_capo_suggestions


def run_cmd(cmd_list):
    res = subprocess.run(cmd_list, capture_output=True, text=True)
    if res.returncode != 0:
        raise RuntimeError(f"Command failed ({' '.join(cmd_list)}):\n{res.stderr}")
    return res.stdout.strip()


def find_file_by_name(query: str):
    out = run_cmd(["gdrive", "readonly", "search", "--name-contains", query, "--json"])
    items = json.loads(out) if out else []
    if not items:
        raise ValueError(f"No file found in Google Drive matching '{query}'")
    return items[0]


def get_file_info(file_id: str):
    out = run_cmd(["gdrive", "readonly", "info", file_id, "--json"])
    return json.loads(out)


def main():
    parser = argparse.ArgumentParser(description="Transpose a Google Drive .docx or Google Doc chord sheet")
    parser.add_argument("target", help="Google Drive File ID or song name to search")
    parser.add_argument("--to", "-t", required=True, help="Comma-separated target keys (e.g. 'C,Eb,G' or 'number')")
    parser.add_argument("--from-key", "-f", default=None, help="Source key (auto-detected if omitted)")
    parser.add_argument("--accidental", "-a", choices=["auto", "sharp", "flat"], default="auto")
    parser.add_argument("--out-dir", "-o", default="/home/antigravity/chord-transposer/output",
                        help="Local directory to save transposed .docx files")
    parser.add_argument("--upload", action="store_true",
                        help="Upload transposed .docx file(s) back to the same Google Drive folder")
    args = parser.parse_args()

    os.makedirs(args.out_dir, exist_ok=True)

    # Resolve file ID
    if len(args.target) >= 20 and " " not in args.target:
        info = get_file_info(args.target)
    else:
        found = find_file_by_name(args.target)
        info = get_file_info(found["id"])

    file_id = info["id"]
    file_name = info["name"]
    mime_type = info.get("mimeType", "")
    parents = info.get("parents", [])
    parent_id = parents[0] if parents else None

    with tempfile.TemporaryDirectory() as tmpdir:
        local_docx = os.path.join(tmpdir, file_name if file_name.endswith(".docx") else f"{file_name}.docx")
        if mime_type == "application/vnd.google-apps.document":
            run_cmd(["gdrive", "readonly", "export", file_id, local_docx, "--format", "docx"])
        else:
            run_cmd(["gdrive", "readonly", "download", file_id, "--out", local_docx])

        base_title = os.path.splitext(file_name)[0]
        import re
        base_clean = re.sub(r"\s+[A-G][#b♯♭]?m?$", "", base_title)

        targets = [k.strip() for k in args.to.split(",") if k.strip()]
        for tk in targets:
            if tk.lower() in ("number", "nashville", "級數", "級數譜"):
                mode = "number"
                dst_key = args.from_key or "C"
                out_name = f"{base_clean} 級數譜.docx"
            else:
                mode = "chord"
                dst_key = tk
                out_name = f"{base_clean} {dst_key}.docx"

            out_path = os.path.join(args.out_dir, out_name)
            res = transpose_docx(local_docx, out_path, dst_key, args.from_key, args.accidental, mode)
            print(f"✅ Generated: {out_path} ({res['source_key']} -> {res['target_key']}, mode={mode})")

            if args.upload:
                upload_cmd = ["gdrive", "mutate", "upload", out_path]
                if parent_id:
                    upload_cmd.extend(["--parent", parent_id])
                up_out = run_cmd(upload_cmd)
                print(f"☁️ Uploaded to Google Drive: {up_out}")


if __name__ == "__main__":
    main()
