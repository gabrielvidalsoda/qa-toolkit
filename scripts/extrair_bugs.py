# -*- coding: utf-8 -*-
"""
Extract BUG-YYYYMMDD-NN entries from a Markdown bug report and generate
a Taiga import JSON (one object per bug).

Usage:
  python scripts/extrair_bugs.py <input.md> [output.json]

If paths are omitted, defaults are:
  input:  ./Relatorio-Testes-Manuais.md
  output: ./taiga-import.json
"""
import json
import os
import re
import sys


def main():
    if len(sys.argv) >= 2:
        in_path = sys.argv[1]
    else:
        in_path = os.path.join(os.getcwd(), "Relatorio-Testes-Manuais.md")

    if len(sys.argv) >= 3:
        out_path = sys.argv[2]
    else:
        base = os.path.splitext(os.path.basename(in_path))[0]
        out_path = os.path.join(os.path.dirname(os.path.abspath(in_path)) or os.getcwd(), f"{base}-taiga-import.json")

    with open(in_path, "r", encoding="utf-8") as f:
        text = f.read()

    # Split by bug headers: ## BUG-YYYYMMDD-NN
    parts = re.split(r"^##\s+(BUG-\d{8}-\d+)\b", text, flags=re.M)
    results = []
    for i in range(1, len(parts), 2):
        bug_id = parts[i]
        body = parts[i + 1] if i + 1 < len(parts) else ""
        # Extract title after bug id on the same line if present
        m_title = re.match(r"[^\n]*\n##\s+([^\n]+)", "## " + bug_id + body)
        # Easier: look for first heading line after split
        lines = body.splitlines()
        subject = bug_id
        # Find a descriptive title (first non-empty line that isn't metadata)
        for ln in lines:
            ln_stripped = ln.strip()
            if not ln_stripped:
                continue
            if ln_stripped.startswith(("**", "-")):
                continue
            if ln_stripped.lower().startswith(("tela", "rota", "severidade", "observado", "esperado", "evidencia", "causa", "nota")):
                continue
            # use as title if short
            subject = f"{bug_id} - {ln_stripped.lstrip('#').strip()}"
            break

        desc_lines = []
        sev = None
        tags = []
        for ln in lines:
            s = ln.strip()
            if not s:
                continue
            if s.startswith(("TAGS:", "Tags:")):
                # extract simple tags
                t = s.split(":", 1)[1]
                tags = [x.strip().strip("[]\"'") for x in re.split(r"[,\s]+", t) if x.strip()]
                continue
            msev = re.search(r"Severidade:\s*(\d+)", s, re.I)
            if msev:
                try:
                    sev = int(msev.group(1))
                except ValueError:
                    pass
            desc_lines.append(s)

        item = {
            "id": bug_id,
            "subject": subject,
            "description": "\n".join(desc_lines).strip(),
        }
        if sev is not None:
            item["severity_scale"] = sev
        if tags:
            item["tags"] = tags
        results.append(item)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"Wrote {len(results)} items to {out_path}")


if __name__ == "__main__":
    main()
