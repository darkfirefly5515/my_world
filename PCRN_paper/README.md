# PCRN Paper — Overleaf-Ready LaTeX Project

Self-contained LaTeX project for the paper:

> **PCRN: A Pocket-derived Co-occurrence Residue Network for Analyzing Protein Pocket Dynamics**

It contains:

- `main.tex` — the full title, abstract, introduction, and related work (English)
- `references.bib` — BibTeX file for all 23 cited references

---

## How to use on Overleaf

1. Go to [overleaf.com](https://www.overleaf.com), click **New Project → Upload Project**.
2. Upload the entire `PCRN_paper` folder as a ZIP, **or** create a new blank project and upload `main.tex` + `references.bib` into it.
3. Make sure the **Compiler** is set to **pdfLaTeX** (Menu → Settings).
4. Click **Recompile**. If citations show as `[?]`, click Recompile **once more** so BibTeX runs again. Overleaf usually does this automatically.
5. The output PDF will show: title block, abstract, Introduction (7 subsections), Related Work (4 subsections), and the numbered reference list.

## How to compile locally (optional)

```bash
pdflatex main.tex
bibtex   main
pdflatex main.tex
pdflatex main.tex
```

## ⚠ Reference verification reminder

A few BibTeX entries are still marked `% TODO verify` because the
authors / volume / year fields could not be fully confirmed in the
research session. Please cross-check with Google Scholar or PubMed
**before final submission**:

- `[10] wang2022pnmavis` — first author needs verification
- `[11] vappd2022` — full author list needs verification
- `[14] d3pockets` — year and journal need verification
- `[15] vis_largescale_pli` — full bibliographic record needs verification
- `[20] rin_clustering_ensembles` — full bibliographic record needs verification
- `[3] meller2025cryptic` — confirm whether it appeared in 2024 or 2025
- `[4] krivak2024comparative` — confirm year and journal

Everything else is from sources I verified during the writing session.

## Switching to a journal template later

When you decide a target venue:

- For **MDPI Applied Sciences / Molecules**, replace the preamble with the
  MDPI class file (`mdpi.cls`) and remove `geometry` + `setspace`.
- For **Bioinformatics (OUP)**, switch to OUP's article template.
- For **IEEE TVCG / EuroVis**, switch to the IEEE/EG class.

The body text (sections, citations) will work without modification because
all references go through `\citep{}` (natbib).
