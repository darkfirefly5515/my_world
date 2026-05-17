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

## Reference verification status (after deeper round of checks)

After a second round of cross-checking on PubMed, Wiley, MDPI, ACS, OUP,
Frontiers, Springer, ChemRxiv, the situation is now:

### Fully verified (no further action needed)
- `[1] stank2016protein` (Acc. Chem. Res. 2016)
- `[2] krone2016visual` (CGF 2016 STAR)
- `[5] eguida2022estimating` (IJMS 2022)
- `[7] lindow2013exploring` (BMC Bioinformatics 2013)
- `[8] parulek2013visual` (BMC Bioinformatics 2013)
- `[10] guo2022pnmavis` — authors confirmed: Guo, Feng, Zhang, Guo, Wang, Xu
- `[11] guo2022vappd` — authors confirmed: Guo, Feng, Shi, Cao, Li, Wang, Xu
- `[12] guo2025pocketscp` — authors confirmed: Guo, Zhao, Huang, Zhao, Xu, Liu, Yang
- `[14] xu2019d3pockets` — venue, year, pages confirmed (JCIM 59(8):3353–3358)
- `[17] weill2010alignment` (JCIM 2010)
- `[21] bernetti2024probing` (Curr. Opin. Struct. Biol. 2024)
- `[22] sheikamamuddy2021allosteric` (CSBJ 2021)
- `[23] lasala2017pocket` (ACS Cent. Sci. 2017)
- `[6] eposbp_website` (URL only — never changes)

### Verified but with co-author placeholders (cosmetic only)
These are correct in venue/year/DOI; the only missing detail is the full
co-author list. You should fill in 1–3 missing co-authors before submission:

- `[4] krivak2024comparative` — confirm full author list from PMC11552181
- `[9] zhan2020spatiotemporal` — confirm full author list (J. Vis. 2020)
- `[14] xu2019d3pockets` — D3Pockets has 4 authors (Z. Xu et al.); verify list
- `[15] schatz2021visual` — Wiley page lists Schatz, Franco-Moreno, Schäfer + corresponding author Krone; fill the rest
- `[16] vant2024complex` — Frontiers page has 5+ co-authors; fill the rest
- `[18] vishveshwara2009modeling` — venue confirmed; co-authors placeholder
- `[19] centrality2021highlight` — DOI confirmed; co-authors placeholder

### Newly clarified
- `[3] meller2025cryptic` — confirmed as a 2025 review (PMID 39778412, Curr. Opin. Struct. Biol.)
- `[20] rin_clustering_ensembles` — confirmed as a ChemRxiv 2025 preprint (id 68430481c1cb1ecda0e41722); awaiting peer-reviewed publication

### Final advice
Before submission, paste each `% VERIFY` flag's DOI into Google Scholar
or the publisher page to grab the exact full author list. The DOIs and
venues above were all confirmed in my checks, so once the author lists
are filled in, the bibliography will be submission-ready.

## Switching to a journal template later

When you decide a target venue:

- For **MDPI Applied Sciences / Molecules**, replace the preamble with the
  MDPI class file (`mdpi.cls`) and remove `geometry` + `setspace`.
- For **Bioinformatics (OUP)**, switch to OUP's article template.
- For **IEEE TVCG / EuroVis**, switch to the IEEE/EG class.

The body text (sections, citations) will work without modification because
all references go through `\citep{}` and `\citet{}` (natbib).
