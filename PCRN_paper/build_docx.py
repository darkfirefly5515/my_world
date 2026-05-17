#!/usr/bin/env python3
"""
Build a Microsoft Word (.docx) version of the PCRN paper draft.

The content (title, abstract, introduction, related work, references)
mirrors the LaTeX version in main.tex / references.bib so that the
Word and PDF outputs are interchangeable.

Usage:
    python build_docx.py            # creates PCRN_paper.docx in the same dir
"""

from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm, RGBColor

# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------

def add_heading(doc, text, level):
    """Add a heading with consistent style."""
    h = doc.add_heading(text, level=level)
    if level == 0:
        for run in h.runs:
            run.font.size = Pt(18)
            run.font.bold = True
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif level == 1:
        for run in h.runs:
            run.font.size = Pt(14)
            run.font.bold = True
    elif level == 2:
        for run in h.runs:
            run.font.size = Pt(12)
            run.font.bold = True
    return h


def add_para(doc, text, bold=False, italic=False, size=11,
             alignment=None, first_line_indent=None):
    p = doc.add_paragraph()
    if alignment is not None:
        p.alignment = alignment
    if first_line_indent is not None:
        p.paragraph_format.first_line_indent = first_line_indent
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    return p


def add_bullet(doc, text):
    p = doc.add_paragraph(text, style="List Bullet")
    for run in p.runs:
        run.font.size = Pt(11)
    return p


def add_numbered(doc, text):
    p = doc.add_paragraph(text, style="List Number")
    for run in p.runs:
        run.font.size = Pt(11)
    return p


# ---------------------------------------------------------------------
# Document content (mirrors main.tex)
# ---------------------------------------------------------------------

doc = Document()

# Page margins: 1 inch all around
for section in doc.sections:
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.54)
    section.right_margin = Cm(2.54)

# -------- Title --------
title_text = ("PCRN: A Pocket-derived Co-occurrence Residue Network for "
              "Analyzing Protein Pocket Dynamics")
add_heading(doc, title_text, level=0)

# -------- Author block --------
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Anonymous Author(s)\n")
run.font.size = Pt(11)
run = p.add_run("Affiliation\n")
run.font.size = Pt(10)
run.italic = True
run = p.add_run("email@institution.edu")
run.font.size = Pt(10)
run.font.color.rgb = RGBColor(0x40, 0x40, 0x80)

doc.add_paragraph()  # spacer

# -------- Abstract --------
add_heading(doc, "Abstract", level=1)

abstract_text = (
    "Protein pockets are inherently dynamic structures whose stability, "
    "merging, splitting, and allosteric coupling underpin protein function "
    "and drug binding. Existing residue-level networks predominantly encode "
    "physical proximity or non-covalent interactions, capturing structural "
    "contacts but overlooking the functional co-membership of residues "
    "within transient pockets. We introduce PCRN (Pocket-derived "
    "Co-occurrence Residue Network), in which an edge between two residues "
    "is weighted by how frequently they co-occur as members of the same "
    "pocket across a conformational ensemble, with node weights reflecting "
    "per-residue pocket participation frequency. The construction is "
    "alignment-free, builds directly on per-frame pocket detections from "
    "fpocket, and shifts the network's semantic level from physical contact "
    "to functional co-membership. We show that PCRN exposes pocket-relevant "
    "features that contact-based networks miss: tightly co-occurring residue "
    "cliques mark stable pocket cores, low-weight bridges identify hinges "
    "where pockets merge or split, and community structure tracks the "
    "temporal organization of pocket regions. PCRN is computationally "
    "lightweight, complementary to existing residue interaction networks, "
    "and offers a generalizable lens for studying dynamic pocket behavior."
)
add_para(doc, abstract_text, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

p = doc.add_paragraph()
run = p.add_run("Keywords: ")
run.font.bold = True
run.font.size = Pt(11)
p.add_run(
    "protein pocket dynamics; residue co-occurrence network; "
    "alignment-free analysis; molecular dynamics; visual analytics; "
    "allosteric communication."
).font.size = Pt(11)

# =====================================================================
# 1 Introduction
# =====================================================================
add_heading(doc, "1  Introduction", level=1)

# 1.1 ----------------------------------------------------------------
add_heading(doc, "1.1  The Importance of Protein Pockets", level=2)
add_para(doc,
    "Protein pockets — also referred to as cavities or binding sites — are "
    "geometrically concave regions formed by specific residues on the "
    "protein surface or in its interior, and they play central roles in "
    "ligand recognition, catalysis, and allosteric regulation [1, 2]. In "
    "structural biology and drug design, the accurate localization and "
    "characterization of pockets serve as a critical starting point for "
    "hit discovery, virtual screening, and mechanistic interpretation "
    "[3, 4]. Increasing evidence, however, indicates that many functionally "
    "relevant pockets — including cryptic pockets, allosteric sites, and "
    "transient pockets — are not visible in static crystal structures and "
    "can only be revealed by examining the dynamic conformational behavior "
    "of the protein [1, 3]. Consequently, systematic analysis of pockets "
    "and their dynamics has become a central topic in modern "
    "structure-based drug discovery.",
    alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

# 1.2 ----------------------------------------------------------------
add_heading(doc, "1.2  The Development of Static Pocket Detection Algorithms", level=2)
add_para(doc,
    "Pocket detection methods built on static structures have been "
    "developed for decades. Krone et al. [2] classified the mainstream "
    "approaches into four categories: grid-based, Voronoi-based "
    "(e.g., fpocket), surface-based, and probe-based methods. Tools such "
    "as EPOS_BP [6] further enable high-throughput detection of pocket "
    "ensembles on protein surfaces. Building on these algorithms, a "
    "variety of pocket characterization and similarity metrics have been "
    "proposed [5, 17] to support cross-protein comparison. The systematic "
    "evaluation of pocket prediction methods by Krivák et al. [4] shows "
    "that geometry-based and machine-learning-based methods each have "
    "their strengths, but both rely on a static structure as input.",
    alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

# 1.3 ----------------------------------------------------------------
add_heading(doc, "1.3  The Limitations of Static Pocket Detection", level=2)
add_para(doc,
    "Despite the substantial progress made in static pocket detection, "
    "its fundamental limitation lies in neglecting the intrinsic dynamics "
    "of protein conformations. Real proteins undergo continuous thermal "
    "motion, during which pocket shape, volume, accessibility, and "
    "residue composition all vary over time. Detection based on a single "
    "crystal structure (i) misses cryptic pockets that appear only in "
    "specific dynamic conformations [3]; (ii) cannot characterize the "
    "transient expansion, contraction, merging, and splitting of pockets "
    "[1, 7]; and (iii) struggles to distinguish functionally persistent "
    "pockets from transient noise-like cavities [8]. These limitations "
    "directly hinder the application of static pocket analysis to "
    "allosteric regulation studies, drug repurposing, and the targeting "
    "of \u201cundruggable\u201d proteins.",
    alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

# 1.4 ----------------------------------------------------------------
add_heading(doc, "1.4  The Importance of Dynamic Pocket Analysis", level=2)
add_para(doc,
    "With advances in hardware acceleration and enhanced sampling "
    "techniques for molecular dynamics (MD) simulations, researchers can "
    "now observe pocket evolution in conformational space at atomic "
    "resolution [1, 16]. Dynamic pocket analysis can reveal: the opening "
    "and closing mechanisms of cryptic pockets [3], transient "
    "accessibility along ligand entry/exit channels [7], and inter-site "
    "allosteric communication networks [19, 22, 23]. Stank et al. [1] "
    "systematically reviewed multiple dimensions of pocket dynamics "
    "characterization — shape, volume, stability, and inter-pocket "
    "transitions — and emphasized their irreplaceable value in drug "
    "design. Recent reviews on cryptic pockets further indicate that "
    "pocket discovery from a dynamic perspective has become a key avenue "
    "for expanding the druggable proteome [3].",
    alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

# 1.5 ----------------------------------------------------------------
add_heading(doc, "1.5  Limitations of Alignment-Based Approaches for Dynamic Pocket Analysis", level=2)
add_para(doc,
    "In existing dynamic pocket analysis frameworks, structural alignment "
    "is an implicit prerequisite for most methods: all frames of an MD "
    "trajectory must first be superimposed onto a reference structure, "
    "and then pocket density, volume, or shape can be aggregated and "
    "compared in a unified coordinate system [7, 8, 9]. This paradigm "
    "suffers from several issues:",
    alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_numbered(doc,
    "Reference selection is subjective — using different reference "
    "frames or different reference residue sets can substantially alter "
    "the resulting pocket density map [2].")
add_numbered(doc,
    "Alignment is unreliable in flexible regions — the motion of "
    "flexible loops and surface residues is often \u201caveraged out\u201d "
    "by global alignment, masking the true local pocket dynamics [5].")
add_numbered(doc,
    "Coupling between global and local motions — global translation and "
    "rotation of the protein cannot be cleanly separated from local pocket "
    "variation through alignment, and the pocket dynamics signal becomes "
    "contaminated by global pseudo-signals [16].")
add_numbered(doc,
    "Cross-protein comparison is difficult — when comparing pockets from "
    "different proteins or different mutants of the same protein, global "
    "alignment often fails because of differences in fold [5, 17].")

# 1.6 ----------------------------------------------------------------
add_heading(doc, "1.6  The Advantages of Alignment-Free Approaches", level=2)
add_para(doc,
    "To circumvent the issues above, alignment-free approaches to pocket "
    "analysis have attracted growing attention. Weill and Rognan [17] "
    "proposed an early ultra-high-throughput pocket comparison method "
    "based on pharmacophoric fingerprints, demonstrating that "
    "pocket-to-pocket similarity can be computed in milliseconds without "
    "any structural alignment. In a 2022 review, Eguida and Rognan [5] "
    "further pointed out that alignment-free methods offer significant "
    "advantages in large-scale pocket comparison, cross-family comparison, "
    "and machine-learning-oriented feature construction. The core "
    "advantages of alignment-free methods are: (i) inherent rotational "
    "and translational invariance, free from reference-frame bias; "
    "(ii) massive parallelizability, suitable for trajectories of tens of "
    "thousands of frames; and (iii) clear biological semantics — "
    "descriptions based on residue identity or topological invariants "
    "naturally preserve cross-frame correspondence. These advantages are "
    "particularly relevant to the cross-frame pocket analysis on a single "
    "protein studied in this paper, because residue numbering remains "
    "constant throughout the trajectory, making alignment-free analysis "
    "essentially \u201cfree of charge\u201d.",
    alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

# 1.7 ----------------------------------------------------------------
add_heading(doc, "1.7  The Need for Visual Analytics on Long-Timescale MD Data", level=2)
add_para(doc,
    "As MD simulation timescales extend from nanoseconds into the "
    "microsecond and even millisecond regime, a single trajectory may "
    "contain tens of thousands to millions of frames covering hundreds "
    "of thousands of atoms — a high-dimensional spatiotemporal dataset "
    "[16]. Numerical statistics alone (e.g., RMSD, RMSF, pocket-volume "
    "time series) cannot reveal the complex events embedded in pocket "
    "dynamics — opening, closing, merging, splitting, and allosteric "
    "coupling. Visual analytics, by combining automated algorithms with "
    "human perception and cognition, provides a multi-layered, "
    "interactive, and exploratory path to insight [2, 8, 16]. Both the "
    "state-of-the-art report by Krone et al. [2] and the 2024 review of "
    "MD trajectory visualization by Vant et al. [16] emphasize that "
    "dedicated visual analytics systems for dynamic pockets are a key "
    "bridge between raw MD data and biological interpretation. The work "
    "presented in this paper is motivated by this need: we propose a "
    "residue co-occurrence network–based approach to dynamic pocket "
    "analysis and visualization.",
    alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

# =====================================================================
# 2 Related Work
# =====================================================================
add_heading(doc, "2  Related Work", level=1)

# 2.1 ----------------------------------------------------------------
add_heading(doc, "2.1  Geometric and Topological Visual Analytics for Dynamic Pockets", level=2)
add_para(doc,
    "Early work on dynamic pocket visualization was primarily based on "
    "geometric and topological representations. Lindow et al. [7] "
    "proposed a preprocessing pipeline based on Voronoi diagrams and "
    "van der Waals spheres, coupled with an interactive visualization "
    "interface that allows users to dynamically select and observe cavity "
    "evolution along an MD trajectory. Parulek et al. [10] proposed a "
    "pocket extraction method based on implicit-function sampling and "
    "graph algorithms, used 3D graphs and residue-based attributes to "
    "characterize pockets, and applied the method to MD simulations of "
    "Proteinase 3. Krone et al.'s IEEE PacificVis 2014 work [13] further "
    "proposed a multiple-coordinated-view visual analytics system "
    "targeting dynamic protein cavities and binding sites. The 2016 STAR "
    "by Krone et al. [2] provided a comprehensive survey of the field "
    "and proposed the four-class taxonomy of pocket detection methods, "
    "and remains one of the most influential reviews in this area.",
    alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_para(doc,
    "More refined visual analytics systems for dynamic pockets have "
    "appeared in recent years: the spatiotemporal multiscale framework "
    "of Zhan et al. [9] presents pocket evolution at multiple semantic "
    "levels; PNMAVis [10] couples normal mode analysis (NMA) with pocket "
    "dynamics to provide mode-based interpretations of pocket variation; "
    "VAPPD [11] proposes a pocket shape representation that integrates "
    "topological features, together with a high-dimensional similarity "
    "metric and a multi-view analytical system; and the more recent "
    "PocketSCP [12] exploits the spatiotemporal topological properties "
    "of lining residue atoms for visualizing and analyzing dynamic "
    "pockets. At the tool level, D3Pockets [14] provides a systematic "
    "web service for pocket dynamics, while works such as Visual "
    "Analysis of Large-Scale Protein–Ligand Interaction Data [15] "
    "contribute complementary perspectives from the angle of "
    "large-scale protein–ligand interaction visualization.",
    alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

# 2.2 ----------------------------------------------------------------
add_heading(doc, "2.2  Residue Interaction Networks (RINs) and Their Use in MD", level=2)
add_para(doc,
    "Representing a protein structure as a residue interaction network "
    "(RIN) has become an important paradigm for analyzing protein "
    "function and dynamics [18]. In this framework, nodes represent "
    "residues and edges represent residue–residue interactions in some "
    "defined sense, compressing high-dimensional structural information "
    "into a graph that is amenable to network analysis. Centrality "
    "measures computed on RINs (e.g., degree, betweenness, closeness) "
    "have been widely used to identify functionally important residues "
    "and allosteric hubs [19].",
    alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_para(doc,
    "As MD-derived data have become more abundant, researchers have "
    "extended RINs to time-resolved settings. Salamanca-Suárez et al. "
    "[20] extracted ensembles of RINs along MD trajectories and "
    "clustered them, revealing distinct network patterns that correspond "
    "to different conformational families. A recent review by Bernetti "
    "et al. [21] systematically summarizes methodological advances that "
    "combine MD with network analysis for probing allosteric "
    "communication. Sheik Amamuddy et al. [22], working on Falcipain 2, "
    "employed dynamic RIN centrality analysis to identify allosteric "
    "pockets and network hubs associated with drug-resistance mutations, "
    "providing a representative example of how RINs can serve dynamic "
    "pocket analysis.",
    alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

# 2.3 ----------------------------------------------------------------
add_heading(doc, "2.3  Pocket-to-Pocket Communication and \u201cPocket Crosstalk\u201d Networks", level=2)
add_para(doc,
    "A few recent studies have started to construct networks at a higher "
    "abstraction level than residues. La Sala et al. [23] proposed "
    "pocket crosstalk analysis, in which the temporal exchange of atoms "
    "between pockets along an MD trajectory is treated as the edge of "
    "an allosteric communication network. By taking pockets as nodes "
    "and inter-pocket atom flow as edges, this work complements "
    "traditional residue-level networks at a higher hierarchical level "
    "and reveals inter-pocket allosteric coupling patterns that had "
    "been difficult to capture.",
    alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

# 2.4 ----------------------------------------------------------------
add_heading(doc, "2.4  Limitations of Existing Work and the Niche for This Paper", level=2)
add_para(doc,
    "The literature above can be seen as two relatively independent "
    "threads: one focusing on geometry- and topology-driven dynamic "
    "pocket visualization (e.g., VAPPD, PocketSCP, the spatiotemporal "
    "multiscale framework) [9, 11, 12]; and the other focusing on "
    "network-driven residue analysis [18, 19, 20, 21]. However:",
    alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_bullet(doc,
    "Existing residue interaction networks are largely built on physical "
    "distance or non-covalent interactions and lack a network "
    "representation at the level of functional co-membership;")
add_bullet(doc,
    "Existing dynamic pocket visualization approaches are mainly "
    "geometric or topological, and do not fully exploit the strengths of "
    "network analysis for revealing pocket stability, merging/splitting "
    "events, and allosteric communication;")
add_bullet(doc,
    "Cross-frame pocket tracking has long depended on structural "
    "alignment, and the alignment-free paradigm has not yet been "
    "systematically adopted in dynamic pocket analysis [5, 17].")

add_para(doc,
    "To address these gaps, we propose the Pocket-derived Co-occurrence "
    "Residue Network (PCRN): treating the pockets detected per-frame by "
    "fpocket as \u201cfunctional events\u201d, we use the co-occurrence "
    "frequency of residue pairs within the same pocket as edge weights "
    "and the per-residue pocket participation frequency as node weights, "
    "thereby constructing a fully alignment-free, pocket-oriented "
    "residue-level network. Combined with sliding time windows and "
    "community detection, PCRN can characterize the stable cores and "
    "dynamic edges of pockets, the merging and splitting events along "
    "the trajectory, and the allosteric communication paths between "
    "pockets, and presents these dynamic behaviors to domain experts "
    "through a multi-view visual analytics system in an interpretable "
    "manner.",
    alignment=WD_ALIGN_PARAGRAPH.JUSTIFY)

# =====================================================================
# References
# =====================================================================
add_heading(doc, "References", level=1)

REFS = [
    ("[1]", "Stank, A., Kokh, D. B., Fuller, J. C., & Wade, R. C. "
            "(2016). Protein binding pocket dynamics. "
            "Accounts of Chemical Research, 49(5), 809–815. "
            "doi: 10.1021/acs.accounts.5b00516."),
    ("[2]", "Krone, M., Kozlíková, B., Lindow, N., Baaden, M., Baum, D., "
            "Parulek, J., Hege, H.-C., & Viola, I. (2016). Visual "
            "analysis of biomolecular cavities: state of the art. "
            "Computer Graphics Forum, 35(3), 527–551. "
            "doi: 10.1111/cgf.12928."),
    ("[3]", "Meller, A., & Bowman, G. R. (2025). Computational advances "
            "in discovering cryptic pockets for drug discovery. "
            "Current Opinion in Structural Biology. PMID: 39778412."),
    ("[4]", "Krivák, R., Hoksza, D., et al. (2024). Comparative "
            "evaluation of methods for the prediction of protein–ligand "
            "binding sites. Journal of Cheminformatics, 16, 126. "
            "doi: 10.1186/s13321-024-00923-z."),
    ("[5]", "Eguida, M., & Rognan, D. (2022). Estimating the similarity "
            "between protein pockets. International Journal of "
            "Molecular Sciences, 23(20), 12462. "
            "doi: 10.3390/ijms232012462."),
    ("[6]", "Saarland University CBI. EPOS_BP: Ensemble of Pockets on "
            "Protein Surfaces with Ballpass. "
            "https://www-cbi.cs.uni-saarland.de/software/"
            "epos_bp-ensemble-of-pockets-on-protein-surfaces-with-ballpass/  "
            "(accessed 2026-05)."),
    ("[7]", "Lindow, N., Baum, D., & Hege, H.-C. (2013). Exploring "
            "cavity dynamics in biomolecular systems. BMC "
            "Bioinformatics, 14(Suppl 19), S5. "
            "doi: 10.1186/1471-2105-14-S19-S5."),
    ("[8]", "Parulek, J., Turkay, C., Reuter, N., & Viola, I. (2013). "
            "Visual cavity analysis in molecular simulations. BMC "
            "Bioinformatics, 14(Suppl 19), S4. "
            "doi: 10.1186/1471-2105-14-S19-S4."),
    ("[9]", "Zhan, Z., et al. (2020). Spatiotemporal multiscale "
            "molecular cavity visualization and visual analysis. "
            "Journal of Visualization. "
            "doi: 10.1007/s12650-020-00646-x."),
    ("[10]","Guo, D., Feng, L., Zhang, T., Guo, Y., Wang, Y., & Xu, X. "
            "(2022). PNMAVis: visual analysis tool of protein normal "
            "mode for understanding cavity dynamics. Applied Sciences, "
            "12(15), 7919. doi: 10.3390/app12157919."),
    ("[11]","Guo, D., Feng, L., Shi, C., Cao, L., Li, Y., Wang, Y., & "
            "Xu, X. (2022). VAPPD: visual analysis of protein pocket "
            "dynamics. Applied Sciences, 12(20), 10465. "
            "doi: 10.3390/app122010465."),
    ("[12]","Guo, D., Zhao, H., Huang, J., Zhao, J., Xu, X., Liu, Y., & "
            "Yang, Y. (2025). PocketSCP: a method for spatiotemporal "
            "topological visualization and analysis of protein pocket "
            "dynamics. Journal of Chemical Information and Modeling. "
            "doi: 10.1021/acs.jcim.5c00728."),
    ("[13]","Krone, M., Kozlíková, B., et al. (2014). Visual analysis "
            "of dynamic protein cavities and binding sites. In "
            "Proceedings of IEEE Pacific Visualization Symposium "
            "(PacificVis). IEEE."),
    ("[14]","Xu, Z., et al. (2019). D3Pockets: a method and web server "
            "for systematic analysis of protein pocket dynamics. "
            "Journal of Chemical Information and Modeling, 59(8), "
            "3353–3358."),
    ("[15]","Schatz, K., Franco-Moreno, J. J., Schäfer, M., et al. "
            "(2021). Visual analysis of large-scale protein–ligand "
            "interaction data. Computer Graphics Forum, 40(3) "
            "(EuroVis 2021). doi: 10.1111/cgf.14386."),
    ("[16]","Vant, J. W., et al. (2024). From complex data to clear "
            "insights: visualizing molecular dynamics trajectories. "
            "Frontiers in Bioinformatics, 4, 1356659. "
            "doi: 10.3389/fbinf.2024.1356659."),
    ("[17]","Weill, N., & Rognan, D. (2010). Alignment-free "
            "ultra-high-throughput comparison of druggable "
            "protein–ligand binding sites. Journal of Chemical "
            "Information and Modeling, 50(1), 123–135. "
            "doi: 10.1021/ci900349y."),
    ("[18]","Vishveshwara, S., Ghosh, A., & Hansia, P. (2009). Modeling "
            "proteins as residue interaction networks. Current Protein "
            "and Peptide Science."),
    ("[19]","(2021). Centrality measures in residue interaction "
            "networks to highlight amino acids in protein–protein "
            "binding. Frontiers in Bioinformatics, 1, 684970. "
            "doi: 10.3389/fbinf.2021.684970."),
    ("[20]","(2025). Clustering and analyzing ensembles of residue "
            "interaction networks from molecular dynamics simulations. "
            "ChemRxiv preprint. "
            "https://www.chemrxiv.org/engage/chemrxiv/article-details/"
            "68430481c1cb1ecda0e41722"),
    ("[21]","Bernetti, M., Bosio, S., Bresciani, V., & Masetti, M. "
            "(2024). Probing allosteric communication with combined "
            "molecular dynamics simulations and network analysis. "
            "Current Opinion in Structural Biology, 86, 102820. "
            "doi: 10.1016/j.sbi.2024.102820."),
    ("[22]","Sheik Amamuddy, O., Glenister, M., & Tastan Bishop, Ö. "
            "(2021). Allosteric pockets and dynamic residue network "
            "hubs of Falcipain 2 in mutations including those linked "
            "to artemisinin resistance. Computational and Structural "
            "Biotechnology Journal, 19, 5868–5882. "
            "doi: 10.1016/j.csbj.2021.10.011."),
    ("[23]","La Sala, G., Decherchi, S., De Vivo, M., & Rocchia, W. "
            "(2017). Allosteric communication networks in proteins "
            "revealed through pocket crosstalk analysis. ACS Central "
            "Science, 3(9), 949–960. "
            "doi: 10.1021/acscentsci.7b00211."),
]

for marker, text in REFS:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.8)
    p.paragraph_format.first_line_indent = Cm(-0.8)
    run = p.add_run(marker + " ")
    run.font.bold = True
    run.font.size = Pt(10)
    run = p.add_run(text)
    run.font.size = Pt(10)

# ---------------------------------------------------------------------
# Save
# ---------------------------------------------------------------------
out = Path(__file__).parent / "PCRN_paper.docx"
doc.save(out)
print(f"Wrote: {out}  ({out.stat().st_size:,} bytes)")
