#!/usr/bin/env python3
"""
生成中文版 PCRN 论文草稿 (.docx)。

内容与 main.tex / build_docx.py 的英文版完全对应，包括标题、摘要、
引言、相关工作和参考文献。

使用：
    python3 build_docx_zh.py    # 生成 PCRN_paper_zh.docx
"""

from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, Cm, RGBColor

# ---------------------------------------------------------------------
# 字体设置 — 让中文显示为宋体，英文显示为 Times New Roman
# ---------------------------------------------------------------------

CN_FONT = "宋体"          # 正文中文字体
CN_HEAD_FONT = "黑体"     # 标题中文字体
EN_FONT = "Times New Roman"


def set_run_fonts(run, size_pt, *, bold=False, italic=False,
                  cn_font=CN_FONT, en_font=EN_FONT,
                  color=None):
    """同时设置中英文字体，避免中文显示异常。"""
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    run.font.name = en_font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        from docx.oxml import OxmlElement
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), en_font)
    rFonts.set(qn("w:hAnsi"), en_font)
    rFonts.set(qn("w:eastAsia"), cn_font)
    rFonts.set(qn("w:cs"), en_font)


def add_heading(doc, text, level):
    """添加标题；中文用黑体加粗。"""
    h = doc.add_paragraph()
    if level == 0:
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        size = 18
    elif level == 1:
        size = 14
    else:
        size = 12
    run = h.add_run(text)
    set_run_fonts(run, size, bold=True, cn_font=CN_HEAD_FONT)
    h.paragraph_format.space_before = Pt(12)
    h.paragraph_format.space_after = Pt(6)
    return h


def add_para(doc, text, *, size=11, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
             first_line_indent=Cm(0.74), bold=False, italic=False):
    p = doc.add_paragraph()
    p.alignment = alignment
    if first_line_indent:
        p.paragraph_format.first_line_indent = first_line_indent
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(text)
    set_run_fonts(run, size, bold=bold, italic=italic)
    return p


def add_bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.line_spacing = 1.5
    run = p.add_run(text)
    set_run_fonts(run, 11)
    return p


def add_numbered(doc, text):
    p = doc.add_paragraph(style="List Number")
    p.paragraph_format.line_spacing = 1.5
    run = p.add_run(text)
    set_run_fonts(run, 11)
    return p


# ---------------------------------------------------------------------
# 文档构建
# ---------------------------------------------------------------------

doc = Document()

# 页边距 1 英寸
for section in doc.sections:
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.54)
    section.right_margin = Cm(2.54)

# -------- 标题 --------
add_heading(doc,
            "PCRN：一种用于分析蛋白质口袋动态的口袋衍生残基共现网络",
            level=0)

# -------- 作者块 --------
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("匿名作者")
set_run_fonts(run, 11)
p.add_run("\n")
run = p.add_run("作者单位")
set_run_fonts(run, 10, italic=True)
p.add_run("\n")
run = p.add_run("email@institution.edu")
set_run_fonts(run, 10, color=RGBColor(0x40, 0x40, 0x80))

doc.add_paragraph()

# -------- 摘要 --------
add_heading(doc, "摘要", level=1)

abstract_text = (
    "蛋白质口袋本质上是动态结构，其稳定性、合并、分裂以及变构耦合"
    "是蛋白质功能与药物结合的关键基础。然而，目前的残基级网络主要"
    "依据残基之间的物理邻近性或非共价相互作用进行连边，能够刻画结"
    "构性接触，却难以捕捉残基在瞬态口袋中的功能性共归属信息。本文"
    "提出 PCRN（Pocket-derived Co-occurrence Residue Network，口袋"
    "衍生残基共现网络）：以两个残基在同一个口袋中共同出现的频率作"
    "为边权，以单个残基的口袋出现频率作为节点权。该方法完全免对齐"
    "（alignment-free），直接基于 fpocket 在每一帧上检测得到的口袋"
    "构建网络，使得网络的语义层级从“物理接触”跃迁到“功能共归属”。"
    "实验表明，PCRN 能够揭示传统接触网络无法捕获的口袋特性：高密"
    "度共现的残基团对应于稳定的口袋核心；低权值的桥边对应于口袋发"
    "生合并或分裂的铰链区域；而社区结构（community structure）则"
    "反映了口袋区域随时间的组织演化。PCRN 计算开销低、可与已有残"
    "基相互作用网络互补，为研究蛋白质口袋的动态行为提供了一种通用"
    "的分析视角。"
)
add_para(doc, abstract_text)

p = doc.add_paragraph()
p.paragraph_format.first_line_indent = Cm(0.74)
p.paragraph_format.line_spacing = 1.5
run = p.add_run("关键词：")
set_run_fonts(run, 11, bold=True)
run = p.add_run(
    "蛋白质口袋动态；残基共现网络；免对齐分析；分子动力学模拟；"
    "可视分析；变构通信"
)
set_run_fonts(run, 11)

# =====================================================================
# 1 引言
# =====================================================================
add_heading(doc, "1  引言", level=1)

# ---- 1.1 ----
add_heading(doc, "1.1  蛋白质口袋的重要性", level=2)
add_para(doc,
    "蛋白质口袋（pocket / cavity / binding site）是蛋白表面或内"
    "部由特定残基围成的几何凹陷区域，承担着配体识别、催化反应、"
    "变构调控等核心功能 [1, 2]。在结构生物学和药物设计中，对口"
    "袋的精确定位与表征是先导化合物筛选、虚拟筛选与作用机制阐明"
    "的关键起点 [3, 4]。然而，越来越多的证据表明，许多重要的功"
    "能性口袋——包括隐蔽口袋（cryptic pockets）、变构位点（"
    "allosteric sites）和瞬态口袋（transient pockets）——在静态"
    "晶体结构中并不可见，必须通过对蛋白构象动态变化的研究才能揭"
    "示 [1, 3]。因此，对口袋及其动态行为的系统分析已成为现代基"
    "于结构的药物发现的核心议题之一。")

# ---- 1.2 ----
add_heading(doc, "1.2  静态口袋检测算法的发展", level=2)
add_para(doc,
    "围绕静态结构的口袋检测算法已有数十年的发展。Krone 等人 [2] "
    "在其综述中将主流方法归纳为四大类：基于格点（grid-based）、"
    "基于 Voronoi 分解（Voronoi-based，如 fpocket）、基于表面的"
    "（surface-based）和基于探针的（probe-based）方法。EPOS_BP "
    "[6] 等工具进一步提供了对蛋白表面口袋集合的高通量探测。在此"
    "基础上，研究者也发展出多种口袋表征与相似性度量方法 [5, 17]"
    "，以支持跨蛋白的比较分析。Krivák 等人 [4] 对 50 余种主流口"
    "袋预测方法的系统对比表明，基于几何与基于机器学习的方法各有"
    "所长，但都依赖于静态结构作为输入。")

# ---- 1.3 ----
add_heading(doc, "1.3  静态口袋检测的局限性", level=2)
add_para(doc,
    "尽管静态口袋检测取得了显著进展，但其根本局限在于忽视了蛋白"
    "质构象的内在动力学。真实蛋白处于持续的热运动中，口袋的形状"
    "、体积、可及性以及残基组成都会随时间发生变化。仅基于单一晶"
    "体结构的检测：（1）会遗漏 cryptic pocket——只在特定动态构"
    "象下才出现的隐藏位点 [3]；（2）无法表征口袋的瞬态扩张、收"
    "缩、合并与分裂行为 [1, 7]；（3）难以区分功能性持续口袋与短"
    "暂出现的噪声腔体 [8]。这些局限直接限制了静态口袋分析在变构"
    "调控研究、药物再利用以及“难成药靶点”研究中的应用。")

# ---- 1.4 ----
add_heading(doc, "1.4  蛋白质动态口袋分析的重要性", level=2)
add_para(doc,
    "随着分子动力学（Molecular Dynamics, MD）模拟在硬件加速与采"
    "样技术上的进步，研究者得以在原子级别观察蛋白构象空间中的口"
    "袋演化 [1, 16]。动态口袋分析能够揭示：cryptic pocket 的开"
    "启与关闭机制 [3]，配体进出通道的瞬态可及性 [7]，以及位点间"
    "的变构通信网络 [19, 22, 23]。Stank 等人 [1] 系统综述了口袋"
    "动力学的多种刻画维度——形状、体积、稳定性、相互转换——并指"
    "出其对药物设计具有不可替代的价值。近年来的 cryptic pocket "
    "综述进一步表明，动态视角下的口袋发现已成为拓展可成药蛋白质"
    "组（druggable proteome）的关键途径 [3]。")

# ---- 1.5 ----
add_heading(doc, "1.5  传统基于结构对齐分析动态口袋方法的局限性", level=2)
add_para(doc,
    "在已有的动态口袋分析框架中，结构对齐（structural alignment"
    "）是大多数方法的隐式前提：它需要把 MD 轨迹的所有帧叠合到一"
    "个参考结构上，然后在统一坐标系下累积口袋密度、计算口袋体积"
    "或比较口袋形状 [7, 8, 9]。该范式存在以下几方面问题：")

add_numbered(doc,
    "对齐参考的选择具有主观性——选用不同的参考帧或参考残基集合，"
    "会显著改变累积出的“口袋密度图” [2]。")
add_numbered(doc,
    "柔性区域的对齐不可靠——蛋白柔性 loop、表面残基的运动会被全"
    "局对齐“平均掉”，反而掩盖了真正的局部口袋动态 [5]。")
add_numbered(doc,
    "整体运动与局部运动的耦合——蛋白整体的平动 / 转动与口袋局部"
    "变化在对齐中难以分离，导致口袋动态信号被全局伪信号污染 [16]。")
add_numbered(doc,
    "跨蛋白比较困难——当比较来自不同蛋白或同一蛋白不同突变体的"
    "口袋时，全局对齐往往因折叠差异而失败 [5, 17]。")

# ---- 1.6 ----
add_heading(doc, "1.6  免对齐（Alignment-free）方法的优势", level=2)
add_para(doc,
    "为了规避上述问题，免对齐口袋分析方法逐渐受到关注。Weill 与 "
    "Rognan [17] 较早提出基于药效团指纹的超高通量比较方法，证明"
    "无需结构对齐即可在毫秒级完成口袋之间的相似度计算。Eguida "
    "与 Rognan [5] 在 2022 年的综述中进一步指出，免对齐方法在大"
    "规模口袋比较、跨家族比较以及面向机器学习的特征构建上具有显"
    "著优势。免对齐方法的核心优势在于：（1）天然旋转 / 平移不变"
    "，不受参考帧选择影响；（2）可大规模并行，适用于上万帧 MD "
    "轨迹；（3）生物语义清晰——基于残基编号或拓扑量的描述天然保"
    "持跨帧的对应关系。这些优势对于本文所研究的“同一蛋白 MD 帧"
    "间口袋分析”尤其重要，因为残基编号在轨迹中始终不变，免对齐"
    "分析几乎是“免费的午餐”。")

# ---- 1.7 ----
add_heading(doc, "1.7  长时 MD 模拟数据的复杂性需要可视分析", level=2)
add_para(doc,
    "随着 MD 模拟时长从纳秒延伸到微秒甚至毫秒尺度，单条轨迹可能"
    "包含数万至数百万帧、数十万原子的高维时空数据 [16]。仅靠数"
    "值统计指标（如 RMSD、RMSF、口袋体积时间曲线）无法揭示口袋"
    "动力学中的复杂事件——开闭、合并、分裂、变构耦合等。可视分"
    "析（visual analytics）通过将自动化算法与人类感知—认知过程"
    "结合，提供了多层次、可交互、可探索的洞察途径 [2, 8, 16]。"
    "Krone 等人 [2] 的综述与 2024 年的 MD 可视化综述 [16] 共同"
    "强调，针对动态口袋的专用可视分析系统是连接 MD 模拟原始数据"
    "与生物学解释的关键桥梁。本文工作正是在这一背景下，提出基于"
    "残基共现网络的口袋动态分析与可视化方法。")

# =====================================================================
# 2 相关工作
# =====================================================================
add_heading(doc, "2  相关工作", level=1)

# ---- 2.1 ----
add_heading(doc, "2.1  动态口袋的几何与拓扑可视分析方法", level=2)
add_para(doc,
    "早期的动态口袋可视化研究主要基于几何与拓扑表征。Lindow 等"
    "人 [7] 提出基于 Voronoi 图与范德华球的预处理流程，配合交互"
    "式可视化界面，让用户在 MD 轨迹上动态选择并观察空腔的演化。"
    "Parulek 等人 [8] 提出基于隐函数采样与图算法的口袋提取方法"
    "，用 3D 图与残基属性表征口袋，并应用于 Proteinase 3 的 MD "
    "模拟分析。Krone 等人 [13] 在 IEEE PacificVis 2014 的工作中"
    "进一步提出了针对动态口袋与结合位点的多视图可视分析系统。"
    "Krone 等人 2016 年发表的 STAR 报告 [2] 则系统梳理了该方向"
    "的研究现状，并提出了口袋检测方法的四类分类法，至今仍是该领"
    "域最具影响力的综述之一。")

add_para(doc,
    "近年来更精细的动态口袋可视分析系统不断涌现：Zhan 等人 [9] "
    "提出的时空多尺度可视分析框架在不同的语义抽象层级上展示口袋"
    "演化；PNMAVis [10] 将正态模式分析（NMA）与口袋动力学耦合，"
    "提供基于模式的口袋变化解释；VAPPD [11] 提出结合拓扑特征的"
    "口袋形状表征方法，并设计高维相似度度量与多视图分析系统；最"
    "新的 PocketSCP [12] 利用衬里氨基酸的时空拓扑性质对动态口袋"
    "进行可视化与分析。在工具层面，D3Pockets [14] 提供了一套面"
    "向口袋动态的系统性 Web 服务，而 Visual Analysis of "
    "Large-Scale Protein-Ligand Interaction Data [15] 等工作则"
    "从大规模口袋—配体相互作用数据可视化的角度提供了补充思路。")

# ---- 2.2 ----
add_heading(doc, "2.2  残基相互作用网络（RIN）及其在 MD 中的应用", level=2)
add_para(doc,
    "将蛋白结构表示为残基相互作用网络（Residue Interaction "
    "Network, RIN）已成为分析蛋白功能与动力学的重要范式 [18]。"
    "在该框架中，节点为残基，边为某种意义下的相互作用关系，从而"
    "把高维结构信息压缩到便于网络分析的图结构中。在 RIN 上计算"
    "的中心性指标（degree、betweenness、closeness 等）已被广泛"
    "用于识别功能关键残基与变构枢纽 [19]。")

add_para(doc,
    "随着 MD 模拟数据的丰富，研究者将 RIN 拓展到时变情境。"
    "Salamanca-Suárez 等人 [20] 从 MD 轨迹中提取一系列 RIN 并"
    "进行聚类，揭示了蛋白构象族对应的不同网络模式。Bernetti 等"
    "人 [21] 最新综述系统总结了将 MD 与网络分析结合用于变构通信"
    "探测的方法学进展。Sheik Amamuddy 等人 [22] 针对 Falcipain "
    "2 系统，使用动态 RIN 中心性分析识别了变构口袋及与抗药性突"
    "变相关的网络枢纽，提供了 RIN 服务于口袋分析的一个典型范例。")

# ---- 2.3 ----
add_heading(doc, "2.3  口袋间通信与“口袋串扰”网络", level=2)
add_para(doc,
    "少数前沿工作开始尝试在比“残基”更高的抽象层级上构建网络。"
    "La Sala 等人 [23] 提出“pocket crosstalk 分析”——把 MD 轨"
    "迹中口袋之间在时间上交换原子的行为作为变构通信网络的边。该"
    "工作以口袋为节点、以口袋间原子流动为边，与传统的残基级网络"
    "形成层级互补，揭示了 MD 模拟中此前难以捕捉的口袋间变构耦合"
    "模式。")

# ---- 2.4 ----
add_heading(doc, "2.4  现有研究的不足与本工作的切入点", level=2)
add_para(doc,
    "综合以上文献可以看出，目前的研究存在两条相对独立的脉络：一"
    "条是几何 / 拓扑驱动的口袋动态可视分析（如 VAPPD、"
    "PocketSCP、Spatiotemporal Multiscale）[9, 11, 12]；另一条"
    "是网络驱动的残基分析 [18, 19, 20, 21]。然而：")

add_bullet(doc,
    "既有的残基相互作用网络主要基于物理距离或非共价作用，缺少"
    "“功能性共归属”层级的网络表征；")
add_bullet(doc,
    "既有的动态口袋可视分析方法主要基于几何 / 拓扑表征，未充分"
    "利用网络分析在揭示口袋稳定性、合并 / 分裂事件以及变构通信"
    "上的优势；")
add_bullet(doc,
    "跨帧口袋追踪长期依赖结构对齐，免对齐范式在动态口袋分析中尚"
    "未被系统采纳 [5, 17]。")

add_para(doc,
    "针对上述不足，本文提出口袋衍生残基共现网络（Pocket-derived "
    "Co-occurrence Residue Network, PCRN）：将 fpocket 逐帧检测"
    "出的口袋作为“功能事件”，把残基对在同一口袋中的共现频率作为"
    "边权、把残基的口袋出现频率作为节点权，构建一个完全免对齐、"
    "面向口袋动态的残基级网络。结合滑动时间窗口与社区检测，PCRN "
    "能够刻画口袋的稳定核心、动态边缘、合并 / 分裂事件以及口袋"
    "间的变构通信路径，并通过多视图可视分析系统将这些动力学行为"
    "以可解释的方式呈现给生物学专家。")

# =====================================================================
# 参考文献
# =====================================================================
add_heading(doc, "参考文献", level=1)

REFS = [
    ("[1]", "Stank, A., Kokh, D. B., Fuller, J. C., & Wade, R. C. "
            "(2016). Protein binding pocket dynamics. Accounts of "
            "Chemical Research, 49(5), 809–815. "
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
            "（访问日期 2026-05）"),
    ("[7]", "Lindow, N., Baum, D., & Hege, H.-C. (2013). Exploring "
            "cavity dynamics in biomolecular systems. BMC Bioinformatics, "
            "14(Suppl 19), S5. doi: 10.1186/1471-2105-14-S19-S5."),
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
    p.paragraph_format.line_spacing = 1.4
    run = p.add_run(marker + " ")
    set_run_fonts(run, 10, bold=True)
    run = p.add_run(text)
    set_run_fonts(run, 10)

# ---------------------------------------------------------------------
# 保存
# ---------------------------------------------------------------------
out = Path(__file__).parent / "PCRN_paper_zh.docx"
doc.save(out)
print(f"已生成：{out}（{out.stat().st_size:,} 字节）")
