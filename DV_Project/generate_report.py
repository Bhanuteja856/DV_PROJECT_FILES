"""
Generate a formatted Word document for the Data Visualization Project Report.
Global Military Firepower 2025 – Defence Analysis
"""

from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import os

OUTPUT_PATH = r"d:\DV_PROJECT\DV_Project\DV_Project_Report.docx"

# ── Helpers ──────────────────────────────────────────────────────────────────

def set_cell_shading(cell, color_hex):
    """Set background shading of a table cell."""
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def add_formatted_table(doc, headers, rows, col_widths=None, header_color="1F3864"):
    """Add a styled table with header row shading."""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header row
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run(header)
        run.bold = True
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_shading(cell, header_color)

    # Data rows
    for r_idx, row_data in enumerate(rows):
        for c_idx, cell_text in enumerate(row_data):
            cell = table.rows[r_idx + 1].cells[c_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(str(cell_text))
            run.font.size = Pt(10)
            # Alternate row shading
            if r_idx % 2 == 0:
                set_cell_shading(cell, "D6E4F0")

    # Column widths
    if col_widths:
        for row in table.rows:
            for idx, width in enumerate(col_widths):
                row.cells[idx].width = Cm(width)

    doc.add_paragraph()  # spacer
    return table


def heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
    return h


def body(doc, text):
    p = doc.add_paragraph(text)
    p.style = doc.styles["Normal"]
    return p


def bullet(doc, text, bold_prefix=None):
    p = doc.add_paragraph(style="List Bullet")
    if bold_prefix:
        run = p.add_run(bold_prefix)
        run.bold = True
        run.font.size = Pt(11)
        p.add_run(text)
    else:
        p.add_run(text)
    return p


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    doc = Document()

    # ── Page setup ───────────────────────────────────────────────────────
    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(2.54)
        section.right_margin = Cm(2.54)

    # ── Default font ─────────────────────────────────────────────────────
    style = doc.styles["Normal"]
    font = style.font
    font.name = "Times New Roman"
    font.size = Pt(12)
    style.paragraph_format.space_after = Pt(6)
    style.paragraph_format.line_spacing = 1.5

    # ══════════════════════════════════════════════════════════════════════
    # TITLE PAGE
    # ══════════════════════════════════════════════════════════════════════
    for _ in range(6):
        doc.add_paragraph()

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("DATA VISUALIZATION PROJECT REPORT")
    run.bold = True
    run.font.size = Pt(26)
    run.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run("Global Military Firepower 2025\nDefence Analysis")
    run.font.size = Pt(18)
    run.font.color.rgb = RGBColor(0x4A, 0x6F, 0xA5)

    doc.add_paragraph()

    # Info block
    info = doc.add_paragraph()
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for line in [
        "Subject: Data Visualization",
        "Batch: 2",
        "Academic Year: 2025–2026",
        "Tools Used: Microsoft Excel  |  Microsoft Power BI",
        "Dataset: Global Military Firepower 2025",
    ]:
        run = info.add_run(line + "\n")
        run.font.size = Pt(12)

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════
    # TABLE OF CONTENTS (Manual)
    # ══════════════════════════════════════════════════════════════════════
    heading(doc, "Table of Contents", level=1)
    toc_items = [
        "1.  Abstract",
        "2.  Introduction",
        "3.  Literature Review",
        "4.  Objectives",
        "5.  Dataset Description",
        "6.  Tools and Technologies Used",
        "7.  Methodology",
        "    7.1  Data Collection",
        "    7.2  Data Cleaning",
        "    7.3  Pivot Table Analysis",
        "8.  Dashboard Design and Visualizations",
        "9.  Results and Analysis",
        "10. Conclusion",
        "11. Future Scope",
        "12. References",
    ]
    for item in toc_items:
        p = doc.add_paragraph(item)
        p.paragraph_format.space_after = Pt(2)
        for run in p.runs:
            run.font.size = Pt(12)

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════
    # 1. ABSTRACT
    # ══════════════════════════════════════════════════════════════════════
    heading(doc, "1. Abstract", level=1)
    body(doc,
        "Defence capability analysis is critical for understanding global geopolitical dynamics. "
        "This project presents a comprehensive data visualization study of the Global Military "
        "Firepower 2025 dataset, which ranks 145 countries across parameters such as total population, "
        "military manpower, active personnel, reserve forces, aircraft, naval fleet, land coverage, "
        "sea coverage, and natural resource consumption (coal, oil, natural gas)."
    )
    body(doc,
        "Using Microsoft Excel for data cleaning and pivot table analysis, and Microsoft Power BI "
        "for interactive dashboard creation, we transform raw numerical data into meaningful visual "
        "narratives. The dashboard includes a Map Chart for geographical distribution, a Line Chart "
        "for comparing naval assets (submarines, tankers, bunkers) across countries, a Bar Chart for "
        "country-wise aircraft counts, and a Pie Chart for coal consumption distribution."
    )
    body(doc,
        "The study aims to assist researchers, analysts, and policy-makers in quickly grasping "
        "comparative defence strengths across nations and regions."
    )

    # ══════════════════════════════════════════════════════════════════════
    # 2. INTRODUCTION
    # ══════════════════════════════════════════════════════════════════════
    heading(doc, "2. Introduction", level=1)
    body(doc,
        "In an era of rapidly evolving global security landscapes, access to clear, comparative "
        "defence data is more important than ever. Raw datasets containing hundreds of rows and "
        "dozens of columns are difficult to interpret in tabular form. Data visualization bridges "
        "this gap by converting complex data into intuitive charts, graphs, and dashboards that "
        "reveal patterns, outliers, and trends at a glance."
    )
    body(doc,
        "The Global Military Firepower (GFP) 2025 index ranks 145 nations on over 60 individual "
        "factors — from troop strength and equipment counts to natural resource reserves and "
        "geographic considerations — to produce a composite \"PowerIndex\" score. For this project, "
        "we selected a focused subset of columns relevant to population, manpower, aircraft, naval "
        "fleet, land and sea coverage, and energy consumption."
    )
    body(doc,
        "This report documents the entire process from data acquisition and cleaning through pivot "
        "table construction to final dashboard design in Power BI. The goal is to demonstrate how "
        "thoughtful visualization transforms defence statistics into actionable insights."
    )

    # ══════════════════════════════════════════════════════════════════════
    # 3. LITERATURE REVIEW
    # ══════════════════════════════════════════════════════════════════════
    heading(doc, "3. Literature Review", level=1)
    body(doc,
        "Data visualization has long been recognized as a powerful technique for communicating "
        "quantitative information. Key references informing this work include:"
    )
    lit_items = [
        ("Edward Tufte, The Visual Display of Quantitative Information (1983) — ",
         "Established foundational principles of data-ink ratio and chart integrity."),
        ("Stephen Few, Show Me the Numbers (2012) — ",
         "Provided practical guidance on choosing chart types for different analytical questions."),
        ("Global Firepower (globalfirepower.com) — ",
         "The authoritative source for military strength rankings and the dataset used in this study."),
        ("Microsoft Power BI Documentation — ",
         "Official guides on creating interactive dashboards, DAX expressions, and map visualizations."),
        ("Kaggle Military Datasets — ",
         "Community-curated datasets used for benchmarking and cross-validation."),
    ]
    for bold_part, normal_part in lit_items:
        bullet(doc, normal_part, bold_prefix=bold_part)

    body(doc,
        "Previous studies have visualized military data using Python (Matplotlib, Seaborn) and Tableau. "
        "This project differentiates itself by combining Excel-based pivot table analysis with Power BI "
        "dashboards, a workflow widely adopted in business and government settings."
    )

    # ══════════════════════════════════════════════════════════════════════
    # 4. OBJECTIVES
    # ══════════════════════════════════════════════════════════════════════
    heading(doc, "4. Objectives", level=1)
    objectives = [
        "Clean and prepare the Global Military Firepower 2025 dataset for analysis.",
        "Create pivot tables in Excel to summarize data across multiple dimensions (country, region, military branch).",
        "Design an interactive Power BI dashboard with diverse chart types to visualize defence metrics.",
        "Derive meaningful insights about global military strength distribution, regional dominance, and resource consumption patterns.",
        "Present findings in a clear, professional format suitable for academic and policy review.",
    ]
    for i, obj in enumerate(objectives, 1):
        p = doc.add_paragraph(style="List Number")
        p.add_run(obj)

    # ══════════════════════════════════════════════════════════════════════
    # 5. DATASET DESCRIPTION
    # ══════════════════════════════════════════════════════════════════════
    heading(doc, "5. Dataset Description", level=1)

    heading(doc, "5.1 Source", level=2)
    body(doc,
        "The dataset was obtained from the Global Firepower (GFP) 2025 ranking, publicly available "
        "at globalfirepower.com. A selected-column version was prepared in CSV format for this project."
    )

    heading(doc, "5.2 Structure", level=2)
    add_formatted_table(doc,
        headers=["Column", "Description", "Data Type"],
        rows=[
            ["rank", "GFP PowerIndex rank (1 = strongest)", "Integer"],
            ["country", "Nation name", "Text"],
            ["region", "Geographic region (Asia, Europe, Africa, etc.)", "Text"],
            ["total_population", "Total population of the country", "Integer"],
            ["total_military_manpower", "Population available for military service", "Integer"],
            ["fit_for_service", "Individuals physically fit for service", "Integer"],
            ["population_reaching_military_age_annually", "New potential recruits per year", "Integer"],
            ["active_personnel", "Currently serving military members", "Integer"],
            ["reserve_personnel", "Reserve military forces", "Integer"],
            ["paramilitary", "Paramilitary / auxiliary forces", "Integer"],
        ],
        col_widths=[5.5, 8, 2.5],
    )

    heading(doc, "5.3 Scale", level=2)
    bullet(doc, "145 countries across 7 regions: Asia, Europe, Africa, North America, South America, Oceania.")
    bullet(doc,
        "The Excel workbook (global_military_firepower_2025.xlsx) contains additional columns for "
        "aircraft, naval fleet, land coverage, sea coverage, coal data, natural gas consumption, and "
        "oil consumption, which were used for pivot tables and dashboard charts."
    )

    heading(doc, "5.4 Sample Records", level=2)
    add_formatted_table(doc,
        headers=["Rank", "Country", "Region", "Total Population", "Active Personnel"],
        rows=[
            ["1", "United States", "North America", "341,963,408", "1,328,000"],
            ["2", "Russia", "Asia", "140,820,810", "1,320,000"],
            ["3", "China", "Asia", "1,415,043,270", "2,035,000"],
            ["4", "India", "Asia", "1,409,128,296", "1,455,550"],
            ["5", "South Korea", "Asia", "52,081,799", "600,000"],
        ],
        col_widths=[2, 3.5, 3.5, 3.5, 3.5],
    )

    # ══════════════════════════════════════════════════════════════════════
    # 6. TOOLS AND TECHNOLOGIES
    # ══════════════════════════════════════════════════════════════════════
    heading(doc, "6. Tools and Technologies Used", level=1)

    heading(doc, "6.1 Microsoft Excel", level=2)
    bullet(doc, "Data cleaning, removal of duplicates, creation of pivot tables.", bold_prefix="Purpose: ")
    bullet(doc, "Microsoft 365 / Excel 2021.", bold_prefix="Version: ")
    bullet(doc, "Sort & Filter, Remove Duplicates, Pivot Tables, Pivot Charts, Conditional Formatting.", bold_prefix="Features Used: ")

    heading(doc, "6.2 Microsoft Power BI Desktop", level=2)
    bullet(doc, "Interactive dashboard creation and advanced visualization.", bold_prefix="Purpose: ")
    bullet(doc, "Import Data, Data Modeling, Map Chart (ArcGIS / Bing Maps), Line Chart, Bar Chart, Pie Chart, Slicers, Filters, Custom Themes.", bold_prefix="Features Used: ")

    heading(doc, "6.3 Supporting Tools", level=2)
    add_formatted_table(doc,
        headers=["Tool", "Purpose"],
        rows=[
            ["CSV / XLSX file format", "Data interchange"],
            ["Canva", "Dashboard background design"],
            ["Power BI Service (optional)", "Cloud publishing"],
        ],
        col_widths=[5, 11],
    )

    # ══════════════════════════════════════════════════════════════════════
    # 7. METHODOLOGY
    # ══════════════════════════════════════════════════════════════════════
    heading(doc, "7. Methodology", level=1)

    heading(doc, "7.1 Data Collection", level=2)
    body(doc,
        "The raw dataset was downloaded from the Global Firepower website in Excel format "
        "(global_military_firepower_2025.xlsx). A subset of relevant columns was exported as a CSV "
        "file (global_military_firepower_2025-selected-columns.csv) containing 145 records and 10 "
        "key attributes."
    )

    heading(doc, "7.2 Data Cleaning", level=2)
    body(doc, "The following data-cleaning steps were performed in Microsoft Excel:")
    cleaning_steps = [
        ("Duplicate Removal: ", "Duplicate country entries for the Asia region were identified and removed to ensure each country appeared only once."),
        ("Missing Value Check: ", "All cells were inspected for blanks or null values; countries with 0 values in certain columns (e.g., Iceland with 0 active personnel) were retained as they represent legitimate data points."),
        ("Data Type Validation: ", "Numeric columns were verified to be formatted as numbers (not text) to ensure correct aggregation in pivot tables."),
        ("Column Standardization: ", "Column headers were standardized to a consistent naming convention for seamless import into Power BI."),
    ]
    for bold_part, normal_part in cleaning_steps:
        bullet(doc, normal_part, bold_prefix=bold_part)

    heading(doc, "7.3 Pivot Table Analysis", level=2)
    body(doc,
        "Twelve (12) pivot tables were created in Excel to summarize the data from multiple "
        "analytical angles:"
    )
    add_formatted_table(doc,
        headers=["PT #", "Pivot Table Name", "Purpose"],
        rows=[
            ["PT 1", "Country-wise Aircraft", "Compare air force size across nations"],
            ["PT 2", "Region-wise Population & Manpower", "Regional demographic strength"],
            ["PT 3", "Region-wise Aircraft", "Air power by continent"],
            ["PT 4", "Asia — Country-wise Population & Military", "Deep dive into Asia's military demographics"],
            ["PT 5", "Country-wise Attack Aircraft", "Offensive air capability"],
            ["PT 6", "Country-wise Naval Fleet", "Naval strength comparison"],
            ["PT 7", "Country-wise Military Data (Submarines, Rocket Projectors, Naval Fleet, Destroyers)", "Multi-metric military comparison"],
            ["PT 8", "Country-wise Sea Coverage", "Maritime territory"],
            ["PT 9", "Country-wise Land Coverage", "Territorial land area"],
            ["PT 10", "Country-wise Coal Data", "Energy resource — Coal"],
            ["PT 11", "Country-wise Natural Gas Consumption", "Energy resource — Gas"],
            ["PT 12", "Country-wise Oil Consumption", "Energy resource — Oil"],
        ],
        col_widths=[1.5, 7, 7],
    )

    # ══════════════════════════════════════════════════════════════════════
    # 8. DASHBOARD DESIGN AND VISUALIZATIONS
    # ══════════════════════════════════════════════════════════════════════
    heading(doc, "8. Dashboard Design and Visualizations", level=1)

    heading(doc, "8.1 Dashboard Overview", level=2)
    body(doc,
        "The Power BI dashboard (dashboard.pbix) was designed with a custom military-themed "
        "background featuring a world map, fighter jets, a battle tank, naval warships, and "
        "satellite imagery — creating an immersive visual context for the data."
    )
    body(doc, "The dashboard contains four primary visualizations:")

    # Chart 1
    heading(doc, "8.2 Chart 1 — Map Chart (Geographical Distribution)", level=2)
    add_formatted_table(doc,
        headers=["Property", "Detail"],
        rows=[
            ["Chart Type", "Filled Map / Bubble Map"],
            ["Data Field", "Country, Region"],
            ["Metric", "Military Firepower Rank / Active Personnel"],
            ["Purpose", "Geographically plot each country's military strength on a world map"],
        ],
        col_widths=[4, 12],
    )
    body(doc,
        "Design Rationale: A map chart provides instant spatial context. Users can hover over any "
        "country to see its rank, population, and active personnel. Color gradients or bubble sizes "
        "encode the magnitude of military power."
    )

    # Chart 2
    heading(doc, "8.3 Chart 2 — Line Chart (Naval Assets Comparison)", level=2)
    add_formatted_table(doc,
        headers=["Property", "Detail"],
        rows=[
            ["Chart Type", "Multi-line Chart"],
            ["X-Axis", "Country"],
            ["Y-Axis (Lines)", "No. of Submarines, No. of Tankers, No. of Bunkers"],
            ["Purpose", "Compare naval sub-categories across nations"],
        ],
        col_widths=[4, 12],
    )
    body(doc,
        "Design Rationale: A line chart with multiple series allows side-by-side tracking of "
        "different naval asset types. Peaks indicate countries with dominant submarine or tanker "
        "fleets (e.g., United States, Russia, China)."
    )

    # Chart 3
    heading(doc, "8.4 Chart 3 — Bar Chart (Country-wise Aircraft)", level=2)
    add_formatted_table(doc,
        headers=["Property", "Detail"],
        rows=[
            ["Chart Type", "Horizontal / Vertical Bar Chart"],
            ["X-Axis", "Country"],
            ["Y-Axis", "Total Aircraft Count"],
            ["Purpose", "Rank countries by total air force fleet size"],
        ],
        col_widths=[4, 12],
    )
    body(doc,
        "Design Rationale: Bar charts are ideal for comparing a single metric across many categories. "
        "The chart clearly shows the United States, Russia, and China as the top three nations in "
        "total aircraft, with a significant drop-off for subsequent countries."
    )

    # Chart 4
    heading(doc, "8.5 Chart 4 — Pie Chart (Coal Consumption)", level=2)
    add_formatted_table(doc,
        headers=["Property", "Detail"],
        rows=[
            ["Chart Type", "Pie / Donut Chart"],
            ["Segments", "Country or Region"],
            ["Metric", "Coal Consumption"],
            ["Purpose", "Show proportional distribution of coal usage"],
        ],
        col_widths=[4, 12],
    )
    body(doc,
        "Design Rationale: A pie chart effectively communicates part-to-whole relationships. It "
        "highlights that a small number of countries (China, India, United States) account for the "
        "vast majority of global coal consumption, which has direct implications for energy security "
        "and military logistics."
    )

    # ══════════════════════════════════════════════════════════════════════
    # 9. RESULTS AND ANALYSIS
    # ══════════════════════════════════════════════════════════════════════
    heading(doc, "9. Results and Analysis", level=1)

    heading(doc, "9.1 Key Findings", level=2)

    findings = [
        ("Top 5 Military Powers: ",
         "The United States, Russia, China, India, and South Korea occupy the top five ranks, "
         "consistent with the GFP PowerIndex methodology that weighs both quantity and diversity of military assets."),
        ("Asia Dominates Military Manpower: ",
         "Asian countries collectively account for the largest share of total military manpower, "
         "driven by the massive populations of China, India, Pakistan, and Indonesia."),
        ("Aircraft Distribution is Highly Skewed: ",
         "The top 3 countries (USA, Russia, China) possess a disproportionately large share of the "
         "world's military aircraft. Most nations below rank 30 have comparatively modest air forces."),
        ("Naval Strength is Concentrated: ",
         "Submarine and destroyer counts are dominated by fewer than 10 nations, reflecting the high "
         "cost of naval hardware. The United States, Russia, and China lead in submarine counts."),
        ("Energy Consumption Correlates with Military Rank: ",
         "Nations with higher military ranks tend to be larger consumers of coal, oil, and natural gas "
         "— reflecting both the logistical needs of large militaries and the broader industrial base "
         "required to sustain them."),
        ("Regional Insights: ",
         "Europe has many mid-ranked nations with moderate but well-equipped forces. Africa features "
         "many lower-ranked countries with smaller active-duty forces. South America has a relatively "
         "uniform distribution of military strength."),
    ]
    for bold_part, normal_part in findings:
        bullet(doc, normal_part, bold_prefix=bold_part)

    heading(doc, "9.2 Dashboard Interactivity", level=2)
    body(doc, "The Power BI dashboard supports the following interactive features:")
    bullet(doc, "to filter by region or rank range.", bold_prefix="Slicers — ")
    bullet(doc, "clicking on a country in the map chart highlights its data in all other charts.", bold_prefix="Cross-filtering — ")
    bullet(doc, "showing detailed country-level statistics on hover.", bold_prefix="Tooltips — ")

    # ══════════════════════════════════════════════════════════════════════
    # 10. CONCLUSION
    # ══════════════════════════════════════════════════════════════════════
    heading(doc, "10. Conclusion", level=1)
    body(doc,
        "This project demonstrates the power of data visualization in transforming complex, "
        "multi-dimensional defence data into accessible, interactive insights. By combining Excel's "
        "pivot table capabilities with Power BI's interactive charting engine, we produced a dashboard "
        "that enables users to:"
    )
    bullet(doc, "Instantly compare military strength across 145 countries.")
    bullet(doc, "Drill down into specific regions or asset categories.")
    bullet(doc, "Understand the correlation between population, resources, and military capability.")

    body(doc,
        "The workflow — from data cleaning and pivot table creation to dashboard design — follows "
        "industry-standard practices used in business intelligence and government analytics."
    )

    # Note about What-If
    note = doc.add_paragraph()
    run = note.add_run("Note: ")
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)
    run2 = note.add_run(
        "The \"What-If Analysis\" feature was not included in this project as per faculty guidelines."
    )
    run2.font.size = Pt(11)
    run2.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)

    # ══════════════════════════════════════════════════════════════════════
    # 11. FUTURE SCOPE
    # ══════════════════════════════════════════════════════════════════════
    heading(doc, "11. Future Scope", level=1)
    future_items = [
        ("Time-Series Analysis: ", "Incorporating historical GFP data (2020–2025) to visualize trends in military power over time."),
        ("Predictive Modeling: ", "Using machine learning techniques to forecast future military rankings based on economic and demographic indicators."),
        ("Additional Data Dimensions: ", "Including defense budget, cyber capabilities, and space assets for a more holistic analysis."),
        ("Web-Based Dashboard: ", "Publishing the Power BI dashboard to Power BI Service for real-time collaboration and sharing."),
        ("Comparative Index Creation: ", "Developing a custom composite index that weights factors differently for regional vs. global analysis."),
    ]
    for bold_part, normal_part in future_items:
        p = doc.add_paragraph(style="List Number")
        run = p.add_run(bold_part)
        run.bold = True
        p.add_run(normal_part)

    # ══════════════════════════════════════════════════════════════════════
    # 12. REFERENCES
    # ══════════════════════════════════════════════════════════════════════
    heading(doc, "12. References", level=1)
    references = [
        "Global Firepower. (2025). 2025 Military Strength Ranking. Retrieved from https://www.globalfirepower.com",
        "Tufte, E. R. (1983). The Visual Display of Quantitative Information. Graphics Press.",
        "Few, S. (2012). Show Me the Numbers: Designing Tables and Graphs to Enlighten. Analytics Press.",
        "Microsoft. (2025). Power BI Documentation. Retrieved from https://learn.microsoft.com/en-us/power-bi/",
        "Microsoft. (2025). Excel Pivot Tables — Overview. Retrieved from https://support.microsoft.com/en-us/excel",
        "Kirk, A. (2019). Data Visualisation: A Handbook for Data Driven Design. SAGE Publications.",
        "Kaggle. (2025). Military Strength Datasets. Retrieved from https://www.kaggle.com",
    ]
    for i, ref in enumerate(references, 1):
        p = doc.add_paragraph(f"[{i}]  {ref}")
        p.paragraph_format.space_after = Pt(3)

    # ── Save ─────────────────────────────────────────────────────────────
    doc.save(OUTPUT_PATH)
    print(f"\n✅  Report saved to: {OUTPUT_PATH}")
    print(f"    File size: {os.path.getsize(OUTPUT_PATH):,} bytes")


if __name__ == "__main__":
    main()
