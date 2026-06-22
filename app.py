# ==============================================================================
# STEP 1: INITIALIZATION & ENVIRONMENT CONFIGURATION
# ==============================================================================
import os
import sys
import re
import pandas as pd

try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
except ImportError:
    print("Installing python-docx dependency layer...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "python-docx"])
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

# ==============================================================================
# STEP 2: OPENXML STYLING COMPONENT ENGINE
# ==============================================================================
def apply_border(cell, color="000000", size="4", thick=False):
    """
    Directly mutates the OpenXML cell properties schema to draw borders.
    """
    tcPr = cell._element.get_or_add_tcPr()
    border_val = "thick" if thick else "single"
    hex_color = color.lstrip('#')
    tcBorders_xml = (
        f'<w:tcBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="{border_val}" w:sz="{size}" w:color="{hex_color}"/>\n'
        f'  <w:left w:val="{border_val}" w:sz="{size}" w:color="{hex_color}"/>\n'
        f'  <w:bottom w:val="{border_val}" w:sz="{size}" w:color="{hex_color}"/>\n'
        f'  <w:right w:val="{border_val}" w:sz="{size}" w:color="{hex_color}"/>\n'
        f'</w:tcBorders>'
    )
    tcPr.append(parse_xml(tcBorders_xml))

# ==============================================================================
# STEP 3: WORKSPACE SCANNER & BINDING ENGINE
# ==============================================================================
def identify_data_sources():
    print("Scanning directory workspace via strict filename anchor alignment...")
    all_files = os.listdir(".")
    
    wab_path = None
    sra_template_path = None
    inventory_path = None
    vulnerability_path = None
    questions_path = None
    code_review_paths = []

    for f in all_files:
        if f.endswith(".docx") and not f.startswith("~$") and "generated" not in f.lower():
            if "v2" in f.lower() or f == "EDB (SEMIS) SRA report v2.docx":
                sra_template_path = f
            elif "wab" in f.lower() or "assignment" in f.lower():
                wab_path = f

    if not sra_template_path:
        for f in all_files:
            if f.endswith(".docx") and not f.startswith("~$") and "template" in f.lower():
                sra_template_path = f
                
    for f in all_files:
        if f.endswith(".xlsx") and not f.startswith("~$") and "generated" not in f.lower():
            f_lower = f.lower()
            if "questions" in f_lower or "list of questions" in f_lower or "sraa" in f_lower:
                questions_path = f
            elif "asset" in f_lower or "inventory" in f_lower or "valuation" in f_lower:
                inventory_path = f
            elif "vulnerability" in f_lower or "pentest" in f_lower:
                vulnerability_path = f
            elif "code review" in f_lower:
                code_review_paths.append(f)

    return {
        "wab": wab_path, "sra_template": sra_template_path, "inventory": inventory_path,
        "vulnerability": vulnerability_path, "questions": questions_path, "code_reviews": code_review_paths
    }

# ==============================================================================
# STEP 4: TELEMETRY EXTRACTION ENGINE
# ==============================================================================
def process_dynamic_telemetry(paths, company_name, company_abbr):
    print("\nRunning live analytics over classified source components...")
    
    def read_and_align_sheet(file_path, landmarks, sheet_idx=0):
        try:
            df_raw = pd.read_excel(file_path, header=None, sheet_name=sheet_idx)
            header_row_idx = None
            for idx, row in df_raw.iterrows():
                row_str = " ".join(row.fillna("").astype(str))
                if any(x in row_str for x in landmarks):
                    header_row_idx = idx
                    break
            if header_row_idx is None: header_row_idx = 0 
            df_aligned = pd.read_excel(file_path, skiprows=header_row_idx, sheet_name=sheet_idx)
            df_aligned.columns = [str(c).strip() for c in df_aligned.columns]
            return df_aligned
        except Exception:
            return pd.DataFrame()

    testee_name, testee_abbr = "Education Bureau", "EDB"
    sys_abbr, sys_name = "SEMIS", "Special Education Management Information System"
    scope_paragraphs, objective_paragraphs, security_requirements = [], [], []
    system_description_text = ""
    asset_rows = []

    W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    if paths["wab"]:
        try:
            wab_doc = Document(paths["wab"])
            for p in wab_doc.paragraphs:
                text_clean = p.text.strip()
                if "Bureau/Department :" in text_clean:
                    testee_name = text_clean.split(":")[-1].strip()
                    break
            abbr_match = re.findall(r'\(([A-Z]{2,6})\)', testee_name)
            if abbr_match: 
                testee_abbr = abbr_match[0]
                testee_name = testee_name.split('(')[0].strip()

            capture_scope, capture_obj, capture_desc, capture_req = False, False, False, False
            for p in wab_doc.paragraphs:
                txt = "".join(run.text for run in p.runs).strip()
                normalized_txt = txt.replace('\xa0', ' ').strip()
                
                if "SCOPE OF THE SERVICES" in normalized_txt:
                    capture_scope = True
                    continue
                elif "BACKGROUND" in normalized_txt:
                    capture_scope = False

                if "PROJECT OBJECTIVES" in normalized_txt:
                    capture_obj = True
                    continue
                elif "PROJECT REQUIREMENTS" in normalized_txt or "USER REQUIREMENTS" in normalized_txt:
                    capture_obj = False

                if "4.1.1" in normalized_txt and "Current Environment" in normalized_txt:
                    capture_desc = True
                    continue
                elif "4.1.2" in normalized_txt or "Project Management" in normalized_txt:
                    capture_desc = False

                if "GOVERNMENT STANDARDS, METHODOLOGIES AND QUALITY REQUIREMENTS" in normalized_txt:
                    capture_req = True
                    continue
                elif "PROJECT DELIVERABLES, MILESTONES & IMPLEMENTATION SCHEDULE" in normalized_txt:
                    capture_req = False

                if capture_scope and len(normalized_txt) > 0:
                    if normalized_txt.startswith("To"):
                        cleaned_txt = normalized_txt.replace("the systems and their classification specified in Section 2 (“Systems”)", "the selected systems (“Systems”)")
                        cleaned_txt = re.sub(r'^[\t\s]*\([a-z]\)[\t\s]*', '', cleaned_txt, flags=re.IGNORECASE).strip()
                        if cleaned_txt and cleaned_txt not in scope_paragraphs:
                            scope_paragraphs.append(cleaned_txt)

                if capture_obj and len(normalized_txt) > 0:
                    cleaned_txt = re.sub(r'^3\.\d+\s*', '', normalized_txt).strip()
                    cleaned_txt = re.sub(r'^-\s*', '', cleaned_txt).strip()
                    if cleaned_txt and cleaned_txt not in objective_paragraphs:
                        objective_paragraphs.append(cleaned_txt)

                if capture_desc and len(normalized_txt) > 2:
                    if not (normalized_txt.startswith("4.1") or normalized_txt.startswith("4.1.1")):
                        if "USER REQUIREMENTS" not in normalized_txt and "Current Environment Description" not in normalized_txt:
                            if not system_description_text:
                                system_description_text = normalized_txt
                            else:
                                system_description_text += "\n" + normalized_txt

                if capture_req and len(normalized_txt) > 0:
                    if normalized_txt.startswith("Where necessary"):
                        capture_req = False
                        continue
                    pPr = p._p.get_or_add_pPr()
                    numPr = pPr.find(f'{{{W_NAMESPACE}}}numPr')
                    if numPr is not None:
                        ilvl = numPr.find(f'{{{W_NAMESPACE}}}ilvl')
                        if ilvl is not None and ilvl.get(f'{{{W_NAMESPACE}}}val') == "0":
                            if normalized_txt not in security_requirements:
                                security_requirements.append(normalized_txt)
        except Exception: pass

    if not system_description_text:
        system_description_text = "System to support the processing of information of special education support services, and managing the workflow of the application, assessment and approval of special education support services and grants."

    if paths["inventory"]:
        try:
            inv_df = read_and_align_sheet(paths["inventory"], ["Item", "Role Description", "Hostname"])
            if not inv_df.empty:
                col_list = list(inv_df.columns)
                item_col = next((c for c in col_list if "Item" in c or "No" in c), col_list[1])
                desc_col = next((c for c in col_list if "Description" in c or "Role" in c), col_list[2])
                host_col = next((c for c in col_list if "Hostname" in c or "ID" in c or "Tag" in c), col_list[4])
                ip_col = next((c for c in col_list if "IP" in c or "URL" in c), col_list[5])
                os_col = next((c for c in col_list if "OS" in c or "Version" in c), col_list[6])

                for _, row in inv_df.dropna(subset=[item_col]).iterrows():
                    itm = str(row[item_col]).strip()
                    if itm.replace('.0','').isdigit() or len(itm) <= 2:
                        asset_rows.append([
                            itm.replace('.0',''), str(row[desc_col]).strip(), str(row[host_col]).strip(), 
                            str(row[ip_col]).strip().replace("\n", " "), str(row[os_col]).strip()
                        ])
        except Exception: pass

    op_sec_counts = {"High": 0, "Medium": 0, "Low": 0, "AOI": 0}
    vulnerability_rows, penetration_rows = [], []

    if paths["vulnerability"]:
        try:
            xl = pd.ExcelFile(paths["vulnerability"])
            s9_df = read_and_align_sheet(paths["vulnerability"], ["Observe", "Findings#", "Protocol"], sheet_idx=xl.sheet_names[0])
            if not s9_df.empty:
                col_list = list(s9_df.columns)
                id_col = next((c for c in col_list if "Findings#" in c or "Observe" in c), col_list[1])
                asset_col = next((c for c in col_list if "Asset" in c or "System/" in c), col_list[2])
                vun_col = next((c for c in col_list if "Vulnerability" in c or "Observation" in c or "Risk Name" in c), col_list[7])
                threat_col = next((c for c in col_list if "Threat" in c or "Description" in c or "Details" in c), col_list[8])
                rate_col = next((c for c in col_list if "Rating" in c or "Level" in c or "Risk Level" in c), col_list[10])
                action_col = next((c for c in col_list if "Action" in c or "plan" in c), col_list[9])
                domain_col = next((c for c in col_list if "Domain" in c), None)

                if domain_col and rate_col:
                    filter_mask = s9_df[domain_col].astype(str).str.contains('Opreation|Operation|Infrastructure', na=False, case=False)
                    for rating in op_sec_counts.keys():
                        matched_count = len(s9_df[filter_mask & (s9_df[rate_col].astype(str).str.strip().str.lower() == rating.lower())])
                        op_sec_counts[rating] = int(matched_count)

                for _, row in s9_df.dropna(subset=[id_col]).iterrows():
                    f_id = str(row[id_col]).strip()
                    if f_id == 'nan' or not f_id: continue
                    raw_rtg = str(row[rate_col]).strip() if rate_col in s9_df.columns else "Low"
                    norm_rtg = "AOI" if "aoi" in raw_rtg.lower() or "interest" in raw_rtg.lower() else raw_rtg.capitalize()
                    
                    raw_ips = str(row[asset_col]).strip().replace(" ", "").split(",")
                    formatted_ips = "\n".join(raw_ips)

                    record = [
                        sys_abbr, formatted_ips, str(row[vun_col]).strip(),
                        str(row[threat_col]).strip(), f"{norm_rtg} [{norm_rtg},{norm_rtg}]", str(row[action_col]).strip()
                    ]
                    if f_id.upper().startswith("V"): vulnerability_rows.append(record)
                    elif f_id.upper().startswith("A"): penetration_rows.append(record)
        except Exception: pass

    cr_counts = {"High": 0, "Medium": 0, "Low": 0, "AOI": 0}
    for f in paths["code_reviews"]:
        try:
            xl_file = pd.ExcelFile(f)
            for sheet in xl_file.sheet_names:
                df = read_and_align_sheet(f, ["Observe", "Findings#"], sheet_idx=sheet)
                if not df.empty:
                    rc = [c for c in df.columns if "Rating" in c or "Level" in c]
                    if rc:
                        rc_col = rc[0]
                        for rating in cr_counts.keys():
                            matched_len = len(df[df[rc_col].astype(str).str.strip().str.lower() == rating.lower()])
                            cr_counts[rating] += int(matched_len)
        except Exception: pass

    total_high = op_sec_counts["High"] + cr_counts["High"]
    total_med = op_sec_counts["Medium"] + cr_counts["Medium"]
    total_low = op_sec_counts["Low"] + cr_counts["Low"]
    total_aoi = op_sec_counts["AOI"] + cr_counts["AOI"]

    return {
        "testee_name": testee_name, "testee_abbr": testee_abbr, "sys_name": sys_name, "sys_abbr": sys_abbr,
        "company_name": company_name, "company_abbr": company_abbr, "op_sec": op_sec_counts, "cr": cr_counts,
        "totals": {"High": total_high, "Medium": total_med, "Low": total_low, "AOI": total_aoi},
        "grand_total": total_high + total_med + total_low + total_aoi, "vulnerability_rows": vulnerability_rows,
        "penetration_rows": penetration_rows, "scope_paragraphs": scope_paragraphs, "objective_paragraphs": objective_paragraphs,
        "system_description": system_description_text, "asset_rows": asset_rows, "security_requirements": security_requirements
    }

# ==============================================================================
# STEP 5: COMPILATION & STRUCTURAL CLEANING PIPELINE
# ==============================================================================
def output_recreation_report(paths, data, filename="Generated_SRA_Report_Final.docx"):
    print(f"\nAssembling document parameters -> destination filename: '{filename}'")
    doc = Document(paths["sra_template"])
    
    risk_tier_string = f"{data['totals']['Medium']} Medium, {data['totals']['Low']} low risk and {data['totals']['AOI']} AOI items"

    replacements = {
        "information system name (information system abbreviation)": f"{data['sys_name']} ({data['sys_abbr']})",
        "(INFORMATION SYSTEM ABBREVIATION)": data["sys_abbr"],
        "information system name": data["sys_name"],
        "SEMISSpW23": data["sys_abbr"],
        "amount of total items": f"{data['grand_total']} items",
        "grand total items": risk_tier_string,
        "company name (company abbreviation)": f"{data['company_name']} ({data['company_abbr']})",
        "company name": data["company_name"],
        "company abbreviation": data["company_abbr"],
        "COMPANY ABBREVIATION": data["company_abbr"],
        "testee name (testee abbreviation)": f"{data['testee_name']} ({data['testee_abbr']})",
        "testee abbreviation": data["testee_abbr"]
    }

    def update_cell_text_preserving_runs(cell, new_text, font_name="Times New Roman", size_pt=10, align_center=False):
        if len(cell.paragraphs) == 0:
            cell.add_paragraph()
        p = cell.paragraphs[0]
        if align_center:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if len(p.runs) == 0:
            p.add_run(new_text)
        else:
            p.runs[0].text = new_text
            for r in p.runs[1:]: r.text = ""
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.name = font_name
                run.font.size = Pt(size_pt)
                run.font.color.rgb = RGBColor(0, 0, 0)

    def format_cell_text(cell, font_name="Times New Roman", size_pt=10, align_center=False, bold=False, color_rgb=None):
        for paragraph in cell.paragraphs:
            if align_center:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.font.name = font_name
                run.font.size = Pt(size_pt)
                run.bold = bold
                if color_rgb is not None: run.font.color.rgb = color_rgb

    def remove_asterisk_pointers(text):
        """Removes asterisk-enclosed instructions and structural guides."""
        text = re.sub(r'\*.*?\*', '', text)
        return text.strip()

    def replace_text_in_paragraph(p, lookup_dict):
        full_text = "".join(run.text for run in p.runs)
        if "4." in full_text and any(term in full_text.lower() for term in ["methodology", "asset", "valuation", "threat", "mapping"]):
            return
        for k, v in lookup_dict.items():
            if k in full_text:
                full_text = full_text.replace(k, v)
        full_text = remove_asterisk_pointers(full_text)
        if len(p.runs) > 0:
            p.runs[0].text = full_text
            for r in p.runs[1:]: r.text = ""
        else:
            p.text = full_text

    def insert_paragraph_after(paragraph, text, style=None):
        new_p = doc.add_paragraph(style=style)
        new_p.text = text
        p_element = paragraph._p
        p_element.addnext(new_p._p)
        return new_p

    # 1. Paragraph Modifications Layer
    for p_idx, p in enumerate(list(doc.paragraphs)):
        full_p_text = "".join(run.text for run in p.runs)
        normalized_p_text = full_p_text.replace('\n', ' ').strip()
        
        if 145 <= p_idx <= 165:
            continue

        if "copy from WAB, 4.1.1 Current Environment Description" in normalized_p_text or "copy from WAB, 9.1 System Description" in normalized_p_text:
            p.text = remove_asterisk_pointers(data["system_description"])
            
        elif "insert all under brief of work assignment under category b, 1. Scope of the services, 1.1" in normalized_p_text:
            if len(data["scope_paragraphs"]) > 0:
                p.text = remove_asterisk_pointers(data["scope_paragraphs"][0])
                p.style = 'List Bullet'
                current_anchor = p
                for msg in data["scope_paragraphs"][1:]:
                    current_anchor = insert_paragraph_after(current_anchor, remove_asterisk_pointers(msg), style='List Bullet')
            else:
                p.text = ""
                
        elif "copy from work assignment under category b, all under 3. Project objectives" in normalized_p_text:
            if len(data["objective_paragraphs"]) > 0:
                p.text = remove_asterisk_pointers(data["objective_paragraphs"][0])
                p.style = 'List Bullet'
                current_anchor = p
                for msg in data["objective_paragraphs"][1:]:
                    current_anchor = insert_paragraph_after(current_anchor, remove_asterisk_pointers(msg), style='List Bullet')
            else:
                p.text = ""
                
        elif "copy from WAB section 4.4" in normalized_p_text:
            p.text = f"{data['company_abbr']} referred to the following documents in conducting the security risk assessment:"
            current_anchor = p
            for msg in data["security_requirements"]:
                current_anchor = insert_paragraph_after(current_anchor, remove_asterisk_pointers(msg), style='List Bullet')
                
        elif "copy from follow up plan" in normalized_p_text.lower() or "copy from this document table" in normalized_p_text.lower():
            p.text = ""
        else:
            replace_text_in_paragraph(p, replacements)

    # 2. DOCUMENT CONTROL TABLES UPDATES
    try:
        t_signoff = doc.tables[0]
        t_signoff.cell(0, 0).text = "Version 0.1"
        t_signoff.cell(1, 0).text = "Role"; t_signoff.cell(1, 1).text = "Name"; t_signoff.cell(1, 2).text = "Action"; t_signoff.cell(1, 3).text = "Date"
        t_signoff.cell(2, 0).text = "Created by"; t_signoff.cell(2, 1).text = f"{data['company_name']}"; t_signoff.cell(2, 2).text = "Document Revision"; t_signoff.cell(2, 3).text = "22-05-2026"
        t_signoff.cell(3, 0).text = "Approved by"; t_signoff.cell(3, 1).text = ""; t_signoff.cell(3, 2).text = "Document Approval"; t_signoff.cell(3, 3).text = ""                 
    except Exception: pass

    try:
        t_revision = doc.tables[1]
        t_revision.cell(0, 0).text = "Version"; t_revision.cell(0, 1).text = "Date"; t_revision.cell(0, 2).text = "Author"; t_revision.cell(0, 3).text = "Summary of Changes"
        t_revision.cell(1, 0).text = "0.1"; t_revision.cell(1, 1).text = "22-05-2026"; t_revision.cell(1, 2).text = f"{data['company_name']}"; t_revision.cell(1, 3).text = "Document Creation"
    except Exception: pass

    vulnerability_table_count = 0

    # 3. EXPLICIT INDEX SEGMENTATION FOR CLASSIFIED TABLES
    for table_idx, table in enumerate(doc.tables):
        if len(table.rows) == 0 or len(table.rows[0].cells) == 0: continue
        
        if table_idx in [3, 4, 5, 6, 7]:
            continue

        # Heavy Black Border Profile for Control Blocks
        elif table_idx in [0, 1]:
            for row in table.rows:
                for cell in row.cells:
                    apply_border(cell, color="000000", size="12", thick=True)

        # Standard Office Blue Profile (#4472c4) for Risk Matrices
        elif table_idx in [2, 10]:
            for row in table.rows:
                lbl = row.cells[0].text.lower()
                if "operations security" in lbl or "1.1 risks assessed" in lbl:
                    row.cells[0].text = "Operations Security"
                    update_cell_text_preserving_runs(row.cells[1], str(data["op_sec"]["High"]), align_center=True)
                    update_cell_text_preserving_runs(row.cells[2], str(data["op_sec"]["Medium"]), align_center=True)
                    update_cell_text_preserving_runs(row.cells[3], str(data["op_sec"]["Low"]), align_center=True)
                    update_cell_text_preserving_runs(row.cells[4], str(data["op_sec"]["AOI"]), align_center=True)
                elif "system acquisition" in lbl:
                    row.cells[0].text = "System Acquisition, Development and Maintenance"
                    update_cell_text_preserving_runs(row.cells[1], str(data["cr"]["High"]), align_center=True)
                    update_cell_text_preserving_runs(row.cells[2], str(data["cr"]["Medium"]), align_center=True)
                    update_cell_text_preserving_runs(row.cells[3], str(data["cr"]["Low"]), align_center=True)
                    update_cell_text_preserving_runs(row.cells[4], str(data["cr"]["AOI"]), align_center=True)
                elif "grand total" in lbl:
                    row.cells[0].text = "Grand Total"
                    update_cell_text_preserving_runs(row.cells[1], str(data["totals"]["High"]), align_center=True)
                    update_cell_text_preserving_runs(row.cells[2], str(data["totals"]["Medium"]), align_center=True)
                    update_cell_text_preserving_runs(row.cells[3], str(data["totals"]["Low"]), align_center=True)
                    update_cell_text_preserving_runs(row.cells[4], str(data["totals"]["AOI"]), align_center=True)
            
            for r_idx, row in enumerate(table.rows):
                for c_idx, cell in enumerate(row.cells):
                    apply_border(cell, color="4472c4", size="4", thick=False)
                    if r_idx > 0:
                        if c_idx > 0:
                            update_cell_text_preserving_runs(cell, cell.text.strip(), font_name="Times New Roman", size_pt=10, align_center=True)
                        else:
                            update_cell_text_preserving_runs(cell, cell.text.strip(), font_name="Times New Roman", size_pt=10, align_center=False)

        # Findings Projection Matrices
        elif table_idx in [11, 12]:
            for row in list(table.rows):
                cell_text_clean = row.cells[0].text.lower()
                if "copy from follow up plan" in cell_text_clean or "v numbered rows" in cell_text_clean or "a numbered rows" in cell_text_clean:
                    table._element.remove(row._element)

            if vulnerability_table_count == 0:
                target_dataset = data["vulnerability_rows"]
                vulnerability_table_count += 1
            else:
                target_dataset = data["penetration_rows"]
            
            for row_data in target_dataset:
                r_cells = table.add_row().cells
                for i, val in enumerate(row_data): r_cells[i].text = val

            for r_idx, row in enumerate(table.rows):
                for c_idx, cell in enumerate(row.cells):
                    apply_border(cell, color="000000", size="4", thick=False)
                    if r_idx > 0:
                        center_flag = (c_idx in [3, 4])
                        format_cell_text(cell, font_name="Times New Roman", size_pt=9, align_center=center_flag)
                        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        else:
            if "role description" in table.rows[0].cells[1].text.lower():
                table.cell(0, 0).text = "Item"; table.cell(0, 1).text = "Asset Description"; table.cell(0, 2).text = "Hostname / ID"; table.cell(0, 3).text = "IP Address / URL"; table.cell(0, 4).text = "OS Version"
                for asset_row in data["asset_rows"]:
                    r_cells = table.add_row().cells
                    for i, val in enumerate(asset_row): r_cells[i].text = val
            elif "role" in table.rows[0].cells[0].text.lower() and "name" in table.rows[0].cells[1].text.lower() and len(table.rows) <= 2:
                team_rows = [["Project Manager", f"Representative, {data['company_name']}"], ["Senior Security Auditor", f"Lead Consultant, {data['company_name']}"], ["Technical Auditor Expert", f"Auditor, {data['company_name']}"]]
                for t_row in team_rows:
                    r_cells = table.add_row().cells
                    r_cells[0].text, r_cells[1].text = t_row[0], t_row[1]

            for row in table.rows:
                for cell in row.cells:
                    apply_border(cell, color="000000", size="4", thick=False)
                    cell.text = remove_asterisk_pointers(cell.text)
                    for p in cell.paragraphs:
                        replace_text_in_paragraph(p, replacements)

    doc.save(filename)
    print(f"Compilation finished cleanly -> saved to '{filename}'")

# ==============================================================================
# STEP 6: INTERACTIVE CONFIGURATION & EXECUTION RUNNER
# ==============================================================================
if __name__ == "__main__":
    print("="*80)
    print("        SRA REPORT COMPILE WORKSPACE ENVIRONMENT CONTROLLER             ")
    print("="*80)
    
    # Input sections for custom deployment parameters
    input_company = input("Enter Company Name [Default: SunnyVision Limited]: ").strip()
    input_abbr = input("Enter Company Abbreviation [Default: SV]: ").strip()
    
    company_name = input_company if input_company else "SunnyVision Limited"
    company_abbr = input_abbr if input_abbr else "SV"
    
    nodes = identify_data_sources()
    telemetry = process_dynamic_telemetry(nodes, company_name, company_abbr)
    output_recreation_report(nodes, telemetry)
