import re
lines = open('src/tool.py', encoding='utf-8').read().split('\n')

starts = {}
for i, l in enumerate(lines, 1):
    m = re.match(r'^    (?:async )?def ([a-zA-Z_]\w*)', l)
    if m:
        starts[m.group(1)] = i

names = list(starts)
ends = {}
for idx, n in enumerate(names):
    ends[n] = (starts[names[idx+1]] - 1) if idx+1 < len(names) else len(lines)

core = ["__init__","_resolve_file","_save_and_link","_progress","version_diff","webhook_trigger","import_from_api","import_from_url","version_file","tool_stats","cleanup_files","save_template","use_template","list_templates","schedule_cleanup","smart_template","audit_log","retention_policy","scheduled_report","auto_backup","preview_file","file_search","bulk_folder_ops","convert_format","convert_data","compliance_check","ai_analyze","ai_summarize","export_to_markdown","export_to_html"]
read = ["read_file","_read_xlsx","_read_xls","_read_docx","_read_pptx","_parse_csv_rows"]
write = ["create_file","generate_document","generate_slides","generate_spreadsheet","add_content","add_chart","add_watermark","add_alt_text","add_speaker_notes","add_qr_code","add_data_validation","add_named_range","add_slide_transitions","add_pivot_table","sql_to_spreadsheet","fill_pdf_form","document_assembly","conditional_format","smart_fill"]
edit = ["replace_text","update_cells","modify_rows","protect_file","merge_sheets","batch_process","merge_pdfs","split_pdf","mail_merge","edit_metadata","check_accessibility","compare_documents","grammar_check","translate_document","classify_document","upload_to_drive","ocr_extract","translate_errors","document_stats"]
comments = ["add_comment"]

moved = set(read+write+edit+comments)
unassigned = [n for n in names if n not in core and n not in moved]
print("UNASSIGNED METHODS:", unassigned)

targets = ['_b64_mod','copy','traceback','platform','_office_plugins','register_office_plugin','_call_office_plugins','_encode_filename','_decode_filename','_read_odf','_cell_value','_xls_to_xlsx','_DB_PATH','_UPLOAD_DIR','_EXPORT_DIR','_P_NS','_P14_NS','_R_NS','_PKG_REL_NS','_CT_NS','_CM_REL_TYPE','_CT_MODERN','_CT_AUTHORS','_get_owui_data_dir','_resolve_file_path','_read_file_bytes','_detect_type','_format_text','_parse_inline_md','_add_callout_box','_add_professional_table','_render_content_slide','_data_dir']
for target in targets:
    use_in = []
    for n in moved:
        body = '\n'.join(lines[starts[n]-1:ends[n]])
        if re.search(r'\b' + re.escape(target) + r'\b', body):
            use_in.append(n)
    if use_in:
        print(f"{target}: used in {use_in}")

# check for comment-only class-level lines immediately above each moved def
print("\n--- comment lines above moved defs ---")
for n in moved:
    s = starts[n]
    # walk up from s-2 (0-indexed s-2 => line s-1)
    collected = []
    i = s - 2  # 0-indexed line before def
    while i >= 0:
        l = lines[i]
        if l.strip() == '' or re.match(r'^    #', l):
            collected.append((i+1, l))
            i -= 1
        else:
            break
    if collected:
        print(f"--- above {n} (def line {s}):")
        for ln, txt in reversed(collected):
            print(f"  {ln}: {txt}")

# any non-def, non-comment, non-blank class-level lines between methods? (stray code)
print("\n--- stray class-level code between methods ---")
for idx, n in enumerate(names):
    s = starts[n]
    prev_end = (ends[names[idx-1]] + 1) if idx > 0 else 44  # Valves ends ~43
    for i in range(prev_end, s):
        l = lines[i-1]
        if l.strip() and not re.match(r'^    #', l) and not re.match(r'^    (?:async )?def ', l):
            print(f"  {i}: {l}")

print("\n--- module-level names in moved bodies not in any import list ---")
# crude: find identifiers that look like module-level constants/functions defined in constants/utils
const_utils_names = set()
cu = open('src/constants.py', encoding='utf-8').read() + open('src/utils.py', encoding='utf-8').read()
for m in re.finditer(r'^(?:def |class |[A-Za-z_]\w*\s*=)', cu, re.M):
    pass
