from docx import Document
from lxml import etree
import zipfile
import shutil
import os
import tempfile
from datetime import datetime
import uuid

WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W = f'{{{WORD_NS}}}'

def extract_tracked_changes(docx_path):
    """
    Extract tracked changes with text and type (insert/delete) from English DOCX.
    Returns list of dicts with 'type' and 'text'.
    """
    changes = []

    with zipfile.ZipFile(docx_path) as docx_zip:
        xml_content = docx_zip.read('word/document.xml')
    tree = etree.fromstring(xml_content)

    for ins in tree.findall('.//w:ins', namespaces={'w': WORD_NS}):
        text = ''.join(ins.xpath('.//w:t/text()', namespaces={'w': WORD_NS})).strip()
        if text:
            changes.append({'type': 'insert', 'text': text})

    for deletion in tree.findall('.//w:del', namespaces={'w': WORD_NS}):
        text = ''.join(deletion.xpath('.//w:t/text()', namespaces={'w': WORD_NS})).strip()
        if text:
            changes.append({'type': 'delete', 'text': text})

    return changes

def _create_run(text):
    """
    Create a <w:r> element with <w:t> text child
    """
    r = etree.Element(W + 'r')
    t = etree.SubElement(r, W + 't')
    # Preserve spaces if any
    if text.strip() != text:
        t.set(f'{W}space', 'preserve')
    t.text = text
    return r

def _wrap_tracked_change(change_type, run_element, author='AutoScript', date=None):
    """
    Wrap a run element with <w:ins> or <w:del> with metadata for tracked change
    """
    if date is None:
        date = datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ')

    elem = etree.Element(W + ('ins' if change_type == 'insert' else 'del'))
    elem.set(W + 'author', author)
    elem.set(W + 'date', date)
    elem.append(run_element)
    return elem

def apply_tracked_changes_to_chinese_doc(ch_docx_path, changes_map, output_path):
    """
    Applies tracked changes (insert/delete) to Chinese DOCX.
    Each change in changes_map:
     - 'type': 'insert' or 'delete'
     - 'chinese_text': text to mark
    """

    tmp_dir = tempfile.mkdtemp()
    with zipfile.ZipFile(ch_docx_path, 'r') as zip_ref:
        zip_ref.extractall(tmp_dir)

    doc_xml_path = os.path.join(tmp_dir, 'word/document.xml')
    parser = etree.XMLParser(ns_clean=True, recover=True)
    tree = etree.parse(doc_xml_path, parser)
    root = tree.getroot()

    paragraphs = root.findall('.//w:p', namespaces={'w': WORD_NS})

    for change in changes_map:
        c_text = change['chinese_text']
        c_type = change['type']

        applied = False

        for p in paragraphs:
            # Rebuild paragraph full text including runs and keep track of indices
            runs = p.findall('.//w:r', namespaces={'w': WORD_NS})

            # Aggregate run texts
            run_texts = [ ''.join(r.xpath('.//w:t/text()', namespaces={'w': WORD_NS})) or '' for r in runs ]
            para_text = ''.join(run_texts)

            # Find start index of chinese_text in para_text (first occurrence only)
            idx = para_text.find(c_text)
            if idx == -1:
                continue

            # We need to replace the runs covering c_text substring
            # Find which runs cover the substring [idx : idx + len(c_text)]
            current_pos = 0
            start_run_idx = None
            end_run_idx = None
            start_offset = None
            end_offset = None

            for i, rt in enumerate(run_texts):
                run_len = len(rt)
                run_start = current_pos
                run_end = current_pos + run_len

                if start_run_idx is None and run_start <= idx < run_end:
                    start_run_idx = i
                    start_offset = idx - run_start
                if start_run_idx is not None and run_start < idx + len(c_text) <= run_end:
                    end_run_idx = i
                    end_offset = idx + len(c_text) - run_start
                    break

                current_pos += run_len

            # If end_run_idx not found, assume ends at last run
            if end_run_idx is None:
                end_run_idx = len(run_texts) - 1
                end_offset = len(run_texts[end_run_idx])

            # Rebuild runs:
            # Runs before start_run_idx remain unchanged
            # Runs after end_run_idx remain unchanged
            # Runs between start_run_idx and end_run_idx replaced by wrapped tracked change run

            # Extract parts before and after in start and end runs
            # Create new runs for parts before, changed text, parts after

            # Before text in start run
            before_text = run_texts[start_run_idx][:start_offset]
            # Changed text spanning runs
            changed_text_parts = []
            for run_i in range(start_run_idx, end_run_idx + 1):
                rt = run_texts[run_i]
                if run_i == start_run_idx:
                    part = rt[start_offset:]
                elif run_i == end_run_idx:
                    part = rt[:end_offset]
                else:
                    part = rt
                changed_text_parts.append(part)
            changed_text = ''.join(changed_text_parts)
            # After text in end run
            after_text = run_texts[end_run_idx][end_offset:]

            parent = runs[start_run_idx].getparent()

            # Remove old runs
            for run_i in range(start_run_idx, end_run_idx + 1):
                parent.remove(runs[run_i])

            # Insert before text run if exists
            insert_pos = start_run_idx
            if before_text:
                r_before = _create_run(before_text)
                parent.insert(insert_pos, r_before)
                insert_pos += 1

            # Insert tracked change run
            r_changed = _create_run(changed_text)
            wrapped = _wrap_tracked_change(c_type, r_changed)
            parent.insert(insert_pos, wrapped)
            insert_pos += 1

            # Insert after text run if exists
            if after_text:
                r_after = _create_run(after_text)
                parent.insert(insert_pos, r_after)

            applied = True
            break

        if not applied:
            print(f"Warning: Could not find Chinese text '{c_text}' to apply {c_type} change.")

    # Write back XML
    tree.write(doc_xml_path, encoding='utf-8', xml_declaration=True)

    # Rezip the docx
    shutil.make_archive(output_path.replace('.docx',''), 'zip', tmp_dir)
    shutil.move(output_path.replace('.docx','') + '.zip', output_path)
    shutil.rmtree(tmp_dir)
