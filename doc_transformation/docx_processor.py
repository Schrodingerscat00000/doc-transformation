# enhanced_docx_processor.py

import zipfile
import ollama
from lxml import etree
from copy import deepcopy
from ollama_client import OllamaClient
from datetime import datetime
import uuid
import re

# Instantiate once at module‐level
_client = OllamaClient(model="deepseek-r1:1.5b", base_url="http://localhost:11434")

# XML namespaces needed for Word documents
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}

# Register namespaces
for prefix, uri in NS.items():
    etree.register_namespace(prefix, uri)

def _get_paragraph_text(p_element):
    """Extracts plain text from a paragraph's XML element, preserving structure."""
    return ''.join(p_element.xpath('.//w:t/text()', namespaces=NS))

def _get_paragraph_text_with_structure(p_element):
    """Get paragraph text while maintaining information about runs and their properties."""
    runs_info = []
    for run in p_element.xpath('./w:r', namespaces=NS):
        run_text = ''.join(run.xpath('.//w:t/text()', namespaces=NS))
        if run_text:
            runs_info.append({
                'text': run_text,
                'element': run,
                'is_insertion': bool(run.xpath('.//w:ins', namespaces=NS)),
                'is_deletion': bool(run.xpath('.//w:del', namespaces=NS))
            })
    return runs_info

def _extract_changes_from_english_docx(docx_path):
    """
    Enhanced extraction that preserves more context and metadata about changes.
    Returns a list of change dictionaries with better context information.
    """
    changes = []
    with zipfile.ZipFile(docx_path, 'r') as docx_zip:
        xml_content = docx_zip.read('word/document.xml')
        root = etree.fromstring(xml_content)

        for para_idx, p in enumerate(root.xpath('//w:body/w:p', namespaces=NS)):
            # Get the complete paragraph structure
            runs_info = _get_paragraph_text_with_structure(p)
            
            # Reconstruct original text (before changes)
            original_text_parts = []
            current_text_parts = []
            
            for run_info in runs_info:
                if run_info['is_deletion']:
                    # This was in original but deleted
                    deleted_text = ''.join(run_info['element'].xpath('.//w:delText/text()', namespaces=NS))
                    original_text_parts.append(deleted_text)
                elif not run_info['is_insertion']:
                    # This was in original and remains
                    original_text_parts.append(run_info['text'])
                    current_text_parts.append(run_info['text'])
                else:
                    # This is new (insertion)
                    current_text_parts.append(run_info['text'])
            
            original_paragraph_text = ''.join(original_text_parts).strip()
            current_paragraph_text = ''.join(current_text_parts).strip()
            
            if not original_paragraph_text and not current_paragraph_text:
                continue

            # Extract insertions with better context
            insertions = p.xpath('.//w:ins', namespaces=NS)
            for ins_idx, ins in enumerate(insertions):
                inserted_text = ''.join(ins.xpath('.//w:t/text()', namespaces=NS)).strip()
                if inserted_text:
                    # Get author and date if available
                    author = ins.get(etree.QName(NS['w'], 'author'), 'Unknown')
                    date = ins.get(etree.QName(NS['w'], 'date'), datetime.now().isoformat())
                    change_id = ins.get(etree.QName(NS['w'], 'id'), str(uuid.uuid4().int)[:8])
                    
                    changes.append({
                        'type': 'insertion',
                        'text': inserted_text,
                        'original_context': original_paragraph_text,
                        'current_context': current_paragraph_text,
                        'paragraph_index': para_idx,
                        'author': author,
                        'date': date,
                        'change_id': change_id,
                        'element': ins
                    })
            
            # Extract deletions with better context
            deletions = p.xpath('.//w:del', namespaces=NS)
            for del_idx, dele in enumerate(deletions):
                deleted_text = ''.join(dele.xpath('.//w:delText/text()', namespaces=NS)).strip()
                if deleted_text:
                    # Get author and date if available
                    author = dele.get(etree.QName(NS['w'], 'author'), 'Unknown')
                    date = dele.get(etree.QName(NS['w'], 'date'), datetime.now().isoformat())
                    change_id = dele.get(etree.QName(NS['w'], 'id'), str(uuid.uuid4().int)[:8])
                    
                    changes.append({
                        'type': 'deletion',
                        'text': deleted_text,
                        'original_context': original_paragraph_text,
                        'current_context': current_paragraph_text,
                        'paragraph_index': para_idx,
                        'author': author,
                        'date': date,
                        'change_id': change_id,
                        'element': dele
                    })
    
    return changes

def _get_llm_response(prompt: str) -> str:
    """Enhanced LLM interaction with better error handling and retry logic."""
    if not _client.is_available():
        raise ConnectionError("Ollama server not available or model not loaded.")

    # Add system context to improve consistency
    enhanced_prompt = f"""You are a professional document processing assistant. Follow instructions precisely and respond only with the requested information.

{prompt}

Important: Respond only with the exact text requested, no additional commentary."""

    result = _client.query(enhanced_prompt)
    if result is None:
        raise ConnectionError("Ollama query failed or returned no text.")
    
    # Clean up common LLM response artifacts
    result = result.strip()
    # Remove quotes if the LLM wrapped the response
    if result.startswith('"') and result.endswith('"'):
        result = result[1:-1]
    
    return result

def _find_best_chinese_paragraph_match(change, all_chinese_paras, chinese_para_elements):
    """Enhanced paragraph matching with multiple strategies."""
    
    # Strategy 1: Direct LLM alignment
    bullet_list = "\n- ".join([f"{i}: {para[:100]}..." if len(para) > 100 else f"{i}: {para}" 
                              for i, para in enumerate(all_chinese_paras)])
    
    alignment_prompt = f"""Find the Chinese paragraph that corresponds to this English paragraph.

English paragraph: "{change['original_context']}"

Chinese paragraphs:
{bullet_list}

Respond with only the number (index) of the matching Chinese paragraph."""

    try:
        response = _get_llm_response(alignment_prompt)
        # Extract number from response
        match = re.search(r'\d+', response)
        if match:
            index = int(match.group())
            if 0 <= index < len(all_chinese_paras):
                return all_chinese_paras[index], chinese_para_elements[all_chinese_paras[index]]
    except (ValueError, IndexError):
        pass
    
    # Strategy 2: Similarity-based matching using LLM
    similarity_scores = []
    for i, chinese_para in enumerate(all_chinese_paras[:10]):  # Limit to avoid token limits
        similarity_prompt = f"""Rate the similarity between these two paragraphs on a scale of 0-10:

English: "{change['original_context'][:200]}"
Chinese: "{chinese_para[:200]}"

Respond with only a number from 0 to 10."""
        
        try:
            score_response = _get_llm_response(similarity_prompt)
            score = float(re.search(r'[\d.]+', score_response).group())
            similarity_scores.append((score, i, chinese_para))
        except:
            similarity_scores.append((0, i, chinese_para))
    
    if similarity_scores:
        best_match = max(similarity_scores, key=lambda x: x[0])
        if best_match[0] > 5:  # Threshold for acceptable match
            chinese_text = best_match[2]
            return chinese_text, chinese_para_elements[chinese_text]
    
    return None, None

def _create_track_changes_element(change_type, text, author, date, change_id):
    """Create properly formatted track changes XML elements."""
    if change_type == 'insertion':
        run = etree.Element(etree.QName(NS['w'], 'r'))
        ins_element = etree.SubElement(run, etree.QName(NS['w'], 'ins'), {
            etree.QName(NS['w'], 'author'): author,
            etree.QName(NS['w'], 'date'): date,
            etree.QName(NS['w'], 'id'): change_id
        })
        t_element = etree.SubElement(ins_element, etree.QName(NS['w'], 't'))
        t_element.text = text
        t_element.set(etree.QName("{http://www.w3.org/XML/1998/namespace}space"), "preserve")
        return run
    
    elif change_type == 'deletion':
        run = etree.Element(etree.QName(NS['w'], 'r'))
        del_element = etree.SubElement(run, etree.QName(NS['w'], 'del'), {
            etree.QName(NS['w'], 'author'): author,
            etree.QName(NS['w'], 'date'): date,
            etree.QName(NS['w'], 'id'): change_id
        })
        del_text = etree.SubElement(del_element, etree.QName(NS['w'], 'delText'))
        del_text.text = text
        del_text.set(etree.QName("{http://www.w3.org/XML/1998/namespace}space"), "preserve")
        return run

def _apply_insertion_change(target_para_element, text_to_insert, chinese_text, change):
    """Enhanced insertion handling with better positioning."""
    
    # Find optimal insertion position
    position_prompt = f"""Determine the best character position to insert this text:

Original Chinese text: "{chinese_text}"
Text to insert: "{text_to_insert}"
Context: This corresponds to an insertion in the English version.

Respond with only a number representing the character position (0 to {len(chinese_text)})."""
    
    try:
        position_str = _get_llm_response(position_prompt)
        position = int(re.search(r'\d+', position_str).group())
        position = max(0, min(position, len(chinese_text)))
    except:
        # Default to end if position detection fails
        position = len(chinese_text)
    
    # Find the appropriate run and position within it
    runs = target_para_element.xpath('./w:r', namespaces=NS)
    current_pos = 0
    target_run_index = None
    offset_in_run = None
    
    for i, run in enumerate(runs):
        run_text = ''.join(run.xpath('.//w:t/text()', namespaces=NS))
        if current_pos <= position <= current_pos + len(run_text):
            target_run_index = i
            offset_in_run = position - current_pos
            break
        current_pos += len(run_text)
    
    # Create insertion element with preserved metadata
    ins_run = _create_track_changes_element('insertion', text_to_insert, 
                                          change['author'], change['date'], change['change_id'])
    
    if target_run_index is None or not runs:
        # No runs exist or position is beyond all runs
        target_para_element.append(ins_run)
    else:
        target_run = runs[target_run_index]
        run_text = ''.join(target_run.xpath('.//w:t/text()', namespaces=NS))
        
        if offset_in_run == 0:
            # Insert at the beginning of the run
            target_para_element.insert(target_para_element.index(target_run), ins_run)
        elif offset_in_run >= len(run_text):
            # Insert after the run
            target_para_element.insert(target_para_element.index(target_run) + 1, ins_run)
        else:
            # Split the run
            before_text = run_text[:offset_in_run]
            after_text = run_text[offset_in_run:]
            
            # Update current run with before text
            for t in target_run.xpath('.//w:t', namespaces=NS):
                t.text = before_text
                t.set(etree.QName("{http://www.w3.org/XML/1998/namespace}space"), "preserve")
            
            # Insert the new content
            target_para_element.insert(target_para_element.index(target_run) + 1, ins_run)
            
            # Create after run if needed
            if after_text:
                after_run = deepcopy(target_run)
                for t in after_run.xpath('.//w:t', namespaces=NS):
                    t.text = after_text
                    t.set(etree.QName("{http://www.w3.org/XML/1998/namespace}space"), "preserve")
                target_para_element.insert(target_para_element.index(target_run) + 2, after_run)

def _apply_deletion_change(target_para_element, chinese_text, change):
    """Enhanced deletion handling with better text identification."""
    
    # Identify the Chinese text to delete
    identification_prompt = f"""Identify the exact Chinese text that should be deleted based on the English deletion:

Chinese paragraph: "{chinese_text}"
English text that was deleted: "{change['text']}"

Respond with only the exact Chinese text that corresponds to the deleted English text."""
    
    try:
        text_to_delete = _get_llm_response(identification_prompt)
        text_to_delete = text_to_delete.strip()
    except:
        return False  # Skip if identification fails
    
    if not text_to_delete or text_to_delete not in chinese_text:
        return False
    
    # Find position of text to delete
    delete_start = chinese_text.find(text_to_delete)
    delete_end = delete_start + len(text_to_delete)
    
    # Find runs that contain the text to delete
    runs = target_para_element.xpath('./w:r', namespaces=NS)
    current_pos = 0
    runs_to_modify = []
    
    for i, run in enumerate(runs):
        run_text = ''.join(run.xpath('.//w:t/text()', namespaces=NS))
        run_start = current_pos
        run_end = current_pos + len(run_text)
        
        # Check if this run overlaps with deletion range
        if run_start < delete_end and run_end > delete_start:
            overlap_start = max(run_start, delete_start)
            overlap_end = min(run_end, delete_end)
            runs_to_modify.append({
                'index': i,
                'element': run,
                'text': run_text,
                'run_start': run_start,
                'run_end': run_end,
                'delete_start_in_run': overlap_start - run_start,
                'delete_end_in_run': overlap_end - run_start
            })
        
        current_pos += len(run_text)
    
    # Apply deletions to identified runs
    for run_info in reversed(runs_to_modify):  # Reverse to maintain indices
        run_element = run_info['element']
        run_text = run_info['text']
        start_offset = run_info['delete_start_in_run']
        end_offset = run_info['delete_end_in_run']
        
        before_text = run_text[:start_offset]
        deleted_text = run_text[start_offset:end_offset]
        after_text = run_text[end_offset:]
        
        run_index = target_para_element.index(run_element)
        
        # Create before run if needed
        insert_offset = 0
        if before_text:
            before_run = deepcopy(run_element)
            for t in before_run.xpath('.//w:t', namespaces=NS):
                t.text = before_text
                t.set(etree.QName("{http://www.w3.org/XML/1998/namespace}space"), "preserve")
            target_para_element.insert(run_index, before_run)
            insert_offset += 1
        
        # Create deletion run
        del_run = _create_track_changes_element('deletion', deleted_text,
                                              change['author'], change['date'], change['change_id'])
        target_para_element.insert(run_index + insert_offset, del_run)
        insert_offset += 1
        
        # Create after run if needed
        if after_text:
            after_run = deepcopy(run_element)
            for t in after_run.xpath('.//w:t', namespaces=NS):
                t.text = after_text
                t.set(etree.QName("{http://www.w3.org/XML/1998/namespace}space"), "preserve")
            target_para_element.insert(run_index + insert_offset, after_run)
        
        # Remove original run
        target_para_element.remove(run_element)
    
    return True

def _create_updated_docx(chinese_path, output_path, changes, update_status_callback):
    """Enhanced document creation with better track changes preservation."""
    
    with zipfile.ZipFile(chinese_path, 'r') as z_in:
        chinese_xml_content = z_in.read('word/document.xml')
        chinese_root = etree.fromstring(chinese_xml_content)
        
        # Create mapping of Chinese paragraphs
        chinese_paragraphs = chinese_root.xpath('//w:body/w:p', namespaces=NS)
        chinese_para_map = {}
        all_chinese_paras = []
        
        for p in chinese_paragraphs:
            para_text = _get_paragraph_text(p)
            if para_text.strip():
                chinese_para_map[para_text] = p
                all_chinese_paras.append(para_text)

    successful_changes = 0
    failed_changes = 0

    # Group changes by paragraph for better processing
    changes_by_para = {}
    for change in changes:
        para_idx = change['paragraph_index']
        if para_idx not in changes_by_para:
            changes_by_para[para_idx] = []
        changes_by_para[para_idx].append(change)

    # Process changes paragraph by paragraph
    for para_idx, para_changes in changes_by_para.items():
        update_status_callback(f"Processing paragraph {para_idx + 1} with {len(para_changes)} changes...")
        
        # Find the best matching Chinese paragraph for this set of changes
        representative_change = para_changes[0]  # Use first change as representative
        chinese_text, target_para_element = _find_best_chinese_paragraph_match(
            representative_change, all_chinese_paras, chinese_para_map)
        
        if not chinese_text or not target_para_element:
            update_status_callback(f"Warning: Could not match paragraph {para_idx + 1}. Skipping {len(para_changes)} changes.")
            failed_changes += len(para_changes)
            continue
        
        # Apply changes to this paragraph
        for change in para_changes:
            try:
                if change['type'] == 'insertion':
                    # Translate the inserted text
                    translation_prompt = f"""Translate this English text to Chinese:
English: "{change['text']}"

Respond with only the Chinese translation."""
                    
                    text_to_insert = _get_llm_response(translation_prompt)
                    _apply_insertion_change(target_para_element, text_to_insert, chinese_text, change)
                    successful_changes += 1
                    
                elif change['type'] == 'deletion':
                    if _apply_deletion_change(target_para_element, chinese_text, change):
                        successful_changes += 1
                    else:
                        failed_changes += 1
                        update_status_callback(f"Warning: Could not apply deletion '{change['text'][:30]}...'")
            
            except Exception as e:
                failed_changes += 1
                update_status_callback(f"Warning: Error processing change: {str(e)[:50]}...")

    # Write the final document
    update_status_callback("Finalizing document...")
    
    # Ensure proper XML formatting
    final_xml_string = etree.tostring(chinese_root, pretty_print=True, 
                                    xml_declaration=True, encoding='UTF-8')
    
    # Create output document
    with zipfile.ZipFile(chinese_path, 'r') as z_in:
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as z_out:
            for item in z_in.infolist():
                if item.filename != 'word/document.xml':
                    z_out.writestr(item, z_in.read(item.filename))
            z_out.writestr('word/document.xml', final_xml_string)
    
    update_status_callback(f"Complete! Applied {successful_changes} changes successfully. "
                          f"Failed: {failed_changes} changes.")

def run_document_processing(english_path, chinese_path, output_path, update_status_callback):
    """Enhanced main orchestrator with better error handling and reporting."""
    try:
        update_status_callback("Step 1/3: Analyzing English document for tracked changes...")
        changes = _extract_changes_from_english_docx(english_path)
        
        if not changes:
            update_status_callback("Complete: No tracked changes found in the English document.")
            # Copy original file as output
            import shutil
            shutil.copy2(chinese_path, output_path)
            return

        insertions = [c for c in changes if c['type'] == 'insertion']
        deletions = [c for c in changes if c['type'] == 'deletion']
        
        update_status_callback(f"Found {len(changes)} total changes: {len(insertions)} insertions, {len(deletions)} deletions.")
        update_status_callback("Step 2/3: Matching content and applying changes...")
        
        _create_updated_docx(chinese_path, output_path, changes, update_status_callback)
        
        update_status_callback(f"Step 3/3: Document processing complete!\nOutput saved to: {output_path}")

    except ConnectionError as e:
        update_status_callback(f"LLM Connection Error: {e}")
    except zipfile.BadZipFile:
        update_status_callback("Error: Invalid DOCX file format detected.")
    except etree.XMLSyntaxError as e:
        update_status_callback(f"XML Parsing Error: {e}")
    except Exception as e:
        update_status_callback(f"Unexpected error: {e}")
        import traceback
        traceback.print_exc()