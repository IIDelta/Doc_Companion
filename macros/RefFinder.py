import requests
import json
import time
import os
import re

# --- Configuration (Kept for backend, email is important) ---
UNPAYWALL_EMAIL = 'jaustind1@gmail.com' # User's email
CROSSREF_API_URL = 'https://api.crossref.org/works'
UNPAYWALL_API_URL_BASE = 'https://api.unpaywall.org/v2/'
SEMANTIC_SCHOLAR_API_URL_BASE = 'https://api.semanticscholar.org/graph/v1/paper/'

REQUEST_TIMEOUT_API = 15
REQUEST_TIMEOUT_DOWNLOAD = 60
USER_AGENT = 'ReferenceCheckerApp/1.7_GUICopySummary (mailto:{})'.format(UNPAYWALL_EMAIL) # Version bump

# --- Backend Helper Functions (sanitize_filename, get_doi_from_crossref, etc.) ---
# These functions (sanitize_filename, get_doi_from_crossref, 
# get_open_access_info_unpaywall, get_pdf_info_from_semantic_scholar, 
# download_pdf, process_references_backend) should remain the same as 
# in the script version that introduced the live summary Treeview.
# For brevity, I will not repeat them here but ensure they are correctly defined in your script.
# Make sure they use `log_callback` for GUI updates as established previously.


# Placeholder for your existing helper and backend functions:
def sanitize_filename(name):
    name = str(name) 
    name = re.sub(r'[<>:"/\\|?*]', '', name) 
    name = re.sub(r'\s+', '_', name) 
    return name[:100] 


def get_doi_from_crossref(reference_string, log_callback): # From your script
    headers = {'User-Agent': USER_AGENT} 
    params = {'query.bibliographic': reference_string, 'rows': 1} 
    try:
        response = requests.get(CROSSREF_API_URL, params=params, headers=headers, timeout=REQUEST_TIMEOUT_API) 
        response.raise_for_status() 
        data = response.json() 
        if data['message']['items']: 
            item = data['message']['items'][0] 
            doi = item.get('DOI') 
            title_list = item.get('title', []) 
            title = title_list[0] if title_list else "No title found" 
            return doi, title 
        return None, None 
    except requests.exceptions.RequestException as e:
        log_callback(f"  [CrossRef Error] Could not query CrossRef for '{reference_string[:50]}...': {e}") 
        return None, None 
    except (json.JSONDecodeError, KeyError) as e:
        log_callback(f"  [CrossRef Error] Problem with CrossRef response for '{reference_string[:50]}...': {e}") 
        return None, None 

def get_open_access_info_unpaywall(doi, log_callback): # From your script
    if not doi: return None 
    headers = {'User-Agent': USER_AGENT} 
    url = f"{UNPAYWALL_API_URL_BASE}{doi}?email={UNPAYWALL_EMAIL}" 
    try:
        response = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT_API) 
        response.raise_for_status() 
        return response.json() 
    except requests.exceptions.RequestException as e:
        log_callback(f"  [Unpaywall Error] Could not query Unpaywall for DOI '{doi}': {e}") 
        return None 
    except json.JSONDecodeError as e:
        log_callback(f"  [Unpaywall Error] Could not decode JSON from Unpaywall for DOI '{doi}': {e}") 
        return None 

def get_pdf_info_from_semantic_scholar(doi, log_callback): # From your script
    if not doi: return None, None 
    fields = "isOpenAccess,openAccessPdf,title,externalIds" 
    url = f"{SEMANTIC_SCHOLAR_API_URL_BASE}{doi}?fields={fields}" 
    headers = {'User-Agent': USER_AGENT} 
    try:
        response = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT_API) 
        response.raise_for_status() 
        data = response.json() 
        title = data.get('title', "No title found") 
        if data.get('isOpenAccess') and data.get('openAccessPdf') and data['openAccessPdf'].get('url'): 
            pdf_url = data['openAccessPdf']['url'] 
            return pdf_url, title 
        return None, title 
    except requests.exceptions.HTTPError as e: 
        if e.response.status_code == 404: 
            log_callback(f"  [Semantic Scholar] DOI '{doi}' not found.") 
        elif e.response.status_code == 429: 
            log_callback(f"  [Semantic Scholar] Rate limit exceeded for DOI '{doi}'. Try again later or use an API key.") 
        else:
            log_callback(f"  [Semantic Scholar] HTTP error for DOI '{doi}': {e}") 
        return None, None 
    except requests.exceptions.RequestException as e: 
        log_callback(f"  [Semantic Scholar] Request error for DOI '{doi}': {e}") 
        return None, None 
    except json.JSONDecodeError: 
        log_callback(f"  [Semantic Scholar] Could not decode JSON response for DOI '{doi}'.") 
        return None, None 

def download_pdf(pdf_url, chosen_output_folder, doi, title, source_service="Unknown", log_callback=print): # From your script
    if not pdf_url: 
        return None, "No PDF URL provided." 
    if not os.path.exists(chosen_output_folder): 
        log_callback(f"    ERROR: Output folder {chosen_output_folder} does not exist. Please select a valid folder.") 
        return None, f"Output folder {chosen_output_folder} does not exist." 

    browser_user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36' 
    local_headers = {'User-Agent': browser_user_agent} 
    filepath = None 
    try:
        doi_filename_part = doi.replace('/', '_').replace('.', '-') 
        if title and title not in ["No title found", "Unknown Title"]: 
            title_words = str(title).split(' ') 
            short_title_part = '_'.join(title_words[0:min(len(title_words), 5)]) 
        else:
            short_title_part = "untitled" 
        sanitized_title_short = sanitize_filename(short_title_part) 
        filename = f"{doi_filename_part}__{sanitized_title_short}.pdf" 
        filepath = os.path.join(chosen_output_folder, filename) 

        log_callback(f"    Attempting to download PDF from {source_service}: {pdf_url}") 
        pdf_response = requests.get(pdf_url, stream=True, timeout=REQUEST_TIMEOUT_DOWNLOAD, headers=local_headers) 
        if pdf_response.status_code == 403: 
            log_callback(f"    ❌ Server returned 403 Forbidden. Response text (first 200 chars): {pdf_response.text[:200]}") 
            pdf_response.raise_for_status() 
        pdf_response.raise_for_status() 

        content_type = pdf_response.headers.get('content-type', '').lower() 
        if 'application/pdf' not in content_type: 
            if not pdf_url.lower().endswith('.pdf') and '.pdf?' not in pdf_url.lower(): 
                 log_callback(f"    ⚠️ Warning: Content-type is '{content_type}'. URL doesn't end with .pdf. Proceeding, but may not be a PDF.") 
        
        file_written_successfully = False 
        with open(filepath, 'wb') as f: 
            for chunk in pdf_response.iter_content(chunk_size=8192): 
                if chunk: 
                    f.write(chunk) 
            if os.path.exists(filepath) and os.path.getsize(filepath) > 0: 
                file_written_successfully = True 
            elif os.path.exists(filepath): 
                 log_callback(f"    ⚠️ PDF created but is empty (0 bytes): {filepath}") 
                 return None, f"PDF is empty (0 bytes) for {pdf_url}" 
            else: 
                log_callback(f"    ❌ File not found after write attempt: {filepath}") 
                return None, f"File not found after write attempt for {pdf_url}" 

        if file_written_successfully: 
            file_size = os.path.getsize(filepath) 
            log_callback(f"    ✅ PDF downloaded successfully via {source_service} as: {filepath} ({file_size} bytes)") 
            return filepath, "Downloaded successfully." 
        else: 
            log_callback(f"    ❌ PDF download (via {source_service}) reported success, but file integrity check failed for: {filepath}") 
            return None, "Download succeeded but file check failed." 
    except requests.exceptions.Timeout: 
        log_callback(f"    ❌ Download timed out (via {source_service}) for {pdf_url}.") 
        return None, f"Download timed out for {pdf_url}." 
    except requests.exceptions.HTTPError as e: 
        log_callback(f"    ❌ HTTP error (via {source_service}) during PDF download from {pdf_url}: {e}") 
        return None, f"Failed to download PDF from {pdf_url}: {e}" 
    except requests.exceptions.RequestException as e: 
        log_callback(f"    ❌ Request exception (via {source_service}) during PDF download from {pdf_url}: {e}") 
        return None, f"Failed to download PDF from {pdf_url}: {e}" 
    except IOError as e: 
        log_callback(f"    ❌ File system error (via {source_service}) for {filepath if filepath else 'unknown path'}: {e}") 
        return None, f"File system error for {filepath if filepath else 'unknown path'}: {e}" 
    except Exception as e: 
        log_callback(f"    ❌ An unexpected error (via {source_service}) occurred during PDF download ({pdf_url}): {e}") 
        return None, f"An unexpected error occurred during PDF download: {e}" 

def process_references_backend(reference_list, chosen_download_folder, gui_update_callback): # From your script
    if not UNPAYWALL_EMAIL or UNPAYWALL_EMAIL == 'YOUR_EMAIL@example.com': 
        gui_update_callback({'type': 'log', 'message': "🛑 ERROR: Please set your email address in the UNPAYWALL_EMAIL variable (top of the script)."}) 
        return 

    try:
        os.makedirs(chosen_download_folder, exist_ok=True) 
    except OSError as e: 
        gui_update_callback({'type': 'log', 'message': f"🛑 ERROR: Could not create download folder '{chosen_download_folder}': {e}"}) 
        return 

    gui_update_callback({'type': 'log', 'message': f"PDFs will be saved to '{chosen_download_folder}/' folder."}) 
    gui_update_callback({'type': 'log', 'message': f"\nProcessing {len(reference_list)} references...\n"}) 

    for i, ref_string in enumerate(reference_list): 
        gui_update_callback({'type': 'log', 'message': f"Processing reference {i+1}/{len(reference_list)}: \"{ref_string[:70]}...\""}) 
        doi, crossref_title = get_doi_from_crossref(ref_string, lambda msg: gui_update_callback({'type': 'log', 'message': msg})) 

        base_entry_info = {"original_reference": ref_string, "doi": doi} 

        if not doi: 
            gui_update_callback({'type': 'log', 'message': f"  DOI not found for: \"{ref_string[:70]}...\""}) 
            summary_item = {**base_entry_info, "title": ref_string[:60] + "...", "status_info": "DOI not found", "url_info": "N/A"} 
            gui_update_callback({'type': 'summary_add', 'category_key': 'doi_not_found', 'item_data': summary_item}) 
            time.sleep(0.25) 
            continue 

        article_title_for_processing = crossref_title if crossref_title and crossref_title != "No title found" else "Unknown Title" 
        gui_update_callback({'type': 'log', 'message': f"  Found DOI: {doi} (Title: {article_title_for_processing})"}) 
        
        base_entry_info["title"] = article_title_for_processing 

        unpaywall_data = get_open_access_info_unpaywall(doi, lambda msg: gui_update_callback({'type': 'log', 'message': msg})) 
        time.sleep(0.25) 

        publisher_landing_url = f"https://doi.org/{doi}" 
        final_title = article_title_for_processing 
        entry_for_summary = {} 

        if not unpaywall_data: 
            gui_update_callback({'type': 'log', 'message': f"  Could not retrieve data from Unpaywall for DOI: {doi}. Trying Semantic Scholar as fallback."}) 
            ss_pdf_url, ss_title = get_pdf_info_from_semantic_scholar(doi, lambda msg: gui_update_callback({'type': 'log', 'message': msg})) 
            time.sleep(0.25) 
            
            final_title = ss_title or article_title_for_processing 
            entry_for_summary = { 
                **base_entry_info, "title": final_title, 
                "status_info": "Unpaywall failed. Semantic Scholar: " + ("PDF found" if ss_pdf_url else "No PDF"), 
                "url_info": ss_pdf_url or publisher_landing_url 
            }
            category_key_to_use = 'unpaywall_data_unavailable' 

            if ss_pdf_url: 
                gui_update_callback({'type': 'log', 'message': f"  [Semantic Scholar] Found PDF for Unpaywall-failed item: {ss_pdf_url}"}) 
                downloaded_filepath, download_status = download_pdf(ss_pdf_url, chosen_download_folder, doi, final_title, source_service="Semantic Scholar", log_callback=lambda msg: gui_update_callback({'type': 'log', 'message': msg})) 
                if downloaded_filepath: 
                    entry_for_summary["status_info"] = f"Downloaded via Semantic Scholar: {os.path.basename(downloaded_filepath)}" 
                    entry_for_summary["url_info"] = ss_pdf_url 
                    category_key_to_use = 'downloaded_pdf' 
                else: 
                    entry_for_summary["status_info"] = f"Semantic Scholar PDF found but download failed: {download_status}" 
                    category_key_to_use = 'open_access_no_pdf_link' 
            gui_update_callback({'type': 'summary_add', 'category_key': category_key_to_use, 'item_data': entry_for_summary}) 
            continue 
        
        final_title = unpaywall_data.get('title') or article_title_for_processing 
        publisher_landing_url = unpaywall_data.get('doi_url') or f"https://doi.org/{doi}" 
        
        entry_for_summary = { 
            **base_entry_info, "title": final_title, 
            "status_info": f"Unpaywall OA Status: {unpaywall_data.get('oa_status', 'unknown')}", 
            "url_info": publisher_landing_url 
        }
        category_key_to_use = 'paywalled' 

        pdf_download_path_from_unpaywall = None 
        unpaywall_pdf_url = None 

        if unpaywall_data.get('is_oa'): 
            best_oa_location = unpaywall_data.get('best_oa_location', {}) 
            unpaywall_pdf_url = best_oa_location.get('url_for_pdf') if best_oa_location else None 
            oa_content_landing_url = best_oa_location.get('url') if best_oa_location else None 
            
            entry_for_summary["url_info"] = oa_content_landing_url or publisher_landing_url 

            if unpaywall_pdf_url: 
                entry_for_summary["url_info"] = unpaywall_pdf_url 
                temp_dl_path, temp_dl_status = download_pdf(unpaywall_pdf_url, chosen_download_folder, doi, final_title, source_service="Unpaywall", log_callback=lambda msg: gui_update_callback({'type': 'log', 'message': msg})) 
                if temp_dl_path: 
                    pdf_download_path_from_unpaywall = temp_dl_path 
                    entry_for_summary["status_info"] = f"Downloaded via Unpaywall: {os.path.basename(pdf_download_path_from_unpaywall)}" 
                    category_key_to_use = 'downloaded_pdf' 
                else: 
                    entry_for_summary["status_info"] = f"Unpaywall PDF download failed: {temp_dl_status}" 
                    category_key_to_use = 'open_access_no_pdf_link' 
            else: 
                entry_for_summary["status_info"] = f"Unpaywall OA (Status: {entry_for_summary['status_info']}), no direct PDF." 
                category_key_to_use = 'open_access_no_pdf_link' 

            if not pdf_download_path_from_unpaywall: 
                gui_update_callback({'type': 'log', 'message': f"  Attempting Semantic Scholar backup for {doi}..."}) 
                ss_pdf_url, ss_title = get_pdf_info_from_semantic_scholar(doi, lambda msg: gui_update_callback({'type': 'log', 'message': msg})) 
                time.sleep(0.25) 
                if ss_pdf_url: 
                    gui_update_callback({'type': 'log', 'message': f"  [Semantic Scholar] Found alternative PDF: {ss_pdf_url}"}) 
                    title_for_ss_download = ss_title if (ss_title and ss_title != "No title found") else final_title 
                    downloaded_filepath, download_status = download_pdf(ss_pdf_url, chosen_download_folder, doi, title_for_ss_download, source_service="Semantic Scholar", log_callback=lambda msg: gui_update_callback({'type': 'log', 'message': msg})) 
                    entry_for_summary["url_info"] = ss_pdf_url 
                    if downloaded_filepath: 
                        entry_for_summary["status_info"] = f"Downloaded via Semantic Scholar: {os.path.basename(downloaded_filepath)}" 
                        category_key_to_use = 'downloaded_pdf' 
                    else: 
                        entry_for_summary["status_info"] = f"Unpaywall attempt failed. Semantic Scholar PDF found but download failed: {download_status}" 
                        # category_key_to_use remains 'open_access_no_pdf_link'
                else: 
                    gui_update_callback({'type': 'log', 'message': f"  [Semantic Scholar] No PDF link found as backup for {doi}."}) 
                    entry_for_summary["status_info"] += " Semantic Scholar found no PDF." 
        else: 
            entry_for_summary["status_info"] = f"Paywalled (Unpaywall: {unpaywall_data.get('oa_status', 'unknown')})" 
            category_key_to_use = 'paywalled' 
        
        gui_update_callback({'type': 'summary_add', 'category_key': category_key_to_use, 'item_data': entry_for_summary}) 
            
    gui_update_callback({'type': 'log', 'message': "\n--- Processing Complete ---"}) 

