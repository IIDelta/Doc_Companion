# macros/CleanDocument.py
import sys
import os
import win32com.client as win32
import time
import re

def generate_clean_filename(file_path):
    """
    Generates a new filename with '_clean' added after the YYYY.MM.DD date.
    If no date is found, adds '_clean' before the extension.
    Uses the logic from the provided script.
    """
    dir_name = os.path.dirname(file_path)
    base_name = os.path.basename(file_path)
    file_root, file_ext = os.path.splitext(base_name)

    # Search for the YYYY.MM.DD pattern
    match = re.search(r'(\d{4}\.\d{2}\.\d{2})', file_root)

    if match:
        date_end_index = match.end(1)
        part_before_and_date = file_root[:date_end_index]
        # Build the new root name (as per original script logic)
        new_file_root = f"{part_before_and_date}_clean"
    else:
        # No date found, just add _clean at the end
        new_file_root = f"{file_root}_clean"

    # Combine back into a full path
    new_filename = f"{new_file_root}{file_ext}"
    return os.path.join(dir_name, new_filename)


def process_word_document(file_path, password):
    """
    Opens a Word document, performs processing, and saves it as a copy.
    Returns a tuple (success: bool, messages: list[str]).
    """
    word = None
    doc = None
    absolute_path = os.path.abspath(file_path)
    messages = []

    messages.append("-" * 50)
    messages.append(f"Starting processing for: {absolute_path}")
    messages.append("-" * 50)

    if not os.path.exists(absolute_path):
        messages.append(f"Error: File not found at {absolute_path}")
        return False, messages

    new_save_path = generate_clean_filename(absolute_path)
    messages.append(f"Output file will be: {new_save_path}")

    try:
        messages.append("Starting Word application...")
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False

        WD_ALLOW_ONLY_REVISIONS = 0  # As requested (standard is 2)
        WD_NO_PROTECTION = -1
        messages.append(f"Using Protection Type = {WD_ALLOW_ONLY_REVISIONS}")

        messages.append(f"Opening document: {os.path.basename(absolute_path)}...")
        doc = word.Documents.Open(absolute_path)

        messages.append("Checking document protection status...")
        if doc.ProtectionType != WD_NO_PROTECTION:
            messages.append(f"Document is protected (Type: {doc.ProtectionType}). Attempting to unprotect...")
            try:
                doc.Unprotect(Password=password)
                messages.append("Document successfully unprotected.")
            except Exception as unprotect_err:
                messages.append(f"!!! CRITICAL ERROR: Could not unprotect document: {unprotect_err}")
                messages.append("    Aborting processing.")
                doc.Close(SaveChanges=False)
                word.Quit()
                return False, messages
        else:
            messages.append("Document is not protected. Proceeding.")

        messages.append("Updating all fields...")
        for story in doc.StoryRanges:
            story.Fields.Update()
            current_story = story
            while current_story.NextStoryRange:
                current_story = current_story.NextStoryRange
                current_story.Fields.Update()
        messages.append("Fields updated.")
        time.sleep(1)

        messages.append("Deleting all comments...")
        comment_count = doc.Comments.Count
        if comment_count > 0:
            doc.DeleteAllComments()
            messages.append(f"{comment_count} comments deleted.")
        else:
            messages.append("No comments found to delete.")
        time.sleep(1)

        messages.append("Accepting all tracked revisions...")
        revision_count = doc.Revisions.Count
        if revision_count > 0:
            doc.AcceptAllRevisions()
            messages.append(f"{revision_count} revisions accepted.")
        else:
            messages.append("No revisions found to accept.")
        time.sleep(1)

        messages.append(f"Applying protection (Using Type={WD_ALLOW_ONLY_REVISIONS})...")
        doc.Protect(Type=WD_ALLOW_ONLY_REVISIONS, Password=password)
        messages.append("Document protected.")

        messages.append(f"Saving document as: {os.path.basename(new_save_path)}...")
        doc.SaveAs(new_save_path)
        messages.append("Closing original document...")
        doc.Close(SaveChanges=False)
        doc = None

        messages.append("-" * 50)
        messages.append(f"Processing completed successfully! Output: {new_save_path}")
        messages.append("-" * 50)
        return True, messages

    except Exception as e:
        messages.append("\n" + "=" * 50)
        messages.append(f"!!! An Error Occurred !!!")
        messages.append(f"    Error Type: {type(e).__name__}")
        messages.append(f"    Details: {e}")
        if hasattr(e, 'excepinfo'):
            messages.append(f"    COM Error Info: {e.excepinfo}")
        messages.append("=" * 50 + "\n")
        return False, messages

    finally:
        if doc:
            try: doc.Close(SaveChanges=False)
            except Exception: pass
        if word:
            try: word.Quit()
            except Exception: pass
        word = None
        doc = None