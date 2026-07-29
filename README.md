# Doc\_Companion

Doc\_Companion is a helpful desktop utility designed to streamline common document processing tasks within Microsoft Word. It provides a user-friendly interface to automate actions like finding and defining acronyms, performing bulk text replacements, and cleaning up documents.

## Features

Doc\_Companion offers several tools to enhance your document workflow:

1.  **Acronym Finder & Table Generator:**
    * Scans your active Word document to identify potential acronyms.
    * Categorizes findings into 'Likely', 'Possible', and 'Unlikely' tabs for easy review.
    * Fetches definitions from a base list (updated online) and a customizable user list.
    * Allows you to review, edit definitions, and select acronyms.
    * Generates a clean, sorted "List of Acronyms and Abbreviations" table in a new Word document.
    * Supports custom include and exclude lists for finer control.

2.  **Replace Values in Selection:**
    * Performs find-and-replace operations within a selected portion of your Word document.
    * Uses an Excel file as the source for find/replace pairs (Column A: Find, Column B: Replace).
    * Optionally supports Microsoft Word's wildcard characters for advanced search patterns.

3.  **Clean & Protect Document:**
    * Processes a chosen Word document (\*.docx or \*.doc).
    * Updates all document fields.
    * Deletes all comments.
    * Accepts all tracked revisions.
    * Applies protection to the document, allowing only revisions (requires a predefined password for unprotection within the tool).
    * Saves the processed file as a new copy with a `_clean` suffix added to its name.

4.  **Reference/Open-Access Finder:**
    * **Locates DOIs** from a list of references.
    * **Checks for Open Access (OA) availability** via Unpaywall and Semantic Scholar.
    * **Downloads available PDFs** to a specified output folder.
    * Provides a **live log** of the processing steps.
    * Displays a **categorized summary** of results (Downloaded PDFs, OA no PDF, Paywalled, DOI Not Found, etc.).
    * **New:** Accepts references either from a **`.txt` file** or directly from your **clipboard**.
    * **New:** Includes a **"Cancel" button** to stop the processing mid-way if needed.
    * Allows **copying the full summary** to your clipboard for easy sharing or record-keeping.

## Usage

1.  **Main Window:**
    * The main window displays the currently active Word document.
    * Use the "Stay on Top" checkbox to keep the window visible.
    * Click the buttons to access different features.

2.  **Replace Values:**
    * Click "Replace Values".
    * Choose your Excel file containing find/replace pairs.
    * Select the text in your Word document where you want the replacements to occur.
    * Check "Use Wildcards" if your Excel 'Find' column uses Word's wildcard syntax.
    * Click "Run Macro".

3.  **Acronyms Table:**
    * Click "Acronyms Table".
    * Click "Find Acronyms" to scan the active Word document.
    * Review the 'Likely', 'Possible', and 'Unlikely' tabs.
    * Edit definitions directly in the table (these will be saved to your user list).
    * Use the checkboxes in the 'Include' column to select acronyms for your table.
    * Use the "Check/Uncheck All" buttons for quick selection.
    * Click "Generate Table" and choose a location to save your new Word document.

4.  **Reference/Open-Access Finder:**
    * Click "Reference/Open-Access Finder".
    * **Choose an output folder** where downloaded PDFs will be saved.
    * **To process from a file:** Click "Select Input File (.txt)" and choose your references file, then click "Start Processing (File)".
    * **To process from clipboard:** Copy your list of references to your system clipboard, then click "Process from Clipboard".
    * **To stop an ongoing process:** Click the "Cancel Process" button.
    * Once processing is complete (or cancelled), click "Copy Summary" to get a detailed overview of the results in your clipboard.

5.  **Clean & Protect:**
    * Click "Clean & Protect Document".
    * Select the Word document you wish to process.
    * The tool will perform the cleaning actions and save a new file. A log of actions will be displayed.

## Configuration

* **User Acronyms:** A `user_acronyms.txt` file is automatically created in `C:\Users\<YourUsername>\.doc_companion\`. You can manually edit this file (using Tab as a separator) or let the Acronyms window update it when you edit definitions.
* **Base Acronyms:** A base list is fetched from GitHub and cached locally.
* **Include/Exclude Lists (Advanced):** You can create `user_exclude.txt` and `user_include.txt` files in the `C:\Users\<YourUsername>\.doc_companion\` directory to force certain words to be ignored or always considered (even if not following standard patterns). Add one word/phrase per line.

### Prepare PleaseReview Copy

Creates a safe review copy of a Word document containing tracked changes.

- The source document is never modified.
- Text insertions are accepted and highlighted yellow.
- Text deletions are rejected, restoring the deleted text, then marked using the
  `DC PleaseReview Deleted Text` custom character style.
- By default, restored deletions appear as red strikethrough text.
- Optional: restored deletions can also receive yellow highlight.
- Headers, footers, footnotes, endnotes, text boxes, and other accessible Word
  story ranges are processed where revisions are available.
- Unsupported revision types are retained and listed in an adjacent JSON audit log.
- Processing can be cancelled safely between revisions.
- A partial output created after cancellation is clearly marked in the audit log.

#### Removing redactions after review

1. Open the prepared PleaseReview copy in Word.
2. Press `Ctrl+H`.
3. Select **More**.
4. Choose **Format** → **Style**.
5. Select `DC PleaseReview Deleted Text`.
6. Leave **Replace with** blank.
7. Use **Replace All** only after final review and approval.

> Always retain the original tracked-changes protocol and the generated audit JSON
> until the review process is complete.
