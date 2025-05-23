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

4.  **Clean & Protect:**
    * Click "Clean & Protect Document".
    * Select the Word document you wish to process.
    * The tool will perform the cleaning actions and save a new file. A log of actions will be displayed.

## Configuration

* **User Acronyms:** A `user_acronyms.txt` file is automatically created in `C:\Users\<YourUsername>\.doc_companion\`. You can manually edit this file (using Tab as a separator) or let the Acronyms window update it when you edit definitions.
* **Base Acronyms:** A base list is fetched from GitHub and cached locally.
* **Include/Exclude Lists (Advanced):** You can create `user_exclude.txt` and `user_include.txt` files in the `C:\Users\<YourUsername>\.doc_companion\` directory to force certain words to be ignored or always considered (even if not following standard patterns). Add one word/phrase per line.