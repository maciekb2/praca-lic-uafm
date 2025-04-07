# Google Docs Script: Praca Lic. UAFM (Quote and Cleanup Tools)

## Description

This Google Apps Script, adds useful Polish language editing and cleanup tools to your Google Docs documents. It provides functionalities accessible through a custom menu named **`Praca Lic. UAFM`**.

### Features (Menu Items)

The script adds the following options under the **`Praca Lic. UAFM`** menu in Google Docs:

* **`Sprawdź parzystość cudzysłowów` (Check Quote Parity):** Analyzes the entire document (including headers, footers, and footnotes) and counts all found quotation marks of various types (`"`, `“`, `”`, `„`, `«`, `»`). It informs you if their total number is even. If the number is odd (which usually indicates an error in the text), it highlights the last found quote in red to help you locate it.
* **`Zamień cudzysłowy na polskie` (Replace Quotes with Polish):** Finds all specified quotation marks (`"`, `“`, `”`, `„`, `«`, `»`) throughout the document. If their total count is even (it checks this automatically), it replaces them with the correct Polish typographic quotes – opening `„` and closing `”`, ensuring they alternate correctly. It asks for confirmation before making changes.
* **`Podświetl treść (w akapicie)` (Highlight Content (within paragraph)):** Finds pairs of any quotation marks (`"..."`, `„...”`, etc.) that are located within the same paragraph (or other text block) and highlights the text *between* these quotes with a yellow background. This helps in visually verifying quotes. It doesn't work for quote pairs split across paragraphs.
* **`Wyczyść podświetlenia` (Clear Highlights):** Removes *all* background colors (yellow, red, or others) from the entire text in the document (including headers, footers, and footnotes). Useful for removing markings added by this script's functions.
* **`Wstaw twarde spacje po sierotkach` (Insert non-breaking spaces after orphans):** Searches the **main body text** of the document (ignores headers, footers, footnotes) for all short words (one or two letters, e.g., `i`, `w`, `na`, `o`, `że` in Polish context) followed by a space and replaces that regular space with a "non-breaking space" (`&nbsp;`). A non-breaking space prevents the short word from being moved to the next line alone, improving text aesthetics. It asks for confirmation before running.
* **`usuń podwójne spacje, spacje przed interpunkcją ...` (remove double spaces, spaces before punctuation...):** Performs a set of automatic copy-editing cleanups on the **main body text** (ignores headers, footers, footnotes). The operations include:
    * Replacing three consecutive dots (`...`) with the correct ellipsis character (`…`).
    * Removing double or multiple spaces and replacing them with a single space.
    * Removing spaces that appear immediately before punctuation marks (`,`, `.`, `;`, `:`, `?`, `!`).
    * Replacing a standard hyphen (`-`) with an en dash (`–`) in number ranges (e.g., `10-20` becomes `10–20`).
    * Inserting a non-breaking space after numbers followed by common units or abbreviations (e.g., `10 kg`, `5 zł`, `2023 r.`, `8 %`).
    It asks for confirmation before running.

## Installation / Setup

Follow these steps to add the script to your Google Docs document:

**Before you start:**
* You need edit permissions for the Google Docs document where you want to add the script.
* The script code is located in this repository: [https://github.com/maciekb2/praca-lic-uafm](https://github.com/maciekb2/praca-lic-uafm)

**Steps:**

1.  **Copy the Script Code from GitHub:**
    * Go to the repository page: [https://github.com/maciekb2/praca-lic-uafm](https://github.com/maciekb2/praca-lic-uafm)
    * In the list of files, find and click on the file named `kod.js`.
    * In the top right corner of the code view, click the **Raw** button. This will display the plain code text.
    * Select **all** the code text on the page (keyboard shortcut: `Ctrl+A` or `Cmd+A`).
    * Copy the selected code to your clipboard (`Ctrl+C` or `Cmd+C`).

2.  **Open Your Google Document and the Script Editor:**
    * Open the Google Docs document where you want to use the script.
    * In the top menu, click **Extensions** > **Apps Script**. A new browser tab will open with the script editor.

3.  **Paste the Copied Code:**
    * In the script editor tab, select any default code present (`Ctrl+A` or `Cmd+A`).
    * Delete the selected code (`Delete` or `Backspace` key).
    * Paste the code you copied from GitHub (`Ctrl+V` or `Cmd+V`).

4.  **Save the Script:**
    * Click the **Save project** icon (looks like a floppy disk) at the top of the editor.
    * If prompted, give the project a name (e.g., "**Praca Lic UAFM Script**") and click **Rename**.

5.  **Activate the Script:**
    * Go back to your Google Docs document browser tab.
    * **Refresh the page** (`F5` key or your browser's refresh button) or close and reopen the document. This is necessary for the script to create its menu.
    * After the document reloads, a new menu named **`Praca Lic. UAFM`** should appear in the top menu bar.

6.  **Authorize the Script (First Time Use Only):**
    * Click the new **`Praca Lic. UAFM`** menu.
    * Choose any option from the menu (e.g., `Sprawdź parzystość cudzysłowów`).
    * An **Authorization required** window will appear. The script needs your permission to run. Click **Continue**.
    * Choose your Google Account.
    * You might see a **"Google hasn’t verified this app"** screen.
        * **This is normal** for scripts not installed from the Google Workspace Marketplace. Since you are installing the script from a known source (this GitHub repository), it's generally safe to proceed if you trust the source.
        * Click **Advanced**.
        * Click **Go to [Your Script Project Name] (unsafe)**. (The name will be what you entered in Step 4).
    * Review the permissions the script needs (e.g., access your documents where it's installed, display user interface elements).
    * Click **Allow**.

## Usage

Once installed and authorized, simply use the functions available under the **`Praca Lic. UAFM`** menu in your Google Docs document. Remember that some functions ask for confirmation before modifying your document text.

## Important Notes

* Only install scripts from sources you trust. Granting permission allows the script to perform actions on your document as described in the permission request.
* This `README.md` file may contain additional information or updates.
* If the `Praca Lic. UAFM` menu doesn't appear after refreshing your document, double-check that you saved the script project correctly (Step 4) and try refreshing the document page again.
* This script is installed on a per-document basis. If you want to use it in other Google Docs files, you will need to repeat the installation steps for each document.

---
*This README file structure was generated on Monday, April 7, 2025, 6:56:59 PM CEST, in Kraków, Lesser Poland Voivodeship, Poland.*
