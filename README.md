# InDesign DOCX Batch Import Script

This repository contains `word-files-to-indesign.jsx`, an Adobe InDesign script that places multiple DOCX files into a document using predefined master pages. It automates creating spreads for each DOCX, inserting the text into threaded frames, and adding optional separator pages.

## Key Features
- Adds one spread per DOCX based on your main master page.
- Places DOCX content into threaded text frames, preserving graphics where possible.
- Inserts separator pages based on a second master page (optional after the final file).
- Keeps your existing layout untouched outside the created spreads.

## Prerequisites
- Adobe InDesign with scripting enabled (tested with ExtendScript / JSX).
- A document prepared with:
  - Master page `A` containing the target spread and threaded text frames.
  - Master page `B` containing a single separator page design.
- A folder of DOCX files that follow the naming order you want (they are imported alphabetically).

## Installation
1. Copy `word-files-to-indesign.jsx` into your InDesign Scripts panel folder. Typical location:
   En: Window > Utilities > Scripts
   De: Fenster > Hilfsprogramme > **Skripte**
2. Open the **User** folder.
3. Right click and select **Show in Finder**.
4. Copy `word-files-to-indesign.jsx` into this folder. Probably called **Scripts Panel**

## Usage
1. Open the InDesign document that should receive the DOCX content.
2. Ensure the master pages referenced at the top of the script are set correctly:
   ```javascript
   var MAIN_MASTER_NAME = "A";          // master that defines the content spread
   var SEPARATOR_MASTER_NAME = "B";     // master that defines the separator page
   var INSERT_SEPARATOR_AFTER_LAST = false; // set true to add a separator after the final DOCX
   ```
   Adjust these variables directly in the script if your document uses different master names or you want a separator after the final file.
3. Run the script from the Scripts panel.
4. Confirm the setup checklist displayed by the script.
5. Select the folder containing your DOCX files when prompted. The script will:
   - Sort the DOCX files alphabetically.
   - Add a new spread using the main master for each file.
   - Place the DOCX content into the threaded text frames on the new spread.
   - Insert a separator page between entries using the separator master (and optionally after the last entry).

## Notes
- Existing pages in the document remain untouched; the script adds new pages at the end.
- Word import preferences are temporarily adjusted for the duration of the script and restored afterward.
- Sample DOCX files used for testing are available in `test-texts/`.

## Troubleshooting
- **Missing master spreads**: If the script alerts that it cannot find master `A` or `B`, rename the masters in your document or update the variables at the top of the script.
- **Empty text frames**: Ensure the master spread contains threaded text frames. The script clears the first frame in the thread before placing content.
- **Import settings**: If you need different Word import behavior (styles, bullets, etc.), adjust `configureWordImportPreferences()` inside the script.

Use this script as a starting point for automating similar DOCX-to-InDesign workflows.
