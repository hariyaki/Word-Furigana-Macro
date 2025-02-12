# Word Furigana Macro

This repository contains a **VBA macro**—**FuriganaMaker**—that automates adding **furigana** (ruby text) to Kanji in Microsoft Word documents. It also unlinks existing fields to avoid field-code expansions that can interfere with the process. The macro processes your document from the end to the start, preventing re-selection of newly inserted furigana fields.

## Table of Contents
1. [Features](#features)
2. [Macro Creation Notes](#macro-creation-notes)
3. [Installation](#installation)
4. [Usage](#usage)
5. [Notes](#notes)
6. [License](#license)

---

## Features
- **Unlink Fields**: Converts any leftover field codes (e.g., from previous macros or equation editors) into plain text.
- **Backward Scanning**: Processes from the document’s end to its start, avoiding repeated selections of newly created furigana.
- **Kanji Detection**: Only applies furigana to recognized Kanji code points (`U+4E00–U+9FFF`, plus Extension A if needed).
- **Auto “Phonetic Guide”**: Leverages Word’s built-in phonetic guide dialog, pressing “OK” for you via a simulated Enter keystroke.

---

## Macro Creation Notes
- **Remove Existing Furigana**: If a document already has phonetic guides, it may cause issues. It’s best to remove them first or start with a fresh document.
- **Macro Window**: Press **Alt + F8** in Word to open the **Macros** window.
- **Forced Stop**: If the macro is running too long or appears stuck, press **Ctrl + Alt + B** to force-stop it.
- **Avoid Interaction**: Do not use the mouse or keyboard while the macro is running—the “SendKeys” approach is fragile and may lose focus.
- **Compound Words**: Ensure compound kanji words have **no spaces or line breaks** separating the characters; otherwise, the macro may treat them as separate runs.

---

## Installation

You can install **FuriganaMaker.bas** either as a **global** module (available in all documents) or as a **local** module (affects only the current document):

### Option A: Global Installation (Normal.dotm)
1. **Open Word** and press **Alt + F11** to open the VBA Editor.
2. In the **Project** pane (usually on the left), expand **Normal** → **Modules**.
   - If no Modules folder exists, right-click **Normal** and choose **Insert → Module**.
3. **Copy** the contents of [`FuriganaMaker.bas`](FuriganaMaker.bas) into that module window.
4. **Close** the VBA Editor and **confirm** saving changes to `Normal.dotm` if prompted.
5. Now the macro is available in every Word document you open.

### Option B: Local Installation (Current Document)
1. **Open Word** and press **Alt + F11** to open the VBA Editor.
2. In the **Project** pane, find your current document (e.g., **Document1**).
3. Right-click the document’s project, choose **Insert → Module**.
4. **Copy** the contents of [`FuriganaMaker.bas`](FuriganaMaker.bas) into that new module window.
5. **Close** the VBA Editor. The macro is now stored in **this** document only.

---

## Usage
1. **Alt + F8**: In Word, press **Alt + F8** to open the **Macros** dialog.
2. **Select “FuriganaMaker”**
3. **Run**: The macro will:
   - Unlink all fields (removing them as fields but preserving other formatting).
   - Set the document’s language to Japanese.
   - Iterate from the end to the start, applying furigana to consecutive Kanji runs.
4. **Forced Stop** (if needed): Press **Ctrl + Alt + B** to abort.

---

## Notes
- **Timing**: If Word is very large or slow, increase the wait time (`WAIT_SECONDS`) in the code.
- **Extended Blocks**: If your text contains uncommon Kanji outside `U+4E00–U+9FFF` or Extension A, update the `IsKanjiChar` function accordingly.
- **No Interaction**: The macro relies on **SendKeys** to confirm the Phonetic Guide dialog. Avoid other keyboard/mouse usage during its run.

---

## License
[MIT License](LICENSE).  
