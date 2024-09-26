
DOCX Variable Replacement Library
=================================

This Node.js library provides an efficient solution for processing DOCX files, extracting placeholders, and replacing them with dynamic values, even in cases where the placeholders span across multiple text runs or paragraphs. Built on top of `docx4js`, it ensures that placeholders are located and filled correctly, regardless of how they are split in the document.

Features
--------

*   **Extract Placeholders (Variables)**: Identify and list all placeholders embedded in DOCX files, even if they are split across different runs or paragraphs.
*   **Replace Placeholders with Dynamic Content**: Replace identified placeholders with actual values based on a provided mapping.
*   **Handles Complex DOCX Structure**: Correctly processes placeholders that are split across multiple text runs or paragraphs, a common challenge with DOCX files.
*   **Save Modified DOCX**: Save the updated DOCX file with filled-in content, ready for use.

Installation
------------

1.  Clone this repository to your local machine.
2.  Install the required dependencies using npm:
    
        npm install docx4js cross-blob
    

Usage
-----

### 1\. Extract Variables (Placeholders) from a DOCX File

You can extract all placeholders formatted as `$({variable_name})` from a DOCX file as follows:

    const { docGetVars } = require('./path_to_library');
    
    docGetVars('input.docx').then(vars => {
        console.log(vars);  // Output: Array of placeholders found in the document
    });
    

### 2\. Replace Placeholders and Save the Document

To replace placeholders with real values, pass a key-value map where keys are placeholder names and values are the text you want to insert:

    const { docFill } = require('./path_to_library');
    
    const replacements = {
        name: "John Doe",
        date: "2024-09-25",
        location: "Tunis, Tunisia"
    };
    
    docFill('input.docx', 'output.docx', replacements).then(() => {
        console.log("Document saved with replacements.");
    });
    

How It Solves Technical Challenges
----------------------------------

### Handling Split Placeholders Across Runs and Paragraphs

In DOCX files, text can be split into multiple _runs_ or _paragraphs_ due to formatting or internal structuring, making simple string replacements impossible. For example, a placeholder like `$(name)` could be broken into separate segments in the DOCX structure:

    <w:t>$</w:t><w:t>(na</w:t><w:t>me)</w:t>
    

This library solves that by:

1.  **Accurately Detecting Split Variables**: The library's custom handler (`MyModelhandler`) processes the document in a way that recognizes placeholders, even when they are divided across different text runs.
2.  **Collision Detection Logic**: The `collisionArea1D` function calculates overlap between text runs and placeholders, ensuring that the placeholder is replaced correctly, even if it spans multiple runs or paragraphs.
3.  **Efficient Replacement Process**: The `buildTextFromMutations` function combines segments back together, applying the necessary replacements while maintaining document structure. This ensures that any placeholder, no matter how fragmented, is replaced seamlessly.

### Example of a Complex Placeholder Replacement:

If a DOCX file contains this:

    <w:t>$</w:t>
    <w:t>(</w:t>
    <w:t>name</w:t>
    <w:t>)</w:t>
    

The library will detect it as the complete placeholder `$(name)` and replace it with the corresponding value, such as _John Doe_, without breaking the document structure.

### Why This Matters:

*   **DOCX Structure Complexity**: In Microsoft Word, text is often split across multiple runs due to styling, font changes, or internal file formatting, making direct text replacement difficult.
*   **Preserving Formatting**: By working at the level of runs and paragraphs, this library ensures that replacements are made without altering the original document's format or layout.
*   **Robust Replacement Process**: Even when placeholders are split across different sections of a document, this library manages to locate and replace them, ensuring accurate and reliable content injection.

API Reference
-------------

### `docGetVars(inputPath)`

*   **Description**: Extracts all variables (placeholders) from the specified DOCX file.
*   **Input**: `inputPath` - Path to the DOCX file.
*   **Output**: Returns a Promise that resolves to an array of variable names found in the file.

### `docFill(inputPath, outputPath, map)`

*   **Description**: Replaces placeholders in the DOCX file with values from the `map` object and saves the updated document.
*   **Input**:
    *   `inputPath` - Path to the DOCX file to be processed.
    *   `outputPath` - Path where the modified DOCX file will be saved.
    *   `map` - Object where keys are the placeholder names and values are the replacement text.
*   **Output**: Returns a Promise that resolves when the document is saved.

Example
-------

Given a DOCX template that contains:

    Dear $(name),
    
    Your appointment is on $(date) at $(location).
    

You can use the following code to fill the placeholders:

    const replacements = {
        name: "John Doe",
        date: "2024-09-25",
        location: "Tunis, Tunisia"
    };
    
    docFill('template.docx', 'output.docx', replacements).then(() => {
        console.log("Document saved with dynamic content.");
    });
    

The output DOCX file will contain:

    Dear John Doe,
    
    Your appointment is on 2024-09-25 at Tunis, Tunisia.
    

Technical Details
-----------------

*   **DOCX Parsing**: Uses `docx4js` to parse the DOCX document, extract the text content, and handle the replacement process.
*   **Custom Model Handler**: The `MyModelhandler` class extends the functionality of `docx4js` to allow precise control over text replacement in runs and paragraphs.
*   **Blob Support**: Utilizes `cross-blob` for handling document saving, ensuring cross-environment compatibility.

Requirements
------------

*   Node.js v12 or higher
*   `docx4js` for DOCX manipulation
*   `cross-blob` for handling Blobs

License
-------

This project is licensed under the MIT License.
