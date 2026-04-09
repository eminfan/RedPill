---
name: "docx-reader"
description: "Extracts and reads text content from Word (.docx) files. Invoke when user wants to read or analyze Word document content."
---

# DOCX Reader Skill

This skill extracts plain text content from Microsoft Word (.docx) files for reading and analysis.

## When to Invoke

Invoke this skill when:
- User asks to read a Word document
- User wants to extract content from .docx files
- User needs to analyze Word document content

## How to Use

Run the PowerShell script `read-docx.ps1` located in the project root:

```powershell
powershell -ExecutionPolicy Bypass -File "d:\projects\天道\read-docx.ps1" "<docx-file-path>" "<output-txt-path>"
```

Then read the output text file using the Read tool.

## Example

```powershell
powershell -ExecutionPolicy Bypass -File "d:\projects\天道\read-docx.ps1" "d:\projects\天道\基因群体演化论\基因群体演化论核心内容.docx" "d:\projects\天道\temp-output.txt"
```

Then:
```
Read("d:\projects\天道\temp-output.txt")
```

## Technical Details

The script extracts text from `<w:t>` tags in the `word/document.xml` file inside the .docx archive (which is a ZIP file). This method:
- Works without Microsoft Word installed
- Preserves text content
- Does not preserve formatting (bold, italic, etc.)
- Does not extract images or tables properly

## Limitations

- Formatting (bold, italic, colors) is not preserved
- Tables may lose structure
- Images and embedded objects are not extracted
- Headers/footers may not be included
