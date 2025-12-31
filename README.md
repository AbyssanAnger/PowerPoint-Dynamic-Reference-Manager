# PowerPoint Dynamic Reference Manager

[![Support me](https://img.shields.io/badge/Donate-PayPal-blue.svg)](https://paypal.me/ThatDudeSebastian?locale.x=de_DE&country.x=DE)
[![Feedback](https://img.shields.io/badge/Feedback-Welcome-green.svg)](https://github.com/AbyssanAnger/PowerPoint-Dynamic-Reference-Manager/issues)

> **Support this project**: If you find this tool helpful, consider [supporting its development](https://paypal.me/ThatDudeSebastian?locale.x=de_DE&country.x=DE). I love to hear feedback and will implement improvements when possible!

## 📘 User Guide

### Installation

1. Download the "Dynamic_References_Manager.ppam" File from the "Add-In"-Folder inside of the repository.
2. Drag the Add-In File: "Dynamic_References_Manager.ppam" into the Microsoft Add-In Folder: "C:\Users\User\AppData\Roaming\Microsoft\AddIns"
3. **Open PowerPoint** go into "Options" &#8594; "Add-Ins" &#8594; "Manage: PowerPoint-Add-Ins" &#8594; "Go" and choose "Dynamic_References_Manager.ppam".

### Usage

#### Step 1: Load Sources

1. When you select "create bibliography", a file dialog will open automatically.
2. Select your `.bib` file (e.g., exported from Zotero/Citavi).
3. Alternatively, you can use the included `test_references.bib` for testing.

> **Note**: If you do not select a file, the bibliography will remain empty.

#### Step 2: Insert Citations into Slides

Insert citations in your PowerPoint slides using the format `[Key]`, where `Key` corresponds to the citation key in your BibTeX file:

**Example:**

```
"According to [Smith2020], machine learning is..."
```

#### Step 3: Update Citations

For when you add new slides with citations or when you add new citations to existing slides. Or for when the order of the citations changes.

1. Select "Update Citations".
2. Click `Run`.

#### Result

- All citations are replaced by numbers: `[1]`, `[2]`, etc.
- A bibliography slide is created at the end.
- Identical sources receive the same number.

### Example

**Before:**

```
Slide 1: "Machine Learning [Smith2020] is important"
Slide 2: "Other studies [Jones2021] and [Smith2020] show..."
```

**After:**

```
Slide 1: "Machine Learning [1] is important"
Slide 2: "Other studies [2] and [1] show..."

Bibliography Slide:
[1] Smith, J. et al. (2020). Machine Learning Fundamentals. IEEE Transactions on Neural Networks.
[2] Jones, A. & Brown, B. (2021). Advanced AI Techniques in Modern Computing. Nature Machine Intelligence.
```

### Troubleshooting

**Issue**: "Unknown citation" warning in Debug window  
**Solution**: Add the missing source to `LoadSourceRegistry()`.

**Issue**: Macro does not work  
**Solution**: Ensure macros are enabled (`File` → `Options` → `Trust Center`).

**Issue**: Formatting is lost  
**Solution**: The macro only replaces the text; formatting should be preserved. If not, contact support.

### About & Support

This project is open source. I love to hear feedback and will implement improvements when possible.

If you find this tool helpful and want to support its development:
[Donate via PayPal](https://paypal.me/ThatDudeSebastian?locale.x=de_DE&country.x=DE)

---

## 🛠 Developer Guide

### Modifying the Source Code

The source code for the Add-In is located in the `src/` folder.

To modify the Add-In:

1.  Import the `.bas` files from `src/` (e.g., `ReferenceManager.bas`) into the PowerPoint VBA Editor (`Alt + F11`).
2.  Make your changes to the logic (parsing, reference management, etc.).
3.  Export the updated modules back to the `src/` folder to keep the repository updated.

### Customizations

#### Change Citation Format

In the `ProcessTextContent()` function, you can adjust the Regex pattern:

```vba
' Current: [SourceID]
regex.Pattern = "\[([A-Za-z0-9]+)\]"

' Alternative: (SourceID)
regex.Pattern = "\(([A-Za-z0-9]+)\)"
```

#### Use External Source File

You can also load sources from a CSV file. Example Code:

```vba
Private Sub LoadSourceRegistryFromFile()
    Dim filePath As String
    Dim fileNum As Integer
    Dim line As String
    Dim parts() As String

    filePath = "C:\Path\to\sources.csv"
    fileNum = FreeFile

    Open filePath For Input As fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        parts = Split(line, "|")

        If UBound(parts) >= 4 Then
            Call AddSource(parts(0), parts(1), parts(2), parts(3), parts(4))
        End If
    Loop
    Close fileNum
End Sub
```
