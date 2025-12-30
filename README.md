# PowerPoint VBA Reference Manager - User Guide

## Installation

1. **Open PowerPoint** and load your presentation.
2. **Open VBA Editor**: Press `Alt + F11`.
3. **Import Module**:
   - `File` → `Import File...`
   - Select `ReferenceManager.bas`
   - Click `Open`

## Usage

### Step 1: Load Sources

1. When you start the macro, a file dialog will open automatically.
2. Select your `.bib` file (e.g., exported from Zotero/Citavi).
3. Alternatively, you can use the included `test_references.bib` for testing.

> **Note**: If you do not select a file, the bibliography will remain empty.

### Step 2: Insert Citations into Slides

Insert citations in your PowerPoint slides using the format `[Key]`, where `Key` corresponds to the citation key in your BibTeX file:

**Example:**
```
"According to [Smith2020], machine learning is..."
```

### Step 3: Run Macro

1. Press `Alt + F8`.
2. Select `UpdateReferences`.
3. Click `Run`.

### Result

- All citations are replaced by numbers: `[1]`, `[2]`, etc.
- A bibliography slide is created at the end.
- Identical sources receive the same number.

## Example

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

## Customizations

### Change Citation Format

In the `ProcessTextContent()` function, you can adjust the Regex pattern:

```vba
' Current: [SourceID]
regex.Pattern = "\[([A-Za-z0-9]+)\]"

' Alternative: (SourceID)
regex.Pattern = "\(([A-Za-z0-9]+)\)"
```

### Use External Source File

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

## Troubleshooting

**Issue**: "Unknown citation" warning in Debug window  
**Solution**: Add the missing source to `LoadSourceRegistry()`.

**Issue**: Macro does not work  
**Solution**: Ensure macros are enabled (`File` → `Options` → `Trust Center`).

**Issue**: Formatting is lost  
**Solution**: The macro only replaces the text; formatting should be preserved. If not, contact support.
