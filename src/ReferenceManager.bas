Attribute VB_Name = "ReferenceManager"
' ============================================================================
' PowerPoint Dynamic Reference Manager
' ============================================================================
' Automatically manages citations and generates bibliography slides
' Based on the UpdateReferences flowchart design
'
' Usage:
'   1. Add citations in your slides using format: [SourceID]
'      Example: "According to [Smith2020], machine learning..."
'   2. Run UpdateReferences macro
'   3. Citations will be replaced with numbers: [1], [2], etc.
'   4. Bibliography slide will be generated at the end
' ============================================================================

Option Explicit

' Global variables for reference tracking
Private referenceMap As Object      ' Dictionary: SourceID -> Reference Number
Private referenceCounter As Long    ' Counter for next reference number
Private sourceRegistry As Object    ' Dictionary: SourceID -> Source Details

' ============================================================================
' RIBBON CALLBACKS
' ============================================================================
Public Sub OnBtnUpdateReferences(control As IRibbonControl)
    Call UpdateReferences
End Sub

Public Sub OnBtnDonate(control As IRibbonControl)
    ActivePresentation.FollowHyperlink "https://paypal.me/ThatDudeSebastian?locale.x=de_DE&country.x=DE"
End Sub

Public Sub OnBtnResetHistory(control As IRibbonControl)
    Call ResetCitationHistory
End Sub

' ============================================================================
' MAIN ENTRY POINT
' ============================================================================
Public Sub UpdateReferences()
    On Error GoTo ErrorHandler
    
    Dim startTime As Double
    Dim errorLocation As String
    startTime = Timer
    
    ' Step 1: Load Source Registry
    errorLocation = "Loading Source Registry"
    Call LoadSourceRegistry
    
    ' Step 2: Initialize Reference Map and Counter
    errorLocation = "Initializing Reference Map"
    Set referenceMap = CreateObject("Scripting.Dictionary")
    
    ' Initialize counter based on existing tags to ensure new numbers don't conflict
    ' and to enable bibliography generation even if no new sources are added
    referenceCounter = GetNextAvailableNumber()
    
    ' Step 3-6: Iterate Slides and Process Citations
    errorLocation = "Processing Slides"
    Call ProcessAllSlides
    
    ' Step 7: Generate Bibliography Slide
    errorLocation = "Generating Bibliography"
    Call GenerateBibliographySlide
    
    ' Success message
    MsgBox "References updated successfully!" & vbCrLf & _
           "Total unique references: " & (referenceCounter - 1) & vbCrLf & _
           "Processing time: " & Format(Timer - startTime, "0.00") & " seconds", _
           vbInformation, "Update Complete"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error at: " & errorLocation & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical, "Error Details"
End Sub

' ============================================================================
' RESET MODULE
' ============================================================================
Public Sub ResetCitationHistory()
    Dim i As Long
    Dim count As Long
    Dim tagName As String
    
    If MsgBox("This will clear all stored reference numbers and allow fresh re-numbering." & vbCrLf & _
              "Are you sure?", vbYesNo + vbQuestion, "Reset History") = vbNo Then Exit Sub
              
    count = 0
    ' Delete all tags starting with BIB_REF_
    For i = ActivePresentation.Tags.Count To 1 Step -1
        tagName = ActivePresentation.Tags.Name(i)
        If Left(tagName, 8) = "BIB_REF_" Then
            ActivePresentation.Tags.Delete tagName
            count = count + 1
        End If
    Next i
    
    MsgBox "Reset complete. " & count & " stored references cleared." & vbCrLf & _
           "Run UpdateReferences to re-number everything from scratch.", vbInformation
End Sub

' ============================================================================
' SOURCE REGISTRY MODULE
' ============================================================================
Private Sub LoadSourceRegistry()
    Dim fd As FileDialog
    Dim selectedFile As String
    
    ' Initialize source registry dictionary
    Set sourceRegistry = CreateObject("Scripting.Dictionary")
    
    ' Create FileDialog object
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Select BibCoA/BibLaTeX File"
        .Filters.Clear
        .Filters.Add "BibTeX Files", "*.bib"
        .Filters.Add "All Files", "*.*"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            selectedFile = .SelectedItems(1)
            ' Load sources from the selected file
            Call LoadSourcesFromBibFile(selectedFile)
        Else
            MsgBox "No file selected. Using empty registry.", vbExclamation
        End If
    End With
End Sub

Private Sub LoadSourcesFromBibFile(ByVal filePath As String)
    Dim fileNum As Integer
    Dim line As String
    Dim currentSource As Object
    Dim currentID As String
    Dim fieldName As String
    Dim fieldValue As String
    Dim regexEntry As Object
    Dim regexField As Object
    Dim matches As Object
    
    fileNum = FreeFile
    Open filePath For Input As fileNum
    
    ' Setup Regex for detecting entries: @type{ID,
    Set regexEntry = CreateObject("VBScript.RegExp")
    regexEntry.Pattern = "^\s*@\w+\s*\{\s*([^,]+),"
    regexEntry.IgnoreCase = True
    
    ' Setup Regex for detecting fields: name = {value} or name = "value"
    Set regexField = CreateObject("VBScript.RegExp")
    regexField.Pattern = "^\s*([a-zA-Z0-9_]+)\s*=\s*[{\""']?(.*?)[}\""']?,?\s*$"
    regexField.IgnoreCase = True
    
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        
        ' Check for new entry start
        If regexEntry.Test(line) Then
            Set matches = regexEntry.Execute(line)
            currentID = matches(0).SubMatches(0)
            
            ' Save previous source if exists
            If Not currentSource Is Nothing And Len(currentID) > 0 Then
                ' Add basic fallback fields if missing
                If Not currentSource.Exists("Author") Then currentSource("Author") = "Unknown"
                If Not currentSource.Exists("Year") Then currentSource("Year") = "????"
                If Not currentSource.Exists("Title") Then currentSource("Title") = "No Title"
                If Not currentSource.Exists("Journal") Then currentSource("Journal") = ""
                
                Set sourceRegistry(currentSource("ID")) = currentSource
            End If
            
            ' Start new source
            Set currentSource = CreateObject("Scripting.Dictionary")
            currentSource("ID") = currentID
            
        ElseIf Not currentSource Is Nothing Then
            ' Parse fields
            If regexField.Test(line) Then
                Set matches = regexField.Execute(line)
                fieldName = LCase(matches(0).SubMatches(0))
                fieldValue = CleanBibTeXField(matches(0).SubMatches(1))
                
                ' Map BibTeX fields to our schema
                Select Case fieldName
                    Case "author"
                        currentSource("Author") = fieldValue
                    Case "year", "date"
                        ' Extract year from date if needed (simple check)
                        If Len(fieldValue) >= 4 Then
                            currentSource("Year") = Left(fieldValue, 4)
                        Else
                            currentSource("Year") = fieldValue
                        End If
                    Case "title"
                        currentSource("Title") = fieldValue
                    Case "journal", "booktitle", "publisher"
                        currentSource("Journal") = fieldValue
                End Select
            ElseIf InStr(line, "}") > 0 And Trim(line) = "}" Then
                ' End of entry (simple heuristic)
                If Not currentSource Is Nothing Then
                    Set sourceRegistry(currentSource("ID")) = currentSource
                    Set currentSource = Nothing
                End If
            End If
        End If
    Loop
    
    ' Add last source if file ended
    If Not currentSource Is Nothing Then
        Set sourceRegistry(currentSource("ID")) = currentSource
    End If
    
    Close fileNum
End Sub

Private Function CleanBibTeXField(ByVal text As String) As String
    ' Remove curly braces {} often used in BibTeX
    text = Replace(text, "{", "")
    text = Replace(text, "}", "")
    ' Remove quotes if remaining
    text = Replace(text, """", "")
    ' Basic cleanup
    CleanBibTeXField = Trim(text)
End Function

Private Sub AddSource(ByVal sourceID As String, _
                     ByVal author As String, _
                     ByVal year As String, _
                     ByVal title As String, _
                     ByVal journal As String)
    ' Create source object
    Dim source As Object
    Set source = CreateObject("Scripting.Dictionary")
    
    source("ID") = sourceID
    source("Author") = author
    source("Year") = year
    source("Title") = title
    source("Journal") = journal
    
    ' Add to registry
    Set sourceRegistry(sourceID) = source
End Sub

' ============================================================================
' CITATION PROCESSING MODULE
' ============================================================================
Private Sub ProcessAllSlides()
    Dim sld As Slide
    Dim slideIndex As Long
    
    ' Iterate through all slides in ascending order
    For slideIndex = 1 To ActivePresentation.Slides.Count
        Set sld = ActivePresentation.Slides(slideIndex)
        
        ' Skip the last slide if it's the bibliography (will be recreated)
        If slideIndex = ActivePresentation.Slides.Count And sld.Shapes.Count > 0 Then
            If sld.Shapes(1).HasTextFrame Then
                If sld.Shapes(1).TextFrame.HasText Then
                    If InStr(1, sld.Shapes(1).TextFrame.TextRange.Text, "References", vbTextCompare) > 0 Then
                        Exit For
                    End If
                End If
            End If
        End If
        
        ' Process all text shapes on this slide
        Call ProcessSlideShapes(sld)
    Next slideIndex
End Sub

Private Sub ProcessSlideShapes(ByVal sld As Slide)
    Dim shp As Shape
    
    ' Iterate through all shapes on the slide
    For Each shp In sld.Shapes
        ' Check if shape has text
        If shp.HasTextFrame Then
            If shp.TextFrame.HasText Then
                ' Process text content for citations
                Call ProcessTextContent(shp.TextFrame.TextRange)
            End If
        End If
    Next shp
End Sub

Private Sub ProcessTextContent(ByVal textRange As TextRange)
    Dim text As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim token As String
    Dim sourceID As String
    Dim refNumber As Long
    Dim usedSources As Object
    
    text = textRange.text
    
    ' Create regex object to find citations [SourceID] or [Number]
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = False
    regex.Pattern = "\[([A-Za-z0-9]+)\]"
    
    ' Find all citations in text
    Set matches = regex.Execute(text)
    
    ' Process each citation found
    For Each match In matches
        token = match.SubMatches(0)  ' Extract content tokens
        
        ' Case A: Token is a number (e.g. "1") - Existing Reference
        If IsNumeric(token) Then
            refNumber = CLng(token)
            ' Check if we have a mapping for this number (Reverse Lookup)
            sourceID = GetSourceIDByNumber(refNumber)
            
            If Len(sourceID) > 0 Then
                ' Confirm this ID exists in our registry
                If sourceRegistry.Exists(sourceID) Then
                    ' Mark this source as actively used in this run
                    If Not referenceMap.Exists(sourceID) Then
                        referenceMap(sourceID) = refNumber
                    End If
                End If
            End If
            
        ' Case B: Token is a Citation Key (e.g. "Smith2020") - New/Raw Reference
        ElseIf sourceRegistry.Exists(token) Then
            sourceID = token
            
            ' Check if we already assigned a number to this ID (Historical or Current)
            refNumber = GetStoredReferenceNumber(sourceID)
            
            If refNumber = 0 Then
                ' Truly new reference: Assign next available number
                refNumber = GetNextAvailableNumber()
                ' Store mapping persistently
                Call StoreReferenceNumber(sourceID, refNumber)
            End If
            
            ' Add to current execution map
            If Not referenceMap.Exists(sourceID) Then
                referenceMap(sourceID) = refNumber
            End If
            
            ' Replace placeholder with reference number
            text = Replace(text, "[" & sourceID & "]", "[" & refNumber & "]")
        Else
            ' Unknown citation
            Debug.Print "Warning: Unknown citation [" & token & "] found"
        End If
    Next match
    
    ' Update text if changes were made
    If text <> textRange.text Then
        textRange.text = text
    End If
End Sub

' ============================================================================
' PERSISTENCE HELPERS
' ============================================================================
Private Function GetStoredReferenceNumber(ByVal sourceID As String) As Long
    Dim tagName As String
    tagName = "BIB_REF_" & sourceID
    
    ' Check presentation tags
    Dim val As String
    val = ActivePresentation.Tags.Item(tagName)
    
    If Len(val) > 0 Then
        GetStoredReferenceNumber = CLng(val)
    Else
        GetStoredReferenceNumber = 0
    End If
End Function

Private Sub StoreReferenceNumber(ByVal sourceID As String, ByVal refNumber As Long)
    Dim tagName As String
    tagName = "BIB_REF_" & sourceID
    ActivePresentation.Tags.Add tagName, CStr(refNumber)
End Sub

Private Function GetSourceIDByNumber(ByVal refNum As Long) As String
    ' Reverse lookup: Find ID for a given number from Tags
    ' Note: This is inefficient but runs only for existing numbered citations
    Dim i As Long
    Dim tagName As String
    Dim tagVal As String
    Dim prefixLen As Long
    
    prefixLen = Len("BIB_REF_")
    
    For i = 1 To ActivePresentation.Tags.Count
        tagName = ActivePresentation.Tags.Name(i)
        If Left(tagName, prefixLen) = "BIB_REF_" Then
            tagVal = ActivePresentation.Tags.Value(i)
            If IsNumeric(tagVal) Then
                If CLng(tagVal) = refNum Then
                    GetSourceIDByNumber = Mid(tagName, prefixLen + 1)
                    Exit Function
                End If
            End If
        End If
    Next i
    
    GetSourceIDByNumber = ""
End Function

Private Function GetNextAvailableNumber() As Long
    ' Find the highest used number in Tags to ensure we don't reuse fields
    ' even if they were deleted from text but remain in Tags history
    Dim i As Long
    Dim maxNum As Long
    Dim tagVal As String
    Dim currentNum As Long
    
    maxNum = 0
    
    For i = 1 To ActivePresentation.Tags.Count
        If Left(ActivePresentation.Tags.Name(i), 8) = "BIB_REF_" Then
            tagVal = ActivePresentation.Tags.Value(i)
            If IsNumeric(tagVal) Then
                currentNum = CLng(tagVal)
                If currentNum > maxNum Then maxNum = currentNum
            End If
        End If
    Next i
    
    GetNextAvailableNumber = maxNum + 1
End Function

' ============================================================================
' BIBLIOGRAPHY GENERATOR MODULE
' ============================================================================
Private Sub GenerateBibliographySlide()
    Dim bibSlide As Slide
    Dim titleShape As Shape
    Dim textShape As Shape
    Dim bibText As String
    Dim sourceID As Variant
    Dim refNum As Long
    Dim source As Object
    Dim i As Long
    
    ' Remove existing bibliography slide if present
    Call RemoveExistingBibliography
    
    ' Check if we have any references to add
    ' Check if we have any references to add
    If referenceMap.Count = 0 Then
        ' No references found, don't create bibliography
        Exit Sub
    End If
    
    ' Create new blank slide at the end
    Set bibSlide = ActivePresentation.Slides.Add( _
        ActivePresentation.Slides.Count + 1, ppLayoutBlank)
    
    ' Create title shape manually
    Set titleShape = bibSlide.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, _
        50, 50, _
        ActivePresentation.PageSetup.SlideWidth - 100, 80)
    
    titleShape.TextFrame.TextRange.Text = "References"
    titleShape.TextFrame.TextRange.Font.Size = 44
    titleShape.TextFrame.TextRange.Font.Bold = True
    titleShape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
    
    ' Build bibliography text
    bibText = ""
    
    ' Create array of references sorted by number
    Dim refArray() As String
    Dim maxRefNum As Long
    
    ' Find maximum reference number used in this run
    maxRefNum = 0
    For Each sourceID In referenceMap.Keys
        If referenceMap(sourceID) > maxRefNum Then
            maxRefNum = referenceMap(sourceID)
        End If
    Next sourceID
    
    ' Check if we have any references
    If maxRefNum = 0 Then Exit Sub
    
    ReDim refArray(1 To maxRefNum)
    
    ' Populate array
    For Each sourceID In referenceMap.Keys
        refNum = referenceMap(sourceID)
        Set source = sourceRegistry(sourceID)
        
        ' Format: [1] Author (Year). Title. Journal.
        refArray(refNum) = "[" & refNum & "] " & _
                          source("Author") & " (" & source("Year") & "). " & _
                          source("Title") & ". " & _
                          source("Journal") & "."
    Next sourceID
    
    ' Combine into single text
    For i = 1 To UBound(refArray)
        If Len(refArray(i)) > 0 Then
            bibText = bibText & refArray(i) & vbCrLf & vbCrLf
        End If
    Next i
    
    ' Create text shape manually
    Set textShape = bibSlide.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, _
        50, 150, _
        ActivePresentation.PageSetup.SlideWidth - 100, _
        ActivePresentation.PageSetup.SlideHeight - 200)
    
    textShape.TextFrame.TextRange.Text = bibText
    textShape.TextFrame.TextRange.Font.Size = 14
    textShape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
    textShape.TextFrame.WordWrap = msoTrue
End Sub

Private Sub RemoveExistingBibliography()
    Dim sld As Slide
    Dim i As Long
    
    ' Check last few slides for bibliography
    For i = ActivePresentation.Slides.Count To 1 Step -1
        Set sld = ActivePresentation.Slides(i)
        
        ' Check if slide title contains "References" or "Bibliography"
        If sld.Shapes.Count > 0 Then
            If sld.Shapes(1).HasTextFrame Then
                If sld.Shapes(1).TextFrame.HasText Then
                    If InStr(1, sld.Shapes(1).TextFrame.TextRange.Text, "References", vbTextCompare) > 0 Or _
                       InStr(1, sld.Shapes(1).TextFrame.TextRange.Text, "Bibliography", vbTextCompare) > 0 Then
                        sld.Delete
                        Exit For
                    End If
                End If
            End If
        End If
        
        ' Only check last 3 slides
        If ActivePresentation.Slides.Count - i > 2 Then Exit For
    Next i
End Sub