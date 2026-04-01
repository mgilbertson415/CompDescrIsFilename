Attribute VB_Name = "CompDescrIsFilename1"
Dim swApp           As SldWorks.SldWorks
Dim bDryRun         As Boolean ' True means "do not edit anything, just report"
Dim sDir            As String
Dim sOutputFilename As String

Sub main()

    ' purpose: open all parts and assemblies in a given folder, and set the description property to the filename without the extension
    ' original use case: ILSCO legacy Solidworks data slated for Enovia migration had description issues, and this was the simplest description info to use

    Dim sSomeFile   As String
    Dim vDryRun     As Variant
    
    Set swApp = Application.SldWorks
    InitializeDocumentManager
    
    sOutputFilename = swApp.GetCurrentMacroPathFolder & "\CompDescrIsFilename.csv"
    InitializeOutputFile sOutputFilename
    ' write header line
    WriteToFile sOutputFilename, "Component,Result"
    
    vDryRun = MsgBox("Use 'Dry Run' mode, where no changes are actually saved?", vbYesNo)
    bDryRun = IIf(vDryRun = vbYes, True, False)
    sDir = BrowseButton("Browse to folder containing Solidworks files to adjust:", "17")
    
    ' loop through all drawings
    sSomeFile = Dir(sDir & "\*.sld*")
    Do While sSomeFile <> ""
        If Right(UCase(sSomeFile), 3) <> "DRW" Then ProcessComp sDir & "\" & sSomeFile
        sSomeFile = Dir
    Loop
    
    MsgBox "All done! See execution details in " & sOutputFilename
    
End Sub

Sub ProcessComp(sCompFilename As String)
    
    Dim dmDoc           As SwDMDocument29
    Dim sFilenameShort  As String ' to use as description value
    Dim i               As Integer
    Dim bFoundDescrProp As Boolean
    Dim vPropNames      As Variant
    
    Set dmDoc = dmApp.GetDocument(sCompFilename, swDmDocumentUnknown, False, Empty)
    If dmDoc Is Nothing Then
        sResult = "file could not be opened"
        WriteToFile sOutputFilename, sCompFilename & "," & sResult
        Exit Sub
    End If
    
    bFoundDescrProp = False
    dmDoc.GetAllCustomPropertyNamesAndValues vPropNames, Empty, Empty, Empty
    If Not IsEmpty(vPropNames) Then
        For i = 0 To UBound(vPropNames)
            If UCase(vPropNames(i)) = "DESCRIPTION" Then
                bFoundDescrProp = True
                Exit For
            End If
        Next i
    End If
    
    ' prepare value
    sFilenameShort = Trim(Mid(sCompFilename, InStrRev(sCompFilename, "\") + 1, InStrRev(sCompFilename, ".") - InStrRev(sCompFilename, "\") - 1))
    
    ' write value, either into existing property or on a new one
    If bFoundDescrProp Then
        dmDoc.SetCustomProperty2 "Description", sFilenameShort
        sResult = "found Description property - set value to " & Chr(34) & sFilenameShort & Chr(34)
    Else
        dmDoc.AddCustomProperty "Description", swDmCustomInfoText, sFilenameShort
        sResult = "Description property missing - added it and initialized to " & Chr(34) & sFilenameShort & Chr(34)
    End If
    If bDryRun = False Then dmDoc.Save
    
    dmDoc.CloseDoc
    WriteToFile sOutputFilename, sCompFilename & "," & sResult
    
End Sub
