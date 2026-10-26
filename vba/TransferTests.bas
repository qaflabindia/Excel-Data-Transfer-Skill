Attribute VB_Name = "TransferTests"
Option Explicit
Private root As String, configRow As Long, checks As Long

Public Sub RunTransferTests()
    Dim a As Workbook, b As Workbook, output As Workbook, c As Worksheet, w As Worksheet
    Dim oldPath As String, templatePath As String, outPath As String, rerunPath As String
    Dim priorEvents As Boolean, oldBytes As Long, i As Long, j As Long
    On Error GoTo Failed
    checks = 0
    priorEvents = Application.EnableEvents
    root = Environ$("TEMP") & "\ExcelTransferTests_" & Format$(Now, "yyyymmdd_hhnnss")
    MkDir root
    oldPath = root & "\OLD.xlsx": templatePath = root & "\NEW.xlsx"
    outPath = root & "\Result.xlsx": rerunPath = root & "\Rerun.xlsx"
    SetupTransfer
    Set c = ThisWorkbook.Worksheets("Transfer Config")
    c.Range("A11:N12").ClearContents: configRow = 11
    Set a = Workbooks.Add(xlWBATWorksheet): a.Worksheets(1).Name = "Data"
    Set b = Workbooks.Add(xlWBATWorksheet): b.Worksheets(1).Name = "Data"
    SimpleRegion a.Worksheets("Data"), "A1", False
    SimpleRegion b.Worksheets("Data"), "A1", True
    'Second region repeats row/column labels, but has an independent stable identity.
    SimpleRegion a.Worksheets("Data"), "F1", False
    SimpleRegion b.Worksheets("Data"), "H10", True
    AddPair "Data", "main", "A1:D6", "A1:D6", "B2:D6", "B2:D6", 1, 1
    AddPair "Data", "second", "F1:I6", "H10:K15", "G2:I6", "I11:K15", 1, 1
    With a.Worksheets("Data")
        .Range("A6").Value2 = "SourceOnly"
        .Range("B2").ClearContents
        .Range("C2").Value2 = 0
        .Range("D2").Value2 = "'00123"
        .Range("B3").Value2 = CDbl(DateSerial(2026, 10, 26))
        .Range("B3").NumberFormat = "yyyy-mm-dd"
        .Range("C3").Value2 = True
        .Range("D3").Value2 = CVErr(xlErrNA)
    End With
    With b.Worksheets("Data")
        .Range("A6").Value2 = "DestinationOnly"
        .Range("B4").Locked = True
        .Range("C4").Formula = "=40+2"
        .Range("D4").Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="0", Formula2:="1"
        .Protect Password:="fixture", UserInterfaceOnly:=False
    End With
    'Duplicate within one region: entire region is unresolved, never first-match wins.
    Set w = AddSheet(a, "Duplicate"): SimpleRegion w, "A1", False: w.Range("A3").Value2 = w.Range("A2").Value2
    Set w = AddSheet(b, "Duplicate"): SimpleRegion w, "A1", False
    AddPair "Duplicate", "dup", "A1:D6", "A1:D6", "B2:D6", "B2:D6", 1, 1
    'Merged hierarchical headings and two-column row identifiers.
    Hierarchy AddSheet(a, "Hierarchy"), False
    Hierarchy AddSheet(b, "Hierarchy"), True
    AddPair "Hierarchy", "hier", "A1:F5", "A1:F5", "C3:F5", "C3:F5", 2, 2
    'Colour fallback: title row makes positional interpretation unreliable.
    ColourFixture AddSheet(a, "Colour")
    ColourFixture AddSheet(b, "Colour")
    AddPair "Colour", "colour", "A1:D5", "A1:D5", "B3:D5", "B3:D5", 0, 0
    For i = configRow - 2 To configRow - 1
        c.Cells(i, 9).Value2 = RGB(180, 210, 240): c.Cells(i, 13).Value2 = "AUTO"
    Next i
    'Missing region and sheet are deliberate coverage exceptions.
    SimpleRegion AddSheet(a, "OnlyOld"), "A1", False
    AddConfig "OLD", "OnlyOld", "orphan", "A1:D6", "B2:D6", 1, 1
    AddConfig "NEW", "AbsentSheet", "absent", "A1:D6", "B2:D6", 1, 1
    a.SaveAs Filename:=oldPath, FileFormat:=xlOpenXMLWorkbook
    b.SaveAs Filename:=templatePath, FileFormat:=xlOpenXMLWorkbook
    a.Close False: b.Close False
    oldBytes = FileLen(oldPath)
    c.Cells(2, 2).Value2 = oldPath: c.Cells(3, 2).Value2 = templatePath: c.Cells(4, 2).Value2 = outPath
    c.Cells(5, 2).Value2 = True
    TransferTestMode = True
    RunTransfer
    Set output = OpenResult(outPath)
    Assert Not output.Worksheets("Data").Range("C3").HasFormula, "Dry run did not write"
    Assert LogHas(output, "missing worksheet"), "Missing worksheet logged"
    Assert LogHas(output, "duplicate header"), "Duplicate header logged"
    output.Close SaveChanges:=True
    c.Cells(4, 2).Value2 = root & "\Executed.xlsx": c.Cells(5, 2).Value2 = False
    'Inject one recoverable write failure through a controller-only named cell.
    c.Range("T1").Value2 = "Data!$I$11"
    ThisWorkbook.Names.Add Name:="TRANSFER_TEST_FAIL", RefersTo:="='Transfer Config'!$T$1"
    RunTransfer
    Set output = OpenResult(CStr(c.Cells(4, 2).Value2))
    With output.Worksheets("Data")
        Assert .Range("C3").HasFormula, "Reordered Alpha/Jan linked"
        Assert .Range("C3").Value2 = "", "Physical source blank displays empty string"
        Assert Not IsEmpty(.Range("C3").Value2), "Link is not physically empty"
        Assert .Range("D3").Value2 = 0, "Zero preserved"
        Assert VarType(.Range("B3").Value2) = vbString And .Range("B3").Value2 = "00123", "Leading zero text preserved"
        Assert .Range("C5").Value2 = CDbl(DateSerial(2026, 10, 26)), "Date serial preserved"
        Assert VarType(.Range("D5").Value2) = vbBoolean, "Boolean preserved"
        Assert .Range("C4").Formula = "=40+2", "Business formula preserved"
        Assert .Range("B4").Value2 = 99, "Locked constant preserved"
        Assert .ProtectContents, "Protection remains enabled"
        Assert .Range("D4").Validation.Type = xlValidateWholeNumber, "Validation preserved"
        Assert .Range("I11").Value2 = 99, "Failed write restored"
        Assert .Range("J12").HasFormula, "Other region continues after failure"
    End With
    Assert LogHas(output, "missing source"), "Destination reconciliation logged"
    Assert LogHas(output, "missing destination"), "Source reconciliation logged"
    Assert LogHas(output, "source error"), "Source error flagged"
    Assert LogHas(output, "locked destination"), "Locked destination logged"
    Assert LogHas(output, "existing formula"), "Existing formula logged"
    Assert LogHas(output, "invalid value"), "Validation rejection logged"
    Assert LogHas(output, "write failure"), "Recoverable failure logged"
    Assert output.Worksheets("Hierarchy").Range("C1:D1").MergeCells, "Merged heading preserved"
    Assert output.Worksheets("Hierarchy").Range("C3").HasFormula, "Hierarchy transferred"
    Assert output.Worksheets("Colour").Range("B3").HasFormula, "Colour fallback transferred"
    Assert FileLen(oldPath) = oldBytes, "OLD file byte length unchanged"
    output.Save: output.Close False
    ThisWorkbook.Names("TRANSFER_TEST_FAIL").Delete
    'Safe rerun: use prior output as NEW template; ownership travels with internal audit.
    c.Cells(3, 2).Value2 = c.Cells(4, 2).Value2: c.Cells(4, 2).Value2 = rerunPath
    RunTransfer
    Set output = OpenResult(rerunPath)
    Assert SuccessHas(output, "already correct"), "Rerun recognizes audited links"
    Assert output.Worksheets("Data").Range("C4").Formula = "=40+2", "Rerun preserves unrelated formula"
    output.Close True
    'Blocked structural additions: external audit must be selected, never unprotect.
    Set b = Workbooks.Open(templatePath, UpdateLinks:=0)
    b.Protect Password:="structure", Structure:=True
    b.SaveAs Filename:=root & "\Blocked.xlsx", FileFormat:=xlOpenXMLWorkbook
    b.Close False
    c.Cells(3, 2).Value2 = root & "\Blocked.xlsx": c.Cells(4, 2).Value2 = root & "\BlockedResult.xlsx"
    c.Cells(8, 2).Value2 = root & "\ExternalAudit.xlsx"
    RunTransfer
    Set output = OpenResult(CStr(c.Cells(4, 2).Value2))
    Assert output.ProtectStructure, "Structure protection retained"
    Assert Dir$(CStr(c.Cells(8, 2).Value2)) <> "", "External audit exists"
    output.Close True
    'Values-only semantics tested independently on a fresh copy.
    c.Cells(3, 2).Value2 = templatePath: c.Cells(4, 2).Value2 = root & "\Values.xlsx"
    c.Cells(8, 2).ClearContents: c.Cells(6, 2).Value2 = True
    RunTransfer
    Set output = OpenResult(CStr(c.Cells(4, 2).Value2))
    Assert IsEmpty(output.Worksheets("Data").Range("C3").Value2), "Values-only blank physically empty"
    Assert output.Worksheets("Data").Range("B3").Value2 = "00123", "Values-only text preserved"
    Assert Not output.Worksheets("Data").Range("B3").HasFormula, "Values-only has no new formula"
    output.Close True
    TransferTestMode = False
    Assert Application.EnableEvents = priorEvents, "Application events restored"
    MsgBox CStr(checks) & " assertions passed. Fixtures: " & root, vbInformation
    Exit Sub
Failed:
    TransferTestMode = False
    MsgBox "Test failure: " & Err.Description & vbCrLf & "Fixtures retained: " & root, vbCritical
End Sub
Private Function AddSheet(ByVal wb As Workbook, ByVal name As String) As Worksheet
    Set AddSheet = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    AddSheet.Name = name
End Function
Private Sub SimpleRegion(ByVal w As Worksheet, ByVal addr As String, ByVal reordered As Boolean)
    Dim a As Range, i As Long, j As Long, rows As Variant, cols As Variant
    Set a = w.Range(addr).Resize(6, 4)
    If reordered Then
        rows = Array("Gamma", "Alpha", "Delta", "Beta", "Extra")
        cols = Array("Mar", "Jan", "Feb")
    Else
        rows = Array("Alpha", "Beta", "Gamma", "Delta", "Extra")
        cols = Array("Jan", "Feb", "Mar")
    End If
    a.Cells(1, 1).Value2 = "Item"
    For j = 0 To 2: a.Cells(1, j + 2).Value2 = cols(j): Next j
    For i = 0 To 4
        a.Cells(i + 2, 1).Value2 = rows(i)
        For j = 2 To 4: a.Cells(i + 2, j).Value2 = IIf(reordered, 99, (i + 1) * 10 + j): Next j
    Next i
    a.Offset(1, 1).Resize(5, 3).Locked = False
    a.Font.Name = "Calibri": a.Interior.Color = RGB(245, 245, 245)
End Sub
Private Sub Hierarchy(ByVal w As Worksheet, ByVal dest As Boolean)
    Dim i As Long
    w.Range("C1:D1").Merge: w.Range("C1").Value2 = "FY26"
    w.Range("E1:F1").Merge: w.Range("E1").Value2 = "FY27"
    w.Range("C2:F2").Value = Array("Units", "USD", "Units", "USD")
    For i = 3 To 5
        w.Cells(i, 1).Value2 = "North": w.Cells(i, 2).Value2 = "Store" & CStr(i)
    Next i
    w.Range("C3:F5").Value2 = IIf(dest, 99, 10)
    w.Range("C3:F5").Locked = False
End Sub
Private Sub ColourFixture(ByVal w As Worksheet)
    SimpleRegion w, "A1", False
    w.Range("A1:D6").Clear
    w.Range("A1").Value2 = "Budget title"
    w.Range("A2:D2").Value = Array("Item", "Jan", "Feb", "Mar")
    w.Range("A3").Value2 = "A": w.Range("A4").Value2 = "B": w.Range("A5").Value2 = "C"
    w.Range("B3:D5").Value2 = 7: w.Range("B3:D5").Locked = False
    w.Range("A2:D2").Interior.Color = RGB(180, 210, 240)
    w.Range("A3:A5").Interior.Color = RGB(180, 210, 240)
End Sub
Private Sub AddPair(ByVal sheet As String, ByVal id As String, ByVal oldArea As String, ByVal newArea As String, ByVal oldInputs As String, ByVal newInputs As String, ByVal hr As Long, ByVal hc As Long)
    AddConfig "OLD", sheet, id, oldArea, oldInputs, hr, hc
    AddConfig "NEW", sheet, id, newArea, newInputs, hr, hc
End Sub
Private Sub AddConfig(ByVal side As String, ByVal sheet As String, ByVal id As String, ByVal area As String, ByVal inputs As String, ByVal hr As Long, ByVal hc As Long)
    Dim c As Worksheet
    Set c = ThisWorkbook.Worksheets("Transfer Config")
    c.Cells(configRow, 1).Resize(1, 14).Value = Array(side, sheet, id, area, hr, hc, inputs, id, "", False, True, True, "EXPLICIT", "fixture")
    configRow = configRow + 1
End Sub
Private Sub Assert(ByVal condition As Boolean, ByVal description As String)
    If Not condition Then Err.Raise vbObjectError + 500, , description
    checks = checks + 1
End Sub
Private Function OpenResult(ByVal path As String) As Workbook
    Dim w As Workbook
    For Each w In Application.Workbooks
        If StrComp(w.FullName, path, vbTextCompare) = 0 Then Set OpenResult = w: Exit Function
    Next w
    Set OpenResult = Workbooks.Open(path, UpdateLinks:=0)
End Function
Private Function LogHas(ByVal wb As Workbook, ByVal category As String) As Boolean
    Dim w As Worksheet, i As Long
    Set w = wb.Worksheets("missing source to link")
    For i = 2 To w.Cells(w.Rows.Count, 1).End(xlUp).Row
        If w.Cells(i, 17).Value2 = category Then LogHas = True: Exit Function
    Next i
End Function
Private Function SuccessHas(ByVal wb As Workbook, ByVal outcome As String) As Boolean
    Dim w As Worksheet, i As Long
    Set w = wb.Worksheets("transfer audit")
    For i = 2 To w.Cells(w.Rows.Count, 1).End(xlUp).Row
        If w.Cells(i, 9).Value2 = outcome Then SuccessHas = True: Exit Function
    Next i
End Function
