Attribute VB_Name = "Transfer"
Option Explicit

'Late-bound dictionaries: no reference installation required. Desktop Excel for Windows.
Private Const EX As String = "missing source to link"
Private Const AU As String = "transfer audit"
Private Const IV As String = "region inventory"
Private Const PL As String = "transfer plan"
Private Const SU As String = "transfer summary"
Private cfg As Worksheet, auditBook As Workbook, oldBook As Workbook, newBook As Workbook
Private runID As String, caseSensitive As Boolean, valuesOnly As Boolean
Private exceptions As Long, unknowns As Long, successes As Long, already As Long
Private srcCount As Long, dstCount As Long, dry As Boolean
Private outcomes As Object, targets As Object, owned As Object, aliases As Object
Private snapshots As Object, plans As Collection, logsReady As Boolean
Private oldEvents As Boolean, oldCalc As XlCalculation, oldScreen As Boolean
Private oldSecurity As Long, stateSaved As Boolean, writing As Boolean
Private oldAutoFill As Boolean
Public TransferTestMode As Boolean
Private fatal As Boolean, sourceRegions As Collection, destRegions As Collection

Private Function Dict() As Object
    Set Dict = CreateObject("Scripting.Dictionary")
    Dict.CompareMode = vbBinaryCompare
End Function

Public Sub SetupTransfer()
    Dim w As Worksheet
    Set w = GetOrAdd(ThisWorkbook, "Transfer Config")
    w.Cells.Clear
    w.Range("A1:B8").Value = Array2D(Array( _
        Array("Setting", "Value"), Array("OLD path", ""), Array("NEW template path", ""), _
        Array("Output path (must not exist)", ""), Array("Dry run", True), _
        Array("Values only", False), Array("Case sensitive", False), _
        Array("External audit path (optional; must not exist)", "")))
    w.Range("A10:N10").Value = Array("Side", "Worksheet", "Stable region ID", "Region address/name/table", _
        "Column header rows", "Row header columns", "Input address", "Section identity", _
        "Header fill RGB integer", "Use DisplayFormat", "Include hidden", "Include sheet", "Header mode", "Note")
    w.Range("A11:N11").Value = Array("OLD", "Sheet1", "Budget", "A1:D10", 1, 1, "B2:D10", "Budget", "", False, True, True, "EXPLICIT", "Replace example")
    w.Range("A12:N12").Value = Array("NEW", "Sheet1", "Budget", "A1:D10", 1, 1, "B2:D10", "Budget", "", False, True, True, "EXPLICIT", "Replace example")
    w.Range("P10:R10").Value = Array("Alias axis (R/C)", "Raw label", "Approved canonical label")
    w.Columns("A:R").AutoFit
End Sub

Private Function Array2D(ByVal rows As Variant) As Variant
    Dim a() As Variant, i As Long, j As Long
    ReDim a(1 To UBound(rows) + 1, 1 To UBound(rows(0)) + 1)
    For i = 0 To UBound(rows)
        For j = 0 To UBound(rows(i)): a(i + 1, j + 1) = rows(i)(j): Next j
    Next i
    Array2D = a
End Function

Public Sub SelectWorkbooks()
    Dim p As Variant, w As Worksheet
    Set w = ThisWorkbook.Worksheets("Transfer Config")
    p = Application.GetOpenFilename("Excel files,*.xls*")
    If VarType(p) = vbBoolean Then Exit Sub
    w.Cells(2, 2).Value2 = CStr(p)
    p = Application.GetOpenFilename("Excel files,*.xls*")
    If VarType(p) = vbBoolean Then Exit Sub
    w.Cells(3, 2).Value2 = CStr(p)
    p = Application.GetSaveAsFilename(InitialFileName:="Transferred.xlsx", FileFilter:="Excel workbook,*.xlsx")
    If VarType(p) <> vbBoolean Then w.Cells(4, 2).Value2 = CStr(p)
End Sub

Public Sub RunTransfer()
    Dim status As String, p As Variant, k As Variant, n As Long, msg As String
    On Error GoTo Failed
    ResetRun
    Set cfg = ThisWorkbook.Worksheets("Transfer Config")
    dry = CBool(cfg.Cells(5, 2).Value2): valuesOnly = CBool(cfg.Cells(6, 2).Value2)
    caseSensitive = CBool(cfg.Cells(7, 2).Value2)
    SaveApplicationState
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.AutoCorrect.AutoFillFormulasInLists = False
    Application.AutomationSecurity = 3 'msoAutomationSecurityForceDisable
    ValidatePaths
    Set oldBook = Workbooks.Open(CStr(cfg.Cells(2, 2).Value2), UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    'Copy bytes; never open the template writable, never convert its file type.
    FileCopy CStr(cfg.Cells(3, 2).Value2), CStr(cfg.Cells(4, 2).Value2)
    Set newBook = Workbooks.Open(CStr(cfg.Cells(4, 2).Value2), UpdateLinks:=0, ReadOnly:=False, AddToMru:=False)
    If newBook.ReadOnly Then Err.Raise vbObjectError + 1, , "Output is read-only"
    PrepareAudit
    LoadAliases
    LoadOwnership
    Set sourceRegions = DiscoverDataRegions(oldBook, "OLD")
    Set destRegions = DiscoverDataRegions(newBook, "NEW")
    SnapshotWorkbook
    BuildDryRun
    LogSummary "Preflight complete"
    auditBook.Save
    If Not dry Then
        writing = True
        For Each p In plans
            ExecuteMapping p
        Next p
        writing = False
    End If
    VerifyPreservation
    Reconcile
    If dry Then
        status = "Dry run only; no input cells written"
    ElseIf exceptions = 0 And unknowns = 0 Then
        status = "Completed and verified"
    Else
        status = "Completed with exceptions" & ChrW(&H2014) & "review 'missing source to link'"
    End If
    LogSummary status
    newBook.Save
    If Not auditBook Is newBook Then auditBook.Save
    GoTo Finish
Failed:
    n = Err.Number: msg = Err.Description: fatal = True
    status = "Stopped" & ChrW(&H2014) & "partial output; review audit log"
    If logsReady Then EmergencyLog n, msg, status
Finish:
    CloseSource
    RestoreApplicationState
    If Not TransferTestMode Then MsgBox status & vbCrLf & IIf(fatal, msg, "Run ID: " & runID), IIf(fatal, vbExclamation, vbInformation)
    If TransferTestMode And fatal Then Err.Raise vbObjectError + 99, , msg
End Sub

Private Sub ResetRun()
    runID = Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(Timer * 100, "00000000")
    Set outcomes = Dict(): Set targets = Dict(): Set owned = Dict(): Set aliases = Dict()
    Set snapshots = Dict(): Set plans = New Collection
    exceptions = 0: unknowns = 0: successes = 0: already = 0: srcCount = 0: dstCount = 0
    logsReady = False: fatal = False: stateSaved = False: writing = False
    Set auditBook = Nothing: Set oldBook = Nothing: Set newBook = Nothing
End Sub

Private Sub ValidatePaths()
    Dim a As String, b As String, c As String, w As Workbook
    a = CStr(cfg.Cells(2, 2).Value2): b = CStr(cfg.Cells(3, 2).Value2): c = CStr(cfg.Cells(4, 2).Value2)
    If a = "" Or b = "" Or c = "" Then Err.Raise 5, , "Set all workbook paths"
    If Dir$(a) = "" Or Dir$(b) = "" Then Err.Raise 53, , "Input file missing"
    If Dir$(c) <> "" Then Err.Raise 58, , "Output already exists; choose a fresh copy path"
    If LCase$(Mid$(b, InStrRev(b, "."))) <> LCase$(Mid$(c, InStrRev(c, "."))) Then Err.Raise 5, , "Output extension must match template"
    For Each w In Application.Workbooks
        If StrComp(w.FullName, a, vbTextCompare) = 0 Or StrComp(w.FullName, b, vbTextCompare) = 0 Then Err.Raise 5, , "Close OLD and template before starting"
    Next w
End Sub

Private Sub SaveApplicationState()
    oldEvents = Application.EnableEvents: oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating: oldSecurity = Application.AutomationSecurity
    oldAutoFill = Application.AutoCorrect.AutoFillFormulasInLists
    stateSaved = True
End Sub
Private Sub RestoreApplicationState()
    If Not stateSaved Then Exit Sub
    Application.AutoCorrect.AutoFillFormulasInLists = oldAutoFill
    Application.AutomationSecurity = oldSecurity
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    stateSaved = False
End Sub
Private Sub CloseSource()
    If oldBook Is Nothing Then Exit Sub
    oldBook.Close SaveChanges:=False
End Sub

Private Function GetSheet(ByVal wb As Workbook, ByVal name As String) As Worksheet
    Dim w As Worksheet
    For Each w In wb.Worksheets
        If StrComp(w.Name, name, vbTextCompare) = 0 Then Set GetSheet = w: Exit Function
    Next w
End Function
Private Function GetOrAdd(ByVal wb As Workbook, ByVal name As String) As Worksheet
    Set GetOrAdd = GetSheet(wb, name)
    If Not GetOrAdd Is Nothing Then Exit Function
    If wb.ProtectStructure Then Err.Raise 1004, , "Structure protection prevents adding " & name
    Set GetOrAdd = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    GetOrAdd.Name = name
End Function
Private Function AuditName(ByVal s As String) As Boolean
    AuditName = (s = EX Or s = AU Or s = IV Or s = PL Or s = SU)
End Function
Private Function CanLogIn(ByVal wb As Workbook) As Boolean
    Dim v As Variant, w As Worksheet
    CanLogIn = False
    For Each v In Array(EX, AU, IV, PL, SU)
        Set w = GetSheet(wb, CStr(v))
        If w Is Nothing Then
            If wb.ProtectStructure Then Exit Function
        Else
            'Conservative: a protected audit sheet is unsupported, even if some cells are unlocked.
            If w.ProtectContents Then Exit Function
        End If
    Next v
    CanLogIn = True
End Function
Private Sub PrepareAudit()
    Dim v As Variant, w As Worksheet, path As String, headers As Variant, j As Long
    If CanLogIn(newBook) Then
        Set auditBook = newBook
    Else
        path = CStr(cfg.Cells(8, 2).Value2)
        If path = "" Then path = newBook.FullName & ".audit.xlsx"
        If Dir$(path) <> "" Then Err.Raise 58, , "External audit path already exists"
        Set auditBook = Workbooks.Add(xlWBATWorksheet)
        auditBook.SaveAs Filename:=path, FileFormat:=xlOpenXMLWorkbook
    End If
    For Each v In Array(EX, AU, IV, PL, SU)
        Set w = GetOrAdd(auditBook, CStr(v))
        If w.ProtectContents Then Err.Raise 1004, , "Audit sheet is protected"
        If w.Cells(1, 1).Value2 <> "" And w.Cells(1, 1).Value2 <> "Run ID" Then Err.Raise 5, , "Unrecognised audit schema: " & CStr(v)
        Select Case CStr(v)
            Case EX
                headers = Array("Run ID", "Timestamp", "Stage", "Source workbook", "Source sheet", "Source region", "Source address", "Destination workbook", "Destination sheet", "Destination region", "Destination address", "Raw row", "Raw column", "Row key", "Column key", "Header method", "Category", "Error number", "Description", "Content outcome", "Corrective action", "Coverage")
            Case AU
                headers = Array("Run ID", "Timestamp", "Destination sheet", "Destination address", "Source full path", "Source sheet", "Source address", "Formula", "Outcome", "Mode")
            Case IV
                headers = Array("Run ID", "Side", "Worksheet", "Region ID", "Address", "Rows", "Columns", "Discovery", "Column headers", "Row headers", "Status")
            Case PL
                headers = Array("Run ID", "Source sheet", "Source address", "Destination sheet", "Destination address", "Row key", "Column key", "Proposed formula/mode")
            Case SU
                headers = Array("Run ID", "Metric", "Value")
        End Select
        For j = 0 To UBound(headers)
            If w.Cells(1, 1).Value2 <> "" Then
                If CStr(w.Cells(1, j + 1).Value2) <> CStr(headers(j)) Then Err.Raise 5, , "Audit schema mismatch: " & CStr(v)
            End If
        Next j
        If w.Cells(1, 1).Value2 = "" Then
            For j = 0 To UBound(headers): w.Cells(1, j + 1).Value2 = headers(j): Next j
        End If
    Next v
    logsReady = True
End Sub
Private Sub AppendRow(ByVal sheetName As String, ByVal vals As Variant)
    Dim w As Worksheet, r As Long, i As Long
    Set w = auditBook.Worksheets(sheetName)
    r = w.Cells(w.Rows.Count, 1).End(xlUp).Row + 1
    If r > w.Rows.Count Then Err.Raise 1004, , "Audit full"
    'Force text: labels/formulas must not execute in the log.
    For i = 0 To UBound(vals)
        w.Cells(r, i + 1).NumberFormat = "@"
        w.Cells(r, i + 1).Value2 = CStr(vals(i))
    Next i
End Sub
Private Sub Issue(ByVal stage As String, ByVal cat As String, ByVal s As Object, ByVal d As Object, _
    ByVal sc As Range, ByVal dc As Range, ByVal detail As String, Optional ByVal errNo As Long = 0, _
    Optional ByVal changed As String = "unchanged", Optional ByVal coverage As String = "known")
    Dim sw As String, sr As String, sa As String, dw As String, dr As String, da As String
    Dim rr As String, cc As String, rk As String, ck As String, method As String
    exceptions = exceptions + 1
    If coverage = "unknown" Then unknowns = unknowns + 1
    If Not s Is Nothing Then sw = s("sheet"): sr = s("id"): method = s("method")
    If Not d Is Nothing Then dw = d("sheet"): dr = d("id")
    If Not sc Is Nothing Then
        sa = sc.Address
        If Not s Is Nothing Then
            If s("valid") Then GetCellHeaderKeys s, sc.Row, sc.Column, rr, cc, rk, ck
        End If
    End If
    If Not dc Is Nothing Then da = dc.Address
    AppendRow EX, Array(runID, Format$(Now, "yyyy-mm-dd hh:nn:ss"), stage, oldBook.FullName, sw, sr, sa, newBook.FullName, dw, dr, da, rr, cc, rk, ck, method, cat, CStr(errNo), detail, changed, "Review configuration, headers, eligibility or validation; rerun on a fresh copy", coverage)
End Sub
Private Sub EmergencyLog(ByVal n As Long, ByVal msg As String, ByVal status As String)
    On Error GoTo CannotLog
    Issue "Fatal", "write failure", Nothing, Nothing, Nothing, Nothing, msg, n, "see prior cell outcomes", "unknown"
    LogSummary status
    auditBook.Save
    Exit Sub
CannotLog:
    MsgBox "Reliable audit logging failed: " & Err.Description & vbCrLf & "Output remains open; do not treat it as verified.", vbCritical
End Sub

Private Sub LoadAliases()
    Dim i As Long, k As String, axis As String
    For i = 11 To cfg.Cells(cfg.Rows.Count, 16).End(xlUp).Row
        axis = UCase$(CStr(cfg.Cells(i, 16).Value2))
        If axis <> "R" And axis <> "C" Then Err.Raise 5, , "Alias axis must be R or C"
        k = axis & "|" & Normalize(CStr(cfg.Cells(i, 17).Value2))
        If aliases.Exists(k) Then Err.Raise 5, , "Duplicate alias"
        aliases.Add k, Normalize(CStr(cfg.Cells(i, 18).Value2))
    Next i
End Sub
Public Function Normalize(ByVal s As String) As String
    s = Replace(s, ChrW(160), " ")
    s = Replace(Replace(Replace(s, vbCrLf, " "), vbCr, " "), vbLf, " ")
    s = Trim$(s)
    If Not caseSensitive Then s = LCase$(s)
    Normalize = s
End Function
Private Function Canon(ByVal axis As String, ByVal label As String) As String
    Dim k As String
    k = Normalize(label)
    If aliases.Exists(axis & "|" & k) Then k = aliases(axis & "|" & k)
    Canon = k
End Function
Private Function Label(ByVal c As Range) As String
    Dim t As Range, v As Variant
    Set t = c
    If c.MergeCells Then Set t = c.MergeArea.Cells(1, 1)
    v = t.Value2
    If IsError(v) Then Exit Function
    'Text labels preserve leading zeros. Numeric headings use stable serials, not locale formatting.
    If IsEmpty(v) Then Exit Function
    Label = CStr(v)
End Function
Private Function Part(ByVal s As String) As String
    Part = CStr(Len(s)) & ":" & s
End Function
Public Sub GetCellHeaderKeys(ByVal r As Object, ByVal row As Long, ByVal col As Long, _
    ByRef rawRow As String, ByRef rawCol As String, ByRef rowKey As String, ByRef colKey As String)
    Dim a As Range, i As Long, t As String
    Set a = r("range")
    rawRow = "": rawCol = "": rowKey = "": colKey = ""
    For i = a.Column To a.Column + r("hc") - 1
        t = Label(a.Worksheet.Cells(row, i)): rawRow = rawRow & Part(t): rowKey = rowKey & Part(Canon("R", t))
    Next i
    For i = a.Row To a.Row + r("hr") - 1
        t = Label(a.Worksheet.Cells(i, col)): rawCol = rawCol & Part(t): colKey = colKey & Part(Canon("C", t))
    Next i
End Sub

Public Function DiscoverDataRegions(ByVal wb As Workbook, ByVal side As String) As Collection
    Dim result As New Collection, seen As Object, w As Worksheet, i As Long, r As Object
    Dim a As Range, inp As Range, t As ListObject, nm As Name, boundary As Range, cell As Range, candidate As Range
    Dim key As String, included As Boolean, configured As Boolean
    Set seen = Dict()
    For Each w In wb.Worksheets
        If Not AuditName(w.Name) Then
            included = True
            For i = 11 To cfg.Cells(cfg.Rows.Count, 1).End(xlUp).Row
                If UCase$(CStr(cfg.Cells(i, 1).Value2)) = side And CStr(cfg.Cells(i, 2).Value2) = w.Name Then
                    If Not CBool(cfg.Cells(i, 12).Value2) Then included = False
                End If
            Next i
            If included Then
                For i = 11 To cfg.Cells(cfg.Rows.Count, 1).End(xlUp).Row
                    If UCase$(CStr(cfg.Cells(i, 1).Value2)) = side And CStr(cfg.Cells(i, 2).Value2) = w.Name Then
                        Set a = ResolveRange(w, CStr(cfg.Cells(i, 4).Value2))
                        Set inp = ResolveRange(w, CStr(cfg.Cells(i, 7).Value2))
                        Set r = Dict()
                        r.Add "sheet", w.Name: r.Add "id", CStr(cfg.Cells(i, 3).Value2)
                        r.Add "section", Normalize(CStr(cfg.Cells(i, 8).Value2))
                        r.Add "range", a: r.Add "inputs", inp
                        r.Add "hr", CLng(Val(cfg.Cells(i, 5).Value2)): r.Add "hc", CLng(Val(cfg.Cells(i, 6).Value2))
                        r.Add "method", "configured": r.Add "valid", False: r.Add "failure", "missing header"
                        r.Add "hidden", CBool(cfg.Cells(i, 11).Value2)
                        DetectHeaderBands r, cfg.Cells(i, 9).Value2, CBool(cfg.Cells(i, 10).Value2), UCase$(CStr(cfg.Cells(i, 13).Value2))
                        If r("valid") Then ValidateRegion r
                        result.Add r
                        Inventory side, r, "configured", IIf(r("valid"), "validated", "review required")
                        If Not a Is Nothing Then seen(w.Name & "!" & a.Address) = True
                        If Not r("valid") Then Issue "Discovery", r("failure"), IIfObject(side = "OLD", r, Nothing), IIfObject(side = "NEW", r, Nothing), Nothing, Nothing, "Invalid configured region/header/input range", 0, "unchanged", "unknown"
                    End If
                Next i
                'Metadata candidates come before the content fallback.
                For Each t In w.ListObjects
                    RecordCandidate side, w, t.Range, "Table:" & t.Name, seen
                Next t
                For Each nm In wb.Names
                    Set a = NamedRange(nm)
                    If Not a Is Nothing Then
                        If a.Worksheet Is w Then RecordCandidate side, w, a, "Name:" & nm.Name, seen
                    End If
                Next nm
                Set boundary = ContentBoundary(w)
                If Not boundary Is Nothing Then
                    For Each cell In boundary.Cells
                        If Not IsEmpty(cell.Value2) Or cell.HasFormula Then
                            key = w.Name & "!" & cell.Address
                            If Not Covered(cell, result) Then
                                If w.ProtectContents Then
                                    Set candidate = ContentBlock(cell, boundary)
                                    RecordCandidate side, w, candidate, "protected read-only blank-boundary", seen
                                Else
                                    Set candidate = cell.CurrentRegion
                                    RecordCandidate side, w, candidate, "CurrentRegion", seen
                                End If
                            End If
                        End If
                    Next cell
                End If
            End If
        End If
    Next w
    For i = 11 To cfg.Cells(cfg.Rows.Count, 1).End(xlUp).Row
        If UCase$(CStr(cfg.Cells(i, 1).Value2)) = side Then
            If GetSheet(wb, CStr(cfg.Cells(i, 2).Value2)) Is Nothing Then
                Issue "Discovery", "missing worksheet", Nothing, Nothing, Nothing, Nothing, side & ": " & CStr(cfg.Cells(i, 2).Value2), 0, "unchanged", "unknown"
            End If
        End If
    Next i
    Set DiscoverDataRegions = result
End Function
Private Function IIfObject(ByVal test As Boolean, ByVal a As Object, ByVal b As Object) As Object
    If test Then Set IIfObject = a Else Set IIfObject = b
End Function
Private Function NamedRange(ByVal nm As Name) As Range
    On Error GoTo NotRange
    Set NamedRange = nm.RefersToRange
NotRange:
End Function
Private Function ResolveRange(ByVal w As Worksheet, ByVal s As String) As Range
    Dim t As ListObject
    If s = "" Then Exit Function
    For Each t In w.ListObjects
        If t.Name = s Then Set ResolveRange = t.Range: Exit Function
    Next t
    On Error GoTo Invalid
    Set ResolveRange = w.Range(s)
    If Not ResolveRange.Worksheet Is w Then Set ResolveRange = Nothing
Invalid:
End Function
Private Function ContentBoundary(ByVal w As Worksheet) As Range
    Dim u As Range, first As Range, last As Range, left As Range, right As Range
    Set u = w.UsedRange
    Set first = u.Find(What:="*", After:=u.Cells(u.Rows.Count, u.Columns.Count), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, MatchByte:=False, SearchFormat:=False)
    If first Is Nothing Then Exit Function
    Set last = u.Find(What:="*", After:=u.Cells(1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, MatchByte:=False, SearchFormat:=False)
    Set left = u.Find(What:="*", After:=u.Cells(u.Rows.Count, u.Columns.Count), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, MatchByte:=False, SearchFormat:=False)
    Set right = u.Find(What:="*", After:=u.Cells(1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False, MatchByte:=False, SearchFormat:=False)
    Set ContentBoundary = w.Range(w.Cells(first.Row, left.Column), w.Cells(last.Row, right.Column))
End Function
Private Function ContentBlock(ByVal seed As Range, ByVal b As Range) As Range
    Dim top As Long, bot As Long, l As Long, rt As Long, w As Worksheet
    Set w = seed.Worksheet: top = seed.Row: bot = top: l = seed.Column: rt = l
    Do While top > b.Row
        If Application.CountA(w.Range(w.Cells(top - 1, b.Column), w.Cells(top - 1, b.Column + b.Columns.Count - 1))) = 0 Then Exit Do
        top = top - 1
    Loop
    Do While bot < b.Row + b.Rows.Count - 1
        If Application.CountA(w.Range(w.Cells(bot + 1, b.Column), w.Cells(bot + 1, b.Column + b.Columns.Count - 1))) = 0 Then Exit Do
        bot = bot + 1
    Loop
    Do While l > b.Column
        If Application.CountA(w.Range(w.Cells(top, l - 1), w.Cells(bot, l - 1))) = 0 Then Exit Do
        l = l - 1
    Loop
    Do While rt < b.Column + b.Columns.Count - 1
        If Application.CountA(w.Range(w.Cells(top, rt + 1), w.Cells(bot, rt + 1))) = 0 Then Exit Do
        rt = rt + 1
    Loop
    Set ContentBlock = w.Range(w.Cells(top, l), w.Cells(bot, rt))
End Function
Private Function Covered(ByVal c As Range, ByVal regions As Collection) As Boolean
    Dim r As Object, a As Range
    For Each r In regions
        Set a = r("range")
        If Not a Is Nothing Then
            If a.Worksheet Is c.Worksheet Then
                If Not Application.Intersect(a, c) Is Nothing Then Covered = True: Exit Function
            End If
        End If
    Next r
End Function
Private Sub RecordCandidate(ByVal side As String, ByVal w As Worksheet, ByVal a As Range, ByVal method As String, ByVal seen As Object)
    Dim k As String, r As Object
    k = w.Name & "!" & a.Address
    If seen.Exists(k) Then Exit Sub
    seen.Add k, True
    Set r = Dict(): r.Add "sheet", w.Name: r.Add "id", "?": r.Add "range", a
    r.Add "hr", 0: r.Add "hc", 0: r.Add "method", method: r.Add "valid", False
    Inventory side, r, method, "candidate only; configure logical boundary and inputs"
    Issue "Discovery", "ambiguous region", IIfObject(side = "OLD", r, Nothing), IIfObject(side = "NEW", r, Nothing), Nothing, Nothing, method & " " & a.Address & "; automatic boundaries are not proof of logical tables", 0, "unchanged", "unknown"
End Sub
Private Sub Inventory(ByVal side As String, ByVal r As Object, ByVal method As String, ByVal status As String)
    Dim a As Range, addr As String, nr As Long, nc As Long, ch As String, rh As String
    Set a = r("range")
    If Not a Is Nothing Then
        addr = a.Address: nr = a.Rows.Count: nc = a.Columns.Count
        If r("hr") > 0 And r("hr") < nr And r("hc") > 0 And r("hc") < nc Then
            ch = a.Resize(r("hr"), nc).Address: rh = a.Resize(nr, r("hc")).Address
        End If
    End If
    AppendRow IV, Array(runID, side, r("sheet"), r("id"), addr, CStr(nr), CStr(nc), method, ch, rh, status)
End Sub

Public Sub DetectHeaderBands(ByVal r As Object, ByVal colour As Variant, ByVal display As Boolean, ByVal mode As String)
    Dim a As Range, i As Long, j As Long, hits As Long, rowBand As Long, colBand As Long
    Set a = r("range")
    If a Is Nothing Then Exit Sub
    If a.Areas.Count <> 1 Then Exit Sub
    If mode = "EXPLICIT" Then
        r("method") = "explicit bands; content/uniqueness validated"
    Else
        'First row/column evaluated first. A strict simple-label test is intentionally conservative.
        If SimpleBands(a) Then
            r("hr") = 1: r("hc") = 1: r("method") = "positional + meaningful text + uniqueness"
        ElseIf Not IsEmpty(colour) And CStr(colour) <> "" Then
            For i = 1 To a.Rows.Count - 1
                hits = 0
                For j = 2 To a.Columns.Count
                    If HeaderColour(a.Cells(i, j), CLng(colour), display) And Len(Label(a.Cells(i, j))) > 0 Then hits = hits + 1
                Next j
                If hits = a.Columns.Count - 1 Then
                    If rowBand <> 0 Then Exit Sub
                    rowBand = i
                End If
            Next i
            For j = 1 To a.Columns.Count - 1
                hits = 0
                For i = rowBand + 1 To a.Rows.Count
                    If HeaderColour(a.Cells(i, j), CLng(colour), display) And Len(Label(a.Cells(i, j))) > 0 Then hits = hits + 1
                Next i
                If hits = a.Rows.Count - rowBand And hits > 0 Then
                    If colBand <> 0 Then Exit Sub
                    colBand = j
                End If
            Next j
            If rowBand = 0 Or colBand = 0 Then Exit Sub
            'Title/note rows before the found band are excluded in memory.
            Set a = a.Offset(rowBand - 1, colBand - 1).Resize(a.Rows.Count - rowBand + 1, a.Columns.Count - colBand + 1)
            Set r("range") = a: r("hr") = 1: r("hc") = 1
            r("method") = IIf(display, "DisplayFormat", "Interior") & " colour + labels + continuity + uniqueness"
        Else
            Exit Sub
        End If
    End If
    If r("hr") < 1 Or r("hc") < 1 Or r("hr") >= a.Rows.Count Or r("hc") >= a.Columns.Count Then Exit Sub
    r("valid") = True
End Sub
Private Function HeaderColour(ByVal c As Range, ByVal colour As Long, ByVal display As Boolean) As Boolean
    If display Then HeaderColour = (c.DisplayFormat.Interior.Color = colour) Else HeaderColour = (c.Interior.Color = colour)
End Function
Private Function SimpleBands(ByVal a As Range) As Boolean
    Dim c As Range
    If a.Rows.Count < 3 Or a.Columns.Count < 3 Then Exit Function
    For Each c In a.Rows(1).Offset(0, 1).Resize(1, a.Columns.Count - 1).Cells
        If c.MergeCells Or c.HasFormula Or VarType(c.Value2) <> vbString Or Len(Trim$(Label(c))) = 0 Then Exit Function
    Next c
    For Each c In a.Columns(1).Offset(1).Resize(a.Rows.Count - 1, 1).Cells
        If c.MergeCells Or c.HasFormula Or VarType(c.Value2) <> vbString Or Len(Trim$(Label(c))) = 0 Then Exit Function
    Next c
    SimpleBands = True
End Function
Private Sub ValidateRegion(ByVal r As Object)
    Dim a As Range, inp As Range, body As Range, rows As Object, cols As Object
    Dim i As Long, rr As String, cc As String, rk As String, ck As String
    Set a = r("range"): Set inp = r("inputs")
    r("valid") = False
    If inp Is Nothing Then Exit Sub
    Set body = a.Offset(r("hr"), r("hc")).Resize(a.Rows.Count - r("hr"), a.Columns.Count - r("hc"))
    If Not inp.Worksheet Is a.Worksheet Then Exit Sub
    If Application.Intersect(inp, body) Is Nothing Then Exit Sub
    If Application.Intersect(inp, body).CountLarge <> inp.CountLarge Then Exit Sub
    Set rows = Dict(): Set cols = Dict()
    For i = body.Row To body.Row + body.Rows.Count - 1
        GetCellHeaderKeys r, i, body.Column, rr, cc, rk, ck
        If rows.Exists(rk) Then r("failure") = "duplicate header": Exit Sub
        If Replace(rr, "0:", "") = "" Then Exit Sub
        rows.Add rk, i
    Next i
    For i = body.Column To body.Column + body.Columns.Count - 1
        GetCellHeaderKeys r, body.Row, i, rr, cc, rk, ck
        If cols.Exists(ck) Then r("failure") = "duplicate header": Exit Sub
        If Replace(cc, "0:", "") = "" Then Exit Sub
        cols.Add ck, i
    Next i
    r.Add "body", body: r.Add "rows", rows: r.Add "cols", cols
    r("valid") = True
End Sub

Public Function MatchRegion(ByVal s As Object, ByRef category As String) As Object
    Dim d As Object, count As Long
    For Each d In destRegions
        'Stable ID is an explicit cross-sheet mapping. IDs must be globally unique per side.
        If d("id") = s("id") And d("valid") Then
            count = count + 1: Set MatchRegion = d
        End If
    Next d
    If count = 0 Then category = "missing region"
    If count > 1 Then category = "ambiguous region"
    If count <> 1 Then Set MatchRegion = Nothing
End Function
Private Function EscapeFind(ByVal s As String) As String
    EscapeFind = Replace(Replace(Replace(s, "~", "~~"), "*", "~*"), "?", "~?")
End Function
Public Function ScopedFind(ByVal r As Object, ByVal source As Object, ByVal sc As Range, ByVal axis As String, ByVal wanted As String) As Long
    Dim a As Range, band As Range, hit As Range, first As String, c As Range, candidates As Object
    Dim i As Long, rr As String, cc As String, rk As String, ck As String, idx As Variant, answer As Long, count As Long
    Set a = r("range"): Set candidates = Dict()
    If axis = "R" Then Set band = a.Resize(a.Rows.Count, r("hc")) Else Set band = a.Resize(r("hr"), a.Columns.Count)
    'Find every raw destination label; then verify complete normalized composite key.
    'This also supports aliases, formula labels and merged spans without assuming raw equality.
    For Each c In band.Cells
        If Len(Label(c)) > 0 Then
            Set hit = band.Find(What:=EscapeFind(Label(c)), After:=band.Cells(band.Cells.CountLarge), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=caseSensitive, MatchByte:=False, SearchFormat:=False)
            If Not hit Is Nothing Then
                first = hit.Address
                Do
                    If axis = "R" Then
                        If hit.Row >= a.Row + r("hr") Then candidates(CStr(hit.Row)) = True
                    Else
                        If hit.Column >= a.Column + r("hc") Then candidates(CStr(hit.Column)) = True
                    End If
                    Set hit = band.Find(What:=EscapeFind(Label(c)), After:=hit, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=caseSensitive, MatchByte:=False, SearchFormat:=False)
                    If hit Is Nothing Then Exit Do
                Loop While hit.Address <> first
            End If
            'Merged heading spans supply indexes absent from Find's top-left result.
            If c.MergeCells Then
                If axis = "C" Then
                    For i = c.MergeArea.Column To c.MergeArea.Column + c.MergeArea.Columns.Count - 1
                        If i >= a.Column + r("hc") And i <= a.Column + a.Columns.Count - 1 Then candidates(CStr(i)) = True
                    Next i
                Else
                    For i = c.MergeArea.Row To c.MergeArea.Row + c.MergeArea.Rows.Count - 1
                        If i >= a.Row + r("hr") And i <= a.Row + a.Rows.Count - 1 Then candidates(CStr(i)) = True
                    Next i
                End If
            End If
        End If
    Next c
    For Each idx In candidates.Keys
        If axis = "R" Then
            GetCellHeaderKeys r, CLng(idx), a.Column + r("hc"), rr, cc, rk, ck
            If rk = wanted Then count = count + 1: answer = CLng(idx)
        Else
            GetCellHeaderKeys r, a.Row + r("hr"), CLng(idx), rr, cc, rk, ck
            If ck = wanted Then count = count + 1: answer = CLng(idx)
        End If
    Next idx
    If count = 1 Then ScopedFind = answer
End Function
Private Function CellID(ByVal c As Range) As String
    CellID = c.Worksheet.Name & "!" & c.Address
End Function
Private Function IncludedInput(ByVal r As Object, ByVal c As Range) As Boolean
    IncludedInput = True
    If Not r("hidden") Then IncludedInput = Not (c.EntireRow.Hidden Or c.EntireColumn.Hidden Or c.Worksheet.Visible <> xlSheetVisible)
End Function

Public Sub BuildDryRun()
    Dim s As Object, d As Object, sc As Range, dc As Range, cat As String, rr As String, cc As String, rk As String, ck As String
    Dim row As Long, col As Long, p As Object, reason As String, cache As Object, key As String
    Set cache = Dict()
    InventoryInvalidInputs sourceRegions, True
    InventoryInvalidInputs destRegions, False
    For Each s In sourceRegions
        If s("valid") Then
            cat = "": Set d = MatchRegion(s, cat)
            For Each sc In s("inputs").Cells
                If IncludedInput(s, sc) Then
                    srcCount = srcCount + 1
                    key = CellID(sc)
                    If outcomes.Exists(key) Then Err.Raise 5, , "Overlapping source input areas: " & key
                    outcomes.Add key, "unresolved"
                    If d Is Nothing Then
                        Issue "Mapping", cat, s, Nothing, sc, Nothing, "Require a unique stable region ID"
                    Else
                        GetCellHeaderKeys s, sc.Row, sc.Column, rr, cc, rk, ck
                        row = CachedFind(cache, d, s, sc, "R", rk)
                        col = CachedFind(cache, d, s, sc, "C", ck)
                        If row = 0 Or col = 0 Then
                            Issue "Mapping", "missing destination", s, d, sc, Nothing, "Missing or ambiguous complete row/column header key"
                        Else
                            Set dc = newBook.Worksheets(d("sheet")).Cells(row, col)
                            If Application.Intersect(dc, d("body")) Is Nothing Or Application.Intersect(dc, d("inputs")) Is Nothing Then
                                Issue "Mapping", "missing destination", s, d, sc, dc, "Resolved cell outside designated destination body/inputs"
                            ElseIf Not IncludedInput(d, dc) Then
                                Issue "Mapping", "missing destination", s, d, sc, dc, "Destination excluded by hidden configuration"
                            ElseIf targets.Exists(CellID(dc)) Then
                                'Invalidate the earlier proposal as well: no first-writer wins.
                                Set p = targets(CellID(dc)): p("valid") = False
                                Issue "Mapping", "duplicate header", s, d, sc, dc, "Multiple sources claim destination; all claims blocked"
                            Else
                                Set p = Dict(): p.Add "s", s: p.Add "d", d: p.Add "sc", sc: p.Add "dc", dc
                                p.Add "formula", LinkFormula(sc): p.Add "valid", True
                                p.Add "previous", dc.Formula: p.Add "previousValue", dc.Value2: p.Add "previousHasFormula", CBool(dc.HasFormula)
                                targets.Add CellID(dc), p
                                reason = Eligibility(sc, dc, p("formula"))
                                If reason <> "" Then
                                    p("valid") = False
                                    Issue "Preflight", reason, s, d, sc, dc, "Eligibility failed"
                                ElseIf Not ValidationAccepts(sc, dc) Then
                                    p("valid") = False
                                    Issue "Preflight", "invalid value", s, d, sc, dc, "Validation fails or cannot be safely evaluated"
                                End If
                                plans.Add p
                                AppendRow PL, Array(runID, sc.Worksheet.Name, sc.Address, dc.Worksheet.Name, dc.Address, rk, ck, IIf(valuesOnly, "values-only", p("formula")))
                            End If
                        End If
                    End If
                End If
            Next sc
        End If
    Next s
    For Each d In destRegions
        If d("valid") Then
            For Each dc In d("inputs").Cells
                If IncludedInput(d, dc) Then
                    dstCount = dstCount + 1
                    key = "D|" & CellID(dc)
                    If outcomes.Exists(key) Then Err.Raise 5, , "Overlapping destination input areas"
                    outcomes.Add key, "unresolved"
                    If Not targets.Exists(CellID(dc)) Then Issue "Reconciliation", "missing source", Nothing, d, Nothing, dc, "No identified source input maps here"
                End If
            Next dc
        End If
    Next d
End Sub
Private Sub InventoryInvalidInputs(ByVal regions As Collection, ByVal sourceSide As Boolean)
    Dim r As Object, a As Range, c As Range, k As String
    For Each r In regions
        If Not r("valid") Then
            Set a = r("inputs")
            If Not a Is Nothing Then
                For Each c In a.Cells
                    If IncludedInput(r, c) Then
                        If sourceSide Then
                            k = CellID(c): srcCount = srcCount + 1
                        Else
                            k = "D|" & CellID(c): dstCount = dstCount + 1
                        End If
                        If outcomes.Exists(k) Then Err.Raise 5, , "Overlapping invalid input designation: " & k
                        outcomes.Add k, "unresolved"
                        If sourceSide Then
                            Issue "Inventory", r("failure"), r, Nothing, c, Nothing, "Input belongs to an unvalidated region"
                        Else
                            Issue "Inventory", r("failure"), Nothing, r, Nothing, c, "Input belongs to an unvalidated region"
                        End If
                    End If
                Next c
            End If
        End If
    Next r
End Sub
Private Function CachedFind(ByVal cache As Object, ByVal d As Object, ByVal s As Object, ByVal sc As Range, ByVal axis As String, ByVal wanted As String) As Long
    Dim k As String
    k = Part(d("sheet")) & Part(d("id")) & axis & Part(wanted)
    If Not cache.Exists(k) Then cache.Add k, ScopedFind(d, s, sc, axis, wanted)
    CachedFind = cache(k)
End Function

Private Sub LoadOwnership()
    Dim w As Worksheet, i As Long, key As String
    'Only audit in the copied template can establish ownership. External logs are never inferred.
    Set w = GetSheet(newBook, AU)
    If w Is Nothing Then Exit Sub
    If w.Cells(1, 1).Value2 <> "Run ID" Then Err.Raise 5, , "Unrecognised success audit"
    For i = 2 To w.Cells(w.Rows.Count, 1).End(xlUp).Row
        If (w.Cells(i, 9).Value2 = "written" Or w.Cells(i, 9).Value2 = "already correct") And w.Cells(i, 10).Value2 = "link" Then
            key = CStr(w.Cells(i, 3).Value2) & "!" & CStr(w.Cells(i, 4).Value2)
            owned(key) = CStr(w.Cells(i, 8).Value2)
        End If
    Next i
End Sub
Public Function LinkFormula(ByVal sc As Range) As String
    Dim ref As String, prefix As String
    prefix = sc.Worksheet.Parent.Path & Application.PathSeparator & "[" & sc.Worksheet.Parent.Name & "]" & sc.Worksheet.Name
    ref = "'" & Replace(prefix, "'", "''") & "'!" & sc.Address(True, True, xlA1)
    LinkFormula = "=IF(ISBLANK(" & ref & "),""""," & ref & ")"
End Function
Private Function Eligibility(ByVal sc As Range, ByVal dc As Range, ByVal formula As String) As String
    If sc.MergeCells Or dc.MergeCells Then Eligibility = "incompatible merged range": Exit Function
    If dc.Locked Then Eligibility = "locked destination": Exit Function
    If dc.HasArray Then Eligibility = "existing formula": Exit Function
    If dc.HasFormula Then
        If Not owned.Exists(CellID(dc)) Then Eligibility = "existing formula": Exit Function
        If CStr(dc.Formula) <> owned(CellID(dc)) Then Eligibility = "existing formula": Exit Function
    End If
    If IsError(sc.Value2) Then Eligibility = "source error": Exit Function
    'Date serials are preserved. Different workbook epochs require explicit conversion, unsupported.
    If oldBook.Date1904 <> newBook.Date1904 Then Eligibility = "invalid value": Exit Function
    If valuesOnly And VarType(sc.Value2) = vbString Then
        If Len(sc.Value2) > 32767 Then Eligibility = "invalid value"
    End If
End Function

Private Function ValidationType(ByVal c As Range) As Long
    On Error GoTo NoValidation
    ValidationType = c.Validation.Type
    Exit Function
NoValidation:
    If Err.Number <> 1004 Then Err.Raise Err.Number, , Err.Description
    ValidationType = -1
End Function
Private Function ValidationAccepts(ByVal sc As Range, ByVal dc As Range) As Boolean
    Dim typ As Long, v As Variant, op As Long, lo As Variant, hi As Variant, f As String, items As Variant, x As Variant
    On Error GoTo Unsupported
    typ = ValidationType(dc)
    If typ = -1 Or typ = xlValidateInputOnly Then ValidationAccepts = True: Exit Function
    v = sc.Value2
    If IsEmpty(v) Or (VarType(v) = vbString And Len(v) = 0) Then
        ValidationAccepts = dc.Validation.IgnoreBlank: Exit Function
    End If
    f = dc.Validation.Formula1
    If typ = xlValidateList Then
        'Literal lists only. Range/formula/custom lists require a reviewed adapter.
        If Left$(f, 1) = "=" Then Exit Function
        items = Split(f, Application.International(xlListSeparator))
        For Each x In items
            If CStr(v) = CStr(x) Then ValidationAccepts = True: Exit Function
        Next x
        Exit Function
    End If
    If typ = xlValidateCustom Then Exit Function
    If Left$(f, 1) = "=" Or Not IsNumeric(f) Then Exit Function
    lo = CDbl(f): hi = lo
    op = dc.Validation.Operator
    If op = xlBetween Or op = xlNotBetween Then
        f = dc.Validation.Formula2
        If Left$(f, 1) = "=" Or Not IsNumeric(f) Then Exit Function
        hi = CDbl(f)
    End If
    If typ = xlValidateTextLength Then
        v = Len(CStr(v))
    Else
        If Not IsNumeric(v) Or VarType(v) = vbBoolean Then Exit Function
        v = CDbl(v)
        If typ = xlValidateWholeNumber And v <> Fix(v) Then Exit Function
    End If
    Select Case op
        Case xlBetween: ValidationAccepts = (v >= lo And v <= hi)
        Case xlNotBetween: ValidationAccepts = (v < lo Or v > hi)
        Case xlEqual: ValidationAccepts = (v = lo)
        Case xlNotEqual: ValidationAccepts = (v <> lo)
        Case xlGreater: ValidationAccepts = (v > lo)
        Case xlLess: ValidationAccepts = (v < lo)
        Case xlGreaterEqual: ValidationAccepts = (v >= lo)
        Case xlLessEqual: ValidationAccepts = (v <= lo)
    End Select
    Exit Function
Unsupported:
    ValidationAccepts = False
End Function

Public Sub ExecuteMapping(ByVal p As Object)
    Dim sc As Range, dc As Range, reason As String, n As Long, msg As String, outcome As String, auditPhase As Boolean
    If Not p("valid") Then Exit Sub
    Set sc = p("sc"): Set dc = p("dc")
    On Error GoTo Failed
    reason = Eligibility(sc, dc, p("formula"))
    If reason <> "" Then Issue "Write", reason, p("s"), p("d"), sc, dc, "Eligibility changed since preflight": Exit Sub
    If Not ValidationAccepts(sc, dc) Then Issue "Write", "invalid value", p("s"), p("d"), sc, dc, "Validation changed": Exit Sub
    If Not valuesOnly And dc.HasFormula And CStr(dc.Formula) = p("formula") Then
        dc.Calculate
        If Not Equivalent(sc, dc) Then Err.Raise vbObjectError + 11, , "Existing link value mismatch"
        outcome = "already correct"
    Else
        If TestWriteFailure(dc) Then Err.Raise 1004, , "Injected recoverable test failure"
        If valuesOnly Then SetValue dc, sc.Value2 Else dc.Formula = p("formula")
        dc.Calculate
        If Not valuesOnly Then
            If CStr(dc.Formula) <> p("formula") Then Err.Raise vbObjectError + 11, , "Link formula mismatch"
        End If
        If Not Equivalent(sc, dc) Then Err.Raise vbObjectError + 11, , "Result mismatch"
        outcome = "written"
    End If
    outcomes(CellID(sc)) = outcome: outcomes("D|" & CellID(dc)) = outcome
    VerifyPreservation
    If outcome = "written" Then successes = successes + 1 Else already = already + 1
    auditPhase = True
    AppendRow AU, Array(runID, Format$(Now, "yyyy-mm-dd hh:nn:ss"), dc.Worksheet.Name, dc.Address, oldBook.FullName, sc.Worksheet.Name, sc.Address, IIf(valuesOnly, "", p("formula")), outcome, IIf(valuesOnly, "values", "link"))
    Exit Sub
Failed:
    n = Err.Number: msg = Err.Description
    'Audit failures must be fatal; do not disguise them as recoverable writes.
    If auditPhase Then Err.Raise n, , msg
    outcomes(CellID(sc)) = "unresolved": outcomes("D|" & CellID(dc)) = "unresolved"
    If RestoreCell(p) Then
        Issue "Write/verify", IIf(n = vbObjectError + 11, "verification mismatch", "write failure"), p("s"), p("d"), sc, dc, msg, n, "restored"
        If n = vbObjectError + 20 Then Err.Raise n, , msg
    Else
        Issue "Rollback", "write failure", p("s"), p("d"), sc, dc, msg, n, "restoration failed", "unknown"
        Err.Raise vbObjectError + 12, , "Cannot restore failed write; stop"
    End If
End Sub
Private Sub SetValue(ByVal c As Range, ByVal v As Variant)
    If IsEmpty(v) Then
        c.ClearContents
    ElseIf VarType(v) = vbString Then
        'Leading apostrophe is Excel's text-entry prefix, not part of Value2.
        c.Value2 = "'" & CStr(v)
    Else
        c.Value2 = v
    End If
End Sub
Private Function RestoreCell(ByVal p As Object) As Boolean
    Dim c As Range
    On Error GoTo Failed
    Set c = p("dc")
    If p("previousHasFormula") Then
        c.Formula = p("previous")
    Else
        SetValue c, p("previousValue")
    End If
    If p("previousHasFormula") Then
        RestoreCell = (CStr(c.Formula) = CStr(p("previous")))
    Else
        RestoreCell = (ValueToken(c.Value2) = ValueToken(p("previousValue")))
    End If
    Exit Function
Failed:
    RestoreCell = False
End Function
Private Function Equivalent(ByVal s As Range, ByVal d As Range) As Boolean
    Dim a As Variant, b As Variant
    a = s.Value2: b = d.Value2
    If IsError(a) Or IsError(b) Then Exit Function
    If IsEmpty(a) Then
        If valuesOnly Then Equivalent = IsEmpty(b) Else Equivalent = EmptyString(b)
    ElseIf EmptyString(a) Then
        Equivalent = (IsEmpty(b) Or EmptyString(b))
    Else
        Equivalent = (ValueToken(a) = ValueToken(b))
    End If
End Function
Private Function EmptyString(ByVal v As Variant) As Boolean
    If VarType(v) = vbString Then EmptyString = (Len(v) = 0)
End Function
Private Function ValueToken(ByVal v As Variant) As String
    If IsError(v) Then ValueToken = "Error:" & CStr(v): Exit Function
    If IsEmpty(v) Then ValueToken = "Empty": Exit Function
    ValueToken = CStr(VarType(v)) & ":" & CStr(v)
End Function
Private Function TestWriteFailure(ByVal c As Range) As Boolean
    'Test-only opt-in name on controller; default absent. Never modifies destination metadata.
    Dim nm As Name, a As Range
    For Each nm In ThisWorkbook.Names
        If nm.Name = "TRANSFER_TEST_FAIL" Then
            Set a = NamedRange(nm)
            If Not a Is Nothing Then TestWriteFailure = (CStr(a.Value2) = CellID(c))
        End If
    Next nm
End Function

Private Sub SnapshotWorkbook()
    Dim w As Worksheet, c As Range, r As Range
    snapshots.Add "WORKBOOK", WorkbookSignature
    For Each w In newBook.Worksheets
        If Not AuditName(w.Name) Then
            snapshots.Add "S|" & w.Name, SheetSignature(w)
            Set r = w.UsedRange
            If r.CountLarge > 1000000 Then Err.Raise 5, , "Preservation snapshot exceeds 1,000,000 cells; reduce inflated UsedRange manually on a separate template"
            For Each c In r.Cells
                snapshots.Add "M|" & CellID(c), Metadata(c)
                snapshots.Add "V|" & CellID(c), Content(c)
            Next c
        End If
    Next w
End Sub
Private Function WorkbookSignature() As String
    WorkbookSignature = CStr(newBook.ProtectStructure) & "|" & CStr(newBook.ProtectWindows) & "|" & CStr(newBook.Date1904) & "|" & CStr(newBook.Sheets.Count)
End Function
Private Function SheetSignature(ByVal w As Worksheet) As String
    With w.Protection
        SheetSignature = CStr(w.ProtectContents) & "|" & CStr(w.ProtectDrawingObjects) & "|" & CStr(w.ProtectScenarios) & "|" & CStr(w.Visible) & "|" & CStr(w.ProtectionMode) & "|" & _
        CStr(.AllowFormattingCells) & CStr(.AllowFormattingColumns) & CStr(.AllowFormattingRows) & CStr(.AllowInsertingColumns) & CStr(.AllowInsertingRows) & CStr(.AllowInsertingHyperlinks) & CStr(.AllowDeletingColumns) & CStr(.AllowDeletingRows) & CStr(.AllowSorting) & CStr(.AllowFiltering) & CStr(.AllowUsingPivotTables)
    End With
End Function
Private Function Content(ByVal c As Range) As String
    If c.HasFormula Then Content = "F|" & CStr(c.Formula) Else Content = "V|" & ValueToken(c.Value2)
End Function
Private Function Metadata(ByVal c As Range) As String
    Dim s As String, b As Variant, typ As Long
    s = Part(CStr(c.Style)) & Part(CStr(c.NumberFormat)) & CStr(c.Locked) & CStr(c.FormulaHidden) & CStr(c.MergeCells)
    If c.MergeCells Then s = s & Part(c.MergeArea.Address)
    With c
        s = s & Part(CStr(.Interior.Color)) & Part(CStr(.Interior.Pattern)) & Part(CStr(.Interior.PatternColor))
        s = s & Part(CStr(.Font.Name)) & Part(CStr(.Font.Size)) & CStr(.Font.Bold) & CStr(.Font.Italic) & Part(CStr(.Font.Color)) & CStr(.Font.Underline) & CStr(.Font.Strikethrough)
        s = s & Part(CStr(.HorizontalAlignment)) & Part(CStr(.VerticalAlignment)) & CStr(.WrapText) & CStr(.ShrinkToFit) & Part(CStr(.Orientation)) & Part(CStr(.IndentLevel))
        s = s & Part(CStr(.RowHeight)) & Part(CStr(.ColumnWidth)) & CStr(.EntireRow.Hidden) & CStr(.EntireColumn.Hidden) & CStr(.FormatConditions.Count)
    End With
    For Each b In Array(xlEdgeLeft, xlEdgeRight, xlEdgeTop, xlEdgeBottom, xlDiagonalUp, xlDiagonalDown)
        With c.Borders(CLng(b)): s = s & Part(CStr(.LineStyle)) & Part(CStr(.Weight)) & Part(CStr(.Color)): End With
    Next b
    typ = ValidationType(c): s = s & Part(CStr(typ))
    If typ <> -1 Then
        With c.Validation
            s = s & Part(.Formula1) & Part(.Formula2) & Part(CStr(.Operator)) & CStr(.IgnoreBlank) & CStr(.InCellDropdown) & CStr(.ShowError) & CStr(.ShowInput) & Part(.ErrorTitle) & Part(.ErrorMessage) & Part(.InputTitle) & Part(.InputMessage) & CStr(.AlertStyle)
        End With
    End If
    Metadata = s
End Function
Public Sub VerifyPreservation()
    Dim w As Worksheet, c As Range, p As Object, permitted As Boolean
    If snapshots("WORKBOOK") <> WorkbookSignature Then Err.Raise vbObjectError + 20, , "Workbook structure/protection changed"
    For Each w In newBook.Worksheets
        If Not AuditName(w.Name) Then
            If snapshots("S|" & w.Name) <> SheetSignature(w) Then Err.Raise vbObjectError + 20, , "Worksheet protection changed: " & w.Name
            For Each c In w.UsedRange.Cells
                If Not snapshots.Exists("M|" & CellID(c)) Then Err.Raise vbObjectError + 20, , "Unexpected UsedRange expansion"
                If snapshots("M|" & CellID(c)) <> Metadata(c) Then Err.Raise vbObjectError + 20, , "Cell metadata changed: " & CellID(c)
                permitted = False
                If targets.Exists(CellID(c)) Then
                    Set p = targets(CellID(c))
                    permitted = p("valid") And outcomes("D|" & CellID(c)) <> "unresolved"
                End If
                If Not permitted Then
                    If snapshots("V|" & CellID(c)) <> Content(c) Then Err.Raise vbObjectError + 20, , "Non-permitted content changed: " & CellID(c)
                End If
            Next c
        End If
    Next w
End Sub
Public Sub Reconcile()
    Dim k As Variant, sUn As Long, dUn As Long, p As Object
    For Each k In outcomes.Keys
        If outcomes(k) = "unresolved" Then
            If Left$(k, 2) = "D|" Then dUn = dUn + 1 Else sUn = sUn + 1
        End If
    Next k
    For Each p In plans
        If p("valid") And Not dry Then
            If Not Equivalent(p("sc"), p("dc")) Then Err.Raise vbObjectError + 11, , "Final value mismatch"
            If Not valuesOnly Then
                If CStr(p("dc").Formula) <> p("formula") Then Err.Raise vbObjectError + 11, , "Final link mismatch"
            End If
        End If
    Next p
    AppendRow SU, Array(runID, "Source unresolved inputs", CStr(sUn))
    AppendRow SU, Array(runID, "Destination unresolved inputs", CStr(dUn))
    If Not dry And (sUn > 0 Or dUn > 0) And exceptions = 0 Then Err.Raise 5, , "Reconciliation has unexplained outcomes"
End Sub
Private Sub LogSummary(ByVal status As String)
    Dim side As Variant, w As Worksheet, i As Long, candidates As Long, validated As Long
    Set w = auditBook.Worksheets(IV)
    For Each side In Array("OLD", "NEW")
        candidates = 0: validated = 0
        For i = 2 To w.Cells(w.Rows.Count, 1).End(xlUp).Row
            If w.Cells(i, 1).Value2 = runID And w.Cells(i, 2).Value2 = side Then
                candidates = candidates + 1
                If w.Cells(i, 11).Value2 = "validated" Then validated = validated + 1
            End If
        Next i
        AppendRow SU, Array(runID, CStr(side) & " candidate regions", CStr(candidates))
        AppendRow SU, Array(runID, CStr(side) & " validated regions", CStr(validated))
    Next side
    AppendRow SU, Array(runID, "Status", status)
    AppendRow SU, Array(runID, "Source intended inputs", CStr(srcCount))
    AppendRow SU, Array(runID, "Destination intended inputs", CStr(dstCount))
    AppendRow SU, Array(runID, "Successful transfers", CStr(successes))
    AppendRow SU, Array(runID, "Already-correct links", CStr(already))
    AppendRow SU, Array(runID, "Exceptions", CStr(exceptions))
    AppendRow SU, Array(runID, "Regions/worksheets with unknown coverage", CStr(unknowns))
    AppendRow SU, Array(runID, "Audit workbook", auditBook.FullName)
End Sub
