Attribute VB_Name = "AFS_FootnoteCompare"
Option Explicit

'=====================================================================
'  AFS FOOTNOTE COMPARE  (bold-heading version, Word PDF import)
'
'  Compares the footnote text of Annual Financial Statements across two
'  years for many funds at once. PDFs are paired by a fund code (XXXX-XX)
'  in the filename. Bold paragraphs are treated as note headings.
'
'  Sheets created in THIS workbook:
'    Config      - folder paths (edit B1 / B2 to change)
'    Data_2024   - one row per fund, one column per bold heading
'    Data_2025   - same
'    Compare     - 4 rows per fund: 2025 / 2024 / Same? / Changed
'    Summary     - one row per fund
'    Log         - skipped files, errors
'
'  Run: CompareFootnotes.  Requires Word 2013 or later on the machine.
'=====================================================================

'---------------------------- SETTINGS -------------------------------
Private Const YEAR_PRIOR As String = "2024"
Private Const YEAR_CURRENT As String = "2025"
Private Const FUND_CODE_PATTERN As String = "[A-Za-z0-9]{4}-[A-Za-z0-9]{2}"

' The page header line. Only pages containing it are read.
Private Const NOTES_HEADER_PATTERN As String = _
    "^notes? to (the )?(consolidated )?financial statements( \(continued\))?$"

' Lines dropped as page furniture: page numbers, "(continued)", date lines.
Private Const NOISE_LINE_PATTERN As String = _
    "^(page\s+\d+(\s+of\s+\d+)?|[-\s]*\d{1,3}[-\s]*|\(?continued\)?|" & _
    "(january|february|march|april|may|june|july|august|september|october|november|december)" & _
    "\s+\d{1,2},?\s+\d{4}(\s+and\s+(\d{4}|\d{1,2},?\s+\d{4}))?)$"

Private Const IGNORE_NUMBERS As Boolean = True       ' numbers become "#" for the Same? test
Private Const SPLIT_RUNIN_HEADINGS As Boolean = True ' "Bold lead-in - text..." counts as a heading
Private Const REPEAT_FURNITURE As Long = 3           ' bold text seen this many times = page header, ignore
Private Const INTRO_LABEL As String = "(Intro)"      ' text on notes pages before the first heading
Private Const FIRST_SECTION_COL As Long = 5          ' Data sheets: A code, B file, C when, D remark, E.. sections
Private Const MAX_CELL As Long = 32000
Private Const MAX_WORDDIFF_CELLS As Long = 250000

'---------------------------- INTERNALS ------------------------------
Private rxCode As Object, rxHeader As Object, rxNoise As Object, rxNum As Object, rxWs As Object, rxCont As Object
Private gWordApp As Object
Private gLog As Collection

'=====================================================================
'  ENTRY POINT
'=====================================================================
Public Sub CompareFootnotes()
    Dim folderP As String, folderC As String
    Dim mapP As Object, mapC As Object, doneP As Object, doneC As Object
    Dim wsDP As Worksheet, wsDC As Worksheet
    Dim codes() As String, pending() As String, nPend As Long, i As Long
    Dim batch As Long, t0 As Single, processed As Long, code As String

    InitRegex
    Set gLog = New Collection

    folderP = GetFolder(YEAR_PRIOR, 1): If folderP = "" Then Exit Sub
    folderC = GetFolder(YEAR_CURRENT, 2): If folderC = "" Then Exit Sub

    Set mapP = BuildFileMap(folderP, YEAR_PRIOR)
    Set mapC = BuildFileMap(folderC, YEAR_CURRENT)
    codes = UnionSorted(mapP, mapC)
    If UBound(codes) < 0 Then
        MsgBox "No PDF filenames containing a fund code (" & FUND_CODE_PATTERN & ") found in either folder.", vbExclamation
        Exit Sub
    End If

    Set wsDP = GetDataSheet(YEAR_PRIOR)
    Set wsDC = GetDataSheet(YEAR_CURRENT)
    Set doneP = LoadDoneCodes(wsDP)
    Set doneC = LoadDoneCodes(wsDC)

    ReDim pending(0 To UBound(codes))
    For i = 0 To UBound(codes)
        If (mapP.Exists(codes(i)) And Not doneP.Exists(codes(i))) Or _
           (mapC.Exists(codes(i)) And Not doneC.Exists(codes(i))) Then
            pending(nPend) = codes(i): nPend = nPend + 1
        End If
    Next i

    batch = AskBatch(UBound(codes) + 1, nPend)
    If batch < 0 Then Exit Sub
    If batch > nPend Then batch = nPend

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    t0 = Timer

    If batch > 0 Then
        If Not StartWord() Then
            MsgBox "Could not start Microsoft Word, which is needed to read the PDFs.", vbCritical
            GoTo Finish
        End If
        For i = 0 To batch - 1
            code = pending(i)
            Application.StatusBar = "Extracting " & (i + 1) & " of " & batch & "  (" & code & ")   elapsed " & _
                                    Format$((Timer - t0) / 86400, "hh:nn:ss")
            DoEvents
            If mapP.Exists(code) And Not doneP.Exists(code) Then
                If ExtractAndStore(code, mapP(code)(0), mapP(code)(1), wsDP, wsDC) Then doneP.Add code, True
            End If
            If mapC.Exists(code) And Not doneC.Exists(code) Then
                If ExtractAndStore(code, mapC(code)(0), mapC(code)(1), wsDC, wsDP) Then doneC.Add code, True
            End If
            processed = processed + 1
        Next i
        ShutdownWord
    End If

    Application.StatusBar = "Building Compare and Summary sheets..."
    DoEvents
    BuildCompare mapP, mapC, wsDP, wsDC
    WriteLog

Finish:
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    On Error Resume Next
    ThisWorkbook.Worksheets("Summary").Activate
    On Error GoTo 0
    MsgBox processed & " fund(s) extracted this run in " & Format$((Timer - t0) / 86400, "hh:nn:ss") & "." & vbCrLf & _
           (nPend - processed) & " fund(s) still pending." & vbCrLf & vbCrLf & _
           "Compare and Summary have been rebuilt from everything stored so far.", vbInformation
End Sub

'=====================================================================
'  BATCH PROMPT / CONFIG
'=====================================================================
Private Function AskBatch(ByVal total As Long, ByVal nPend As Long) As Long
    Dim s As String
    If nPend = 0 Then
        If MsgBox("All " & total & " fund(s) are already extracted." & vbCrLf & vbCrLf & _
                  "Rebuild the Compare and Summary sheets now?", vbYesNo + vbQuestion) = vbYes Then AskBatch = 0 Else AskBatch = -1
        Exit Function
    End If
    Do
        s = InputBox(total & " fund code(s) found in the folders." & vbCrLf & _
                     nPend & " still need extracting." & vbCrLf & vbCrLf & _
                     "How many funds do you want to process now?" & vbCrLf & _
                     "Enter 1, 5, 10, 25, 50, 100, 200 or ALL." & vbCrLf & _
                     "Enter 0 to skip extraction and only rebuild Compare / Summary.", "AFS Footnote Compare", "10")
        If s = "" Then AskBatch = -1: Exit Function
        s = UCase$(Trim$(s))
        If s = "ALL" Then AskBatch = nPend: Exit Function
        Select Case s
            Case "0", "1", "5", "10", "25", "50", "100", "200"
                AskBatch = CLng(s): Exit Function
        End Select
        MsgBox "Please enter one of: 0, 1, 5, 10, 25, 50, 100, 200, ALL", vbExclamation
    Loop
End Function

Private Function GetFolder(ByVal yearLabel As String, ByVal cfgRow As Long) As String
    Dim ws As Worksheet, p As String
    Set ws = GetOrCreateSheet("Config")
    ws.Cells(cfgRow, 1).Value = yearLabel & " folder"
    p = Trim$(CStr(ws.Cells(cfgRow, 2).Value))
    If Len(p) > 0 Then
        If Len(Dir(p, vbDirectory)) = 0 Then p = ""
    End If
    If Len(p) = 0 Then
        p = PickFolder("Select the " & yearLabel & " financial statements folder")
        If Len(p) = 0 Then Exit Function
        ws.Cells(cfgRow, 2).Value = p
    End If
    ws.Columns("A:B").AutoFit
    GetFolder = StripSlash(p)
End Function

'=====================================================================
'  FILE DISCOVERY
'=====================================================================
Private Function BuildFileMap(ByVal folder As String, ByVal yearLabel As String) As Object
    Dim d As Object, f As String, full As String, code As String
    Set d = CreateObject("Scripting.Dictionary")
    f = Dir(folder & "\*.pdf")
    Do While Len(f) > 0
        full = folder & "\" & f
        If rxCode.Test(f) Then
            code = UCase$(rxCode.Execute(f)(0).Value)
            If d.Exists(code) Then
                If FileDateTime(full) > FileDateTime(d(code)(0)) Then
                    d(code) = Array(full, "multiple " & yearLabel & " files for this code; used newest")
                Else
                    d(code) = Array(d(code)(0), "multiple " & yearLabel & " files for this code; used newest")
                End If
            Else
                d.Add code, Array(full, "")
            End If
        Else
            gLog.Add yearLabel & " file skipped (no fund code in name): " & f
        End If
        f = Dir
    Loop
    Set BuildFileMap = d
End Function

Private Function UnionSorted(ParamArray dicts() As Variant) As String()
    Dim all As Object, k As Variant, arr() As String, i As Long, j As Long, tmp As String, d As Variant
    Set all = CreateObject("Scripting.Dictionary")
    For Each d In dicts
        For Each k In d.Keys: all(k) = 1: Next
    Next d
    If all.Count = 0 Then
        ReDim arr(-1 To -1): UnionSorted = arr: Exit Function
    End If
    ReDim arr(0 To all.Count - 1)
    i = 0
    For Each k In all.Keys: arr(i) = k: i = i + 1: Next
    For i = 1 To UBound(arr)
        tmp = arr(i): j = i - 1
        Do While j >= 0
            If arr(j) <= tmp Then Exit Do
            arr(j + 1) = arr(j): j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i
    UnionSorted = arr
End Function

'=====================================================================
'  DATA SHEETS (stored extractions)
'=====================================================================
Private Function GetDataSheet(ByVal yearLabel As String) As Worksheet
    Dim ws As Worksheet, isNew As Boolean
    isNew = Not SheetExists("Data_" & yearLabel)
    Set ws = GetOrCreateSheet("Data_" & yearLabel)
    If isNew Then
        ws.Cells.NumberFormat = "@"
        ws.Range("A1:D1").Value = Array("Fund Code", "File", "Extracted", "Remark")
        ws.Rows(1).Font.Bold = True
        ws.Rows(1).Interior.Color = RGB(217, 225, 242)
        ws.Activate: ws.Range("B2").Select: ActiveWindow.FreezePanes = True
    End If
    Set GetDataSheet = ws
End Function

Private Function LoadDoneCodes(ws As Worksheet) As Object
    Dim d As Object, r As Long, last As Long, c As String
    Set d = CreateObject("Scripting.Dictionary")
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To last
        c = UCase$(Trim$(CStr(ws.Cells(r, 1).Value)))
        If Len(c) > 0 Then If Not d.Exists(c) Then d.Add c, r
    Next r
    Set LoadDoneCodes = d
End Function

' key -> column number, from row 1
Private Function LoadHeaderMap(ws As Worksheet) As Object
    Dim d As Object, c As Long, last As Long, k As String
    Set d = CreateObject("Scripting.Dictionary")
    last = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = FIRST_SECTION_COL To last
        k = NormalizeKey(CStr(ws.Cells(1, c).Value))
        If Len(k) > 0 Then If Not d.Exists(k) Then d.Add k, c
    Next c
    Set LoadHeaderMap = d
End Function

Private Function EnsureColumn(ws As Worksheet, hdr As Object, ByVal key As String, ByVal display As String) As Long
    Dim c As Long
    If hdr.Exists(key) Then EnsureColumn = hdr(key): Exit Function
    c = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
    If c < FIRST_SECTION_COL Then c = FIRST_SECTION_COL
    ws.Cells(1, c).Value = display
    ws.Cells(1, c).Font.Bold = True
    ws.Cells(1, c).Interior.Color = RGB(217, 225, 242)
    hdr.Add key, c
    EnsureColumn = c
End Function

Private Function ExtractAndStore(ByVal code As String, ByVal pdfPath As String, ByVal fileRemark As String, _
                                 wsData As Worksheet, wsOther As Worksheet) As Boolean
    Dim secs As Object, remark As String, r As Long, k As Variant, v As Variant
    Dim hdr As Object, hdrOther As Object, c As Long, txt As String
    On Error GoTo Fail
    Set secs = ExtractSections(pdfPath, remark)
    If secs.Count = 0 Then
        gLog.Add code & ": no text extracted from " & FileNameOnly(pdfPath) & " (" & remark & ")"
        Exit Function
    End If
    Set hdr = LoadHeaderMap(wsData)
    Set hdrOther = LoadHeaderMap(wsOther)
    r = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row + 1
    wsData.Cells(r, 1).Value = code
    wsData.Cells(r, 2).Value = FileNameOnly(pdfPath)
    wsData.Cells(r, 3).Value = Format$(Now, "yyyy-mm-dd hh:nn")
    wsData.Cells(r, 4).Value = JoinNonEmpty(fileRemark, remark)
    For Each k In secs.Keys
        v = secs(k)
        c = EnsureColumn(wsData, hdr, CStr(k), CStr(v(0)))
        EnsureColumn wsOther, hdrOther, CStr(k), CStr(v(0))     ' keep both sheets' headers aligned
        txt = CStr(v(1))
        If Len(txt) > MAX_CELL Then txt = Left$(txt, MAX_CELL) & " ...[TRUNCATED]"
        wsData.Cells(r, c).Value = txt
    Next k
    ExtractAndStore = True
    Exit Function
Fail:
    gLog.Add code & ": ERROR extracting " & FileNameOnly(pdfPath) & " - " & Err.Number & " " & Err.Description
    ExtractAndStore = False
End Function

'=====================================================================
'  WORD EXTRACTION: paragraphs + bold flags -> sections by bold heading
'=====================================================================
Private Function ExtractSections(ByVal pdfPath As String, ByRef remark As String) As Object
    Dim secs As Object, doc As Object, para As Object, rng As Object
    Dim n As Long, i As Long, b As Variant, lead As String
    Dim pText() As String, pKind() As Integer, pPage() As Long, pLead() As String
    Set secs = CreateObject("Scripting.Dictionary")

    Set doc = gWordApp.Documents.Open(pdfPath, False, True, False)   ' ConfirmConversions, ReadOnly, AddToRecent
    n = doc.Paragraphs.Count
    If n = 0 Then doc.Close 0: Set ExtractSections = secs: Exit Function
    ReDim pText(1 To n): ReDim pKind(1 To n): ReDim pPage(1 To n): ReDim pLead(1 To n)

    ' Pass 1: pull text, page number, bold state out of Word
    i = 0
    For Each para In doc.Paragraphs
        i = i + 1
        Set rng = para.Range
        pText(i) = CollapseWs(rng.Text)
        If Len(pText(i)) > 0 Then
            pPage(i) = rng.Information(3)        ' wdActiveEndPageNumber
            b = rng.Font.Bold
            If b = True Then
                pKind(i) = 2                     ' whole paragraph bold = heading
            ElseIf b <> False Then               ' mixed
                lead = CollapseWs(BoldPrefix(doc, rng))
                If Len(lead) >= Len(pText(i)) - 1 Then
                    pKind(i) = 2
                ElseIf SPLIT_RUNIN_HEADINGS And Len(lead) >= 3 Then
                    pKind(i) = 1: pLead(i) = lead ' bold lead-in, then body text
                End If
            End If
        End If
    Next para
    doc.Close 0
    Set doc = Nothing

    ' Pass 2: which pages carry the notes header; which bold lines repeat (page furniture)
    Dim notesPages As Object, repeats As Object, k As String, anyHeader As Boolean
    Set notesPages = CreateObject("Scripting.Dictionary")
    Set repeats = CreateObject("Scripting.Dictionary")
    For i = 1 To n
        If Len(pText(i)) > 0 Then
            If rxHeader.Test(pText(i)) Then notesPages(pPage(i)) = True
            If pKind(i) = 2 Then
                k = NormalizeKey(CleanHeading(pText(i)))
                repeats(k) = repeats(k) + 1
            End If
        End If
    Next i
    anyHeader = (notesPages.Count > 0)

    ' Pass 3: walk the notes pages, bucket text under the current bold heading
    Dim curHead As String, lastWasHeading As Boolean, body As String, nHead As Long
    curHead = INTRO_LABEL
    For i = 1 To n
        If Len(pText(i)) = 0 Then GoTo NextPara
        If anyHeader Then If Not notesPages.Exists(pPage(i)) Then GoTo NextPara
        If rxHeader.Test(pText(i)) Then GoTo NextPara
        If rxNoise.Test(pText(i)) Then GoTo NextPara
        Select Case pKind(i)
            Case 2
                If repeats(NormalizeKey(CleanHeading(pText(i)))) >= REPEAT_FURNITURE Then GoTo NextPara
                If lastWasHeading Then
                    curHead = curHead & " " & CleanHeading(pText(i))       ' heading wrapped over two lines
                Else
                    curHead = CleanHeading(pText(i)): nHead = nHead + 1
                End If
                lastWasHeading = True
            Case 1
                curHead = CleanHeading(pLead(i)): nHead = nHead + 1
                body = Trim$(Mid$(pText(i), Len(pLead(i)) + 1))
                body = CleanLeadPunct(body)
                If Len(body) > 0 Then AppendSection secs, curHead, body
                lastWasHeading = False
            Case Else
                AppendSection secs, curHead, pText(i)
                lastWasHeading = False
        End Select
NextPara:
    Next i

    remark = IIf(anyHeader, notesPages.Count & " notes page(s)", "NOTES HEADER NOT FOUND - whole document used") & _
             "; " & nHead & " bold heading(s)"
    Set ExtractSections = secs
End Function

' Longest all-bold prefix of a paragraph, found by binary search (few COM calls).
Private Function BoldPrefix(doc As Object, rng As Object) As String
    Dim s As Long, lo As Long, hi As Long, m As Long, L As Long
    s = rng.Start: L = rng.End - rng.Start
    If L < 1 Then Exit Function
    If doc.Range(s, s + 1).Font.Bold <> True Then Exit Function
    lo = 1: hi = L
    Do While lo < hi
        m = (lo + hi + 1) \ 2
        If doc.Range(s, s + m).Font.Bold = True Then lo = m Else hi = m - 1
    Loop
    BoldPrefix = doc.Range(s, s + lo).Text
End Function

Private Sub AppendSection(secs As Object, ByVal head As String, ByVal body As String)
    Dim k As String, v As Variant
    k = NormalizeKey(head)
    If Len(k) = 0 Then k = NormalizeKey(INTRO_LABEL): head = INTRO_LABEL
    If secs.Exists(k) Then
        v = secs(k)
        v(1) = v(1) & vbLf & body
        secs(k) = v
    Else
        secs.Add k, Array(head, body)
    End If
End Sub

Private Function CleanHeading(ByVal s As String) As String
    s = rxCont.Replace(s, "")
    s = Trim$(s)
    Do While Len(s) > 0
        If InStr(":-.", Right$(s, 1)) > 0 Or Right$(s, 1) = ChrW(8211) Or Right$(s, 1) = ChrW(8212) Then
            s = Trim$(Left$(s, Len(s) - 1))
        Else
            Exit Do
        End If
    Loop
    CleanHeading = s
End Function

Private Function CleanLeadPunct(ByVal s As String) As String
    Do While Len(s) > 0
        If InStr(":-.", Left$(s, 1)) > 0 Or Left$(s, 1) = ChrW(8211) Or Left$(s, 1) = ChrW(8212) Then
            s = Trim$(Mid$(s, 2))
        Else
            Exit Do
        End If
    Loop
    CleanLeadPunct = s
End Function

Private Function StartWord() As Boolean
    On Error GoTo Fail
    Set gWordApp = CreateObject("Word.Application")
    gWordApp.Visible = False
    gWordApp.DisplayAlerts = 0
    StartWord = True
    Exit Function
Fail:
    StartWord = False
End Function

Private Sub ShutdownWord()
    On Error Resume Next
    If Not gWordApp Is Nothing Then gWordApp.Quit 0
    Set gWordApp = Nothing
End Sub

'=====================================================================
'  COMPARE + SUMMARY (rebuilt from the Data sheets every run)
'=====================================================================
Private Sub BuildCompare(mapP As Object, mapC As Object, wsDP As Worksheet, wsDC As Worksheet)
    Dim heads As Object, dataP As Object, dataC As Object
    Dim wsCmp As Worksheet, wsSum As Worksheet
    Dim codes() As String, i As Long, code As String, r As Long, rs As Long
    Dim hk As Variant, col As Long, tP As String, tC As String, same As Boolean
    Dim hasP As Boolean, hasC As Boolean, nChg As Long, chgList As String
    Dim secP As Object, secC As Object, status As String, anyDiff As Boolean

    Set heads = CreateObject("Scripting.Dictionary")      ' key -> display, in column order
    CollectHeaders wsDP, heads
    CollectHeaders wsDC, heads
    Set dataP = ReadData(wsDP)
    Set dataC = ReadData(wsDC)
    codes = UnionSorted(mapP, mapC, dataP, dataC)

    Set wsCmp = FreshSheet("Compare")
    Set wsSum = FreshSheet("Summary")
    wsCmp.Cells.NumberFormat = "@"

    wsCmp.Range("A1:C1").Value = Array("Fund Code", "Row", "Info")
    col = 4
    For Each hk In heads.Keys
        wsCmp.Cells(1, col).Value = heads(hk)
        col = col + 1
    Next hk
    wsSum.Range("A1:I1").Value = Array("Fund Code", "Status", YEAR_PRIOR & " File", YEAR_CURRENT & " File", _
        "Sections " & YEAR_PRIOR, "Sections " & YEAR_CURRENT, "Sections Changed", "Changed Sections", "Remarks")

    r = 2: rs = 2
    For i = 0 To UBound(codes)
        code = codes(i)
        hasP = dataP.Exists(code): hasC = dataC.Exists(code)
        nChg = 0: chgList = "": anyDiff = False
        If hasP Then Set secP = dataP(code)(2) Else Set secP = Nothing
        If hasC Then Set secC = dataC(code)(2) Else Set secC = Nothing

        wsCmp.Cells(r, 1).Value = code: wsCmp.Cells(r, 2).Value = YEAR_CURRENT
        wsCmp.Cells(r + 1, 1).Value = code: wsCmp.Cells(r + 1, 2).Value = YEAR_PRIOR
        wsCmp.Cells(r + 2, 1).Value = code: wsCmp.Cells(r + 2, 2).Value = "Same?"
        wsCmp.Cells(r + 3, 1).Value = code: wsCmp.Cells(r + 3, 2).Value = "Changed " & YEAR_PRIOR & ">" & YEAR_CURRENT
        wsCmp.Cells(r, 3).Value = IIf(hasC, dataC(code)(0), IIf(mapC.Exists(code), "not extracted yet", "no " & YEAR_CURRENT & " file"))
        wsCmp.Cells(r + 1, 3).Value = IIf(hasP, dataP(code)(0), IIf(mapP.Exists(code), "not extracted yet", "no " & YEAR_PRIOR & " file"))

        col = 4
        For Each hk In heads.Keys
            tP = "": tC = ""
            If hasP Then If secP.Exists(hk) Then tP = secP(hk)
            If hasC Then If secC.Exists(hk) Then tC = secC(hk)
            If Len(tC) > 0 Then wsCmp.Cells(r, col).Value = tC
            If Len(tP) > 0 Then wsCmp.Cells(r + 1, col).Value = tP
            If hasP And hasC And (Len(tP) > 0 Or Len(tC) > 0) Then
                same = (NormalizeKey(tP) = NormalizeKey(tC))
                If same Then
                    wsCmp.Cells(r + 2, col).Value = "Yes"
                    wsCmp.Cells(r + 2, col).Interior.Color = RGB(198, 239, 206)
                Else
                    anyDiff = True: nChg = nChg + 1
                    chgList = JoinNonEmpty(chgList, heads(hk))
                    wsCmp.Cells(r + 2, col).Value = "No"
                    wsCmp.Cells(r + 2, col).Interior.Color = RGB(255, 199, 206)
                    If Len(tP) = 0 Then
                        wsCmp.Cells(r + 3, col).Value = "NEW IN " & YEAR_CURRENT
                    ElseIf Len(tC) = 0 Then
                        wsCmp.Cells(r + 3, col).Value = "REMOVED IN " & YEAR_CURRENT
                    Else
                        wsCmp.Cells(r + 3, col).Value = Left$(WordDiff(tP, tC), MAX_CELL)
                    End If
                    wsCmp.Cells(r + 3, col).Interior.Color = RGB(255, 235, 156)
                End If
            End If
            col = col + 1
        Next hk

        If hasP And hasC Then
            wsCmp.Cells(r + 2, 3).Value = IIf(anyDiff, "No", "Yes")
            wsCmp.Cells(r + 2, 3).Interior.Color = IIf(anyDiff, RGB(255, 199, 206), RGB(198, 239, 206))
            wsCmp.Cells(r + 3, 3).Value = nChg & " section(s) changed"
            status = IIf(anyDiff, "Changes found", "No text changes")
        Else
            wsCmp.Cells(r + 2, 3).Value = "n/a"
            If Not mapP.Exists(code) And Not hasP Then
                status = "Missing " & YEAR_PRIOR & " file"
            ElseIf Not mapC.Exists(code) And Not hasC Then
                status = "Missing " & YEAR_CURRENT & " file"
            Else
                status = "Not yet extracted"
            End If
        End If
        wsCmp.Range(wsCmp.Cells(r, 1), wsCmp.Cells(r, col - 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
        wsCmp.Range(wsCmp.Cells(r, 1), wsCmp.Cells(r + 3, 2)).Font.Bold = True

        With wsSum
            .Cells(rs, 1).Value = code
            .Cells(rs, 2).Value = status
            .Cells(rs, 3).Value = IIf(hasP, dataP(code)(0), "")
            .Cells(rs, 4).Value = IIf(hasC, dataC(code)(0), "")
            .Cells(rs, 5).Value = IIf(hasP, secP.Count, "")
            .Cells(rs, 6).Value = IIf(hasC, secC.Count, "")
            .Cells(rs, 7).Value = IIf(hasP And hasC, nChg, "")
            .Cells(rs, 8).Value = chgList
            .Cells(rs, 9).Value = JoinNonEmpty(IIf(hasP, YEAR_PRIOR & ": " & dataP(code)(1), ""), _
                                               IIf(hasC, YEAR_CURRENT & ": " & dataC(code)(1), ""))
            Select Case status
                Case "Changes found": .Cells(rs, 2).Interior.Color = RGB(255, 235, 156)
                Case "No text changes": .Cells(rs, 2).Interior.Color = RGB(198, 239, 206)
                Case Else: .Cells(rs, 2).Interior.Color = RGB(255, 199, 206)
            End Select
        End With
        r = r + 4: rs = rs + 1
    Next i

    ' formatting
    With wsCmp
        .Rows(1).Font.Bold = True: .Rows(1).Interior.Color = RGB(217, 225, 242)
        .Columns("A").ColumnWidth = 11: .Columns("B").ColumnWidth = 20: .Columns("C").ColumnWidth = 28
        If col > 4 Then .Range(.Cells(1, 4), .Cells(1, col - 1)).EntireColumn.ColumnWidth = 55
        .Cells.WrapText = True
        .Cells.VerticalAlignment = xlTop
        If r > 2 Then .Rows("2:" & (r - 1)).RowHeight = 90
        .Activate: .Range("D2").Select: ActiveWindow.FreezePanes = True
        If r > 2 Then .Range(.Cells(1, 1), .Cells(r - 1, col - 1)).AutoFilter
    End With
    With wsSum
        .Rows(1).Font.Bold = True: .Rows(1).Interior.Color = RGB(217, 225, 242)
        .Columns("A:G").AutoFit
        .Columns("H").ColumnWidth = 60: .Columns("I").ColumnWidth = 60
        .Columns("H:I").WrapText = True
        .Activate: .Range("A2").Select: ActiveWindow.FreezePanes = True
        If rs > 2 Then .Range("A1:I" & (rs - 1)).AutoFilter
    End With
End Sub

Private Sub CollectHeaders(ws As Worksheet, heads As Object)
    Dim c As Long, last As Long, k As String, disp As String
    last = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = FIRST_SECTION_COL To last
        disp = CStr(ws.Cells(1, c).Value)
        k = NormalizeKey(disp)
        If Len(k) > 0 Then If Not heads.Exists(k) Then heads.Add k, disp
    Next c
End Sub

' code -> Array(file, remark, dict(key -> text))
Private Function ReadData(ws As Worksheet) As Object
    Dim d As Object, secs As Object, v As Variant, r As Long, c As Long, lastR As Long, lastC As Long
    Dim code As String, keys() As String
    Set d = CreateObject("Scripting.Dictionary")
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastR < 2 Or lastC < FIRST_SECTION_COL Then Set ReadData = d: Exit Function
    v = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value
    ReDim keys(1 To lastC)
    For c = FIRST_SECTION_COL To lastC: keys(c) = NormalizeKey(CStr(v(1, c))): Next c
    For r = 2 To lastR
        code = UCase$(Trim$(CStr(v(r, 1))))
        If Len(code) > 0 And Not d.Exists(code) Then
            Set secs = CreateObject("Scripting.Dictionary")
            For c = FIRST_SECTION_COL To lastC
                If Len(keys(c)) > 0 And Len(CStr(v(r, c))) > 0 Then secs(keys(c)) = CStr(v(r, c))
            Next c
            d.Add code, Array(CStr(v(r, 2)), CStr(v(r, 4)), secs)
        End If
    Next r
    Set ReadData = d
End Function

'=====================================================================
'  WORD-LEVEL DIFF:  unchanged words as-is, [-removed-] {+added+}
'=====================================================================
Private Function WordDiff(ByVal oldT As String, ByVal newT As String) As String
    Dim a() As String, b() As String, n As Long, m As Long, i As Long, j As Long
    Dim L() As Long, out As String, remBuf As String, addBuf As String
    a = Split(CollapseWs(oldT), " "): b = Split(CollapseWs(newT), " ")
    n = UBound(a) + 1: m = UBound(b) + 1
    If CDbl(n) * CDbl(m) > MAX_WORDDIFF_CELLS Then
        WordDiff = "(section too long for word-level diff; compare the two text rows)": Exit Function
    End If
    ReDim L(0 To n, 0 To m)
    For i = n - 1 To 0 Step -1
        For j = m - 1 To 0 Step -1
            If NormalizeKey(a(i)) = NormalizeKey(b(j)) Then
                L(i, j) = L(i + 1, j + 1) + 1
            ElseIf L(i + 1, j) >= L(i, j + 1) Then
                L(i, j) = L(i + 1, j)
            Else
                L(i, j) = L(i, j + 1)
            End If
        Next j
    Next i
    i = 0: j = 0
    Do While i < n Or j < m
        If i < n And j < m Then
            If NormalizeKey(a(i)) = NormalizeKey(b(j)) Then
                FlushBufs out, remBuf, addBuf
                out = out & " " & b(j)
                i = i + 1: j = j + 1
                GoTo NextTok
            End If
        End If
        If j < m And (i = n Or L(i, j + 1) >= L(i + 1, j)) Then
            addBuf = addBuf & " " & b(j): j = j + 1
        Else
            remBuf = remBuf & " " & a(i): i = i + 1
        End If
NextTok:
    Loop
    FlushBufs out, remBuf, addBuf
    WordDiff = Trim$(out)
End Function

Private Sub FlushBufs(ByRef out As String, ByRef remBuf As String, ByRef addBuf As String)
    If Len(remBuf) > 0 Then out = out & " [-" & Trim$(remBuf) & "-]": remBuf = ""
    If Len(addBuf) > 0 Then out = out & " {+" & Trim$(addBuf) & "+}": addBuf = ""
End Sub

'=====================================================================
'  HELPERS
'=====================================================================
Private Sub InitRegex()
    Set rxCode = NewRegex(FUND_CODE_PATTERN, True, False, False)
    Set rxHeader = NewRegex(NOTES_HEADER_PATTERN, True, False, False)
    Set rxNoise = NewRegex(NOISE_LINE_PATTERN, True, False, False)
    Set rxNum = NewRegex("[\$\(]?\d[\d,\.]*\)?%?", True, False, True)
    Set rxWs = NewRegex("\s+", True, False, True)
    Set rxCont = NewRegex("\(?\s*continued\s*\)?", True, False, True)
End Sub

Private Function NewRegex(ByVal pat As String, ByVal ignoreCase As Boolean, ByVal multiLine As Boolean, ByVal globalMatch As Boolean) As Object
    Set NewRegex = CreateObject("VBScript.RegExp")
    NewRegex.Pattern = pat
    NewRegex.ignoreCase = ignoreCase
    NewRegex.multiLine = multiLine
    NewRegex.Global = globalMatch
End Function

Private Function NormalizeKey(ByVal s As String) As String
    s = LCase$(s)
    s = Replace(s, ChrW(8217), "'"): s = Replace(s, ChrW(8216), "'")
    s = Replace(s, ChrW(8220), """"): s = Replace(s, ChrW(8221), """")
    s = Replace(s, ChrW(8211), "-"): s = Replace(s, ChrW(8212), "-")
    s = Replace(s, ChrW(160), " ")
    If IGNORE_NUMBERS Then s = rxNum.Replace(s, "#")
    s = rxWs.Replace(s, " ")
    NormalizeKey = Trim$(s)
End Function

Private Function CollapseWs(ByVal s As String) As String
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " "): s = Replace(s, vbLf, " ")
    s = Replace(s, Chr$(12), " "): s = Replace(s, Chr$(11), " ")
    CollapseWs = Trim$(rxWs.Replace(s, " "))
End Function

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Private Function GetOrCreateSheet(ByVal nm As String) As Worksheet
    If SheetExists(nm) Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets(nm)
    Else
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateSheet.Name = nm
    End If
End Function

Private Function FreshSheet(ByVal nm As String) As Worksheet
    Application.DisplayAlerts = False
    If SheetExists(nm) Then ThisWorkbook.Worksheets(nm).Delete
    Application.DisplayAlerts = True
    Set FreshSheet = GetOrCreateSheet(nm)
End Function

Private Sub WriteLog()
    Dim ws As Worksheet, i As Long
    Set ws = FreshSheet("Log")
    ws.Cells(1, 1).Value = "Log (" & Format$(Now, "yyyy-mm-dd hh:nn") & ")"
    ws.Rows(1).Font.Bold = True
    For i = 1 To gLog.Count: ws.Cells(i + 1, 1).Value = gLog(i): Next i
    ws.Columns("A").ColumnWidth = 120
End Sub

Private Function PickFolder(ByVal title As String) As String
    With Application.FileDialog(4)   ' msoFileDialogFolderPicker
        .title = title
        .AllowMultiSelect = False
        If .Show = -1 Then PickFolder = .SelectedItems(1)
    End With
End Function

Private Function StripSlash(ByVal p As String) As String
    If Right$(p, 1) = "\" Then StripSlash = Left$(p, Len(p) - 1) Else StripSlash = p
End Function

Private Function FileNameOnly(ByVal p As String) As String
    If Len(p) = 0 Then Exit Function
    FileNameOnly = Mid$(p, InStrRev(p, "\") + 1)
End Function

Private Function JoinNonEmpty(ByVal a As String, ByVal b As String) As String
    If Len(Trim$(a)) = 0 Then
        JoinNonEmpty = Trim$(b)
    ElseIf Len(Trim$(b)) = 0 Then
        JoinNonEmpty = Trim$(a)
    Else
        JoinNonEmpty = Trim$(a) & "; " & Trim$(b)
    End If
End Function
