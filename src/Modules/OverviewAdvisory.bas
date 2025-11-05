Attribute VB_Name = "OverviewAdvisory"
Option Explicit

' =========================
' MAIN ENTRY POINT
' =========================
Public Sub Build_Overview_Advisory_Step1()
    Const SRC1 As String = "All Mandates (Beta)"
    Const SRC2 As String = "All Mandate (Beta)"   ' fallback if typo
    Const SH_OUT As String = "Overview"
    Const ROW_OUT_START As Long = 7

    ' source columns in All Mandates (Beta)
    Const COL_JOIN_LEFT As Long = 3     ' C
    Const COL_JOIN_RIGHT As Long = 5    ' E
    Const COL_TYPE As Long = 8          ' H  (mandate type)
    Const COL_PROFILE As Long = 28      ' AB (investment profile)

    Dim wsSrc As Worksheet, wsOut As Worksheet
    Dim lastRow As Long, r As Long
    Dim buf() As Variant
    Dim n As Long

    On Error GoTo FAIL

    Set wsSrc = GetSheetOrFail(Array(SRC1, SRC2))
    Set wsOut = GetSheetOrFail(Array(SH_OUT))

    lastRow = wsSrc.Cells(wsSrc.Rows.count, 1).End(xlUp).Row

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' clear output area and header rows 5–6
    ClearOutput wsOut, ROW_OUT_START

    If lastRow < 2 Then GoTo DONE

    ' collect advisory rows into memory: [1]=C&" "&E, [2]=Profile(AB)
    ReDim buf(1 To 2, 1 To 1)
    n = 0

    For r = 2 To lastRow
        If Trim$(LCase$(wsSrc.Cells(r, COL_TYPE).Value)) = "advisory mandate" Then
            n = n + 1
            If UBound(buf, 2) < n Then ReDim Preserve buf(1 To 2, 1 To n)
            buf(1, n) = Trim$(wsSrc.Cells(r, COL_JOIN_LEFT).Value) & " " & Trim$(wsSrc.Cells(r, COL_JOIN_RIGHT).Value) ' -> Overview!D
            buf(2, n) = wsSrc.Cells(r, COL_PROFILE).Value                                                                  ' -> Overview!Q
        End If
    Next r

    If n = 0 Then GoTo DONE

    ' sort by custom investment profile order
    CustomSortByProfile buf, 2, 1, n

    ' write to Overview: D=row7.., Q=row7..
    With wsOut
        .Range(.Cells(ROW_OUT_START, "D"), .Cells(ROW_OUT_START + n - 1, "D")).Value = ToColumn(buf, 1, n)
        .Range(.Cells(ROW_OUT_START, "Q"), .Cells(ROW_OUT_START + n - 1, "Q")).Value = ToColumn(buf, 2, n)
    End With

    ' build header + spacer before each profile block, starting with Azionario at rows 5–6
    BuildHeadersBeforeProfileBlocks wsOut, startRow:=ROW_OUT_START, profileCol:="Q", labelCol:="D", valueCol:="E"

DONE:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

FAIL:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Build_Overview_Advisory_Step1 failed: " & Err.Description, vbExclamation
End Sub

' =========================
' HELPERS
' =========================

Private Function GetSheetOrFail(names As Variant) As Worksheet
    Dim i As Long
    For i = LBound(names) To UBound(names)
        On Error Resume Next
        Set GetSheetOrFail = ThisWorkbook.Worksheets(CStr(names(i)))
        On Error GoTo 0
        If Not GetSheetOrFail Is Nothing Then Exit Function
    Next i
    Err.Raise vbObjectError + 701, , "Sheet not found: " & Join(names, " | ")
End Function

' Clears D,E,Q from startRow down and resets bars on rows 5–6
Private Sub ClearOutput(ws As Worksheet, startRow As Long)
    Dim lastRow As Long
    lastRow = Application.Max(ws.Cells(ws.Rows.count, "Q").End(xlUp).Row, ws.Cells(ws.Rows.count, "E").End(xlUp).Row)
    If lastRow < startRow Then lastRow = startRow

    ws.Range("D" & startRow & ":D" & lastRow).ClearContents
    ws.Range("E" & startRow & ":E" & lastRow).ClearContents
    ws.Range("Q" & startRow & ":Q" & lastRow).ClearContents

    ' remove any residual shading/bold in A:F (data area)
    ws.Range(ws.Cells(startRow, "A"), ws.Cells(lastRow, "F")).Interior.ColorIndex = xlNone
    ws.Range(ws.Cells(startRow, "A"), ws.Cells(lastRow, "F")).Font.ColorIndex = xlAutomatic
    ws.Range(ws.Cells(startRow, "A"), ws.Cells(lastRow, "F")).Font.Bold = False

    ' clear and reset rows 5–6
    ws.Range("A5:F6").Interior.ColorIndex = xlNone
    ws.Range("A5:F6").Font.ColorIndex = xlAutomatic
    ws.Range("A5:F6").Font.Bold = False
    ws.Range("D5").ClearContents
    ws.Range("E5").ClearContents
End Sub

Private Function ToColumn(buf As Variant, fieldIndex As Long, count As Long) As Variant
    Dim outArr() As Variant, i As Long
    ReDim outArr(1 To count, 1 To 1)
    For i = 1 To count
        outArr(i, 1) = buf(fieldIndex, i)
    Next i
    ToColumn = outArr
End Function

' Sort buf (2×N) by profile column using custom rank
Private Sub CustomSortByProfile(ByRef buf As Variant, ByVal sortCol As Long, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long
    Dim pivot As Long
    Dim t1 As Variant, t2 As Variant

    i = lo: j = hi
    pivot = ProfileRank(buf(sortCol, (lo + hi) \ 2))

    Do While i <= j
        Do While ProfileRank(buf(sortCol, i)) < pivot: i = i + 1: Loop
        Do While ProfileRank(buf(sortCol, j)) > pivot: j = j - 1: Loop
        If i <= j Then
            t1 = buf(1, i): t2 = buf(2, i)
            buf(1, i) = buf(1, j): buf(2, i) = buf(2, j)
            buf(1, j) = t1: buf(2, j) = t2
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then CustomSortByProfile buf, sortCol, lo, j
    If i < hi Then CustomSortByProfile buf, sortCol, i, hi
End Sub

Private Function ProfileRank(v As Variant) As Long
    Dim s As String
    s = LCase$(Trim$(CStr(v)))
    Select Case s
        Case "azionario": ProfileRank = 1
        Case "bilanciato": ProfileRank = 2
        Case "orientato al guadagno capitale": ProfileRank = 3
        Case "orientato al reddito": ProfileRank = 4
        Case Else: ProfileRank = 9999
    End Select
End Function

' Build headers BEFORE each profile block:
' - Row 5: dark bar "Azionario" in D5 and SUBTOTAL in E5; Row 6 light-grey spacer.
' - For each subsequent block: insert two rows above the block:
'   * first inserted row = dark bar with profile title in D and SUBTOTAL in E
'   * second inserted row = light-grey spacer
Private Sub BuildHeadersBeforeProfileBlocks(ws As Worksheet, ByVal startRow As Long, ByVal profileCol As String, ByVal labelCol As String, ByVal valueCol As String)
    Dim lastDataRow As Long
    Dim blocks As Collection
    Dim r As Long
    Dim curProfile As String
    Dim blockStart As Long

    lastDataRow = ws.Cells(ws.Rows.count, profileCol).End(xlUp).Row
    If lastDataRow < startRow Then Exit Sub

    Set blocks = New Collection

    blockStart = startRow
    curProfile = CStr(ws.Cells(startRow, profileCol).Value)

    For r = startRow + 1 To lastDataRow + 1
        If r > lastDataRow Or CStr(ws.Cells(r, profileCol).Value) <> curProfile Then
            blocks.Add Array(curProfile, blockStart, r - 1)
            If r <= lastDataRow Then
                blockStart = r
                curProfile = CStr(ws.Cells(r, profileCol).Value)
            End If
        End If
    Next r

    If blocks.count = 0 Then Exit Sub

    Dim firstBlk As Variant
    firstBlk = blocks(1)

    FormatHeaderRow ws, 5, True
    ws.Range(labelCol & "5").Value = CStr(firstBlk(0))
    WriteSubtotal ws, headerRow:=5, valueCol:=valueCol, dataStart:=CLng(firstBlk(1)), dataEnd:=CLng(firstBlk(2))
    FormatHeaderRow ws, 6, False

    Dim i As Long
    Dim offset As Long
    Dim insAt As Long
    Dim blk As Variant

    offset = 0

    For i = 2 To blocks.count
        blk = blocks(i)
        insAt = CLng(blk(1)) + offset

        ws.Rows(insAt).Resize(2).Insert shift:=xlDown

        FormatHeaderRow ws, insAt, True
        ws.Range(labelCol & insAt).Value = CStr(blk(0))
        WriteSubtotal ws, headerRow:=insAt, valueCol:=valueCol, _
                      dataStart:=insAt + 2, _
                      dataEnd:=CLng(blk(2)) + offset + 2

        FormatHeaderRow ws, insAt + 1, False

        offset = offset + 2
    Next i
End Sub

' Format a single header row across A:F
Private Sub FormatHeaderRow(ws As Worksheet, ByVal rowIx As Long, ByVal dark As Boolean)
    With ws.Range(ws.Cells(rowIx, "A"), ws.Cells(rowIx, "F"))
        If dark Then
            .Interior.Color = RGB(64, 64, 64)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
        Else
            .Interior.Color = RGB(217, 217, 217)
            .Font.Color = RGB(0, 0, 0)
            .Font.Bold = False
        End If
    End With
End Sub

' Write SUBTOTAL(9, valueCol[dataStart]:valueCol[dataEnd]) into headerRow
Private Sub WriteSubtotal(ws As Worksheet, ByVal headerRow As Long, ByVal valueCol As String, ByVal dataStart As Long, ByVal dataEnd As Long)
    Dim sep As String
    Dim colIndex As Long

    colIndex = ws.Columns(valueCol).Column
    sep = Application.International(xlListSeparator)

    If dataEnd < dataStart Then
        ws.Cells(headerRow, colIndex).ClearContents
        Exit Sub
    End If

    With ws.Cells(headerRow, colIndex)
        .Formula = "=SUBTOTAL(9," & ws.Range(ws.Cells(dataStart, colIndex), ws.Cells(dataEnd, colIndex)).Address(False, False) & ")"
        If sep <> "," Then .Formula = Replace(.Formula, ",", sep)
        If dataEnd >= dataStart Then .NumberFormat = ws.Cells(dataStart, colIndex).NumberFormat
    End With
End Sub

