Attribute VB_Name = "modClientExportMain"
Option Explicit

'=========================================================
' Ribbon callbacks
'=========================================================
Public gRibbon As IRibbonUI

Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set gRibbon = ribbon
End Sub

Public Sub RunClientExportFromRibbon(control As IRibbonControl)
    ShowClientExportDialog
End Sub

'=========================================================
' Show settings dialog
'=========================================================
Public Sub ShowClientExportDialog()
    Dim frm As frmClientExportOptions
    Dim opts As ExportOptions

    Set frm = New frmClientExportOptions
    frm.Show vbModal

    If frm.Cancelled Then
        Unload frm
        Exit Sub
    End If

    opts = frm.GetOptions
    Unload frm

    CreateClientExportWithOptions opts
End Sub

'=========================================================
' Main export routine
'=========================================================
Public Sub CreateClientExportWithOptions(ByRef opts As ExportOptions)
    Dim srcWb As Workbook
    Dim outWb As Workbook
    Dim outPath As String

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Set srcWb = Application.ActiveWorkbook

    If srcWb Is Nothing Then
        Err.Raise vbObjectError + 1000, , "No active workbook found."
    End If

    If LCase$(Right$(srcWb.Name, 5)) = ".xlam" Then
        Err.Raise vbObjectError + 1001, , "The active workbook is the add-in itself. Open the workbook you want to export, then run the tool again."
    End If

    If Len(srcWb.Path) = 0 Then
        Err.Raise vbObjectError + 1002, , "Please save the source workbook first before creating a client export."
    End If

    Dim localPath As String
    localPath = ResolveLocalPath(srcWb.Path)

    If Len(localPath) = 0 Then
        Err.Raise vbObjectError + 1003, , _
            "Could not resolve a local file path for this workbook." & vbCrLf & _
            "If the file is stored in OneDrive, try opening it directly from File Explorer rather than from the OneDrive website."
    End If

    outPath = localPath & "\" & FileBaseName(srcWb.Name) & "_CLIENT_EXPORT.xlsx"

    If Len(Dir(outPath)) > 0 Then
        Kill outPath
    End If

    srcWb.SaveCopyAs outPath
    Set outWb = Workbooks.Open(outPath, UpdateLinks:=0, ReadOnly:=False)

    outWb.Worksheets(1).Activate

    'Always-on cleanup
    RemoveHiddenSheets_Safe outWb
    RemoveCommentsAndNotes outWb
    BreakExternalLinks outWb
    RemoveExternalConnectionsAndQueries outWb
    RemoveExternalDefinedNames outWb

    'User-selected processing
    HandleGroupedRowsAndColumns outWb, opts.GroupedDataAction, opts.GroupedVisibilityAction
    HandleHiddenRowsAndColumns outWb, opts.HiddenDataAction, opts.HiddenVisibilityAction

    'Flatten remaining formulas to values everywhere
    ConvertAllFormulasToValues outWb

    'Apply print-area/object cleanup
    DeleteOutsidePrintableArea_AllSheets outWb, opts.imagePolicy, opts.chartPolicy

    'Always-on cleanup
    RemoveAllHyperlinks outWb
    RemoveCustomDocumentProperties outWb
    RemoveCoreDocProps outWb

    outWb.Save
    outWb.Close False

    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Client export created:" & vbCrLf & outPath, vbInformation
    Exit Sub

CleanFail:
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    If Not outWb Is Nothing Then outWb.Close False
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Client export failed:" & vbCrLf & "Error " & errNum & ": " & errDesc, vbExclamation
End Sub

'=========================================================
' Core cleanup
'=========================================================
Private Sub ConvertAllFormulasToValues(ByVal wb As Workbook)
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.UsedRange.Cells.Count > 0 Then
            ws.UsedRange.Value = ws.UsedRange.Value
        End If
    Next ws
End Sub

Private Sub RemoveCommentsAndNotes(ByVal wb As Workbook)
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        On Error Resume Next
        ws.Cells.ClearComments
        ws.Cells.ClearNotes
        ws.CommentsThreaded.Delete
        On Error GoTo 0
    Next ws
End Sub

Private Sub RemoveHiddenSheets_Safe(ByVal wb As Workbook)
    Dim i As Long
    Dim ws As Worksheet

    If wb.ProtectStructure Then
        Err.Raise vbObjectError + 513, , _
            "Workbook structure is protected. Unprotect workbook structure first (Review > Protect Workbook)."
    End If

    wb.Worksheets(1).Activate

    For i = wb.Worksheets.Count To 1 Step -1
        Set ws = wb.Worksheets(i)
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
            ws.Delete
        End If
    Next i
End Sub

Private Sub BreakExternalLinks(ByVal wb As Workbook)
    Dim links As Variant
    Dim i As Long

    links = wb.LinkSources(Type:=xlLinkTypeExcelLinks)
    If IsEmpty(links) Then Exit Sub

    For i = LBound(links) To UBound(links)
        wb.BreakLink Name:=links(i), Type:=xlLinkTypeExcelLinks
    Next i
End Sub

Private Sub RemoveExternalConnectionsAndQueries(ByVal wb As Workbook)
    Dim cn As WorkbookConnection
    Dim q As Object

    On Error Resume Next
    For Each cn In wb.Connections
        cn.Delete
    Next cn
    On Error GoTo 0

    On Error Resume Next
    For Each q In wb.Queries
        q.Delete
    Next q
    On Error GoTo 0
End Sub

Private Sub RemoveExternalDefinedNames(ByVal wb As Workbook)
    Dim nm As Name
    Dim toDelete As Collection
    Dim i As Long

    Set toDelete = New Collection

    For Each nm In wb.Names
        If InStr(1, nm.RefersTo, "[", vbTextCompare) > 0 Then
            toDelete.Add nm.Name
        End If
    Next nm

    For i = 1 To toDelete.Count
        wb.Names(toDelete(i)).Delete
    Next i
End Sub

Private Sub RemoveAllHyperlinks(ByVal wb As Workbook)
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        ws.Hyperlinks.Delete
    Next ws
End Sub

Private Sub RemoveCustomDocumentProperties(ByVal wb As Workbook)
    Dim props As Object
    Dim p As Object

    On Error Resume Next
    Set props = wb.CustomDocumentProperties
    If props Is Nothing Then Exit Sub

    For Each p In props
        p.Delete
    Next p
    On Error GoTo 0
End Sub

Private Sub RemoveCoreDocProps(ByVal wb As Workbook)
    On Error Resume Next
    With wb.BuiltinDocumentProperties
        .Item("Author").Value = ""
        .Item("Company").Value = ""
        .Item("Manager").Value = ""
        .Item("Title").Value = ""
        .Item("Subject").Value = ""
        .Item("Keywords").Value = ""
        .Item("Comments").Value = ""
    End With
    On Error GoTo 0
End Sub

'=========================================================
' Grouped rows/columns handling
'=========================================================
Private Sub HandleGroupedRowsAndColumns(ByVal wb As Workbook, ByVal dataAction As eDataAction, ByVal visAction As eVisibilityAction)
    Dim ws As Worksheet
    Dim r As Long, c As Long
    Dim firstRow As Long, lastRow As Long
    Dim firstCol As Long, lastCol As Long

    For Each ws In wb.Worksheets
        GetUsedRangeBounds ws, firstRow, lastRow, firstCol, lastCol

        If lastRow >= firstRow Then
            For r = firstRow To lastRow
                If ws.Rows(r).OutlineLevel > 1 Then
                    ApplyActionToRowSlice ws, r, dataAction
                    If visAction = vaExpand Then
                        On Error Resume Next
                        ws.Rows(r).Hidden = False
                        On Error GoTo 0
                    End If
                End If
            Next r
        End If

        If lastCol >= firstCol Then
            For c = firstCol To lastCol
                If ws.Columns(c).OutlineLevel > 1 Then
                    ApplyActionToColumnSlice ws, c, dataAction
                    If visAction = vaExpand Then
                        On Error Resume Next
                        ws.Columns(c).Hidden = False
                        On Error GoTo 0
                    End If
                End If
            Next c
        End If
    Next ws
End Sub

'=========================================================
' Hidden rows/columns handling
' Manual hidden only, not grouped rows/cols
'=========================================================
Private Sub HandleHiddenRowsAndColumns(ByVal wb As Workbook, ByVal dataAction As eDataAction, ByVal visAction As eVisibilityAction)
    Dim ws As Worksheet
    Dim r As Long, c As Long
    Dim firstRow As Long, lastRow As Long
    Dim firstCol As Long, lastCol As Long

    For Each ws In wb.Worksheets
        GetUsedRangeBounds ws, firstRow, lastRow, firstCol, lastCol

        If lastRow >= firstRow Then
            For r = firstRow To lastRow
                If IsManualHiddenRow(ws, r) Then
                    ApplyActionToRowSlice ws, r, dataAction
                    If visAction = vaExpand Then
                        On Error Resume Next
                        ws.Rows(r).Hidden = False
                        On Error GoTo 0
                    End If
                End If
            Next r
        End If

        If lastCol >= firstCol Then
            For c = firstCol To lastCol
                If IsManualHiddenColumn(ws, c) Then
                    ApplyActionToColumnSlice ws, c, dataAction
                    If visAction = vaExpand Then
                        On Error Resume Next
                        ws.Columns(c).Hidden = False
                        On Error GoTo 0
                    End If
                End If
            Next c
        End If
    Next ws
End Sub

Private Function IsManualHiddenRow(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    On Error Resume Next
    IsManualHiddenRow = (ws.Rows(rowNum).Hidden = True And ws.Rows(rowNum).OutlineLevel <= 1)
    On Error GoTo 0
End Function

Private Function IsManualHiddenColumn(ByVal ws As Worksheet, ByVal colNum As Long) As Boolean
    On Error Resume Next
    IsManualHiddenColumn = (ws.Columns(colNum).Hidden = True And ws.Columns(colNum).OutlineLevel <= 1)
    On Error GoTo 0
End Function

Private Sub ApplyActionToRowSlice(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal dataAction As eDataAction)
    Dim rng As Range

    On Error Resume Next
    Set rng = Intersect(ws.UsedRange, ws.Rows(rowNum))
    On Error GoTo 0

    If rng Is Nothing Then Exit Sub
    ApplyDataAction rng, dataAction
End Sub

Private Sub ApplyActionToColumnSlice(ByVal ws As Worksheet, ByVal colNum As Long, ByVal dataAction As eDataAction)
    Dim rng As Range

    On Error Resume Next
    Set rng = Intersect(ws.UsedRange, ws.Columns(colNum))
    On Error GoTo 0

    If rng Is Nothing Then Exit Sub
    ApplyDataAction rng, dataAction
End Sub

Private Sub ApplyDataAction(ByVal rng As Range, ByVal dataAction As eDataAction)
    On Error Resume Next
    Select Case dataAction
        Case daClear
            rng.Clear
        Case daFlatten
            rng.Value = rng.Value
    End Select
    On Error GoTo 0
End Sub

Private Sub GetUsedRangeBounds(ByVal ws As Worksheet, ByRef firstRow As Long, ByRef lastRow As Long, ByRef firstCol As Long, ByRef lastCol As Long)
    Dim ur As Range

    Set ur = ws.UsedRange
    firstRow = ur.Row
    firstCol = ur.Column
    lastRow = ur.Row + ur.Rows.Count - 1
    lastCol = ur.Column + ur.Columns.Count - 1
End Sub

'=========================================================
' Print area + object cleanup
'=========================================================
Private Sub DeleteOutsidePrintableArea_AllSheets(ByVal wb As Workbook, ByVal imagePolicy As eObjectPolicy, ByVal chartPolicy As eObjectPolicy)
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        DeleteOutsidePrintableArea_OneSheet ws, imagePolicy, chartPolicy
    Next ws
End Sub

Private Sub DeleteOutsidePrintableArea_OneSheet(ByVal ws As Worksheet, ByVal imagePolicy As eObjectPolicy, ByVal chartPolicy As eObjectPolicy)
    Dim pr As String
    Dim pa As Range
    Dim ar As Range
    Dim hasPrintArea As Boolean
    Dim i As Long
    Dim shp As Shape
    Dim tl As Range
    Dim outside As Boolean
    Dim shapeKind As Long
    Dim deleteIt As Boolean

    pr = ws.PageSetup.PrintArea
    hasPrintArea = (Len(pr) > 0)

    If hasPrintArea Then
        Set pa = ws.Range(pr)

        ' Step 1: Clear all cells not covered by any print area.
        '         This handles orphan cells that sit at the intersection of
        '         a kept row and a kept column but belong to neither area
        '         (only possible when areas are offset in both dimensions).
        ClearCellsOutsidePrintAreas ws, pa

        ' Step 2 & 3: Delete every row / column not covered by any print area.
        '             Operating on whole rows/cols (not bounding box) means this
        '             works correctly for stacked, side-by-side, and offset layouts.
        DeleteUncoveredRows ws, pa
        DeleteUncoveredCols ws, pa
    End If

    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        shapeKind = GetShapeKind(shp)
        deleteIt = False

        On Error Resume Next
        Set tl = shp.TopLeftCell
        On Error GoTo 0

        outside = False
        If hasPrintArea Then
            If tl Is Nothing Then
                outside = True
            Else
                outside = True
                For Each ar In pa.Areas
                    If tl.Row >= ar.Row And _
                       tl.Row <= ar.Row + ar.Rows.Count - 1 And _
                       tl.Column >= ar.Column And _
                       tl.Column <= ar.Column + ar.Columns.Count - 1 Then
                        outside = False
                        Exit For
                    End If
                Next ar
            End If
        End If

        Select Case shapeKind
            Case 1 'image
                deleteIt = ShouldDeleteObject(imagePolicy, hasPrintArea, outside)
            Case 2 'chart
                deleteIt = ShouldDeleteObject(chartPolicy, hasPrintArea, outside)
            Case Else 'other shapes
                If hasPrintArea And outside Then deleteIt = True
        End Select

        If deleteIt Then shp.Delete
        Set tl = Nothing
    Next i
End Sub

'=========================================================
' Clear cells that are not inside any print area
'=========================================================
Private Sub ClearCellsOutsidePrintAreas(ByVal ws As Worksheet, ByVal pa As Range)
    Dim ur As Range
    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Sub

    Dim urFirstRow As Long: urFirstRow = ur.Row
    Dim urLastRow  As Long: urLastRow  = ur.Row + ur.Rows.Count - 1
    Dim urFirstCol As Long: urFirstCol = ur.Column
    Dim urLastCol  As Long: urLastCol  = ur.Column + ur.Columns.Count - 1

    Dim r As Long
    Dim ar As Range

    For r = urFirstRow To urLastRow
        ' Build the union of column ranges that any print area covers on this row
        Dim coveredCols As Range
        Set coveredCols = Nothing

        For Each ar In pa.Areas
            If r >= ar.Row And r <= ar.Row + ar.Rows.Count - 1 Then
                Dim arSlice As Range
                Set arSlice = ws.Range(ws.Cells(r, ar.Column), _
                                       ws.Cells(r, ar.Column + ar.Columns.Count - 1))
                If coveredCols Is Nothing Then
                    Set coveredCols = arSlice
                Else
                    Set coveredCols = Union(coveredCols, arSlice)
                End If
            End If
        Next ar

        If coveredCols Is Nothing Then
            ' No print area touches this row at all — clear the whole used row
            ws.Range(ws.Cells(r, urFirstCol), ws.Cells(r, urLastCol)).Clear
        Else
            ' Clear column gaps: before first covered col, between covered ranges, after last
            Dim nAr As Long: nAr = coveredCols.Areas.Count
            Dim cF() As Long, cL() As Long
            ReDim cF(1 To nAr), cL(1 To nAr)
            Dim ai As Long
            For ai = 1 To nAr
                cF(ai) = coveredCols.Areas(ai).Column
                cL(ai) = coveredCols.Areas(ai).Column + coveredCols.Areas(ai).Columns.Count - 1
            Next ai
            SortExtentPairs cF, cL, nAr

            If cF(1) > urFirstCol Then
                ws.Range(ws.Cells(r, urFirstCol), ws.Cells(r, cF(1) - 1)).Clear
            End If
            For ai = 1 To nAr - 1
                If cL(ai) + 1 <= cF(ai + 1) - 1 Then
                    ws.Range(ws.Cells(r, cL(ai) + 1), ws.Cells(r, cF(ai + 1) - 1)).Clear
                End If
            Next ai
            If cL(nAr) < urLastCol Then
                ws.Range(ws.Cells(r, cL(nAr) + 1), ws.Cells(r, urLastCol)).Clear
            End If
        End If
    Next r
End Sub

'=========================================================
' Delete every row not covered by any print area
'=========================================================
Private Sub DeleteUncoveredRows(ByVal ws As Worksheet, ByVal pa As Range)
    Dim n As Long: n = pa.Areas.Count
    Dim rF() As Long, rL() As Long
    ReDim rF(1 To n), rL(1 To n)
    Dim i As Long
    For i = 1 To n
        rF(i) = pa.Areas(i).Row
        rL(i) = pa.Areas(i).Row + pa.Areas(i).Rows.Count - 1
    Next i
    SortExtentPairs rF, rL, n

    ' Delete bottom-to-top so indices stay valid
    If rL(n) < ws.Rows.Count Then
        ws.Rows((rL(n) + 1) & ":" & ws.Rows.Count).Delete
    End If
    For i = n - 1 To 1 Step -1
        If rL(i) + 1 <= rF(i + 1) - 1 Then
            ws.Rows((rL(i) + 1) & ":" & (rF(i + 1) - 1)).Delete
        End If
    Next i
End Sub

'=========================================================
' Delete every column not covered by any print area
'=========================================================
Private Sub DeleteUncoveredCols(ByVal ws As Worksheet, ByVal pa As Range)
    Dim n As Long: n = pa.Areas.Count
    Dim cF() As Long, cL() As Long
    ReDim cF(1 To n), cL(1 To n)
    Dim i As Long
    For i = 1 To n
        cF(i) = pa.Areas(i).Column
        cL(i) = pa.Areas(i).Column + pa.Areas(i).Columns.Count - 1
    Next i
    SortExtentPairs cF, cL, n

    ' Delete right-to-left so indices stay valid
    If cL(n) < ws.Columns.Count Then
        ws.Columns(ColLetter(cL(n) + 1) & ":" & ColLetter(ws.Columns.Count)).Delete
    End If
    For i = n - 1 To 1 Step -1
        If cL(i) + 1 <= cF(i + 1) - 1 Then
            ws.Columns(ColLetter(cL(i) + 1) & ":" & ColLetter(cF(i + 1) - 1)).Delete
        End If
    Next i
End Sub

'=========================================================
' Sort two parallel Long arrays by the first array ascending
'=========================================================
Private Sub SortExtentPairs(ByRef arrF() As Long, ByRef arrL() As Long, ByVal n As Long)
    Dim i As Long, j As Long, tmpF As Long, tmpL As Long
    For i = 1 To n - 1
        For j = i + 1 To n
            If arrF(j) < arrF(i) Then
                tmpF = arrF(i): arrF(i) = arrF(j): arrF(j) = tmpF
                tmpL = arrL(i): arrL(i) = arrL(j): arrL(j) = tmpL
            End If
        Next j
    Next i
End Sub

Private Function ShouldDeleteObject(ByVal policy As eObjectPolicy, ByVal hasPrintArea As Boolean, ByVal outsidePrintArea As Boolean) As Boolean
    Select Case policy
        Case opDeleteAll
            ShouldDeleteObject = True
        Case opKeepAll
            ShouldDeleteObject = False
        Case opKeepInsidePrintArea
            If hasPrintArea Then
                ShouldDeleteObject = outsidePrintArea
            Else
                ShouldDeleteObject = False
            End If
    End Select
End Function

'1 = image, 2 = chart, 0 = other
Private Function GetShapeKind(ByVal shp As Shape) As Long
    On Error Resume Next

    If shp.Type = msoChart Then
        GetShapeKind = 2
    ElseIf shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
        GetShapeKind = 1
    Else
        GetShapeKind = 0
    End If

    On Error GoTo 0
End Function

'=========================================================
' Helpers
'=========================================================
Private Function FileBaseName(ByVal fn As String) As String
    Dim p As Long
    p = InStrRev(fn, ".")
    If p > 0 Then
        FileBaseName = Left$(fn, p - 1)
    Else
        FileBaseName = fn
    End If
End Function

Private Function ColLetter(ByVal colNum As Long) As String
    ColLetter = Split(Application.Cells(1, colNum).Address(True, False), "$")(0)
End Function

'=========================================================
' Resolve cloud/OneDrive paths to a local file system path
' Returns empty string if no local path can be determined
'=========================================================
Private Function ResolveLocalPath(ByVal wbPath As String) As String
    ' Already a local path
    If Left$(wbPath, 4) <> "http" Then
        ResolveLocalPath = wbPath
        Exit Function
    End If

    ' OneDrive Personal: https://d.docs.live.net/<cid>/...
    ' OneDrive for Business: https://<tenant>-my.sharepoint.com/...
    ' Try each known OneDrive environment variable and map the URL segment
    Dim envVars(2) As String
    envVars(0) = Environ("OneDriveCommercial")
    envVars(1) = Environ("OneDriveConsumer")
    envVars(2) = Environ("OneDrive")

    Dim i As Long
    Dim localRoot As String
    Dim urlAfterRoot As String
    Dim candidate As String

    ' Strip scheme and host to get the path portion of the URL
    ' e.g. https://d.docs.live.net/abc123/Folder/Sub -> /Folder/Sub
    Dim urlPath As String
    urlPath = wbPath
    Dim slashPos As Long
    ' Remove https://
    If Left$(urlPath, 8) = "https://" Then urlPath = Mid$(urlPath, 9)
    If Left$(urlPath, 7) = "http://" Then urlPath = Mid$(urlPath, 8)
    ' Remove host (everything up to first /)
    slashPos = InStr(urlPath, "/")
    If slashPos > 0 Then urlPath = Mid$(urlPath, slashPos)  ' now "/Folder/Sub/..."

    ' Replace forward slashes with backslashes and strip leading slash
    urlPath = Mid$(urlPath, 2)  ' remove leading /
    urlPath = Join(Split(urlPath, "/"), "\")

    ' For each OneDrive root, try to build a candidate local path.
    ' The URL path typically contains a folder segment that mirrors the local tree.
    ' Strategy: walk from the deepest URL segment upward and look for a matching local folder.
    For i = 0 To 2
        localRoot = envVars(i)
        If Len(localRoot) = 0 Then GoTo NextVar

        ' Attempt: localRoot + last N segments of urlPath
        Dim parts() As String
        Dim j As Long
        parts = Split(urlPath, "\")
        For j = 0 To UBound(parts)
            Dim seg As String
            seg = parts(j)
            Dim k As Long
            For k = j + 1 To UBound(parts)
                seg = seg & "\" & parts(k)
            Next k
            candidate = localRoot & "\" & seg
            If Len(Dir(candidate, vbDirectory)) > 0 Then
                ResolveLocalPath = candidate
                Exit Function
            End If
        Next j

NextVar:
    Next i

    ' Could not resolve
    ResolveLocalPath = ""
End Function

