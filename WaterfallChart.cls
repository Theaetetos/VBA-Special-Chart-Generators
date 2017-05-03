VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WaterfallChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ArrayDimension
    wcGuess = 0
    wcByRow = 1
    wcByColumn = 2
End Enum

Private raw_arr As Variant
Private header_arr As Variant
Private series_arr As Variant
Private wb_loc As Workbook
Private ws_loc As Worksheet
Private rng_loc As Range
Private pub_cht As Chart
Private ttl As String

Property Get ChangesArray() As Variant

ChangesArray = raw_arr

End Property

Property Let ChangesArray(Arr As Variant)

'takes as input an array of one or two dimensions and turns it into a two-dimensional array with one row with:
'   first and last elements being endpoints and middle elements being changes
'if a two-dimensional array is passed in, only the first row/column of the dimension with the
'   lower magnitude is read

'-------------------------------------------------------------------

'param checks:
    'Arr has 1 or 2 dimensions
Dim dimens As Integer
dimens = GetArrDimensions(Arr)
Select Case dimens
    Case Is < 1
        MsgBox _
            "Attempted to set ChangesArray property with a non-array value." & vbNewLine & _
            "Expected two-dimensional array."
        Exit Property
    Case Is > 2
        MsgBox _
            "Attempted to set ChangesArray property with multi-dimensional array." & vbNewLine & _
            "Expected two-dimensional array."
        Exit Property
End Select

'convert Arr
Dim i As Integer, corr As Integer
Select Case dimens
    Case Is = 1
        If LBound(Arr) = 0 Then corr = -1
        ReDim raw_arr(1 To 1, 1 To UBound(Arr) - LBound(Arr) + 1)
        For i = 1 To UBound(raw_arr, 2)
            raw_arr(1, 1) = Arr(i + corr)
        Next i
    Case Is = 2
        If LBound(Arr, 1) = 0 Then corr = -1
        Select Case UBound(Arr, 1)
            Case Is > UBound(Arr, 2)
                ReDim raw_arr(1 To 1, 1 To UBound(Arr, 1) - LBound(Arr, 1) + 1)
                For i = 1 To UBound(raw_arr, 2)
                    raw_arr(1, i) = Arr(i + corr, LBound(Arr, 2))
                Next i
            Case Else
                ReDim raw_arr(1 To 1, 1 To UBound(Arr, 2) - LBound(Arr, 2) + 1)
                For i = 1 To UBound(raw_arr, 2)
                    raw_arr(1, i) = Arr(LBound(Arr, 1), i + corr)
                Next i
        End Select
End Select

waterfallize

End Property

Property Get HeaderArray() As Variant

HeaderArray = header_arr

End Property

Property Let HeaderArray(Arr As Variant)

'ADD: param checks, cleaning


header_arr = Arr

End Property

Property Get SeriesArray() As Variant

SeriesArray = series_arr

End Property

Property Get WorkbookLocation() As Workbook

Set WorkbookLocation = wb_loc

End Property

Property Set WorkbookLocation(wb As Workbook)

Set wb_loc = wb

End Property

Property Get WorksheetLocation() As Worksheet

Set WorksheetLocation = ws_loc

End Property

Property Set WorksheetLocation(ws As Worksheet)

'ADD: param checks, cleaning


Set ws_loc = ws

End Property

Property Get RangeLocation() As Range

Set RangeLocation = rng_loc

End Property

Property Set RangeLocation(Rng As Range)

'ADD: param checks


Set rng_loc = Rng

End Property

Property Get PublishedChart() As Chart

Set PublishedChart = pub_cht

End Property

Property Set PublishedChart(cht As Chart)

Set pub_cht = cht

End Property

Property Get Title() As String

Title = ttl

End Property

Property Let Title(Str As String)

ttl = Str

End Property

Sub ConvertData(Arr As Variant, Optional ReadDimension As ArrayDimension = 0)

'splits a passed-in array of data into ChangesArray and HeaderArray
'   assumes that first row/column is headers and second is endpoints with changes in between
'   ReadDimension parameter allows user to specify which dimension series data is contained in:
'       ByRow means that data belonging to a series is organized by row; ByColumn means by column
'ChangesArray and HeaderArray are overwritten

'-----------------------------------------

'param checks: arr is one-dimensional and arr is >two-dimensional
Select Case GetArrDimensions(Arr)
    Case Is = 1
        MsgBox _
            "One-dimensional array passed into ConvertData." & vbNewLine & _
            "Expected two-dimensional array."
        Exit Sub
    Case Is > 2
        MsgBox _
            "Multi-dimensional array passed into ConvertData." & vbNewLine & _
            "Expected two-dimensional array."
End Select

'get dimension to read by if not specified (error out if neither dimension has a magnitude of two)
If ReadDimension = 0 Then
    If UBound(Arr, 1) = 2 Then
        ReadDimension = wcByRow
    ElseIf UBound(Arr, 2) = 2 Then
        ReadDimension = wcByColumn
    Else
        MsgBox _
            "Neither dimension of array passed into ConvertData has a magnitude of 2." & vbNewLine & _
            "Expected at least one dimension of magnitude 2."
        Exit Sub
    End If
End If

'split out headers and data
Dim i As Integer, corr As Integer
If LBound(Arr, 1) = 0 Then corr = -1
If ReadDimension = wcByRow Then
    ReDim header_arr(1 To 1, 1 To UBound(Arr, 2) - LBound(Arr, 2) + 1)
    For i = 1 To UBound(header_arr, 2)
        header_arr(1, i) = Arr(LBound(Arr, 1), i + corr)
    Next i
    ReDim raw_arr(1 To 1, 1 To UBound(Arr, 2) - LBound(Arr, 2) + 1)
    For i = 1 To UBound(raw_arr, 2)
        raw_arr(1, i) = Arr(LBound(Arr, 1) + 1, i + corr)
    Next i
ElseIf ReadDimension = wcByColumn Then
    ReDim header_arr(1 To 1, 1 To UBound(Arr, 1) - LBound(Arr, 1) + 1)
    For i = 1 To UBound(header_arr, 1)
        header_arr(1, i) = Arr(i + corr, LBound(Arr, 1))
    Next i
    ReDim raw_arr(1 To 1, 1 To UBound(Arr, 1) - LBound(Arr, 1) + 1)
    For i = 1 To UBound(series_arr, 1)
        raw_arr(1, i) = Arr(i + corr, LBound(Arr, 1) + 1)
    Next i
End If

waterfallize

End Sub

Sub PublishChart(Optional ByArray As Boolean = True)

'publishes a waterfall chart to the location stored in object variables
'param behavior:
'   if ByArray=True, then creates chart using stored arrays
'   if False, creates chart using stored range
'   will error out if necessary variables are missing

'--------------------------------------------------

'determine publishing location and create new chart at that location
'   WorksheetLocation and RangeLocation are checked to ensure they are not Nothing
'   If they are, chart is published to a new chart sheet
'   If RangeLocation is not Nothing and WorksheetLocation is, WorksheetLocation defaults to ActiveSheet
If wb_loc Is Nothing Then Set wb_loc = ThisWorkbook
If ws_loc Is Nothing Then
    If Not (rng_loc Is Nothing) Then Set ws_loc = ActiveSheet
    If SheetExists(ws_loc.Name, wb_loc.Name) = False Then
        MsgBox _
            Prompt:="Procedure terminated because of fatal input error:" & vbNewLine & _
                Chr(9) & "Worksheet: '" & ws_loc.Name & "' does not exist in workbook: '" & wb_loc.Name & ".'" & vbNewLine & _
                "Double-check object references.", _
            Buttons:=vbOKOnly
        Exit Sub
    End If
End If
Dim cht As Chart
Set cht = Charts.ADD
If Not (rng_loc Is Nothing) Then
    Set cht = cht.Location(xlLocationAsObject, ws_loc.Name)
    With cht.Parent
        .Top = rng_loc.Top
        .Left = rng_loc.Left
    End With
End If

'set chart characteristics
    'insert series data
set_series cht
    'formatting
        'general and x-axis
format_chart_general cht
        'y-axis
format_chart_y_axis cht

Set pub_cht = cht

End Sub

Sub UpdateSeries()

'updates an existing waterfall chart using current object variables
'errors out if PublishedChart is Nothing

'-----------------------------------

If pub_cht Is Nothing Then
    MsgBox "No chart object in scope."
    Exit Sub
End If

Dim i As Integer

'ensure user has set header_arr
On Error GoTo no_headers:
i = UBound(header_arr, 1)
On Error GoTo 0

'ensure user has set series_arr
On Error GoTo no_series:
i = UBound(series_arr, 1)
On Error GoTo 0

'update series data
set_series pub_cht
'formatting
    'general and x-axis
format_chart_general pub_cht
    'y-axis
format_chart_y_axis pub_cht

Exit Sub

no_headers:
MsgBox "No header data."
Exit Sub

no_series:
MsgBox "No series data."
Exit Sub

End Sub

Private Sub Class_Initialize()

Set wb_loc = ThisWorkbook

End Sub

Private Sub waterfallize()

'creates an array formatted to be used as a series array for a waterfall chart from ChangesArray and stores it in SeriesArray
'takes as input array with first and last elements being endpoints and middle elements being moves
'behavior-will accept array of any dimension as moves; will only read first list as determined by
'   checking which of first two dimensions has more elements

'--------------------------------

'determine read protocol for raw_arr
    'get number of dimensions
Dim dims As Integer
dims = GetArrDimensions(raw_arr)
    'get dimension with more elements if necessary
Dim read_dim As Integer
read_dim = 1
If dims > 1 Then
    If UBound(raw_arr, 1) < UBound(raw_arr, 2) Then read_dim = 2
End If

'create clean_arr to read raw_arr into to simplify later steps-only loop through a single dimension
    'determine dimensions of raw_arr
Dim col_ubound As Integer
col_ubound = UBound(raw_arr, read_dim)
If LBound(raw_arr, read_dim) = 0 Then col_ubound = col_ubound + 1 'handle base 0 case-a base 1 array is created with this variable
Dim raw_arr_lbound As Integer
raw_arr_lbound = LBound(raw_arr, read_dim)
    'fill clean_arr by looping through req elements in raw_arr
        'data is also converted to type Double in this step
Dim clean_arr() As Variant
ReDim clean_arr(1 To col_ubound)
Dim dbl As Double
Dim i As Integer, col_id As Integer
col_id = 1
If dims > 1 Then
    If read_dim = 1 Then
        For i = raw_arr_lbound To UBound(raw_arr, 1)
            If raw_arr(i, 1) = Empty Then
                dbl = 0
            Else
                dbl = raw_arr(i, 1)
            End If
            clean_arr(col_id) = dbl
            col_id = col_id + 1
        Next i
    Else
        For i = raw_arr_lbound To UBound(raw_arr, 2)
            If raw_arr(1, i) = Empty Then
                dbl = 0
            Else
                dbl = raw_arr(1, i)
            End If
            clean_arr(col_id) = dbl
            col_id = col_id + 1
        Next i
    End If
Else
    For i = raw_arr_lbound To UBound(raw_arr, 1)
        If raw_arr(i) = Empty Then
            dbl = 0
        Else
            dbl = raw_arr(i)
        End If
        clean_arr(col_id) = dbl
        col_id = col_id + 1
    Next i
End If

'fill return_arr
ReDim series_arr(1 To 7, 1 To col_ubound)
    'row 1-endpoints
series_arr(1, 1) = clean_arr(1)
For i = 2 To col_ubound
    series_arr(1, i) = 0
Next i
    'row 2-ending value
For i = 1 To col_ubound - 1
    series_arr(2, i) = 0
Next i
series_arr(2, col_ubound) = clean_arr(col_ubound)
    'row 4-positive movements
series_arr(4, 1) = 0
For i = 2 To col_ubound - 1
    If clean_arr(i) > 0 Then
        series_arr(4, i) = clean_arr(i)
    Else
        series_arr(4, i) = 0
    End If
Next i
series_arr(4, col_ubound) = 0
    'row 6-negative movements
series_arr(6, 1) = 0
For i = 2 To col_ubound - 1
    If clean_arr(i) < 0 Then
        series_arr(6, i) = -clean_arr(i)
    Else
        series_arr(6, i) = 0
    End If
Next i
series_arr(6, col_ubound) = 0
    'row 3-invisible part of column
    '   rows 5 and 7 are populated here, as well, and 4 and 6 adjusted
series_arr(3, 1) = 0
series_arr(5, 1) = 0
series_arr(7, 1) = 0
Dim pr_invis As Double, _
    pr_pos As Double, _
    pr_pos_adj As Double, _
    pr_neg As Double, _
    pr_neg_adj As Double, _
    last_abs_end, _
    cur_invis As Double, _
    cur_pos As Double, _
    cur_pos_adj As Double, _
    cur_neg As Double, _
    cur_neg_adj As Double, _
    cur_abs_end As Double
For i = 2 To col_ubound - 1
    pr_invis = series_arr(3, i - 1)
    If i = 2 Then
        If series_arr(1, 1) > 0 Then
            pr_pos = series_arr(1, 1)
            pr_neg = 0
        Else
            pr_neg = series_arr(1, 1)
            pr_pos = 0
        End If
    Else
        pr_pos = series_arr(4, i - 1)
        pr_neg = series_arr(6, i - 1)
    End If
    pr_pos_adj = series_arr(5, i - 1)
    pr_neg_adj = series_arr(7, i - 1)
    If i = 2 Then
        last_abs_end = series_arr(1, 1)
    Else
        If pr_invis > 0 Then
            If pr_pos > 0 Then
                last_abs_end = pr_invis + pr_pos
            Else
                last_abs_end = pr_invis
            End If
        Else
            If pr_neg_adj < 0 Then
                last_abs_end = pr_neg_adj
            ElseIf pr_pos_adj < 0 Then
                last_abs_end = pr_pos
            Else
                last_abs_end = pr_invis + pr_neg
            End If
        End If
    End If
    pr_invis = Abs(pr_invis)
    pr_pos = Abs(pr_pos)
    pr_pos_adj = Abs(pr_pos_adj)
    pr_neg = Abs(pr_neg)
    pr_neg_adj = Abs(pr_neg_adj)
    cur_pos = series_arr(4, i)
    cur_neg = series_arr(6, i)
    cur_abs_end = last_abs_end + cur_pos - cur_neg
    If last_abs_end < 0 Then
        If cur_abs_end < 0 Then
            If pr_pos > 0 Then
                If cur_pos > 0 Then
                    cur_invis = pr_invis - cur_pos
                Else
                    cur_invis = pr_invis
                End If
            Else
                If cur_pos > 0 Then
                    If pr_neg_adj > 0 Then
                        cur_invis = pr_neg_adj - cur_pos
                    Else
                        cur_invis = pr_invis + pr_neg - cur_pos
                    End If
                Else
                    If pr_neg_adj > 0 Then
                        cur_invis = pr_neg_adj
                    Else
                        cur_invis = pr_invis + pr_neg
                    End If
                End If
            End If
            cur_invis = -cur_invis
            series_arr(4, i) = -series_arr(4, i)
            series_arr(6, i) = -series_arr(6, i)
            cur_pos_adj = 0
            cur_neg_adj = 0
        Else
            cur_invis = 0
            cur_neg_adj = 0
            If pr_pos > 0 Then
                cur_pos_adj = pr_invis
            Else
                If pr_neg_adj > 0 Then
                    cur_pos_adj = pr_neg_adj
                Else
                    cur_pos_adj = pr_invis + pr_neg
                End If
            End If
            cur_pos_adj = -cur_pos_adj
            cur_pos = cur_pos + cur_pos_adj
            series_arr(4, i) = cur_pos
        End If
    Else
        If cur_abs_end < 0 Then
            cur_invis = 0
            cur_pos_adj = 0
            If pr_pos > 0 Then
                cur_neg_adj = cur_neg - pr_invis - pr_pos
            Else
                cur_neg_adj = cur_neg - pr_invis
            End If
            cur_neg_adj = -cur_neg_adj
            cur_neg = cur_neg + cur_neg_adj
            series_arr(6, i) = cur_neg
        Else
            If pr_pos > 0 Then
                If pr_pos_adj > 0 Then
                    cur_invis = pr_pos
                Else
                    If cur_pos > 0 Then
                        cur_invis = pr_invis + pr_pos
                    Else
                        cur_invis = pr_invis + pr_pos - cur_neg
                    End If
                End If
            Else
                If cur_pos > 0 Then
                    cur_invis = pr_invis
                Else
                    cur_invis = pr_invis - cur_neg
                End If
            End If
            cur_pos_adj = 0
            cur_neg_adj = 0
        End If
    End If
    series_arr(3, i) = cur_invis
    series_arr(5, i) = cur_pos_adj
    series_arr(7, i) = cur_neg_adj
Next i
series_arr(3, col_ubound) = 0
series_arr(5, col_ubound) = 0
series_arr(7, col_ubound) = 0

End Sub

Private Sub set_series(ByRef cht As Chart)

With cht
    'ensure correct number of series exist for cht
    Dim srs_ct As Integer
    srs_ct = .SeriesCollection.Count
    If srs_ct > 7 Then
        Do Until .SeriesCollection.Count = 7
            .SeriesCollection(8).Delete
        Loop
    ElseIf srs_ct < 7 Then
        Do Until .SeriesCollection.Count = 7
            .SeriesCollection.NewSeries
        Loop
    End If
    'transfer data in series_arr using temp arr
    Dim i As Integer, j As Integer
    For i = 1 To 7
        With .SeriesCollection(i)
            If LBound(series_arr, 1) = 0 Then i = i - 1 'handle base 0 case-operation is reversed at end of loop to avoid incorrect steps
            ReDim temp(LBound(series_arr, 2) To UBound(series_arr, 2))
            For j = LBound(series_arr, 2) To UBound(series_arr, 2)
                temp(j) = series_arr(i, j)
            Next j
            .Values = temp
            .XValues = header_arr
        End With
        If LBound(series_arr, 1) = 0 Then i = i + 1
    Next i
End With

End Sub

Private Sub format_chart_general(ByRef cht As Chart)

With cht
    .ChartType = xlColumnStacked
    .HasTitle = False 'kludge-don't know why it works or why it's necessary, but it is
    If ttl <> "" Then
        .HasTitle = True
        .ChartTitle.Text = ttl
    End If
    On Error Resume Next
    .Legend.Delete
    On Error GoTo 0
    .SeriesCollection(3).Format.Fill.Visible = False
    If (series_arr(1, 1) > 0 And series_arr(2, UBound(series_arr, 2)) > 0) Or _
    (series_arr(1, 1) < 0 And series_arr(2, UBound(series_arr, 2)) < 0) Then
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = .SeriesCollection(1).Format.Fill.ForeColor.RGB
    End If
    .SeriesCollection(5).Format.Fill.ForeColor.RGB = .SeriesCollection(4).Format.Fill.ForeColor.RGB
    .SeriesCollection(7).Format.Fill.ForeColor.RGB = .SeriesCollection(6).Format.Fill.ForeColor.RGB
    With .Axes(xlCategory)
        .TickLabels.Offset = 10
        .TickLabelPosition = xlLow
    End With
End With

End Sub

Private Sub format_chart_y_axis(ByRef cht As Chart)

Dim col_min As Double, col_max As Double
col_min = y_min()
col_max = y_max()
Dim digits As Integer
digits = Len(CStr(Round(col_max)))
col_max = Application.WorksheetFunction.RoundUp(col_max, 2 - digits)
digits = Len(CStr(Abs(Round(col_min))))
col_min = Application.WorksheetFunction.RoundDown(col_min, 2 - digits)

With cht
    With .Axes(xlValue)
        .MinimumScale = col_min
        col_min = -col_min
        If col_min > col_max Then
            digits = Len(CStr(col_min))
        Else
            digits = Len(CStr(col_max))
        End If
        Select Case digits
            Case Is >= 7
                Select Case .MajorUnit
                    Case Is >= 1000000
                        .TickLabels.NumberFormat = "[=0]0;[<>0]$#,##0,,""M"""
                    Case Is >= 1000
                        .TickLabels.NumberFormat = "[=0]0;[<>0]$#,##0.0,,""M"""
                    Case Else
                        .TickLabels.NumberFormat = "[=0]0;[<>0]$#,##0.00,,""M"""
                End Select
            Case Is >= 4
                If .MajorUnit >= 1000 Then
                    .TickLabels.NumberFormat = "[=0]0;[<>0]$#,##0,""K"""
                Else
                    .TickLabels.NumberFormat = "[=0]0;[<>0]$#,##0.0,""K"""
                End If
            Case Else
                .TickLabels.NumberFormat = "[=0]0;[<>0]$#,##0"
        End Select
        .MajorGridlines.Format.Line.Visible = False
    End With
End With

End Sub

Private Function GetArrDimensions(Arr As Variant) As Integer

'returns the number of dimensions of passed-in array

'----------------------------------------

Dim i As Long, dims As Integer
Dim done As Boolean
dims = 1

Do Until done = True
    On Error GoTo dims:
    i = UBound(Arr, dims)
    On Error GoTo 0
    If done = True Then Exit Do
    dims = dims + 1
Loop

Exit Function

dims:
done = True
dims = dims - 1
GetArrDimensions = dims

Resume Next

End Function

Private Function SheetExists(ws_name As String, Optional wb_name As String = "") As Boolean

'checks whether a worksheet of passed-in name exists in wb
'   wb optional; defaults to ThisWorkbook
'   wb_name used because of difficulty of passing var of type Workbook

'-------------------------------------------------------------------------------

Dim wb As Workbook

If wb_name = "" Then
    Set wb = ThisWorkbook
Else
    On Error GoTo wbNotFound:
    Set wb = Workbooks(wb_name)
    On Error GoTo 0
End If

Dim ws As Worksheet

For Each ws In wb.Worksheets
    If ws.Name = ws_name Then
        SheetExists = True
        Exit For
    End If
Next ws

Set ws = Nothing
Set wb = Nothing

Exit Function

wbNotFound:
MsgBox _
    Prompt:="No open workbook named '" & wb_name & "' found.", _
    Buttons:=vbOKOnly
Set ws = Nothing
Set wb = Nothing

End Function

Private Function y_max() As Double

Dim i As Integer, j As Integer
Dim col_height As Double
For i = LBound(series_arr, 2) To UBound(series_arr, 2)
    col_height = 0
    For j = 1 To 7
        If series_arr(j, i) > 0 Then col_height = col_height + series_arr(j, i)
    Next j
    If col_height > y_max Then y_max = col_height
Next i

End Function

Private Function y_min() As Double

Dim i As Integer
Dim col_height As Double
For i = LBound(series_arr, 2) + 1 To UBound(series_arr, 2) - 1
    If series_arr(3, i) > 0 Then
        col_height = series_arr(3, i)
    Else
        If series_arr(5, i) < 0 Or series_arr(7, i) < 0 Then
            col_height = series_arr(5, i) + series_arr(7, i)
        ElseIf series_arr(4, i) < 0 Or series_arr(6, i) < 0 Then
            col_height = series_arr(4, i) + series_arr(6, i)
        End If
    End If
    If i = 2 Then
        y_min = col_height
    Else
        If col_height < y_min Then y_min = col_height
    End If
Next i

End Function
