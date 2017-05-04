VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LimitChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum lcArrayDimension
    lcGuess = 0
    lcByRow = 1
    lcByColumn = 2
End Enum

'below two factors are for calculating position and size parameters of percentage ovals
'   originally set by trial-and-error; may be changed if most common chart size/proportions change
Private Const oval_top_scalar_default As Double = 0.625
Private Const oval_height_scalar_default As Double = 0.125

'chart defaults to marking 90%<usage<100% as yellow
Private Const yellow_bound_default As Integer = 90
Private Const red_bound_default As Integer = 100

Private name_arr As Variant
Private lim_arr As Variant
Private usage_arr As Variant
Private wb_loc As Workbook
Private ws_loc As Worksheet
Private rng_loc As Range
Private pub_cht As Chart
Private ttl As String
Private oval_top_scalar As Double
Private oval_height_scalar As Double
Private yellow_bound As Integer
Private red_bound As Integer

Property Let NameArray(ByRef Arr As Variant)

'takes as input a 1d or 2d array and saves data as a 1d array to name_arr
'if Arr is 2d, only first row/column will be read

'---------------------------------------------------

Dim dimens As Integer
dimens = are_valid_dims(Arr, 1, 2, "Property Let NameArray")
If dimens = 0 Then Exit Property

to_1d Arr, name_arr, dimens

End Property

Property Get NameArray() As Variant

NameArray = name_arr

End Property

Property Let LimitArray(ByRef Arr As Variant)

'takes as input a 1d or 2d array and saves data as a 1d array to name_arr
'if Arr is 2d, only first row/column of larger dimension will be read

'---------------------------------------------------

Dim dimens As Integer
dimens = are_valid_dims(Arr, 1, 2, "Property Let LimitArray")
If dimens = 0 Then Exit Property

to_1d Arr, lim_arr, dimens

End Property

Property Get LimitArray() As Variant

LimitArray = lim_arr

End Property

Property Let UsageArray(Arr As Variant)

'takes as input a 1d or 2d array and saves data as a 1d array to name_arr
'if Arr is 2d, only first row/column of larger dimension will be read

'---------------------------------------------------

Dim dimens As Integer
dimens = are_valid_dims(Arr, 1, 2, "Property Let UsageArray")
If dimens = 0 Then Exit Property

to_1d Arr, usage_arr, dimens

End Property

Property Get UsageArray() As Variant

UsageArray = usage_arr

End Property

Property Set WorkbookLocation(wb As Workbook)

Set wb_loc = wb

End Property

Property Get WorkbookLocation() As Workbook

Set WorkbookLocation = wb_loc

End Property

Property Set WorksheetLocation(ws As Worksheet)

Set ws_loc = ws

End Property

Property Get WorksheetLocation() As Worksheet

Set WorksheetLocation = ws_loc

End Property

Property Set RangeLocation(Rng As Range)

Set rng_loc = Rng

End Property

Property Get RangeLocation() As Range

Set RangeLocation = rng_loc

End Property

Property Set PublishedChart(cht As Chart)

Set pub_cht = cht

End Property

Property Get PublishedChart() As Chart

Set PublishedChart = pub_cht

End Property

Property Let Title(Str As String)

ttl = Str

End Property

Property Get Title() As String

Title = ttl

End Property

Property Let OvalTopScalar(Dbl As Double)

If Dbl > 0 And Dbl < 1 Then
    oval_top_scalar = Dbl
End If

End Property

Property Get OvalTopScalar() As Double

OvalTopScalar = oval_top_scalar

End Property

Property Let OvalHeightScalar(Dbl As Double)

If Dbl > 0 And Dbl < 1 Then
    oval_height_scalar = Dbl
End If

End Property

Property Get OvalHeightScalar() As Double

OvalHeightScalar = oval_height_scalar

End Property

Property Let YellowBound(Intr As Integer)

yellow_bound = Intr

End Property

Property Get YellowBound() As Integer

YellowBound = yellow_bound

End Property

Property Let RedBound(Intr As Integer)

red_bound = Intr

End Property

Property Get RedBound() As Integer

RedBound = red_bound

End Property

Sub ReadData(ByRef Arr As Variant, Optional ReadDimension As lcArrayDimension = lcGuess)

'reads an array containing all data into respective arrays
'like setting source data on a chart
'assumes that data is in this order, regardless of orientation: 1) names, 2) limits, 3) usage

'---------------------------------------------

'ensure Arr has valid number of dimensions
Dim dimens As Integer
dimens = are_valid_dims(Arr, 2, 2, "Method ReadData")
If dimens = 0 Then Exit Sub

'get dimension to read by if not specified (error out if neither dimension has a magnitude of 3)
If ReadDimension = 0 Then
    If UBound(Arr, 1) = 3 Then
        ReadDimension = lcByRow
    ElseIf UBound(Arr, 2) = 3 Then
        ReadDimension = lcByColumn
    Else
        MsgBox _
            "Neither dimension of array passed into ReadData has a magnitude of 3." & vbNewLine & _
            "Expected at least one dimension of magnitude 3."
        Exit Sub
    End If
End If

'read Arr
Dim i As Integer, corr As Integer, l_bound As Integer, u_bound As Integer
l_bound = LBound(Arr, 1)
If l_bound = 0 Then corr = -1 'adjustment in case Arr is base 0
If ReadDimension = lcByColumn Then
    u_bound = UBound(Arr, 1)
    ReDim name_arr(1 To u_bound - l_bound + 1)
    ReDim lim_arr(1 To u_bound - l_bound + 1)
    ReDim usage_arr(1 To u_bound - l_bound + 1)
    For i = LBound(name_arr) To UBound(name_arr)
        name_arr(i) = Arr(i + corr, l_bound)
        lim_arr(i) = Arr(i + corr, l_bound + 1)
        usage_arr(i) = Arr(i + corr, l_bound + 2)
    Next i
Else
    u_bound = UBound(Arr, 2)
    ReDim name_arr(1 To u_bound - l_bound + 1)
    ReDim lim_arr(1 To u_bound - l_bound + 1)
    ReDim usage_arr(1 To u_bound - l_bound + 1)
    For i = LBound(name_arr) To UBound(name_arr)
        name_arr(i) = Arr(l_bound, i + corr)
        lim_arr(i) = Arr(l_bound + 1, i + corr)
        usage_arr(i) = Arr(l_bound + 2, i + corr)
    Next i
End If

End Sub

Sub PublishChart()

'creates limit utilization chart using class vars

'------------------------------------------------

'check for var errors
    'name_arr, lim_arr, and usage_arr all exist
If data_ready = False Then Exit Sub

    'target_ws is not in target_wb and target is provided
    'target_wb defaults to ThisWorkbook
    'target_ws defaults to ActiveSheet, but only if target is not nothing: if it is, then the chart is created on a new chart sheet
If wb_loc Is Nothing Then Set wb_loc = ThisWorkbook
If ws_loc Is Nothing Then
    If Not (rng_loc Is Nothing) Then Set ws_loc = ActiveSheet
    If sheet_exists(ws_loc.Name, wb_loc.Name) = False Then
        MsgBox _
            Prompt:="Procedure terminated because of fatal input error:" & vbNewLine & _
                    Chr(9) & "Worksheet: '" & ws_loc.Name & "' does not exist in workbook: '" & wb_loc.Name & ".'" & vbNewLine & _
                    "Double-check object references.", _
            Buttons:=vbOKOnly
        Exit Sub
    End If
End If

'create cht and put it in correct location
Dim cht As Chart
Set cht = Charts.ADD
If Not (rng_loc Is Nothing) Then
    Set cht = cht.Location(xlLocationAsObject, ws_loc.Name)
    With cht.Parent
        .Top = rng_loc.Top
        .Left = rng_loc.Left
    End With
End If

'populate chart
    'insert series data
set_series cht
    'formatting
format_cht cht
    'percentage ovals
make_perc_ovals cht

End Sub

Sub UpdateChart()

'updates pub_cht using current data arrays

'---------------------------

'check for var errors
    'name_arr, lim_arr, and usage_arr all exist
If data_ready = False Then Exit Sub

'update chart
    'insert new series data
set_series pub_cht
    'update percentage ovals
update_perc_ovals pub_cht

End Sub

Sub ResizeOvals()

'resizes percentage ovals if, e.g., chart was resized
'assumes percentage ovals are only shapes with chart as parent

'-------------------------------

Dim shp As Shape
Dim oval_prop_arr As Variant
Dim i As Integer
i = 1
With pub_cht
    For Each shp In .Shapes
        oval_prop_arr = oval_data(.PlotArea.Left, .PlotArea.Top, .PlotArea.Width, .PlotArea.Height, .Shapes.Count, i)
        With shp
            .Left = oval_prop_arr(0)
            .Top = oval_prop_arr(1)
            .Width = oval_prop_arr(2)
            .Height = oval_prop_arr(3)
        End With
        i = i + 1
    Next shp
End With

End Sub

Private Sub Class_Initialize()

oval_top_scalar = oval_top_scalar_default
oval_height_scalar = oval_height_scalar_default
yellow_bound = yellow_bound_default
red_bound = red_bound_default

End Sub

Private Sub to_1d(ByRef source_arr As Variant, ByRef dest_arr As Variant, dimens As Integer)

'reads source_arr into dest_arr
'source_arr may be 1 or 2d; dest is 1d
'source_arr may be base 0 or 1; dest is base 1
'reads by whichever dimension is larger (by rows if equal)

'---------------------------------------------------------------

Dim i As Integer, corr As Integer

If dimens = 1 Then
    If LBound(source_arr) = 0 Then corr = -1 'map indices to each other in case arrays have different bases
    ReDim dest_arr(1 To UBound(source_arr) - LBound(source_arr) + 1)
    For i = LBound(dest_arr) To UBound(dest_arr)
        dest_arr(i) = source_arr(i + corr)
    Next i
ElseIf dimens = 2 Then
    Dim rd_dim As Integer
    If LBound(source_arr, 1) = 0 Then corr = -1 'map indices to each other in case arrays have different bases
    'determine which dimension to read by
    Select Case UBound(source_arr, 1)
        Case Is > UBound(source_arr, 2)
            rd_dim = 2
        Case Is <= UBound(source_arr, 2)
            rd_dim = 1
    End Select
    'read array
    Select Case rd_dim
        Case Is = 2
            ReDim dest_arr(1 To UBound(source_arr, 1) - LBound(source_arr, 1) + 1)
            For i = LBound(dest_arr) To UBound(dest_arr)
                dest_arr(i) = source_arr(i + corr, LBound(source_arr, 2))
            Next i
        Case Is = 1
            ReDim dest_arr(1 To UBound(source_arr, 2) - LBound(source_arr, 2) + 1)
            For i = LBound(dest_arr) To UBound(dest_arr)
                dest_arr(i) = source_arr(LBound(source_arr, 1), i + corr)
            Next i
    End Select
End If

End Sub

Private Sub set_series(ByRef cht As Chart)

'sets SeriesCollection of passed-in chart in accordance with class vars

'--------------------------------------------------

With cht
'ensure correct number of series exist in chart
    Dim srs_ct As Integer
    srs_ct = .SeriesCollection.Count
    If srs_ct > 2 Then
        Do Until .SeriesCollection.Count = 2
            .SeriesCollection(3).Delete
        Loop
    ElseIf srs_ct < 2 Then
        Do Until .SeriesCollection.Count = 2
            .SeriesCollection.NewSeries
        Loop
    End If
    
'transfer data in class vars to series
    With .SeriesCollection(1)
        .Values = lim_arr
        .XValues = name_arr
    End With
    With .SeriesCollection(2)
        .Values = usage_arr
        .XValues = name_arr
    End With
End With

End Sub

Private Sub format_cht(ByRef cht As Chart)

'formats limit chart

'---------------------------

With cht
    
    .HasTitle = False 'kludge-don't know why it works or why it's necessary, but it is
    If ttl <> "" Then
        .HasTitle = True
        .ChartTitle.Text = ttl
    End If
    With .Axes(xlValue)
        .MajorGridlines.format.Line.Visible = False
        .Delete
    End With
    .ChartType = xlColumnClustered
    .ChartGroups(1).Overlap = 50
    
    With .SeriesCollection(1)
        .Name = "Limit"
        .format.Fill.ForeColor.RGB = RGB(217, 217, 217)
        .ApplyDataLabels
        With .DataLabels
            .NumberFormat = "$#,##0"
            With .format.TextFrame2.TextRange.Font
                .Fill.ForeColor.RGB = RGB(191, 191, 191)
                .Bold = True
            End With
        End With
    End With
    
    With .SeriesCollection(2)
        .Name = "Usage"
        .ApplyDataLabels
        With .DataLabels
            .NumberFormat = "$#,##0"
            With .format.TextFrame2.TextRange.Font
                .Fill.ForeColor.RGB = ThisWorkbook.Theme.ThemeColorScheme.Colors(msoThemeAccent2) 'kludge to match font color with column color
                .Bold = True
            End With
        End With
    End With
    
End With

End Sub

Private Sub make_perc_ovals(ByRef cht As Chart)

'creates and formats ovals displaying percent usage of limits

'----------------------------------

Dim oval As Shape
Dim oval_prop_arr As Variant
Dim pct As Double
Dim num_xvals As Integer, i As Integer, j As Integer
num_xvals = UBound(usage_arr) - LBound(usage_arr) + 1
j = 1

With cht
    For i = LBound(usage_arr) To UBound(usage_arr)
        pct = usage_arr(i) / lim_arr(i)
        pct = Round(pct, 2) * 100
        oval_prop_arr = oval_data(.PlotArea.Left, .PlotArea.Top, .PlotArea.Width, .PlotArea.Height, num_xvals, j)
        Set oval = .Shapes.AddShape(msoShapeOval, oval_prop_arr(0), oval_prop_arr(1), oval_prop_arr(2), oval_prop_arr(3))
        Erase oval_prop_arr
        With oval
            .Fill.ForeColor.RGB = RGB(217, 217, 217)
            Select Case pct
                Case Is <= yellow_bound
                    .Line.ForeColor.RGB = RGB(49, 145, 72)
                Case Is <= red_bound
                    .Line.ForeColor.RGB = RGB(255, 192, 0)
                Case Else
                    .Line.ForeColor.RGB = RGB(255, 0, 0)
            End Select
            With .TextFrame2
                .VerticalAnchor = msoAnchorMiddle
                With .TextRange
                    .ParagraphFormat.Alignment = msoAlignCenter
                    With .Characters
                        .Text = CStr(pct) & "%"
                        With .Font
                            .Bold = msoTrue
                            .Fill.ForeColor.RGB = ThisWorkbook.Theme.ThemeColorScheme.Colors(msoThemeAccent2) 'kludge to match percentage font color with column color
                        End With
                    End With
                End With
            End With
        End With
        j = j + 1
    Next i
End With

End Sub

Private Sub update_perc_ovals(ByRef cht As Chart)

'creates and formats ovals displaying percent usage of limits

'----------------------------------

Dim oval As Shape
Dim oval_prop_arr As Variant
Dim pct As Double
Dim num_xvals As Integer, i As Integer
num_xvals = UBound(usage_arr) - LBound(usage_arr) + 1
i = 1

With cht
    For Each oval In .Shapes
        pct = usage_arr(i) / lim_arr(i)
        pct = Round(pct, 2) * 100
        oval_prop_arr = oval_data(.PlotArea.Left, .PlotArea.Top, .PlotArea.Width, .PlotArea.Height, num_xvals, i)
        With oval
            .Left = oval_prop_arr(0)
            .Top = oval_prop_arr(1)
            .Width = oval_prop_arr(2)
            .Height = oval_prop_arr(3)
            Erase oval_prop_arr
            Select Case pct
                Case Is <= yellow_bound
                    .Line.ForeColor.RGB = RGB(49, 145, 72)
                Case Is <= red_bound
                    .Line.ForeColor.RGB = RGB(255, 192, 0)
                Case Else
                    .Line.ForeColor.RGB = RGB(255, 0, 0)
            End Select
            .TextFrame2.TextRange.Characters.Text = CStr(pct) & "%"
        End With
        i = i + 1
    Next oval
End With

End Sub

Private Function get_arr_dims(ByRef Arr As Variant) As Integer

'returns the number of dimensions of passed-in array

'----------------------------------------

Dim i As Long, dimens As Integer
Dim done As Boolean
dimens = 1

Do Until done = True
    On Error GoTo Dims:
    i = UBound(Arr, dimens)
    On Error GoTo 0
    If done = True Then Exit Do
    dimens = dimens + 1
Loop

Exit Function

Dims:
done = True
dimens = dimens - 1
get_arr_dims = dimens
Resume Next

End Function

Private Function are_valid_dims(ByRef Arr As Variant, l_bound As Integer, u_bound As Integer, caller As String) As Integer

'checks whether passed-in array has dimensions within given bounds (inclusive)
'returns 0 if not, number of dimensions if so
'meant to be used for error handling (see message to user in case returning 0)

'------------------------------------------------------------------------------------------

'param checks:
If l_bound < 0 Then
    MsgBox "is_valid_dims procedure encountered an error:" & vbNewLine & _
           "l_bound parameter is less than 0."
    Exit Function
End If

If l_bound > u_bound Then
    MsgBox "is_valid_dims procedure encountered an error:" & vbNewLine & _
           "l_bound parameter is greater than ubound parameter."
    Exit Function
End If

Dim dimens As Integer
dimens = get_arr_dims(Arr)

If dimens >= l_bound And dimens <= u_bound Then
    are_valid_dims = dimens
Else
    MsgBox caller & " procedure encountered an error:" & vbNewLine & _
           "Dimensions of array argument not within bounds."
End If

End Function

Private Function is_arr(ByRef Arr As Variant) As Boolean

'checks whether a passed-in variable is an array

'-------------------------------------------

If get_arr_dims(Arr) > 0 Then is_arr = True

End Function

Private Function data_ready() As Boolean

'returns True if name_arr, lim_arr, and usage_arr have data in them

'-------------------------------------------------

If is_arr(name_arr) = False Then
    MsgBox "No name data found."
    Exit Function
End If

If is_arr(lim_arr) = False Then
    MsgBox "No limit data found."
    Exit Function
End If

If is_arr(usage_arr) = False Then
    MsgBox "No usage data found."
    Exit Function
End If

data_ready = True

End Function

Private Function sheet_exists(ws_name As String, Optional wb_name As String = "") As Boolean

'checks whether a worksheet of passed-in name exists in wb
'   wb optional; defaults to ThisWorkbook
'   wb_name used because of difficulty of passing var of type Workbook

'---------------------------------------------------------------------------------

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
        sheet_exists = True
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

Private Function oval_data(plot_left As Double, _
                           plot_top As Double, _
                           plot_width As Double, _
                           plot_height As Double, num_xvals As Integer, oval_index As Integer) As Variant

'returns left, top, width, and height properties (in that order) for (oval_index)th oval on lim chart
'   based on passed-in plot area specs, number of xvalues in chart, and oval index (which xvalue it corresponds to)

'----------------------------------------------------------------

Dim return_arr(3) As Variant

Dim xval_width As Double
xval_width = plot_width / num_xvals

Dim oval_width As Double
oval_width = xval_width * 0.75

Dim oval_left As Double
oval_left = plot_left + ((xval_width - oval_width) / 2) + (xval_width * (oval_index - 1))

Dim oval_top As Double
oval_top = plot_top + (plot_height * oval_top_scalar)

Dim oval_height As Double
oval_height = plot_height * oval_height_scalar

return_arr(0) = oval_left
return_arr(1) = oval_top
return_arr(2) = oval_width
return_arr(3) = oval_height

oval_data = return_arr

End Function
