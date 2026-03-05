'===============================================================================
' [FUNCTION] ConvertVariantToMask
'===============================================================================
' Description:
'   Converts a Variant array of boolean-like values to a typed Boolean array.
'   Used for processing mask parameters passed from worksheet ranges.
'
' Parameters:
'   mask : Variant
'       Array or collection of values that can be converted to Boolean
'
' Returns:
'   Boolean()
'       Typed Boolean array
'===============================================================================
Private Function ConvertVariantToMask( _
    mask As Variant _
) As Boolean()
    Dim maskCollection As New Collection
    
    Dim maskElement As Variant
    For Each maskElement In mask
        Call maskCollection.Add(CBool(maskElement))
    Next maskElement
    
    Dim newMask() As Boolean
    ReDim newMask(1 To maskCollection.count)
    
    Dim maskIdx As Long
    For maskIdx = 1 To maskCollection.count
        newMask(maskIdx) = maskCollection.Item(maskIdx)
    Next maskIdx
    
    ConvertVariantToMask = newMask
End Function

'===============================================================================
' [FUNCTION] PrepareData
'===============================================================================
' Description:
'   Prepares all data arrays needed for plotting. Applies optional masks,
'   scaling, and computes target/limit/statistic lines. Returns a Collection
'   containing DataRowCls instances and derived arrays.
'
' Parameters:
'   xDataRow : DataRowCls
'       Data row for x-axis values
'   yDataRow : DataRowCls
'       Data row for y-axis values
'   x2DataRow : Variant, Optional
'       Secondary x-axis data row
'   mask : Variant, Optional
'       Boolean mask to filter data points
'   yConversion : Double, Optional
'       Divisor for y-values (default: 1)
'   yMargins : Double, Optional
'       Fraction of y-range to add as padding (default: 0.05)
'   stat : String, Optional
'       Statistic to compute ("mean" or "median")
'
' Returns:
'   Collection
'       Contains: xDR, yDR, x2DR (DataRowCls), x, y, x2 (arrays),
'       yTarget, yLower, yUpper, yStat (arrays)
'===============================================================================
Private Function PrepareData( _
    xDataRow As DataRowCls, _
    yDataRow As DataRowCls, _
    Optional x2DataRow As Variant, _
    Optional mask As Variant, _
    Optional yConversion As Double = 1#, _
    Optional yMargins As Double = 0.05, _
    Optional stat As String = "" _
) As Collection
    Dim result As New Collection
    
    Dim xDR As DataRowCls
    Dim yDR As DataRowCls
    Dim x2DR As DataRowCls
    
    If Not IsMissing(mask) Or Not Application.WorksheetFunction.And(mask) Then
        Dim mask_() As Boolean
        mask_ = ConvertVariantToMask(mask)
        
        Set xDR = xDataRow.ApplyMask(mask_, inplace:=False)
        Set yDR = yDataRow.ApplyMask(mask_, inplace:=False)
        If Not IsMissing(x2DataRow) Then
            Set x2DR = x2DataRow.ApplyMask(mask_, inplace:=False)
        End If
    Else
        Set xDR = xDataRow
        Set yDR = yDataRow
        If Not IsMissing(x2DataRow) Then
            Set x2DR = x2DataRow
        End If
    End If
    
    Dim x() As String
    Dim y() As Variant
    Dim x2() As String
    x = xDR.txtData
    y = yDR.DblData
    
    If IsMissing(x2DataRow) Then
        ReDim x2(0)
    Else
        x2 = x2DR.txtData
    End If
    
    Dim i As Long
    For i = 1 To UBound(x)
        If x(i) = "" Then x(i) = "?"
        If Not IsMissing(x2DataRow) And x2(i) = "" Then x2(i) = "?"
    Next i
    
    For i = 1 To UBound(y)
        If Not IsError(y(i)) Then y(i) = CDbl(y(i)) / yConversion
    Next i
    
    Dim yLower() As Double
    If IsError(yDR.min) Then
        ReDim yLower(0)
    Else
        ReDim yLower(1 To UBound(y))
        yLower(1) = yDR.min / yConversion
        For i = 2 To UBound(y)
            yLower(i) = yLower(1)
        Next i
    End If
    
    Dim yUpper() As Double
    If IsError(yDR.max) Then
        ReDim yUpper(0)
    Else
        ReDim yUpper(1 To UBound(y))
        yUpper(1) = yDR.max / yConversion
        For i = 2 To UBound(y)
            yUpper(i) = yUpper(1)
        Next i
    End If
    
    Dim yTarget() As Double
    If IsError(yDR.target) Then
        ReDim yTarget(0)
    Else
        ReDim yTarget(1 To UBound(y))
        yTarget(1) = yDR.target / yConversion
        For i = 2 To UBound(y)
            yTarget(i) = yTarget(1)
        Next i
    End If
    
    Dim statValue As Variant
    Select Case Trim(LCase(stat))
        Case "mean"
            statValue = StatUtils.Mean(StatUtils.RemoveNA(y))
        Case "median"
            statValue = StatUtils.Quantile(StatUtils.RemoveNA(y), 0.5)
        Case ""
            statValue = CVErr(xlNAErr)
    End Select
    
    Dim yStat() As Double
    If IsError(statValue) Then
        ReDim yStat(0)
    Else
        ReDim yStat(1 To UBound(y))
        yStat(1) = CDbl(statValue)
        For i = 2 To UBound(y)
            yStat(i) = yStat(1)
        Next i
    End If
    
    Call result.Add(xDR, key:="xDR")
    Call result.Add(yDR, key:="yDR")
    If Not IsMissing(x2DataRow) Then Call result.Add(x2DR, key:="x2DR")
    
    Call result.Add(x, key:="x")
    Call result.Add(y, key:="y")
    If Not IsMissing(x2DataRow) Then Call result.Add(x2, key:="x2")
    
    Call result.Add(yTarget, key:="yTarget")
    Call result.Add(yLower, key:="yLower")
    Call result.Add(yUpper, key:="yUpper")
    Call result.Add(yStat, key:="yStat")
    
    Set PrepareData = result
End Function

'===============================================================================
' [FUNCTION] Run
'===============================================================================
' Description:
'   Creates a scatter plot chart visualizing y-values against x-values (batch
'   labels). Adds optional horizontal lines for target, min/max bounds, and
'   computed statistics. Supports secondary x-axis for dual labeling.
'
' Parameters:
'   xDataRow : DataRowCls
'       Data row containing x-axis labels (e.g., batch numbers)
'   yDataRow : DataRowCls
'       Data row containing y-axis measurement values
'   x2DataRow : Variant, Optional
'       Secondary x-axis data row (e.g., dates)
'   mask : Variant, Optional
'       Boolean mask to filter data points
'   targetSheet : String, Optional
'       Name of worksheet to place chart (default: active sheet)
'   yConversion : Double, Optional
'       Divisor for y-values (default: 1)
'   yMargins : Double, Optional
'       Fraction of y-range for axis padding (default: 0.05)
'   stat : String, Optional
'       Statistic line to add ("mean" or "median")
'   addGrid : Boolean, Optional
'       Whether to add gridlines (default: True)
'   addTitle : Boolean, Optional
'       Whether to add chart title (default: True)
'   addLegend : Boolean, Optional
'       Whether to add legend (default: True)
'===============================================================================
Function Run( _
    xDataRow As DataRowCls, _
    yDataRow As DataRowCls, _
    Optional x2DataRow As Variant, _
    Optional mask As Variant, _
    Optional targetSheet As String, _
    Optional yConversion As Double = 1#, _
    Optional yMargins As Double = 0.05, _
    Optional stat As String = "", _
    Optional addGrid As Boolean = True, _
    Optional addTitle As Boolean = True, _
    Optional addLegend As Boolean = True _
)
    Dim ws As Worksheet
    If IsMissing(targetSheet) Then
        Set ws = ThisWorkbook.ActiveSheet
    Else
        Set ws = ThisWorkbook.Worksheets(targetSheet)
    End If
    
    Dim preparedDataCollection As New Collection
    Set preparedDataCollection = PrepareData( _
        xDataRow, _
        yDataRow, _
        x2DataRow:=x2DataRow, _
        mask:=mask, _
        yConversion:=yConversion, _
        stat:=stat _
    )
    
    Dim x() As String
    x = preparedDataCollection.Item("x")
    
    Dim y() As Variant
    y = preparedDataCollection.Item("y")
    
    If Not IsMissing(x2DataRow) Then
        Dim x2() As String
        x2 = preparedDataCollection.Item("x2")
    End If
    
    Dim yTarget() As Double
    Dim yLower() As Double
    Dim yUpper() As Double
    Dim yStat() As Double
    yTarget = preparedDataCollection.Item("yTarget")
    yLower = preparedDataCollection.Item("yLower")
    yUpper = preparedDataCollection.Item("yUpper")
    yStat = preparedDataCollection.Item("yStat")
    
    Dim chtObj As ChartObject
    Set chtObj = ws.ChartObjects.Add( _
        Left:=60, Top:=40, Width:=500, Height:=320 _
    )
    
    Dim cht As Chart
    Set cht = chtObj.Chart
    cht.ChartType = xlLineMarkers
    
    Do While cht.SeriesCollection.count > 0
        cht.SeriesCollection(1).Delete
    Loop
    
    Dim serData As Series
    Set serData = cht.SeriesCollection.NewSeries
    With serData
        .name = "Data"
        .XValues = x
        .Values = y
        .ChartType = xlXYScatter
        .MarkerStyle = xlMarkerStyleSquare
        .MarkerSize = 7
        .MarkerForegroundColor = RGB(0, 0, 0)
        .MarkerBackgroundColor = RGB(0, 0, 0)
    End With
    
    If LBound(yTarget) = 1 Then
        Dim serTarget As Series
        Set serTarget = cht.SeriesCollection.NewSeries
        With serTarget
            .name = "Target"
            .XValues = x
            .Values = yTarget
            .ChartType = xlLine
            .Format.Line.ForeColor.RGB = RGB(0, 0, 255)
            .Format.Line.Weight = 1.5
            .Format.Line.DashStyle = msoLineSolid
            .MarkerStyle = xlMarkerStyleNone
        End With
    End If
    
    If LBound(yLower) = 1 Then
        Dim serLower As Series
        Set serLower = cht.SeriesCollection.NewSeries
        With serLower
            .name = "Lower bound"
            .XValues = x
            .Values = yLower
            .ChartType = xlLine
            .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
            .Format.Line.Weight = 1.5
            .Format.Line.DashStyle = msoLineSolid
            .MarkerStyle = xlMarkerStyleNone
        End With
    End If
    
    If LBound(yUpper) = 1 Then
        Dim serUpper As Series
        Set serUpper = cht.SeriesCollection.NewSeries
        With serUpper
            .name = "Upper bound"
            .XValues = x
            .Values = yUpper
            .ChartType = xlLine
            .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
            .Format.Line.Weight = 1.5
            .Format.Line.DashStyle = msoLineSolid
            .MarkerStyle = xlMarkerStyleNone
        End With
    End If
    
    If LBound(yStat) = 1 Then
        Dim serStat As Series
        Set serStat = cht.SeriesCollection.NewSeries
        With serStat
            .name = stat & " (" & Format(yStat(1), "#,##0.0") & ")"
            .XValues = x
            .Values = yStat
            .ChartType = xlLine
            .Format.Line.ForeColor.RGB = RGB(0, 255, 0)
            .Format.Line.Weight = 1.5
            .Format.Line.DashStyle = msoLineSolid
            .MarkerStyle = xlMarkerStyleNone
        End With
    End If
    
    If Not IsMissing(x2DataRow) Then
        Dim serX2 As Series
        Set serX2 = cht.SeriesCollection.NewSeries
        With serX2
            .name = ""
            .XValues = x2
            .Values = y
            .ChartType = xlLine
            .AxisGroup = xlSecondary
            .Format.Line.Visible = msoFalse
            .MarkerStyle = xlMarkerStyleNone
        End With
        
        With cht
            .HasAxis(xlCategory, xlSecondary) = True
            .Axes(xlCategory, xlSecondary).TickLabelPosition = _
                xlTickLabelPositionNextToAxis
            .HasAxis(xlValue, xlSecondary) = True
            .Axes(xlValue, xlSecondary).Delete
            .HasLegend = True
            .Legend.LegendEntries(cht.Legend.LegendEntries.count).Delete
        End With
    End If
    
    If addTitle Then
        cht.HasTitle = True
        cht.chartTitle.Text = yDataRow.key
        cht.chartTitle.Format.TextFrame2.TextRange.ParagraphFormat.Alignment _
            = msoAlignLeft
    End If
    
    If addLegend Then
        cht.HasLegend = True
        cht.Legend.Position = xlLegendPositionBottom
    
        Dim i As Long
        Dim lEntry As LegendEntry
        For i = cht.Legend.LegendEntries.count To 1 Step -1
            Set lEntry = cht.Legend.LegendEntries(i)
            If cht.SeriesCollection(i).name = "" Then lEntry.Delete
        Next i
    End If
    
    If addGrid Then
        cht.Axes(xlCategory).HasMajorGridlines = True
        cht.Axes(xlValue).HasMajorGridlines = True
    End If
    
    Dim xAxis As Axis
    Dim yAxis As Axis
    Set xAxis = cht.Axes(xlCategory)
    Set yAxis = cht.Axes(xlValue)
    
    If Not xAxis Is Nothing Then
        xAxis.HasTitle = True
        xAxis.AxisTitle.Text = xDataRow.name
    End If
    
    If Not yAxis Is Nothing Then
        yAxis.HasTitle = True
        
        If yDataRow.unit = "" Then
            yAxis.AxisTitle.Text = yDataRow.name
        Else
            yAxis.AxisTitle.Text = yDataRow.name & " (" & yDataRow.unit & ")"
        End If
    End If
    
    Dim ySorted() As Double
    ySorted = StatUtils.Sort(StatUtils.RemoveNA(y))
    
    Dim a As Double
    Dim b As Double
    a = ySorted(LBound(ySorted))
    b = ySorted(UBound(ySorted))
    
    If LBound(yLower) = 1 Then
        If yLower(1) < a Then a = yLower(1)
        If yLower(1) > b Then b = yLower(1)
    End If
    If LBound(yUpper) = 1 Then
        If yUpper(1) < a Then a = yUpper(1)
        If yUpper(1) > b Then b = yUpper(1)
    End If
    If LBound(yTarget) = 1 Then
        If yTarget(1) < a Then a = yTarget(1)
        If yTarget(1) > b Then b = yTarget(1)
    End If
    If LBound(yStat) = 1 Then
        If yStat(1) < a Then a = yStat(1)
        If yStat(1) > b Then b = yStat(1)
    End If
    
    Dim yRange As Double
    Dim yPadding As Double
    yRange = b - a
    yPadding = yMargins * yRange
    
    With cht.Axes(xlValue)
        .MinimumScale = a - yPadding
        .MaximumScale = b + yPadding
    End With
End Function