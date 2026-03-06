'===============================================================================
' [FUNCTION] ConvertVariantToMask
'===============================================================================
' Description:
'   Converts a Variant array of Boolean-compatible values into a typed
'   Boolean array (1-based). Used internally to normalise the optional
'   mask parameter passed to PrepareData and Run.
'
' Parameters:
'   mask : Variant
'       A Variant array whose elements can be coerced to Boolean
'
' Returns:
'   Boolean()
'       A 1-based Boolean array with the same values
'===============================================================================
Private Function ConvertVariantToMask( _
    mask As Variant _
) As Boolean()
    Dim maskCollection As New Collection
    Dim maskElement As Variant
    Dim maskIdx As Long
    
    For Each maskElement In mask
        Call maskCollection.Add(CBool(maskElement))
    Next maskElement
    
    Dim newMask() As Boolean
    ReDim newMask(1 To maskCollection.count)
    
    For maskIdx = 1 To maskCollection.count
        newMask(maskIdx) = maskCollection.Item(maskIdx)
    Next maskIdx
    
    ConvertVariantToMask = newMask
End Function

'===============================================================================
' [FUNCTION] PrepareData
'===============================================================================
' Description:
'   Assembles all arrays needed for a scatter plot into a single Collection.
'   Applies an optional Boolean mask to filter batches, scales y-values by
'   an optional conversion factor, replaces empty x labels with "?", and
'   computes optional horizontal reference lines (min, max, target, stat).
'
' Parameters:
'   xDataRow : DataRowCls
'       The row providing x-axis category labels (batch numbers)
'   yDataRow : DataRowCls
'       The row providing y-axis measurement values
'   x2DataRow : Variant, Optional
'       A second category row for a secondary x-axis (e.g. start dates)
'   mask : Variant, Optional
'       Boolean array to filter columns; True keeps the column
'   yConversion : Double, Optional
'       Divisor applied to all y-values and bounds. Default: 1.0
'   yMargins : Double, Optional
'       Fractional padding added to the y-axis range. Default: 0.05
'   stat : String, Optional
'       Statistic line to add: "mean", "median", or "" for none.
'       Default: ""
'
' Returns:
'   Collection
'       A Collection containing the following keyed arrays and objects:
'       "xDR", "yDR", ["x2DR"], "x", "y", ["x2"],
'       "yTarget", "yLower", "yUpper", "yStat"
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
    '' Prepares all data for plotting, returning a collection of double-arrays.
    Dim result As New Collection
    
    Dim xDR As DataRowCls, yDR As DataRowCls, x2DR As DataRowCls
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
    
    '' Load data
    Dim x() As String, y() As Variant, x2() As String
    x = xDR.txtData
    y = yDR.DblData
    
    If IsMissing(x2DataRow) Then
        ReDim x2(0)
    Else
        x2 = x2DR.txtData
    End If
    
    '' Replace missing x-axis labels
    Dim i As Long
    For i = 1 To UBound(x)
        If x(i) = "" Then x(i) = "?"
        If Not IsMissing(x2DataRow) And x2(i) = "" Then x2(i) = "?"
    Next i
    
    '' Apply y-value scaling
    For i = 1 To UBound(y)
        If Not IsError(y(i)) Then y(i) = CDbl(y(i)) / yConversion
    Next i
    
    '' Create array of min y-values
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
    '' Create array of max y-values
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
    '' Create array of target y-values
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
    '' Create array of statistic values
    Dim statValue As Variant
    Select Case Trim(LCase(stat))
        Case "mean"
            statValue = StatUtils.Mean(StatUtils.RemoveNA(y))
        Case "median"
            statValue = StatUtils.Quantile(StatUtils.RemoveNA(y), 0.5)
        Case ""
            ' Do nothing
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
    
    '' Add arrays to output
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
'   Creates a formatted scatter plot chart on the specified (or active)
'   worksheet. Plots y-values against x-axis category labels with optional
'   secondary x-axis, reference lines (target, min, max, statistic),
'   gridlines, legend, and axis titles derived from the DataRowCls metadata.
'
'   The y-axis scale is automatically fitted to include all data and
'   reference lines with configurable margin padding.
'
' Parameters:
'   xDataRow : DataRowCls
'       Row providing x-axis category labels (typically batch numbers)
'   yDataRow : DataRowCls
'       Row providing y-axis measurement values
'   x2DataRow : Variant, Optional
'       Row for a secondary x-axis (e.g. start dates shown above the chart)
'   mask : Variant, Optional
'       Boolean filter array to include only selected batches
'   targetSheet : String, Optional
'       Name of the worksheet to place the chart on. Defaults to the
'       active sheet.
'   yConversion : Double, Optional
'       Divisor applied to all y-values and reference lines. Default: 1.0
'   yMargins : Double, Optional
'       Fractional padding on the y-axis. Default: 0.05 (5 %)
'   stat : String, Optional
'       Horizontal statistic line: "mean", "median", or "". Default: ""
'   addGrid : Boolean, Optional
'       If True (default), major gridlines are shown on both axes
'   addTitle : Boolean, Optional
'       If True (default), the chart title is set to yDataRow.key
'   addLegend : Boolean, Optional
'       If True (default), a legend is shown at the bottom of the chart
'
' Example:
'   Call BasicScatterPlot.Run(batchRow, tempRow, _
'       x2DataRow:=dateRow, stat:="mean", _
'       targetSheet:="Plots")
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
    ' Basic plotting routine for plotting x vs y
    
    '' Choose target sheet
    Dim ws As Worksheet
    If IsMissing(targetSheet) Then
        Set ws = ThisWorkbook.ActiveSheet
    Else
        Set ws = ThisWorkbook.Worksheets(targetSheet)
    End If
    
    '' Prepare Data
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
    
    Dim yTarget() As Double, yLower() As Double, yUpper() As Double, yStat() As Double
    yTarget = preparedDataCollection.Item("yTarget")
    yLower = preparedDataCollection.Item("yLower")
    yUpper = preparedDataCollection.Item("yUpper")
    yStat = preparedDataCollection.Item("yStat")
    
    '' Instantiate chart
    Dim chtObj As ChartObject
    Set chtObj = ws.ChartObjects.Add( _
        Left:=60, Top:=40, Width:=500, Height:=320 _
    )
    Dim cht As Chart
    Set cht = chtObj.Chart
    cht.ChartType = xlLineMarkers
    
    '' Remove existing series
    Do While cht.SeriesCollection.count > 0
        cht.SeriesCollection(1).Delete
    Loop
    
    '' Basic scatterplot
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
    
    '' Horizontal lines
    If LBound(yTarget) = 1 Then
        ' Target exists
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
        ' Lower bound exists
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
        ' Upper bound exists
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
        ' Statistic exists
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
    '' Add secondary x-axis
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
        '' Show axis on top
        With cht
            .HasAxis(xlCategory, xlSecondary) = True
            .Axes(xlCategory, xlSecondary).TickLabelPosition = xlTickLabelPositionNextToAxis
            '' Hide secondary y-axis
            .HasAxis(xlValue, xlSecondary) = True
            .Axes(xlValue, xlSecondary).Delete
            '' Remove from legend
            .HasLegend = True
            .Legend.LegendEntries(cht.Legend.LegendEntries.count).Delete
        End With
    End If
    
    '' Add title
    If addTitle Then
        cht.HasTitle = True
        cht.chartTitle.Text = yDataRow.key
        cht.chartTitle.Format.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    End If
    
    '' Add legend
    If addLegend Then
        cht.HasLegend = True
        cht.Legend.Position = xlLegendPositionBottom
    
        Dim lEntry As LegendEntry
        For i = cht.Legend.LegendEntries.count To 1 Step -1
            Set lEntry = cht.Legend.LegendEntries(i)
            If cht.SeriesCollection(i).name = "" Then lEntry.Delete
        Next i
    End If
    '' Add gridlines
    If addGrid Then
        cht.Axes(xlCategory).HasMajorGridlines = True
        cht.Axes(xlValue).HasMajorGridlines = True
    End If
    
    '' ----- Format Axes: add axis names ----- ''
    Dim xAxis As Axis, yAxis As Axis
    
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
    
    '' ----- Format Axes: adjust bounds based on data ----- ''
    Dim ySorted() As Double
    ySorted = StatUtils.Sort(StatUtils.RemoveNA(y))
    
    Dim a As Double, b As Double
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
    
    Dim yRange As Double, yPadding As Double
    yRange = b - a
    yPadding = yMargins * yRange
    
    With cht.Axes(xlValue)
        .MinimumScale = a - yPadding
        .MaximumScale = b + yPadding
    End With
End Function

