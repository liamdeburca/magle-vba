'===============================================================================
' Module: PlottingUtils
'===============================================================================
' Description:
'   Utilities for managing charts in Excel worksheets.
'   Provides functions to remove charts by title or clear all charts from
'   a worksheet.
'===============================================================================

'===============================================================================
' [SUB] RemoveExistingCharts
'===============================================================================
' Description:
'   Deletes all chart objects from the specified worksheet.
'
' Parameters:
'   targetSheet : Worksheet (Optional)
'       The worksheet to clear charts from. Defaults to ActiveSheet.
'===============================================================================
Public Sub RemoveExistingCharts( _
    Optional targetSheet As Worksheet _
)
    Dim ws As Worksheet
    If targetSheet Is Nothing Then
        Set ws = ActiveSheet
    Else
        Set ws = targetSheet
    End If
    
    Dim chtObj As ChartObject
    For Each chtObj In ws.ChartObjects
        chtObj.Delete
    Next chtObj
End Sub

'===============================================================================
' [SUB] RemoveChartByTitle
'===============================================================================
' Description:
'   Deletes a chart with the specified title from a worksheet.
'   Only removes the first matching chart found.
'
' Parameters:
'   chartTitle : String
'       The title of the chart to remove
'   targetSheet : Worksheet (Optional)
'       The worksheet to search. Defaults to ActiveSheet.
'===============================================================================
Public Sub RemoveChartByTitle( _
    chartTitle As String, _
    Optional targetSheet As Worksheet _
)
    Dim ws As Worksheet
    If targetSheet Is Nothing Then
        Set ws = ActiveSheet
    Else
        Set ws = targetSheet
    End If
    
    Dim chtObj As ChartObject
    For Each chtObj In ws.ChartObjects
        If chtObj.Chart.HasTitle Then
            If chtObj.Chart.chartTitle.Text = chartTitle Then
                chtObj.Delete
                Exit Sub
            End If
        End If
    Next chtObj
End Sub

'===============================================================================
' [MACRO] ClearPlots
'===============================================================================
' Description:
'   Removes all charts from the Plots worksheet.
'===============================================================================
Sub ClearPlots()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Plots")
    Call RemoveExistingCharts(ws)
End Sub

