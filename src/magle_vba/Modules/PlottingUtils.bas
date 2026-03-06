' Utilities for plotting data

' ===== REMOVAL OF EXISTING CHARTS =====

'===============================================================================
' [FUNCTION] RemoveExistingCharts
'===============================================================================
' Description:
'   Deletes all chart objects from the specified worksheet (or the active
'   sheet if none is provided). Used to clear a plot area before adding
'   fresh charts.
'
' Parameters:
'   targetSheet : Worksheet, Optional
'       The worksheet to clear. Defaults to the currently active sheet.
'===============================================================================
Public Function RemoveExistingCharts(Optional targetSheet As Worksheet)
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
End Function

'===============================================================================
' [FUNCTION] RemoveChartByTitle
'===============================================================================
' Description:
'   Searches the specified worksheet (or the active sheet) for a chart
'   whose title matches the given string and deletes it. Does nothing if
'   no matching chart is found.
'
' Parameters:
'   chartTitle : String
'       The exact title text of the chart to remove
'   targetSheet : Worksheet, Optional
'       The worksheet to search. Defaults to the currently active sheet.
'===============================================================================
Public Function RemoveChartByTitle(chartTitle As String, Optional targetSheet As Worksheet)
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
                Exit Function
            End If
        End If
    Next chtObj
End Function

'===============================================================================
' [MACRO] ClearPlots
'===============================================================================
' Description:
'   Removes all charts from the dedicated Plots worksheet. Called on
'   workbook close (Workbook_BeforeClose) or via the README sheet interface
'   to leave the workbook in a clean state.
'===============================================================================
Sub ClearPlots()
    ' Removes all charts from the Plots worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Plots")
    Call RemoveExistingCharts(ws)
End Sub

