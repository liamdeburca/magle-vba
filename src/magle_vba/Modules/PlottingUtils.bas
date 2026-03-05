' Utilities for plotting data

' ===== REMOVAL OF EXISTING CHARTS =====
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

Sub ClearPlots()
    ' Removes all charts from the Plots worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Plots")
    Call RemoveExistingCharts(ws)
End Sub

