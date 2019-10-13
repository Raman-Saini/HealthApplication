Option Explicit
'===========================================================+
' Program title: CreateChart
' Description: Generates Chart To Show User The Weight Changes Over Time
'===========================================================+
Sub CreateChart()
    'Get Data From RecordSet
    Range("A1:E7").Select
    'Insert The Chart
    ActiveSheet.Shapes.AddChart2(227, xlLineMarkers).Select
    'Give The Chart Data
    ActiveChart.SetSourceData Source:=Range("Stats!$A$3:$D$7")
    'Give The Chart A Title
    ActiveChart.ChartTitle.Text = "Weight - Height - BMI"
    'ActiveChart.Name = "Chart1"
End Sub
