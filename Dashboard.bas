Attribute VB_Name = "dashboard"
Option Explicit

Public Sub dashboard()
    '
    ' Sets up dashboard of statistics
    '
    
    Call setup_background
    Call setup_names
    Call setup_pivottable

End Sub

Public Sub setup_background()
    '
    ' Sets up background of the dashboard
    '
    
    Dim fileName As String
    Dim sheet As Worksheet

    ' Creates dashboard sheet
    Set sheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook _
                .Worksheets(ActiveWorkbook.Worksheets.Count))
    ActiveSheet.Name = "Dashboard"

    ' Sets background on dashboard sheet
    fileName = "C:\Users\Festa\Desktop\vba macros\background.jpg"
    Sheets("Dashboard").SetBackgroundPicture (fileName)

    ' Erases gridlines on dashboard sheet
    ActiveWindow.DisplayGridlines = False

End Sub

Public Sub setup_names()
    '
    ' Sets up title and option menu
    '
    
    Dim height As Integer
    Dim i As Integer
    Dim options As Variant
    Dim mac_but As Variant
    
    options = Array("Results", "Scores", "Possession", _
                    "Captains", "Expected Goals", "Venue")
    
    mac_but = Array("results_tab", "scores_tab", _
                    "possession_tab", "captains_tab", _
                    "expected_tab", "venue_tab")
    
    ' Creates title for the dashboard
    Sheets("Dashboard").Shapes.AddTextbox _
        (msoTextOrientationHorizontal, _
        430, 20, 262.3750393701, 90.6250393701).Select
    
    With Selection.ShapeRange
        .Fill.Visible = False
        .line.Visible = False
        .TextFrame2.TextRange.Font.Size = 32
        .TextFrame2.TextRange.ParagraphFormat.Alignment = _
            msoAlignCenter
        .TextFrame2.TextRange.Characters.Text = _
            Worksheets(1).Name
        .TextFrame2.TextRange.Font.Fill.ForeColor. _
            ObjectThemeColor = msoThemeColorBackground1
        .TextFrame2.TextRange.Font.Name = "Helvetica"
    End With
    
    ' Creates buttons for the dashboard
    height = 150
    For i = 0 To 5
        Sheets("Dashboard").Shapes.AddShape _
            (msoShapeRoundedRectangle, _
            71.25, height, 96.75, 27).Select
      
        With Selection.ShapeRange
            .Fill.ForeColor.ObjectThemeColor = _
                msoThemeColorAccent6
            .line.ForeColor.ObjectThemeColor = _
                msoThemeColorAccent6
            .TextFrame2.TextRange.Characters.Text = options(i)
            .TextFrame2.TextRange.Font.Size = 14
            .TextFrame2.TextRange.ParagraphFormat.Alignment _
                = msoAlignCenter
        End With
        
        Selection.OnAction = mac_but(i)
    height = height + 40
    Next i

End Sub

Public Sub setup_pivottable()
    '
    ' Creates a pivot table on a different sheet
    '
    
    Dim sheet As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim pc2 As PivotCache
    Dim pt2 As PivotTable
    Dim sc As SlicerCache
    Dim sl As Slicer

    ' Creates pivot cache
    Set pc = ThisWorkbook.PivotCaches.Create( _
             SourceType:=xlDatabase, _
             SourceData:=Sheet1.Range("A1").CurrentRegion.Address, _
             Version:=xlPivotTableVersion15)

    Set pc2 = ThisWorkbook.PivotCaches.Create( _
             SourceType:=xlDatabase, _
             SourceData:=Sheet1.Range("I1").CurrentRegion.Address, _
             Version:=xlPivotTableVersion15)

    ' Creates sheet for pivot table
    Set sheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook. _
                Worksheets(ActiveWorkbook.Worksheets.Count))
    ActiveSheet.Name = "PivotTable"

    ' Creates pivot table
    Set pt = pc.CreatePivotTable( _
             TableDestination:=Range("A3"), _
             TableName:="Pivot")
    
    Set pt2 = pc.CreatePivotTable( _
              TableDestination:=Range("I3"), _
              TableName:="Pivot2")

    ' Creates slicer cache
    Set sc = ThisWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables _
             ("Pivot"), "Comp")
             
    ' Creates slicer on dashboard sheet
    Set sl = sc.Slicers.Add("Dashboard", , "Competitions", _
             "Competitions", Range("J31").Top, Range("J31").Left)

    With ActiveWorkbook.SlicerCaches("Slicer_Comp").Slicers _
        ("Competitions")
            .Style = "SlicerStyleDark6"
            .NumberOfColumns = 2
    End With
        
    Sheets("Dashboard").Select
    With ActiveSheet.Shapes("Competitions")
        .ScaleHeight 0.37, msoFalse, msoScaleFromTopLeft
        .ScaleWidth 1.4, msoFalse, msoScaleFromTopLeft
    End With

    ActiveWorkbook.SlicerCaches("Slicer_Comp").PivotTables.AddPivotTable ( _
        Sheets("PivotTable").PivotTables("Pivot2"))
    
    ' Adjusts style of the slicer
    ActiveWorkbook.TableStyles.Add ("Slicer Style 1")
    
    With ActiveWorkbook.TableStyles("Slicer Style 1")
        .ShowAsAvailablePivotTableStyle = False
        .ShowAsAvailableTableStyle = False
        .ShowAsAvailableSlicerStyle = True
        .ShowAsAvailableTimelineStyle = False
        .TableStyleElements(xlWholeTable).Font.ThemeColor _
            = xlThemeColorDark1

        With .TableStyleElements(xlWholeTable).Interior
            .ThemeColor = xlThemeColorDark2
            .TintAndShade = -0.749961851863155
        End With
    
        .TableStyleElements(xlSlicerSelectedItemWithData).Interior _
            .ThemeColor = xlThemeColorAccent6
            
        With ActiveWorkbook.TableStyles("Slicer Style 1").TableStyleElements( _
            xlSlicerHoveredSelectedItemWithData).Interior
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.799981688894314
        End With
    End With
    
    ActiveWorkbook.SlicerCaches("Slicer_Comp").Slicers("Competitions").Style = _
        "Slicer Style 1"
    
    ' Resets cursor back to the top left
    Range("A1").Select

End Sub

Public Sub DeleteAllCharts()
    '
    ' Clears all charts on the dashboard
    '
    
    Dim chtObj As ChartObject
    
    Sheets("Dashboard").Activate
    For Each chtObj In ActiveSheet.ChartObjects
        chtObj.Delete
    Next
    
End Sub

Public Sub results_tab()
    '
    ' Adjusts pivot table and displays graphs of results
    '

    Dim rng As Range
    Dim rng2 As Range
    Dim bar As ChartObject
    Dim pie As ChartObject
    Dim bar2 As ChartObject
    
    ' Clears all charts
    Call DeleteAllCharts
    
    ' Clears pivot table
    Application.DisplayAlerts = False
    Sheets("PivotTable").PivotTables("Pivot").ClearTable
    Sheets("PivotTable").PivotTables("Pivot2").ClearTable

    ' Inputs fields into pivot table 1
    Sheets("PivotTable").PivotTables("Pivot"). _
        AddDataField Sheets("PivotTable").PivotTables _
        ("Pivot").PivotFields("Result"), _
        "Count of Result", xlCount

    With Sheets("PivotTable").PivotTables("Pivot") _
        .PivotFields("Result")
            .Orientation = xlRowField
            .PivotItems("L").Position = 1
    End With
    
    Set rng = Sheets("PivotTable").Range("A3").CurrentRegion
    
    ' Inputs fields into pivot table 2
    With Sheets("PivotTable").PivotTables("Pivot2")
        .AddDataField Sheets("PivotTable").PivotTables("Pivot2"). _
            PivotFields("Result"), "Count of Result", xlCount
        .PivotFields("Opponent").Orientation = xlRowField
        .PivotFields("Result").Orientation = xlColumnField
        .PivotFields("Result").PivotItems("L").Position = 1
    End With
    
    Set rng2 = Sheets("PivotTable").Range("I3").CurrentRegion

    ' Inserts result bar graph into dashboard
    Set bar = Sheets("Dashboard").ChartObjects.Add( _
                Left:=Range("Q11").Left, _
                Width:=200, _
                Top:=Range("Q11").Top, _
                height:=125)

    bar.Chart.SetSourceData Source:=rng
    bar.Chart.ChartType = xlColumnClustered
    
    With ActiveSheet
        .ChartObjects(1).Name = "ResultsBar"
        .ChartObjects("ResultsBar").Activate
    End With
    
    With ActiveChart
        .ClearToMatchStyle
        .ChartStyle = 209
        .ChartColor = 13
        .SetElement (msoElementChartTitleAboveChart)
        .ChartTitle.Text = "Results D/L/W"
        .ShowReportFilterFieldButtons = False
        .ShowValueFieldButtons = False
        .ShowAxisFieldButtons = False
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Caption = "Games"
        .SetElement (msoElementDataLabelOutSideEnd)
    End With
    
    ' Inserts result pie graph into dashboard
    Set pie = Sheets("Dashboard").ChartObjects.Add( _
                 Left:=Range("Q21").Left, _
                 Width:=200, _
                 Top:=Range("Q21").Top, _
                 height:=125)
            
    pie.Chart.SetSourceData Source:=rng
    pie.Chart.ChartType = xlPie

    With ActiveSheet
        .ChartObjects(2).Name = "ResultsPie"
        .ChartObjects("ResultsPie").Activate
    End With
    
    With ActiveChart
        .ClearToMatchStyle
        .ChartStyle = 257
        .ChartColor = 12
        .SetElement (msoElementChartTitleAboveChart)
        .ChartTitle.Text = "Results percentage"
        .ApplyDataLabels
        .ShowReportFilterFieldButtons = False
        .ShowValueFieldButtons = False
        .ShowAxisFieldButtons = False
    End With
    
    With ActiveChart.FullSeriesCollection(1).DataLabels
        .ShowPercentage = True
        .ShowValue = False
        .Position = xlLabelPositionOutsideEnd
    End With
    
    ' Inserts Results vs. Teams graph into dashboard
    Set rng2 = Sheets("PivotTable").Range("I3").CurrentRegion
    
    Set bar2 = Sheets("Dashboard").ChartObjects.Add( _
                 Left:=Range("E11").Left, _
                 Width:=600, _
                 Top:=Range("E11").Top, _
                 height:=265)
    
    bar2.Chart.SetSourceData Source:=rng2
    bar2.Chart.ChartType = xlColumnClustered
    
    With ActiveSheet
        .ChartObjects(3).Name = "ResultBar2"
        .ChartObjects("ResultBar2").Activate
    End With
    
    With ActiveChart
        .ClearToMatchStyle
        .ChartStyle = 209
        .ChartColor = 12
        .SetElement (msoElementChartTitleAboveChart)
        .ChartTitle.Text = "Results vs. Teams"
        .ShowReportFilterFieldButtons = False
        .ShowValueFieldButtons = False
        .ShowAxisFieldButtons = False
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Caption = "Number of Games"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Caption = "Teams"
        .SetElement (msoElementDataLabelOutSideEnd)
    End With

    ' Resets cursor back to the top left
    Range("A1").Select
    
End Sub

Public Sub scores_tab()
    '
    ' Adjust pivot table and displays graphs of scores
    '

    Dim rng As Range
    Dim rng2 As Range
    Dim line As ChartObject
    Dim line2 As ChartObject
    
    ' Clears all charts
    Call DeleteAllCharts
    
    ' Clears pivot table
    Application.DisplayAlerts = False
    Sheets("PivotTable").PivotTables("Pivot").ClearTable
    Sheets("PivotTable").PivotTables("Pivot2").ClearTable
    
    ' Inputs fields to pivot table 1
    With Sheets("PivotTable").PivotTables("Pivot")
        .PivotFields("Date").Orientation = xlRowField
        .AddDataField Sheets("PivotTable").PivotTables("Pivot") _
            .PivotFields("GF"), "Sum of GF", xlSum
        .AddDataField Sheets("PivotTable").PivotTables("Pivot") _
            .PivotFields("GA"), "Sum of GA", xlSum
    End With
    
    ' Inputs fields to pivot table 2
    With Sheets("PivotTable").PivotTables("Pivot2")
        .PivotFields("Date").Orientation = xlRowField
        .CalculatedFields.Add "Goal Difference", "=GF -GA", True
        .PivotFields("Goal Difference").Orientation = xlDataField
    End With
    
    ' Inserts GF and GA line graph into dashboard
    Set rng = Sheets("PivotTable").Range("A3").CurrentRegion
    
    Set line = Sheets("Dashboard").ChartObjects.Add( _
                Left:=Range("E11").Left, _
                Width:=400, _
                Top:=Range("E11").Top, _
                height:=265)

    line.Chart.SetSourceData Source:=rng
    line.Chart.ChartType = xlLine
    
    With ActiveSheet
        .ChartObjects(1).Name = "ScoreLine"
        .ChartObjects("ScoreLine").Activate
    End With
    
    With ActiveChart
        .ClearToMatchStyle
        .ChartStyle = 209
        .ChartColor = 13
        .SetElement (msoElementChartTitleAboveChart)
        .ChartTitle.Text = "Goals For and Allowed"
        .SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Caption = "Number of Goals"
        .ShowReportFilterFieldButtons = False
        .ShowValueFieldButtons = False
        .ShowAxisFieldButtons = False
    End With
    
    ' Inserts goal difference line graph into dashboard
    Set rng2 = Sheets("PivotTable").Range("I3").CurrentRegion
    
    Set line2 = Sheets("Dashboard").ChartObjects.Add( _
                 Left:=Range("N11").Left, _
                 Width:=400, _
                 Top:=Range("N11").Top, _
                 height:=265)

    line2.Chart.SetSourceData Source:=rng2
    line2.Chart.ChartType = xlLine
    
    With ActiveSheet
        .ChartObjects(2).Name = "ScoreLine2"
        .ChartObjects("ScoreLine2").Activate
    End With
    
    With ActiveChart
        .ClearToMatchStyle
        .ChartStyle = 209
        .ChartColor = 24
        .SetElement (msoElementChartTitleAboveChart)
        .ChartTitle.Text = "Goals Differences"
        .SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Caption = "Number of Goals"
        .ShowReportFilterFieldButtons = False
        .ShowValueFieldButtons = False
        .ShowAxisFieldButtons = False
        .Axes(xlCategory).TickLabelPosition = xlLow
    End With
    
    ' Resets cursor back to the top left
    Range("A1").Select
    
End Sub

Public Sub possession_tab()
    '
    ' Adjust pivot table and displays graphs of possession
    '

    Dim rng As Range
    Dim rng2 As Range
    Dim bar As ChartObject
    Dim bar2 As ChartObject
    
    ' Clears all charts
    Call DeleteAllCharts
    
    ' Clears results into pivot table
    Application.DisplayAlerts = False
    Sheets("PivotTable").PivotTables("Pivot").ClearTable
    Sheets("PivotTable").PivotTables("Pivot2").ClearTable

    ' Inputs fields into pivot table 1
    With Sheets("PivotTable").PivotTables("Pivot")
        .AddDataField Sheets("PivotTable").PivotTables("Pivot"). _
            PivotFields("Result"), "Count of Result", xlCount
        .PivotFields("Poss").Orientation = xlRowField
        .PivotFields("Result").Orientation = xlColumnField
        .PivotFields("Result").PivotItems("L").Position = 1
    End With
    
    ' Inputs fields into pivot table 2
    With Sheets("PivotTable").PivotTables("Pivot2")
        .PivotFields("Poss").Orientation = xlRowField
        .AddDataField Sheets("PivotTable") _
            .PivotTables("Pivot2").PivotFields("GF"), "Count of GF", xlCount
        .AddDataField Sheets("PivotTable") _
            .PivotTables("Pivot2").PivotFields("GA"), "Count of GA", xlCount
        .PivotFields("Count of GF").Function = xlSum
        .PivotFields("Count of GA").Function = xlSum
    End With
    
    ' Inserts poss w/d/l line graph into dashboard
    Set rng = Sheets("PivotTable").Range("A3").CurrentRegion
    
    Set bar = Sheets("Dashboard").ChartObjects.Add( _
                Left:=Range("E11").Left, _
                Width:=400, _
                Top:=Range("E11").Top, _
                height:=265)
                
    bar.Chart.SetSourceData Source:=rng
    bar.Chart.ChartType = xlColumnClustered
    
    With ActiveSheet
        .ChartObjects(1).Name = "PossBar"
        .ChartObjects("PossBar").Activate
    End With
    
    With ActiveChart
        .ClearToMatchStyle
        .ChartStyle = 209
        .ChartColor = 12
        .SetElement (msoElementChartTitleAboveChart)
        .ChartTitle.Text = "Possession vs. Results"
        .ShowReportFilterFieldButtons = False
        .ShowValueFieldButtons = False
        .ShowAxisFieldButtons = False
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Caption = "Number of Games"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Caption = "Possession by %"
    End With
    
    ' Inserts poss gf ga line graph into dashboard
    Set rng2 = Sheets("PivotTable").Range("I3").CurrentRegion
    
    Set bar2 = Sheets("Dashboard").ChartObjects.Add( _
                 Left:=Range("N11").Left, _
                 Width:=400, _
                 Top:=Range("N11").Top, _
                 height:=265)
    
    bar2.Chart.SetSourceData Source:=rng2
    bar2.Chart.ChartType = xlColumnClustered
    
    With ActiveSheet
        .ChartObjects(2).Name = "PossBar2"
        .ChartObjects("PossBar2").Activate
    End With
    
    With ActiveChart
        .ClearToMatchStyle
        .ChartStyle = 209
        .ChartColor = 13
        .SetElement (msoElementChartTitleAboveChart)
        .ChartTitle.Text = "Possession vs. GF/GA"
        .ShowReportFilterFieldButtons = False
        .ShowValueFieldButtons = False
        .ShowAxisFieldButtons = False
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Caption = "Number of Goals"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Caption = "Possession by %"
    End With
    
    ' Resets cursor back to the top left
    Range("A1").Select
            
End Sub

Public Sub captains_tab()
    '
    ' Adjusts pivot table and displays graphs of captains
    '
    
    Dim rng As Range
    Dim bar As ChartObject
    Dim pie As ChartObject
    
    ' Clears all charts
    Call DeleteAllCharts
    
    ' Clears results into pivot table
    Application.DisplayAlerts = False
    Sheets("PivotTable").PivotTables("Pivot").ClearTable
    Sheets("PivotTable").PivotTables("Pivot2").ClearTable

    ' Inputs fields into pivot table 1
    With Sheets("PivotTable").PivotTables("Pivot")
        .AddDataField Sheets("PivotTable").PivotTables("Pivot"). _
            PivotFields("Result"), "Count of Result", xlCount
        .PivotFields("Captain").Orientation = xlRowField
        .PivotFields("Result").Orientation = xlColumnField
        .PivotFields("Result").PivotItems("L").Position = 1
    End With
    
    Set rng = Sheets("PivotTable").Range("A3").CurrentRegion
    
    ' Inserts captains bar graph into dashboard
    Set bar = Sheets("Dashboard").ChartObjects.Add( _
                Left:=Range("E11").Left, _
                Width:=400, _
                Top:=Range("E11").Top, _
                height:=265)

    bar.Chart.SetSourceData Source:=rng
    bar.Chart.ChartType = xlColumnClustered
    
    With ActiveSheet
        .ChartObjects(1).Name = "CaptainsBar"
        .ChartObjects("CaptainsBar").Activate
    End With
    
    With ActiveChart
        .ClearToMatchStyle
        .ChartStyle = 209
        .ChartColor = 13
        .SetElement (msoElementChartTitleAboveChart)
        .ChartTitle.Text = "Results vs. Captains"
        .ShowReportFilterFieldButtons = False
        .ShowValueFieldButtons = False
        .ShowAxisFieldButtons = False
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Caption = "Games"
        .SetElement (msoElementDataLabelOutSideEnd)
    End With
    
    ' Inserts captains pie graph into dashboard
    Set pie = Sheets("Dashboard").ChartObjects.Add( _
                 Left:=Range("N11").Left, _
                 Width:=400, _
                 Top:=Range("N11").Top, _
                 height:=265)
            
    pie.Chart.SetSourceData Source:=rng
    pie.Chart.ChartType = xlPie

    With ActiveSheet
        .ChartObjects(2).Name = "CaptainPie"
        .ChartObjects("CaptainPie").Activate
    End With
    
    With ActiveChart
        .ClearToMatchStyle
        .ChartStyle = 257
        .ChartColor = 13
        .SetElement (msoElementChartTitleAboveChart)
        .ChartTitle.Text = "Amount of games played by captains"
        .ApplyDataLabels
        .ShowReportFilterFieldButtons = False
        .ShowValueFieldButtons = False
        .ShowAxisFieldButtons = False
    End With
    
    With ActiveChart.FullSeriesCollection(1).DataLabels
        .ShowPercentage = True
        .ShowValue = False
        .Position = xlLabelPositionOutsideEnd
    End With

    ' Resets cursor back to the top left
    Range("A1").Select

End Sub

Public Sub expected_tab()
    '
    ' Adjust pivot table and displays graphs of expected goals
    '
    Dim rng As Range
    Dim rng2 As Range
    Dim line As ChartObject
    Dim line2 As ChartObject
    
    ' Clears all charts
    Call DeleteAllCharts
    
    ' Clears pivot table
    Application.DisplayAlerts = False
    Sheets("PivotTable").PivotTables("Pivot").ClearTable
    Sheets("PivotTable").PivotTables("Pivot2").ClearTable
    
    ' Inputs fields to pivot table 1
    With Sheets("PivotTable").PivotTables("Pivot")
        .PivotFields("Date").Orientation = xlRowField
        .AddDataField Sheets("PivotTable").PivotTables("Pivot") _
            .PivotFields("xG"), "Sum of xG", xlSum
        .AddDataField Sheets("PivotTable").PivotTables("Pivot") _
            .PivotFields("xGA"), "Sum of xGA", xlSum
    End With
    
    ' Inputs fields to pivot table 2
    With Sheets("PivotTable").PivotTables("Pivot2")
        .PivotFields("Date").Orientation = xlRowField
        .CalculatedFields.Add "xG Difference", "=xG -xGA", True
        .PivotFields("xG Difference").Orientation = xlDataField
    End With
    
    ' Inserts xG and xGA line graph into dashboard
    Set rng = Sheets("PivotTable").Range("A3").CurrentRegion
    
    Set line = Sheets("Dashboard").ChartObjects.Add( _
                Left:=Range("E11").Left, _
                Width:=400, _
                Top:=Range("E11").Top, _
                height:=265)

    line.Chart.SetSourceData Source:=rng
    line.Chart.ChartType = xlLine
    
    With ActiveSheet
        .ChartObjects(1).Name = "ExpectedLine"
        .ChartObjects("ExpectedLine").Activate
    End With
    
    With ActiveChart
        .ClearToMatchStyle
        .ChartStyle = 209
        .ChartColor = 13
        .SetElement (msoElementChartTitleAboveChart)
        .ChartTitle.Text = "Expected Goals For and Allowed"
        .SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Caption = _
            "Number of Expected Goals"
        .ShowReportFilterFieldButtons = False
        .ShowValueFieldButtons = False
        .ShowAxisFieldButtons = False
    End With
    
    ' Inserts xG difference line graph into dashboard
    Set rng2 = Sheets("PivotTable").Range("I3").CurrentRegion
    
    Set line2 = Sheets("Dashboard").ChartObjects.Add( _
                 Left:=Range("N11").Left, _
                 Width:=400, _
                 Top:=Range("N11").Top, _
                 height:=265)

    line2.Chart.SetSourceData Source:=rng2
    line2.Chart.ChartType = xlLine
    
    With ActiveSheet
        .ChartObjects(2).Name = "ExpectedLine2"
        .ChartObjects("ExpectedLine2").Activate
    End With
    
    With ActiveChart
        .ClearToMatchStyle
        .ChartStyle = 209
        .ChartColor = 24
        .SetElement (msoElementChartTitleAboveChart)
        .ChartTitle.Text = "Expected Goals Differences"
        .SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Caption = _
            "Number of Expected Goals"
        .ShowReportFilterFieldButtons = False
        .ShowValueFieldButtons = False
        .ShowAxisFieldButtons = False
        .Axes(xlCategory).TickLabelPosition = xlLow
    End With
    
    ' Resets cursor back to the top left
    Range("A1").Select
    
End Sub

Public Sub venue_tab()
    '
    ' Adjust pivot table and displays graphs of venues
    '

    Dim rng As Range
    Dim rng2 As Range
    Dim bar As ChartObject
    Dim bar2 As ChartObject
    
    ' Clears all charts
    Call DeleteAllCharts
    
    ' Clears results into pivot table
    Application.DisplayAlerts = False
    Sheets("PivotTable").PivotTables("Pivot").ClearTable
    Sheets("PivotTable").PivotTables("Pivot2").ClearTable

    ' Inputs fields into pivot table 1
    With Sheets("PivotTable").PivotTables("Pivot")
        .AddDataField Sheets("PivotTable").PivotTables("Pivot"). _
            PivotFields("Result"), "Count of Result", xlCount
        .PivotFields("Venue").Orientation = xlRowField
        .PivotFields("Result").Orientation = xlColumnField
        .PivotFields("Result").PivotItems("L").Position = 1
    End With
    
    ' Inputs fields into pivot table 2
    With Sheets("PivotTable").PivotTables("Pivot2")
        .AddDataField Sheets("PivotTable").PivotTables("Pivot2") _
            .PivotFields("Venue"), "Count of Venue", xlCount
        .PivotFields("Attendance").Orientation = xlRowField
        .PivotFields("Venue").Orientation = xlColumnField
    End With

    With Sheets("PivotTable")
        .Range("I5").Group Start:=1, End:=80000, By:=10000
        .PivotTables("Pivot2").PivotFields("Attendance") _
            .PivotItems("<1").Caption = "0"
    End With
    
    ' Inserts venue result bar graph into dashboard
    Set rng = Sheets("PivotTable").Range("A3").CurrentRegion
    
    Set bar = Sheets("Dashboard").ChartObjects.Add( _
                Left:=Range("E11").Left, _
                Width:=400, _
                Top:=Range("E11").Top, _
                height:=265)
                
    bar.Chart.SetSourceData Source:=rng
    bar.Chart.ChartType = xlColumnClustered
    
    With ActiveSheet
        .ChartObjects(1).Name = "VenueBar"
        .ChartObjects("VenueBar").Activate
    End With
    
    With ActiveChart
        .ClearToMatchStyle
        .ChartStyle = 209
        .ChartColor = 12
        .SetElement (msoElementChartTitleAboveChart)
        .ChartTitle.Text = "Venue vs. Results"
        .ShowReportFilterFieldButtons = False
        .ShowValueFieldButtons = False
        .ShowAxisFieldButtons = False
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Caption = "Number of Games"
        .SetElement (msoElementDataLabelOutSideEnd)
    End With
    
    ' Inserts attendance vs venue bar graph into dashboard
    Set rng2 = Sheets("PivotTable").Range("I3").CurrentRegion
    
    Set bar2 = Sheets("Dashboard").ChartObjects.Add( _
                 Left:=Range("N11").Left, _
                 Width:=400, _
                 Top:=Range("N11").Top, _
                 height:=265)
    
    bar2.Chart.SetSourceData Source:=rng2
    bar2.Chart.ChartType = xlColumnClustered
    
    With ActiveSheet
        .ChartObjects(2).Name = "VenueBar2"
        .ChartObjects("VenueBar2").Activate
    End With
    
    With ActiveChart
        .ClearToMatchStyle
        .ChartStyle = 209
        .ChartColor = 11
        .SetElement (msoElementChartTitleAboveChart)
        .ChartTitle.Text = "Venue vs. Attendance"
        .ShowReportFilterFieldButtons = False
        .ShowValueFieldButtons = False
        .ShowAxisFieldButtons = False
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Caption = "Number of Games"
        .SetElement (msoElementDataLabelOutSideEnd)
    End With
    
    ' Resets cursor back to the top left
    Range("A1").Select
    
End Sub
