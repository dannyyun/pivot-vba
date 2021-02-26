Attribute VB_Name = "Module1"


Sub RollUpTimeBy()

    OptimizeCode_Begin

    
    With ActiveSheet.Shapes("Drop Down 1").ControlFormat
        Select Case .ListIndex
            Case Is = 1
                YearView
            Case Is = 2
                MonthView
            Case Is = 3
                WeekView
            Case Is = 4
                DayView
        End Select
    End With
    
    DataLabel
    
    OptimizeCode_End
    
End Sub

Sub YearView()
    
    ClearRow
    
    Dim pt As PivotTable
    Set pt = ActiveSheet.PivotTables(1)
    With pt.PivotFields("Year")
        .Orientation = xlRowField
        .Position = 1
    End With

End Sub

Sub MonthView()
    
    ClearRow
    
    Dim pt As PivotTable
    Set pt = ActiveSheet.PivotTables(1)
    With pt.PivotFields("Month")
        .Orientation = xlRowField
        .Position = 1
    End With

End Sub

Sub WeekView()
    
    ClearRow
    
    Dim pt As PivotTable
    Set pt = ActiveSheet.PivotTables(1)
    With pt.PivotFields("Week")
        .Orientation = xlRowField
        .Position = 1
    End With

End Sub

Sub DayView()
    
    ClearRow
    
    Dim pt As PivotTable
    Set pt = ActiveSheet.PivotTables(1)
    With pt.PivotFields("Date")
        .Orientation = xlRowField
        .Position = 1
    End With

End Sub

Sub GroupBy()
    
    OptimizeCode_Begin
    
    With ActiveSheet.Shapes("Drop Down 2").ControlFormat
    
        Select Case .ListIndex
            Case Is = 1
                PrTypeView
            Case Is = 2
                ProductView
            Case Is = 3
                ChTypeView
            Case Is = 4
                ChannelView
        End Select
        
    End With
    
    DataLabel
    
    OptimizeCode_End

End Sub

Sub PrTypeView()
    
    ClearColumn
    
    Dim pt As PivotTable
    Set pt = ActiveSheet.PivotTables(1)
    With pt.PivotFields("Product_Type")
        .Orientation = xlColumnField
        .Position = 1
    End With
    

End Sub

Sub ProductView()
    
    ClearColumn
    
    Dim pt As PivotTable
    Set pt = ActiveSheet.PivotTables(1)
    With pt.PivotFields("Product_Name")
        .Orientation = xlColumnField
        .Position = 1
    End With

End Sub
Sub ChTypeView()
    
    ClearColumn
    
    Dim pt As PivotTable
    Set pt = ActiveSheet.PivotTables(1)
    With pt.PivotFields("Channel_Type")
        .Orientation = xlColumnField
        .Position = 1
    End With

End Sub


Sub ChannelView()
    
    ClearColumn
    
    Dim pt As PivotTable
    Set pt = ActiveSheet.PivotTables(1)
    With pt.PivotFields("Channel")
        .Orientation = xlColumnField
        .Position = 1
    End With

End Sub


Sub ChtType()
    
    OptimizeCode_Begin


    With ActiveSheet.Shapes("Drop Down 4").ControlFormat

        Select Case .ListIndex
            Case Is = 1
                ActiveSheet.ChartObjects("Chart 1").Activate
                ActiveChart.ChartType = xlLine
                ActiveSheet.Range("A1").Select
                
            Case Is = 2
                ActiveSheet.ChartObjects("Chart 1").Activate
                ActiveChart.ChartType = xlColumnStacked
                ActiveSheet.Range("A1").Select
                
            Case Is = 3
                ActiveSheet.ChartObjects("Chart 1").Activate
                ActiveChart.ChartType = xlColumnStacked100
                ActiveSheet.Range("A1").Select
                
        End Select
    End With
    
    DataLabel
 
    OptimizeCode_End

End Sub


Sub DisplayValue()
    
    OptimizeCode_Begin
    
    With ActiveSheet.Shapes("Drop Down 3").ControlFormat
    
        Select Case .ListIndex
            Case Is = 1
                Qty
            Case Is = 2
                Amt
            Case Is = 3
                ASP
            Case Is = 4
                AOV
        End Select
        
    End With
    
    DataLabel
    
    OptimizeCode_End

End Sub

Sub Amt()

    OptimizeCode_Begin
    
    Dim pt As PivotTable
    Dim pi As PivotItem
    Set pt = ActiveSheet.PivotTables(1)
    
    For Each pi In pt.DataPivotField.PivotItems
        If pi.Name = "Amt" Then
            Exit Sub
        Else
            pi.Visible = False
        End If
    Next pi
    
    With pt.PivotFields("Amt")
        .Orientation = xlDataField
        .Position = 1
        .Function = xlSum
        .NumberFormat = "$#,##0"
        
    End With
    
    OptimizeCode_End

End Sub

Sub Qty()

    OptimizeCode_Begin
    
    Dim pt As PivotTable
    Dim pi As PivotItem
    Set pt = ActiveSheet.PivotTables(1)
    
    For Each pi In pt.DataPivotField.PivotItems
        If pi.Name = "Qty" Then
            Exit Sub
        Else
            pi.Visible = False
        End If
    Next pi

    With pt.PivotFields("Qty")
        .Orientation = xlDataField
        .Position = 1
        .Function = xlSum
        .NumberFormat = "#,##0"
    End With
    
    OptimizeCode_End

End Sub

Sub ASP()

    OptimizeCode_Begin
   
    Dim pt As PivotTable
    Dim pi As PivotItem
    Set pt = ActiveSheet.PivotTables(1)
    
    For Each pi In pt.DataPivotField.PivotItems
        If pi.Name = "ASP" Then
            Exit Sub
        Else
            pi.Visible = False
        End If
    Next pi
    
    With pt.PivotFields("ASP")
        .Orientation = xlDataField
        .Position = 1
        .Function = xlSum
        .NumberFormat = "$#,##0.00"
    End With
    
    OptimizeCode_End

End Sub

Sub AOV()

    OptimizeCode_Begin
    
    Dim pt As PivotTable
    Dim pi As PivotItem
    Set pt = ActiveSheet.PivotTables(1)
    
    For Each pi In pt.DataPivotField.PivotItems
        If pi.Name = "AOV" Then
            Exit Sub
        Else
            pi.Visible = False
        End If
    Next pi
        
    With pt.PivotFields("AOV")
        .Orientation = xlDataField
        .Position = 1
        .Function = xlSum
        .NumberFormat = "$#,##0.00"
    End With
    
    OptimizeCode_End

End Sub


Sub ClearRow()
    
    Dim pt As PivotTable
    Dim fld As Object
    Set pt = ActiveSheet.PivotTables(1)
    For Each fld In pt.RowFields
        Debug.Print fld.Name
        fld.Orientation = xlHidden
    Next fld

End Sub

Sub ClearColumn()
    
    Dim pt As PivotTable
    Dim fld As Object
    Set pt = ActiveSheet.PivotTables(1)
    For Each fld In pt.ColumnFields
        Debug.Print fld.Name
        fld.Orientation = xlHidden
    Next fld

End Sub


Sub DataLabel()

    OptimizeCode_Begin
    
    Dim cbValue As Object
    Set cbValue = ActiveSheet.CheckBoxes("Check Box 1")
    
    Dim chtObj As ChartObject
    Dim sr As Series
    
    With cbValue
        If .Value = 1 Then
            For Each chtObj In ActiveSheet.ChartObjects
                For Each sr In chtObj.Chart.SeriesCollection
                    sr.ApplyDataLabels
                        With sr.DataLabels
          '             .ShowSeriesName = True
                        .ShowValue = True
           '            .Position = xlLabelPositionInsideBase
          '             .Orientation = -90
          '             .Font.Size = 8
                        End With
                Next sr
            Next chtObj
            
        Else
            For Each chtObj In ActiveSheet.ChartObjects
                For Each sr In chtObj.Chart.SeriesCollection
                    sr.ApplyDataLabels
                        With sr.DataLabels
          '             .ShowSeriesName = True
                        .ShowValue = False
           '            .Position = xlLabelPositionInsideBase
          '             .Orientation = -90
          '             .Font.Size = 8
                        End With
                Next sr
            Next chtObj
        
        End If
    End With
    
    OptimizeCode_End

End Sub


Sub OptimizeCode_Begin()
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

End Sub

Sub OptimizeCode_End()
    
    ActiveSheet.DisplayPageBreaks = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub
