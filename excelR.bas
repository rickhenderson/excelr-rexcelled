Option Explicit
' # ExcelR
' * A project to include R-like commands for data analysis in Excel.
'
' The project will basically focus on Exploratory Data Analysis,
' but can be used for data visualization because Excel's chart
' features are quite flexible and good for presentation quality
' graphics.
'
' ## Function List:
' * read_csv(file, header = True, sep =",")
' * plot()
' * blanks_to_na()

Sub blanks_to_na()
    ' Subroutine to convert all blank values
    ' in a datafile to NA for use with R.
    ' Also shades the cells red for easy identification.
    ' When exported as a CSV, color is ignored and lost.
    
    Dim used_cells As Range
    Dim cell As Range
    
    ' Set a variable to represent all the used cells in the worksheet
    Set used_cells = Worksheets("data").UsedRange
    
    ' Go through every cell in the range and
    ' if it is blank, set the value to NA.
    ' Warning: It is also possible that
    '    blank cells from your data source
    '    actually contain a space, and are not empty.
    ' This function sets both to NA.
    
    For Each cell In used_cells
        If cell.Value = "" Or cell.Value = " " Then
            cell.Value = "NA"
            cell.Interior.Color = RGB(255, 0, 0)
        End If
    Next

End Sub

Public Function read_csv(file As String, Optional header As Boolean = True, Optional sep As String = ",") As Boolean
'
' read_csv Macro
' Created by: Rick Henderson
' Created on: March 5, 2016
' Use \t to specify tab seperator as in R read.csv and related functions.
' A macro recorded by using the Data Import tool in Excel.
' This could have been hand-coded, but this allows this type of project to be more easily reproduced.
' March 5, 2016: Code modified to make it work similar to read.csv in R.
'                Currently a Boolean function until it can be reworked to read data into an array.

' Variable Declaration
'
Dim commaDelimiter As Boolean
Dim spaceDelimiter As Boolean
Dim tabDelimiter As Boolean
Dim semicolonDelimiter As Boolean

' Set default values for safety
commaDelimiter = False
spaceDelimiter = False
tabDelimiter = False
semicolonDelimiter = False
    
    ' Turn off Excel alerts
    Application.DisplayAlerts = False
    
    ' Check the sep argument provided by the user
    Select Case sep
        Case ","
            commaDelimiter = True
        Case " "
            spaceDelimiter = True
        Case "\t"
            ' Delimiter was the tab
            tabDelimiter = True
        Case ";"
            semicolonDelimiter = True
        Case Else
            ' Something else happened. TODO maybe through an exception.
            commaDelimiter = True
    End Select
    
    ' Need to fix this.
    If IsMissing(header) Then
        header = False
    End If
    
    ActiveWorkbook.Worksheets.Add
    
    On Error GoTo FileNotFound
        With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & file, Destination:=Range("$A$1"))
    
    
        ' Set the name of the data table
        .Name = file
        
        ' Argument for if file contains header row or not.
        .fieldNames = header
        
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 850
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        
        ' Sets different delimiter options
        .TextFileTabDelimiter = tabDelimiter
        .TextFileSemicolonDelimiter = semicolonDelimiter
        .TextFileCommaDelimiter = commaDelimiter
        .TextFileSpaceDelimiter = spaceDelimiter
        
        .TextFileColumnDataTypes = Array(1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    read_csv = True
    
    ' Turn Excel alerts back on
    Application.DisplayAlerts = True
    
    Exit Function
    
FileNotFound:
    Debug.Print ("Cannot open file " & file & ": No such file or directory.")
    ' Turn Excel alerts back on
    Application.DisplayAlerts = True
    
End Function
Sub plot()
'
' plot Macro
' A basic plot function to retreive data and chart type as parameters.
'

'
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=Range("Sheet2!$A$1:$B$9")
    ActiveChart.FullSeriesCollection(1).Select
    Selection.Delete
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    Selection.Delete
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.ChartArea.Select
    ActiveChart.FullSeriesCollection(1).XValues = "=Sheet2!$A$2:$A$9"
    ActiveChart.Legend.Select
    Selection.Delete
End Sub

Sub TestReadCSV()
    Dim success As Boolean
    
    ' Basic use for .CSV files with Headers in First Row.
    success = read_csv("data.csv")
    
    ' Read a SPACE separated file, not specifiying header argument.
    success = read_csv("data.spa", , " ")  ' Works
    
    ' Read a SEMICOLON separated file, with no header in first row.
    success = read_csv("data.sem", False, ";") ' Works
    
    ' Reads a TAB separated file.
    success = read_csv("data.txt", True, "\t") ' Fails
    
    ' Read a missing file
    success = read_csv("notthere.csv")
End Sub

Sub Cleanup()
    ' Delete all the worksheets that aren't called "Main"
    ' in the THIS workbook.
    
    Dim wksSheet As Worksheet
    
    ' Turn off alerts
    Application.DisplayAlerts = False
    For Each wksSheet In ThisWorkbook.Worksheets
        If wksSheet.Name <> "Main" Then
            wksSheet.Delete
        End If
    Next
    
    ' Turn alerts back on
    Application.DisplayAlerts = True
End Sub
