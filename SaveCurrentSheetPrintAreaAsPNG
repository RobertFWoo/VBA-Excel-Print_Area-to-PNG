Sub SaveCurrentSheetPrintAreaAsPNG()
    Dim ws As Worksheet
    Dim printArea As Range
    Dim chartObj As ChartObject
    Dim imgPath As String
    Dim outputDir As String

    ' Set the output directory
    outputDir = "C:\CalendarPagesRaw\"
    If Dir(outputDir, vbDirectory) = "" Then
        MkDir outputDir
    End If

    ' Get the active sheet
    Set ws = ActiveSheet

    ' Check if the sheet has a print area
    If ws.PageSetup.printArea <> "" Then
        ' Get the print area range
        Set printArea = ws.Range(ws.PageSetup.printArea)

        ' Create a temporary chart to capture the print area
        Set chartObj = ws.ChartObjects.Add(Left:=printArea.Left, Top:=printArea.Top, Width:=printArea.Width, Height:=printArea.Height)
        chartObj.chart.ChartArea.Format.Line.Visible = msoFalse
        chartObj.chart.PlotArea.Format.Line.Visible = msoFalse
        chartObj.chart.Axes(xlCategory).Delete
        chartObj.chart.Axes(xlValue).Delete
        chartObj.chart.ChartArea.Fill.Visible = msoFalse
        chartObj.chart.PlotArea.Fill.Visible = msoFalse

        ' Copy the print area to the chart
        printArea.CopyPicture Appearance:=xlScreen, Format:=xlPicture
        chartObj.chart.Paste

        ' Save the chart as a PNG file
        imgPath = outputDir & ws.Name & ".png"
        chartObj.chart.Export Filename:=imgPath, FilterName:="PNG"

        ' Delete the temporary chart
        chartObj.Delete

        MsgBox "Saved print area of " & ws.Name & " to " & imgPath, vbInformation

        ' Ask the user if they want to open the folder
        If MsgBox("Do you want to open the folder?", vbYesNo + vbQuestion, "Open Folder") = vbYes Then
            Shell "explorer.exe " & outputDir, vbNormalFocus
        End If
    Else
        MsgBox "No print area defined for this sheet.", vbExclamation
    End If
End Sub
