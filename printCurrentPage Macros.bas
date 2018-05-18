Attribute VB_Name = "Module1"
Sub printCurrentPage()

Dim numStart As Long

numStart = Application.ActivePresentation.SlideShowWindow.View.CurrentShowPosition

With Application.ActivePresentation.PrintOptions
    .RangeType = ppPrintSlideRange
    .Ranges.ClearAll
    .Ranges.Add Start:=numStart, End:=numStart
End With

Application.CommandBars.ExecuteMso ("PrintPreviewAndPrint")

End Sub
