Sub Clearcells()
'Updateby Extendoffice 20161008
'Worksheets("raw_data").Range("A3", "A5").Clear
'Worksheets("raw_data").Range("A:Z").ClearContents
Worksheets("raw_data").UsedRange.ClearContents

End Sub


'Turn on Formula Automation
Application.Calculation = xlAutomatic
'Turn off Formula Automation - Manual
Application.Calculation = xlManual


'Find

Dim RangeFound As Range
Set RangeFound = Range("A1:A100").Find(What:="String", lookin:=xlValues)
If Not RangeFound is Nothing Then
  Range("A101").Value = 100
Else
  MsgBox("not found")
End
