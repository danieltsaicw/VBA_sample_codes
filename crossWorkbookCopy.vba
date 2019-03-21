Sub CopyBloombergData()
  MsgBox "Adding..."

  file_name = "AscenderQualityX"


  Application.ScreenUpdating = False

  Dim wb1 As Workbook
  Dim wb2 As Workbook

   Set wb1 = Workbooks.Open("C:\Users\Bloomberg\AppData\Local\bipy\21110120\projects\df8bee244deb42978d64b430b92f6f1d\" & file_name & ".csv")
   Worksheets(file_name).Range("A:F").Copy

  Set wb2 = Workbooks.Open("C:\Users\Bloomberg\Desktop\Bloomberg-Daniel\foo2.xlsm")
  Worksheets("output_bloomberg").Range("A:F").PasteSpecial
  
  wb1.Application.CutCopyMode = False
  wb1.Close
  
End Sub
