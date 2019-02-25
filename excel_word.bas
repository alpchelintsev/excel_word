' StringReplace replaces the text from variable SearchStr by text from ReplaceStr,
' ReplaceAll means to replace all occurrences,
' App is object of Word.Application
Sub StringReplace(App As Object, SearchStr As String, ReplaceStr As String, ReplaceAll As Boolean)
  App.Selection.Find.ClearFormatting
  App.Selection.Find.Text = SearchStr
  App.Selection.Find.Replacement.Text = ReplaceStr
  App.Selection.Find.Forward = True
  App.Selection.Find.Wrap = 1
  App.Selection.Find.Format = False
  App.Selection.Find.MatchCase = False
  App.Selection.Find.MatchWholeWord = False
  App.Selection.Find.MatchWildcards = False
  App.Selection.Find.MatchSoundsLike = False
  App.Selection.Find.MatchAllWordForms = False
  If ReplaceAll Then
    App.Selection.Find.Execute Replace:=2
  Else
    App.Selection.Find.Execute Replace:=1
  End If
End Sub

' InsertFile inserts a document from a file with name in FN
Sub InsertFile(App As Object, FN As String)
  App.Selection.EndKey Unit:=6
  App.Selection.InsertBreak Type:=7
  App.Selection.InsertFile Filename:=FN
End Sub

' MyMacro is macro that replaces the text {numstud} by text from 1-st sheet of the cell (1, 1)
' in the file d:\my.docx, and appends the file d:\doc2.docx
Sub MyMacro()
  Dim oApp As Object
  Set oApp = CreateObject("Word.Application")
  oApp.Documents.Open ("d:\my.docx")
  StringReplace App:=oApp, SearchStr:="{numstud}", ReplaceStr:=Sheets(1).Cells(1, 1), ReplaceAll:=True
  InsertFile App:=oApp, FN:="d:\doc2.docx"
  oApp.ActiveDocument.SaveAs Filename:="d:\1.docx", FileFormat:=16
  oApp.ActiveDocument.Close
  oApp.Quit True
  Set oApp = Nothing
End Sub
