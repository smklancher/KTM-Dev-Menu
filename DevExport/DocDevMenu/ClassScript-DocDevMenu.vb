Option Explicit

' Class script: DocDevMenu

Private Sub Document_BeforeExtract(ByVal pXDoc As CASCADELib.CscXDocument)
   ' Invoke DevMenu by selecting the class, selecting the document, then extracting the document.
   DevMenu_Dialog(Nothing, pXDoc)
End Sub
