Option Explicit

' Class script: DocDevMenu

Private Sub Document_BeforeExtract(ByVal pXDoc As CASCADELib.CscXDocument)
   DevMenu_Show(Nothing, pXDoc)
   Dev_ExportScriptAndLocators()
End Sub
