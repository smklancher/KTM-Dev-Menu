'#Reference {3F4DACA7-160D-11D2-A8E9-00104B365C9F}#5.5#0#C:\Windows\SysWOW64\vbscript.dll\3#Microsoft VBScript Regular Expressions 5.5#VBScript_RegExp_55
'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\SysWOW64\scrrun.dll#Microsoft Scripting Runtime#Scripting
'#Language "WWB-COM"
Option Explicit

Type Param
   ' Represents a parameter to a function as parsed by regex
   Name As String
   ParamType As String
   Array As Boolean
   OptionalParam As Boolean
   DefaultValue As String
End Type

Type ScriptFunction
   ' Represents a script function parsed by regex
   Name As String
   IsSub As Boolean
   Params() As Param
   StartIndex As Long
   EndIndex As Long
   Content As String
   ClassName As String
   ReturnType As String
   Suspect As Boolean
   StringTag As String
End Type

Private Declare Function GetModuleFileNameEx Lib "psapi" Alias "GetModuleFileNameExA" ( ByVal hProcess As Long, ByVal hModule As Long, ByVal FileName As String, ByVal nSize As Long) As Long



Public Function IsDesignMode() As Boolean
   ' When testing Runtime Script Event (lightning bolt, Batch_ and Application_ events), ScriptExecutionMode is NOT set to Design
   ' So this function is required for code in these functions that needs to know the difference between runtime and testing in ProjectBuilder
   Dim FileName As String
   FileName = Space$(256)
   GetModuleFileNameEx(-1,0, FileName, 256)
   FileName = Left$(FileName, InStr(1, FileName, ChrW$(0)))
   Return InStr(1, FileName, "ProjectBuilder") ' This works for old PB (KTM 5.5) as well as new PB (KTM 6.0+, KTA)
End Function


Public Function NewScriptFunction(m As Match) As ScriptFunction
   ' Creates a new ScriptFunction from regex match, implementation tied to regex in ParseScript()

   With NewScriptFunction
      .StartIndex=m.FirstIndex
      .EndIndex=m.FirstIndex + m.Length
      .IsSub=(LCase(m.SubMatches(0))="sub")
      .Name=m.SubMatches(1)
      ' provide the string of params to be parsed out into a dictionary of param types
      .Params=NewParams(m.SubMatches(2))
      .returntype=m.SubMatches(3)
      .Content=m
   End With
End Function

Public Function NewParams(ParamString As String) As Param()
   ' Creates array of Param from a parameter string (everything between parens)

   Dim r As New RegExp, Matches As MatchCollection
   r.Global=True:   r.Multiline=True:   r.IgnoreCase=True
   r.Pattern = "(Optional *?)?(ByVal|ByRef)? *(\w+?)(\( *\))?(?: As ([\w\.]+?)) *(?:= *(.*?))?(?:,|$)"
   Set Matches = r.Execute(ParamString)

   If Matches.Count>0 Then
      Dim MatchIndex As Integer, p As Param, Params() As Param
      ReDim Params(Matches.Count-1)
      For MatchIndex=0 To Matches.Count-1
         p = NewParam(Matches(MatchIndex))
         Params(MatchIndex)=p
      Next
   End If
   Return Params
End Function

Public Function NewParam(m As Match) As Param
   ' Creates a Param from regex match, implementation tied to regex in NewParams()

   With NewParam
      .OptionalParam=Len(Trim(m.SubMatches(0)))>0
      .Name=m.SubMatches(2)
      .Array=Len(Trim(m.SubMatches(3)))>0
      .ParamType=m.SubMatches(4)
      .DefaultValue=m.SubMatches(5)
   End With
End Function



Public Function ParseScript(ByVal Script As String, Optional ByVal ClassName As String = "") As ScriptFunction()
   ' Parse script into array of ScriptFunction

   ' Return from cache if possible
   Static Cache As New Dictionary
   If Cache.Exists(ClassName) Then
      Return Cache.Item(ClassName)
   End If

   Dim r As New RegExp, Matches As MatchCollection
   r.Global=True:   r.Multiline=True:   r.IgnoreCase=True
   'vbscript regexp does not support singleline mode (. matches \n) and no support for named capturing groups
   r.Pattern = "^(?:Public |Private )?(Sub|Function) (.*?)\((.*?)\)\s*(?: As (.+?))?$((?:.|\n)*?)End \1"
   Set Matches = r.Execute(Script)
   Dim SFs() As ScriptFunction

   If Matches.Count=0 Then Return SFs
   ReDim SFs(Matches.Count-1)

   Dim MatchIndex As Long, sf As ScriptFunction
   For MatchIndex=0 To Matches.Count-1
      sf=NewScriptFunction(Matches(MatchIndex))
      sf.ClassName=ClassName
      SFs(MatchIndex)=sf
   Next

   ' Add to cache
   Cache.Add(ClassName, SFs)

   Return SFs
End Function


Public Function DevMenu_FilteredFunctions(InputFunctions() As ScriptFunction, ByRef Filtered() As ScriptFunction, Optional IncludeFolder As Boolean=True, Optional IncludeDoc As Boolean=True, Optional IncludeOther As Boolean=True, Optional ContainsStr As String="") As String()
   ' Out variable Filtered provides a filtered array of ScriptFunctions, function returns array of strings corresponding to the function names to use in a menu/dropdown

   Dim Choices() As String, i As Long, f As ScriptFunction, FilterResults As Long, ValidParams As Boolean
   ReDim Choices(0) : ReDim Filtered(0)

   For i=LBound(InputFunctions) To UBound(InputFunctions)
      f=InputFunctions(i)

      ' Match name filter
      If ContainsStr = "" Or InStr(1,f.Name,ContainsStr)>0 Then
         ValidParams=False

         If UBound(f.Params)=-1 OrElse f.Params(0).OptionalParam Then
            If IncludeOther Then ValidParams=True ' No params or all optional
         End If
         If UBound(f.Params)>-1 Then
            Dim p As Param
            p=f.Params(0)
            ' Check for first param of XDoc or XFolder and if that type is included in filter
            If (Replace(LCase(p.ParamType),"cascadelib.","")="cscxfolder" And IncludeFolder) Or _
               (Replace(LCase(p.ParamType),"cascadelib.","")="cscxdocument" And IncludeDoc) Then

               ' Either a single param, or everything after the first is optional
               If UBound(f.Params)=0 OrElse f.Params(1).OptionalParam Then
                  ValidParams=True
               End If
            End If
         End If

         If ValidParams Then
            ReDim Preserve Choices(FilterResults)
            ReDim Preserve Filtered(FilterResults)

            ' Function passes filter so add to filtered functions and menu list
            Choices(FilterResults)=InputFunctions(i).Name
            Filtered(FilterResults)=f
            FilterResults=FilterResults+1
         End If
      End If
   Next

   Return Choices
End Function





Public Sub DevMenu_Dialog(Optional pXFolder As CscXFolder=Nothing, Optional pXDoc As CscXDocument=Nothing)
   ' Show a menu that allows testing functions from the Project script with the provided Folder or Doc

   If Not IsDesignMode() Then Exit Sub
   Debug.Clear

   Dim AllFunc() As ScriptFunction, FilteredFunc() As ScriptFunction
   AllFunc=ParseScript(Project.ScriptCode,"Project") ' TODO: choice for local class script instead of project

   Dim FuncNames() As String
   FuncNames=DevMenu_FilteredFunctions(AllFunc,FilteredFunc)

   ' collect function name prefixes before an underscore to use as preset filter groups
   Dim Prefixes(1000) As String, CurPf As String, Pre As New Dictionary, sf As ScriptFunction
   Prefixes(Pre.Count)="Show All" : Pre.Add("Show All","Show All")
   Prefixes(Pre.Count)="Custom Filter:" : Pre.Add("Custom Filter:","Custom Filter:")
   For Each sf In FilteredFunc
      If UBound(Split(sf.Name,"_"))>0 Then
         CurPf=Split(sf.Name,"_")(0)
         If Not Pre.Exists(CurPf) Then
            Prefixes(Pre.Count)=CurPf
            ' dictionary is only used to check whether we've added something already, contents not used
            Pre.Add(CurPf,CurPf)
         End If
      End If
   Next
   On Error Resume Next
   ' Keep an exported copy of the project script updated.
   ' Called via eval and ignoring errors so this will work if it is in the project, and no error if it is not.
   Eval("Dev_ExportScriptAndLocators()")
   On Error GoTo 0

   Begin Dialog DevDialog 640,399,"Transformation Script Development Menu",.DevMenu_DialogFunc ' %GRID:10,7,1,1
      GroupBox 10,0,620,49,"Function List Filter",.FunctionListFilter
      DropListBox 20,21,180,112,Prefixes(),.FunctionNameFilter
      DropListBox 10,56,310,308,FuncNames(),.FunctionName,2 'Lists the (filtered) functions that can be run
      TextBox 10,84,620,308,.TextBox1,2
      CheckBox 360,21,100,14,"Document",.IncludeDoc
      CheckBox 470,21,70,14,"Folder",.IncludeFolder
      CheckBox 550,21,70,14,"Other",.IncludeOther
      TextBox 210,21,130,21,.CustomFilter
      PushButton 330,56,90,21,"Execute",.Execute
      OKButton 530,42,90,21
      PushButton 460,56,170,21,"Continue parent function",.ContinueParentEvent
   End Dialog

   Dim dlg As DevDialog
   dlg.IncludeDoc = (Not pXDoc Is Nothing)
   dlg.IncludeFolder = (Not pXFolder Is Nothing)
   dlg.IncludeOther = True

   DevMenu_ActiveContext(True,pXFolder,pXDoc)

   Dim Result As Integer
   Result=Dialog(dlg)

   'Debug.Print("Dialog result: " & Result)

   Select Case Result
      Case -1 ' cancel
         End
      Case 0 ' OK (or X)
         End
      Case 1 ' Execute
         End
      Case 2 ' Continue parent function
   End Select

End Sub

Public Sub DevMenu_ActiveContext(SetActiveVars As Boolean,ByRef pXFolder As CscXFolder, ByRef pXDoc As CscXDocument)
   ' The dialog function processing form events does not allow additional parameters and need a way to get doc/folder
   ' This is essentially the same as using global variables, just a bit more structured and contained
   Static ActiveFolder As CscXFolder
   Static ActiveDoc As CscXDocument

   If SetActiveVars Then
      Set ActiveFolder=pXFolder
      Set ActiveDoc=pXDoc
   Else
      Set pXFolder=ActiveFolder
      Set pXDoc=ActiveDoc
   End If
End Sub

Public Sub DevMenu_DialogInitialize()

End Sub

Public Sub DevMenu_DialogUpdate(AllFunc() As ScriptFunction, ByRef FilteredFunc() As ScriptFunction, ByRef FuncNames() As String, ByRef FilterStr As String, ByRef ScriptStatus As String)
   FilterStr=IIf(DlgValue("FunctionNameFilter")=0,"",IIf(DlgValue("FunctionNameFilter")=1,DlgText("CustomFilter"),DlgText("FunctionNameFilter")))
   FuncNames=DevMenu_FilteredFunctions(AllFunc,FilteredFunc,(DlgValue("IncludeFolder")=1),DlgValue("IncludeDoc"),DlgValue("IncludeOther"),FilterStr)
   ScriptStatus="Total Functions: " & UBound(AllFunc) & ", Filtered Functions: " & UBound(FilteredFunc)
   DlgListBoxArray("FunctionName",FuncNames)
   DlgEnable("CustomFilter", (DlgValue("FunctionNameFilter")=1))
End Sub


Public Function DevMenu_DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
   Static FilteredFunc() As ScriptFunction
   Static AllFunc() As ScriptFunction
   Static FuncNames() As String
   Static FilterStr As String
   Static CurSF As ScriptFunction
   Static pXDoc As CscXDocument
   Static pXFolder As CscXFolder
   Static ParamStatus As String
   Static ScriptStatus As String

   Select Case Action%
      Case 1 ' Dialog box initialization
         Debug.Print("Dialog: Dialog initialized.  Window handle: " & SuppValue)
         ' Get the folder/doc since they cannot be passed as a parameter to the dialog function
         DevMenu_ActiveContext(False,pXFolder,pXDoc)

         If pXFolder Is Nothing And pXDoc Is Nothing Then ParamStatus="Warning: Neither an XDoc nor XFolder were provided"
         If Not pXFolder Is Nothing And Not pXDoc Is Nothing Then ParamStatus="Both XDoc and XFolder were provided: Functions will use applicable parameter"
         If Not pXFolder Is Nothing And pXDoc Is Nothing Then ParamStatus="FOLDER MODE: Document functions will run on each document in the folder"
         If pXFolder Is Nothing And Not pXDoc Is Nothing Then ParamStatus="DOCUMENT MODE: Folder functions will run on the doc's parent folder"
         If Not pXFolder Is Nothing Then ParamStatus=ParamStatus & vbNewLine & "Folder contains " & pXFolder.DocInfos.Count & " documents."
         If Not pXDoc Is Nothing Then
            ParamStatus=ParamStatus & vbNewLine & "Doc " & pXDoc.IndexInFolder & "/" & pXDoc.ParentFolder.DocInfos.Count & ": " & pXDoc.FileName
         End If

         Dim ParentFunction As String
         On Error Resume Next
         ParentFunction = Mid(CallersLine(1),InStr(1,CallersLine(1),"|")+1,InStr(1,CallersLine(1),"#")-InStr(1,CallersLine(1),"|")-1)
         On Error GoTo 0
         ParamStatus="Launched from parent function: " & ParentFunction & vbNewLine & ParamStatus

         Dim msg As String
         msg="This menu allows for easy execution of design time scripts that work on individual documents, whole folders, or neither.  It will dynamically list all of the project script functions that require no parameters or require only an XDoc/XFolder (with any amount of optional parameters allowed). " & vbNewLine
         msg=msg & vbNewLine & "It should be used from Project Builder in either of two ways:" & vbNewLine
         msg=msg & "1. Providing an XFolder as a parameter, from an event like Batch_Open. Execute the event from the Runtime Script Events button (lightning bolt)." & vbNewLine
         msg=msg & "2. Providing an XDocument as a parameter, from a document level extraction event like Document_BeforeExtract, in a separate class not otherwise used in the project. Execute the event by selecting the class, selecting the document, then extracting the document." & vbNewLine
         msg=msg & vbNewLine & "When an XFolder is provided, Document functions will run on each doc in the folder.  When an XDoc is provided, Folder functions will run on the doc's parent folder.  Not all operations will work when using the parent folder from a document event."

         DlgText("TextBox1",ParamStatus & vbNewLine & vbNewLine & msg)
         'DlgEnable("TextBox1", False)

         AllFunc=ParseScript(Project.ScriptCode,"Project")
         DlgVisible("OK", False) ' OK button must exist to be able to close diaglog via X
         DevMenu_DialogUpdate(AllFunc,FilteredFunc,FuncNames,FilterStr,ScriptStatus)

      Case 2 ' Value changing or button pressed (button press will close dialog unless returning true)
         Debug.Print "Dialog: " & DlgItem & " (" & DlgType(DlgItem) & ") ";
         If InStr(1,DlgType(DlgItem),"Button")=0 Then Debug.Print "value changed to " & SuppValue & " (" & DlgText(DlgItem) & ")." ; : Debug.Print

         Select Case DlgItem
            Case "FunctionNameFilter", "IncludeFolder", "IncludeDoc", "IncludeOther"
               DevMenu_DialogUpdate(AllFunc,FilteredFunc,FuncNames,FilterStr,ScriptStatus)
            Case "FunctionName"
               CurSF=FilteredFunc(DlgValue("FunctionName"))
               DlgText("TextBox1",ParamStatus & vbNewLine & vbNewLine & ScriptStatus & vbNewLine & vbNewLine & CurSF.Content)
            Case "Execute"
               If DlgValue("FunctionName")=-1 Then
                  Return True
               Else
                  CurSF=FilteredFunc(DlgValue("FunctionName"))
                  DevMenu_Execute(CurSF, pXFolder, pXDoc)
               End If
         End Select
      Case 3 ' TextBox or ComboBox text changed
         Debug.Print("Dialog: Text of " & DlgType(DlgItem) & " " & DlgItem & " changed by " & SuppValue & " characters to " & DlgText(DlgItem))
         Select Case DlgItem
            Case "CustomFilter"
               DevMenu_DialogUpdate(AllFunc,FilteredFunc,FuncNames,FilterStr,ScriptStatus)
         End Select
      Case 4 ' Focus changed
         If SuppValue>-1 Then Debug.Print("Dialog: Focus changing from " & DlgName(SuppValue) & " to " & DlgItem)

      Case 5 ' Idle
         Return False ' Prevent further idle actions

      Case 6 ' Function key
         Debug.Print "Dialog: " & IIf(SuppValue And &H100,"Shift-","") & IIf(SuppValue And &H200,"Ctrl-","") & IIf(SuppValue And &H400,"Alt-","") & "F" & (SuppValue And &HFF)
      Case Else
         Debug.Print("Unknown event " & Join(Array(DlgItem,Action,SuppValue),", "))
   End Select
End Function



Public Sub DevMenu_Execute(sf As ScriptFunction, Optional pXFolder As CscXFolder, Optional pXDoc As CscXDocument)
   ' Execute the function using the right parameter based on what is needed and available.
   ' Initially this used Eval to call the function directly, however any unhandled error within eval context causes a crash
   ' and breakpoints are not hit.  Using Eval to get a delegate then invoking it outside of Eval solves these problems.

   ' Use Eval to get a delegate of the function
   Dim EvalStr As String, DelegateVar As Variant
   EvalStr="AddressOf " & sf.Name
   Debug.Print "Delegate Eval: " & EvalStr
   ' Declaring a staticly typed delegate would require a fixed signature, which would not allow for open-ended optional params, or open ended return types
   DelegateVar=Eval(EvalStr)

   Dim PerDoc As Boolean
   If UBound(sf.Params)>-1 Then
      ' Get the param based on what is needed and what is available
      ' could consider a loop to provide default values or prompt user for additional params
      Dim Param1 As Object
      Select Case Replace(LCase(sf.Params(0).ParamType),"cascadelib.","")
         Case "cscxdocument"
            If Not pXDoc Is Nothing Then
               Set Param1=pXDoc
            Else
               If Not pXFolder Is Nothing AndAlso pXFolder.DocInfos.Count>0 Then
                  ' execute per document in folder
                  PerDoc=True
               End If
            End If
         Case "cscxfolder"
            If Not pXFolder Is Nothing Then
               Set Param1=pXFolder
            Else
               If Not pXDoc Is Nothing Then
                  Debug.Print("Executing function using parent folder of xdoc (overriding single-document mode and folder access permissions).")
                  ' Normally if you go to the parent folder from a doc level event, then back down through the xdocinfos to the xdocs,
                  ' that would result in an error saying that it is not currently possible to access documents.
                  ' Disabling single doc mode will allow access to the xdocs
                  ' These commands are unsupported and have high potential to cause problems: They should never be touched at runtime.
                  pXDoc.ParentFolder.SetSingleDocumentMode(False)
                  pXDoc.ParentFolder.SetFolderAccessPermission(255)
                  Set Param1=pXDoc.ParentFolder
               End If
            End If
      End Select
   End If


   Dim Result As Variant
   If Not Param1 Is Nothing Then
      ' Single param (whether doc, folder, or folder (parent of doc)
      Result = DynamicInvoke(DelegateVar,Param1)
   Else
      If UBound(sf.Params)=-1 OrElse sf.Params(0).OptionalParam Then
         ' Invoke delegate with no params
         Result = DynamicInvoke(DelegateVar)
      ElseIf PerDoc Then
         Debug.Print("Executing document function on each doc in folder.")
         Dim DocIndex As Integer, DocResult As Variant, DocResults As String
         For DocIndex=0 To pXFolder.GetTotalDocumentCount()-1
            Debug.Print("Executing on document " & DocIndex+1 & "/" & pXFolder.DocInfos.Count())
            DocResult = DynamicInvoke(DelegateVar,pXFolder.DocInfos(DocIndex).XDocument)

            On Error Resume Next
               DocResults=DocResults & "Doc " & (DocIndex+1) & ": " & CStr(Result) & vbNewLine
            On Error GoTo 0
         Next
         Result=DocResults
      Else
         Debug.Print("Could not get a valid " & sf.Params(0).ParamType & ", skipping execution.")
      End If
   End If

   ' Output result if it can be converted to string, otherwise resume next
   If Not sf.IsSub Then
      Dim msg As String
      On Error Resume Next
      If CStr(Result)<>"" Then
         msg=sf.Name & " = " & CStr(Result)
      Else
         msg=sf.Name & " = [" & TypeName(Result) & "]"
      End If
      On Error GoTo 0
      Debug.Print(msg)
      MsgBox(msg)
   End If
End Sub



Public Function DynamicInvoke(DelegateVar As Variant, ParamArray Params() As Variant) As Variant
   ' This allows dynamically invoking a delegate regardless of number of parameters, as well as
   ' helping show where a real error occurs within an invoked function.


   On Error Resume Next
   Dim Result As Variant
   Select Case UBound(Params)
      Case -1
         Result=DelegateVar.Invoke()
      Case 0
         Result=DelegateVar.Invoke(Params(0))
      Case 1
         Result=DelegateVar.Invoke(Params(0), Params(1))
      Case 2
         Result=DelegateVar.Invoke(Params(0), Params(1), Params(2))
      Case Else
         Err.Raise("Define more case statements to handle more parameters")
   End Select

   If Err.Number=0 Then
      'Debug.Print("Invoked function completed without error.")
   Else
      ' A quirk of debugging invoked functions: IDE will stop execution (+highlight & focus) at the invoke when an error occurs in the invoked function,
      ' however the text of the line causing the actual error will still be changed to red and can be seen from Err.Description.
      ' Instead, we intentionally stop execution here and print the error message to draw attention to this.
      ' If needed, navigate to the stated line number and add a breakpoint.  Then test again to actually stop execution within the invoked function.
      Debug.Print("Real error in invoked function: " & Err.Description)
      Stop ' Refer to Err.Description to see the line number of the real error.
   End If

   ' Output result if it can be converted to string, otherwise resume next
   'If CStr(Result)<>"" Then
   '   Debug.Print("Invoked function result = " & CStr(Result))
   'Else
   '   Debug.Print("Invoked function result = " & TypeName(Result))
   'End If

   On Error GoTo 0

   Return Result
End Function


Public Function TestScriptFunctions(Optional NameFilter As String) As String
   Dim AllFunc() As ScriptFunction
   AllFunc=ParseScript(Project.ScriptCode,"Project")

   Dim sf As ScriptFunction, msg As String, p As Param, sfline As String
   For Each sf In AllFunc
      sfline=""
      If (NameFilter="" Or InStr(1, sf.Name, NameFilter)>0) Or sf.Suspect Then
         sfline=sfline & IIf(sf.Suspect,"[SUSPECT] ","") & IIf(sf.IsSub,"Sub ", "Function ") & sf.Name & "("
         For Each p In sf.Params
            sfline=sfline & IIf(p.OptionalParam, "Optional ","") & p.Name & IIf(p.Array, "() "," ") & IIf(Len(p.ParamType)>0,"As " & p.ParamType, "") & IIf(Len(p.DefaultValue)>0,"=" & p.DefaultValue,"") & ", "
         Next
         If UBound(sf.Params)>-1 Then sfline=Mid(sfline,1,Len(sfline)-2)
         sfline=sfline & ")" & IIf(Len(sf.ReturnType)>0, " As " & sf.ReturnType, "") & IIf(Len(sf.StringTag)>0," [Tag: " & sf.StringTag & "]", "")
         Debug.Print(sfline)
         msg=sfline & vbNewLine
      End If
   Next

   Return msg
End Function





' Functions to test from DevMenu

Public Sub FolderAsTiffs(pXFolder As CscXFolder)
   Dim DocIndex As Integer
   For DocIndex=0 To pXFolder.GetTotalDocumentCount()-1
      Debug.Print("Executing on document " & DocIndex+1 & "/" & pXFolder.DocInfos.Count())
      ExportDocAsMultipageTiffs(pXFolder.DocInfos(DocIndex).XDocument)
   Next
End Sub

Public Sub ExportDocAsIndividualTiffs(pXDoc As CscXDocument)
   ExportDocAsTiff(pXDoc, GetExportPath(), False)
End Sub

Public Sub ExportDocAsMultipageTiffs(pXDoc As CscXDocument)
   ExportDocAsTiff(pXDoc, GetExportPath(), True)
End Sub

Public Function GetExportPath() As String
   Dim fso As New FileSystemObject
   Dim ExportPath As String
   ExportPath=Project.ScriptVariables("ExportPath")

   If Not fso.FolderExists(ExportPath) Then ExportPath=fso.BuildPath(fso.GetFile(Project.FileName).ParentFolder, "ExportedImages")
   If Not fso.FolderExists(ExportPath) Then fso.CreateFolder(ExportPath)
   Return ExportPath
End Function

Public Sub ExportDocAsTiff(pXDoc As CscXDocument, ExportPath As String, Optional MultiPage As Boolean=True)
   Dim fso As New FileSystemObject
   Dim DocName As String, TempPath As String, TiffPath As String
   DocName=fso.GetBaseName(pXDoc.CDoc.SourceFiles(0).FileName)
   If MultiPage Then
      ' Multipage document named by the filename of the first source file
      TiffPath=fso.BuildPath(ExportPath,DocName & ".tif")
      TempPath=TiffPath & ".tmp"
   Else
      ' Create a folder for this document named by the filename of the first source file
      ExportPath=fso.BuildPath(ExportPath,DocName)
      If Not fso.FolderExists(ExportPath) Then fso.CreateFolder(ExportPath)
   End If


   Debug.Print("Exporting " & pXDoc.CDoc.Pages.Count & " pages from document # " & pXDoc.IndexInFolder+1 & " (" & DocName & ")")
   Dim PageIndex As Integer, img As CscImage, imgformat As CscImageFileFormat
   For PageIndex=0 To pXDoc.CDoc.Pages.Count-1
      If (PageIndex+1) Mod 10 = 0 Then Debug.Print("  Processing page " & PageIndex + 1 & "/" & pXDoc.CDoc.Pages.Count & " from document # " & pXDoc.IndexInFolder+1 & " (" & DocName & ")")

      If MultiPage Then
         Set img=pXDoc.CDoc.Pages(PageIndex).GetImage()
         imgformat=IIf(img.BitsPerSample=1 And img.SamplesPerPixel=1,CscImageFileFormat.CscImgFileFormatTIFFFaxG4,CscImageFileFormat.CscImgFileFormatTIFFOJPG)
         img.StgFilterControl(imgformat, CscStgControlOptions.CscStgCtrlTIFFKeepFileOpen, TempPath, 0, 0)
         img.StgFilterControl(imgformat, CscStgControlOptions.CscStgCtrlTIFFKeepExistingPages, TempPath, 0, 0)
         img.Save(TempPath, imgformat)
      Else
         pXDoc.CDoc.Pages(PageIndex).GetImage().Save(fso.BuildPath(ExportPath,DocName & "-Page-" & Format(PageIndex+1,"000") & ".tif"),CscImgFileFormatTIFFOJPG)
         pXDoc.CDoc.Pages(PageIndex).UnloadImage()
      End If
   Next

   If MultiPage And pXDoc.CDoc.Pages.Count>0 Then
      ' Close the multipage tiff file that was kept open
      Set img=pXDoc.CDoc.Pages(0).GetImage()
      img.StgFilterControl(CscImageFileFormat.CscImgFileFormatTIFFFaxG4, CscStgControlOptions.CscStgCtrlTIFFCloseFile, TempPath, 0, 0)
      pXDoc.CDoc.Pages(0).UnloadImage()

      ' Delete existing file at destination if needed, then move temp file to destination
      If fso.FileExists(TiffPath) Then fso.DeleteFile(TiffPath)
      fso.MoveFile(TempPath,TiffPath)
   End If

   Debug.Print("Finished " & pXDoc.CDoc.Pages.Count & " pages from document # " & pXDoc.IndexInFolder+1 & " (" & DocName & ")")
End Sub



Private Sub Batch_Open(ByVal pXRootFolder As CASCADELib.CscXFolder)
   ' Invoke DevMenu by testing the Batch_Open function (lightning bolt)
   DevMenu_Dialog(pXRootFolder)
End Sub


'========  START DEV EXPORT ========
Public Function ClassHierarchy(KtmClass As CscClass) As String
   ' Given TargetClass, returns Baseclass\subclass\(etc...)\TargetClass\

   Dim CurClass As CscClass, Result As String
   Set CurClass = KtmClass

   While Not CurClass.ParentClass Is Nothing
      Result=CurClass.Name & "\" & Result
      Set CurClass = CurClass.ParentClass
   Wend
   Result=CurClass.Name & "\" & Result
   Return Result
End Function

Public Sub CreateClassFolders(ByVal BaseFolder As String, Optional KtmClass As CscClass=Nothing)
   ' Creates folders in BaseFolder matching the project class structure

   Dim SubClasses As CscClasses
   If KtmClass Is Nothing Then
      ' Start with the project class, but don't create a folder
      Set KtmClass = Project.RootClass
      Set SubClasses = Project.BaseClasses
   Else
      ' Create folder for this class and become the new base folder
      Dim fso As New Scripting.FileSystemObject, NewBase As String
      BaseFolder=fso.BuildPath(BaseFolder,KtmClass.Name)
      If Not fso.FolderExists(BaseFolder) Then
         fso.CreateFolder(BaseFolder)
      End If
      Set SubClasses = KtmClass.SubClasses
   End If

   ' Subclasses
   Dim ClassIndex As Long
   For ClassIndex=1 To SubClasses.Count
      CreateClassFolders(BaseFolder, SubClasses.ItemByIndex(ClassIndex))
   Next
End Sub



Public Sub Dev_ExportScriptAndLocators()
   ' Exports design info (script, locators) to to folders matching the project class structure
   ' Default to \ProjectFolderParent\DevExport\(Class Folders)
   ' Set script variable Dev-Export-BaseFolder to path to override
   ' Set script variable Dev-Export-CopyName-(ClassName) to save a separate named copy of a class script

   ' Make sure you've added the Microsoft Scripting Runtime reference
   Dim fso As New Scripting.FileSystemObject
   Dim ExportFolder As String, ScriptFolder As String, LocatorFolder As String

   ' Either use the provided path or default to the parent of the project folder
   If fso.FolderExists(Project.ScriptVariables("Dev-Export-BaseFolder")) Then
      ExportFolder=Project.ScriptVariables("Dev-Export-BaseFolder")
   Else
      ExportFolder=fso.GetFile(Project.FileName).ParentFolder.ParentFolder.Path & "\DevExport"
   End If

   ' Create folder structure for project classes
   If Not fso.FolderExists(ExportFolder) Then fso.CreateFolder(ExportFolder)
   CreateClassFolders(ExportFolder)

   ' Here we use class index -1 to represent the special case of the project class
   Dim ClassIndex As Long
   For ClassIndex=-1 To Project.ClassCount-1
      Dim KtmClass As CscClass, ClassName As String, ScriptCode As String, ClassPath As String

      ' Get the script of this class
      If ClassIndex=-1 Then
         Set KtmClass=Project.RootClass
         ScriptCode=Project.ScriptCode
      Else
         Set KtmClass=Project.ClassByIndex(ClassIndex)
         ScriptCode=KtmClass.ScriptCode
      End If

      ' TODO: check if script is "empty": Option Explicit \n\n ' Class script: {classname}

      ' Get the name and file path for the class
      ClassPath = fso.BuildPath(ExportFolder, ClassHierarchy(KtmClass))
      ClassName=IIf(ClassIndex=-1,"Project",KtmClass.Name)

      ' TODO: Possibly change to match the naming conventions used in the KTM 6.1.1+ feature to save scripts.

      ' Export script to file
      Dim ScriptFile As TextStream
      Set ScriptFile=fso.CreateTextFile(ClassPath & "\ClassScript-" & ClassName & ".vb",True,False)
      ScriptFile.Write(ScriptCode)
      ScriptFile.Close()

      ' Save a copy if a name is defined
      Dim CopyName As String
      CopyName=Project.ScriptVariables("Dev-Export-CopyName-" & ClassName)

      If Not CopyName="" Then
         Set ScriptFile=fso.CreateTextFile(ClassPath & "\" & CopyName & ".vb",True,False)
         ScriptFile.Write(ScriptCode)
         ScriptFile.Close()
      End If

      ' Export locators (same as from Project Builder menus)
      Dim FileName As String
      Dim LocatorIndex As Integer
      For LocatorIndex=0 To KtmClass.Locators.Count-1
         If Not KtmClass.Locators.ItemByIndex(LocatorIndex).LocatorMethod Is Nothing Then
            FileName="\" & ClassName & "-" & KtmClass.Locators.ItemByIndex(LocatorIndex).Name & ".loc"
            KtmClass.Locators.ItemByIndex(LocatorIndex).ExportLocatorMethod(ClassPath & FileName, ClassPath)
         End If
      Next
   Next
End Sub
'========  END   DEV EXPORT ========


