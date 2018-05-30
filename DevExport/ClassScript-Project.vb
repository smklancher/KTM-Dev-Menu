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
      .Content=m.SubMatches(4)
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



Public Function ParseScript(ByVal Script As String, Optional ByVal ClassName As String = "") As Dictionary
   ' Parse script into dictionary of FunctionName, ScriptFunction
   ' To work around a limitiation, the ScriptFunction is wrapped in an array

   Dim r As New RegExp, Matches As MatchCollection
   r.Global=True:   r.Multiline=True:   r.IgnoreCase=True
   'vbscript regexp does not support singleline mode (. matches \n) and no support for named capturing groups
   r.Pattern = "^(?:Public |Private )?(Sub|Function) (.*?)\((.*?)\)\s*(?: As (.+?))?$((?:.|\n)*?)End \1"
   Set Matches = r.Execute(Script)

   Dim MatchIndex As Long, sf As ScriptFunction, ScriptFunctions As New Dictionary
   For MatchIndex=0 To Matches.Count-1
      sf=NewScriptFunction(Matches(MatchIndex))
      sf.ClassName=ClassName
      ScriptFunctions.Add(sf.Name,Array(sf)) ' A UDT cannot be stored as a variant (no error but no data), wrap in array as workaround
   Next

   Return ScriptFunctions
End Function

Public Function DevMenu_Items() As Dictionary
   ' Parse the functions in the Project script, using StringTag to mark functions with valid parameters to be called from the menu
   ' Returning a dictionary of MenuText, ScriptFunction

   Dim Functions As Dictionary,f As ScriptFunction, v As Variant
   Set Functions = ParseScript(Project.ScriptCode, "Project")
   Dim Items As New Dictionary

   ' filter functions to those that can be called from menu
   For Each v In Functions.Items
      f=v(0) ' A UDT cannot be stored as a variant (no error but no data), wrap in array as workaround
      If UBound(f.Params)=-1 OrElse f.Params(0).OptionalParam Then
         f.StringTag="NoParam"
         Items.Add(f.StringTag & ": " & f.Name,Array(f))
      End If
      If UBound(f.Params)>-1 Then
         Dim p As Param
         p=f.Params(0)
         ' Check for first param of XDoc or XFolder
         If Replace(p.ParamType,"CASCADELib.","")="CscXFolder" Or Replace(p.ParamType,"CASCADELib.","")="CscXDocument" Then
            f.StringTag=Replace(Replace(p.ParamType,"CASCADELib.",""),"CscX","") ' Tag as Document or Folder

            ' Either a single param, or everything after the first is optional
            If UBound(f.Params)=0 OrElse f.Params(1).OptionalParam Then
               Items.Add(f.StringTag & ": " & f.Name,Array(f))
            End If
         End If
      End If
   Next
   Return Items
End Function

Public Sub DevMenu_Show(Optional pXFolder As CscXFolder=Nothing, Optional pXDoc As CscXDocument=Nothing)
   ' Show a menu that allows testing functions from the Project script with the provided Folder or Doc

   If Not IsDesignMode() Then Exit Sub

   ' Keep an exported copy of the project script updated.
   ' Called via eval and ignoring errors so this will work if it is in the project, and no error if it is not.
   On Error Resume Next
   Eval("Dev_ExportScriptAndLocators()")
   On Error GoTo 0

   Debug.Clear
   Dim Choices(1000) As String, Choice As Integer, Items As Dictionary
   Set Items=DevMenu_Items()

   Choices(0)="      Transformation Script Development Menu"
   If pXFolder Is Nothing And pXDoc Is Nothing Then Choices(1)="Warning: Neither an XDoc nor XFolder were provided"
   If Not pXFolder Is Nothing And Not pXDoc Is Nothing Then Choices(1)="Both XDoc and XFolder were provided: Functions will use applicable parameter"
   If Not pXFolder Is Nothing And pXDoc Is Nothing Then Choices(1)="FOLDER MODE: Document functions will run on each document in the folder"
   If pXFolder Is Nothing And Not pXDoc Is Nothing Then Choices(1)="DOCUMENT MODE: Folder functions will run on the doc's parent folder"
   Choices(2)="About this menu..."
   Choices(3)="-"
   Dim ChoiceOffset As Integer
   ChoiceOffset=4

   ' Set descriptions of functions
   For Choice=0 To Items.Count-1
      Choices(Choice+ChoiceOffset)=Items.Keys(Choice)
   Next

   Choice=ShowPopupMenu(Choices,vbPopupLeftTopAlign,300,0)
   If Choice>-1 Then Debug.Print("DevMenu selection: " & Choices(Choice) & " (#" & Choice & ")")
   Select Case Choice
      Case 0,1,2,3
         Dim msg As String
         msg="This menu allows for easy execution of design time scripts that work on individual documents, whole folders, or neither.  It will dynamically list all of the project script functions that require no parameters or require only an XDoc/XFolder (with any amount of optional parameters allowed). " & vbNewLine
         msg=msg & vbNewLine & "It should be used from Project Builder in either of two ways:" & vbNewLine
         msg=msg & "1. Providing an XFolder as a parameter, from an event like Batch_Open. Execute the event from the Runtime Script Events button (lightning bolt)." & vbNewLine
         msg=msg & "2. Providing an XDocument as a parameter, from a document level extraction event like Document_BeforeExtract, in a separate class not otherwise used in the project. Execute the event by selecting the class, selecting the document, then extracting the document." & vbNewLine
         msg=msg & vbNewLine & "When an XFolder is provided, Document functions will run on each doc in the folder.  When an XDoc is provided, Folder functions will run on the doc's parent folder.  Not all operations will work when using the parent folder from a document event."
         MsgBox(msg,,"Transformation Script Development Menu")
      Case Else
         If Choice>ChoiceOffset Then DevMenu_Execute(Items.Items(Choice-ChoiceOffset)(0),pXFolder,pXDoc)
   End Select
End Sub


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


   ' Invoke delegate with correct params
   Dim Result As Variant
   If UBound(sf.Params)=-1 Then
      DynamicInvoke(DelegateVar)
   Else
      ' Get the param based on what is needed and what is available
      Dim Param1 As Object
      Select Case sf.StringTag
         Case "Document"
            If Not pXDoc Is Nothing Then
               Set Param1=pXDoc
            Else
               If Not pXFolder Is Nothing AndAlso pXFolder.DocInfos.Count>0 Then
                  Debug.Print("Executing document function on each doc in folder.")

                  Dim DocIndex As Integer
                  For DocIndex=0 To pXFolder.GetTotalDocumentCount()-1
                     Debug.Print("Executing on document " & DocIndex+1 & "/" & pXFolder.DocInfos.Count())
                     DynamicInvoke(DelegateVar,pXFolder.DocInfos(DocIndex).XDocument)
                  Next
                  Exit Sub

               End If
            End If
         Case "Folder"
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

      If Param1 Is Nothing Then
         Debug.Print("Could not get a valid " & sf.StringTag & ", skipping execution.")
      Else
         DynamicInvoke(DelegateVar,Param1)
      End If
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
         Err.Raise("Define more cases statements to handle more parameters")
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
   If CStr(Result)<>"" Then
      Debug.Print("Invoked function result = " & CStr(Result))
   Else
      Debug.Print("Invoked function result = " & TypeName(Result))
   End If

   On Error GoTo 0

   Return Result
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
   DevMenu_Show(pXRootFolder)
   Dev_ExportScriptAndLocators()
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


