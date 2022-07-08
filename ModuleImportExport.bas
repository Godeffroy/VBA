Option Explicit

'Within Excel you need to set a reference to the VB script run-time library. The relevant file is usually located at \Windows\System32\scrrun.dll
'
'    To reference this file, load the Visual Basic Editor (ALT+F11)
'    Select Tools > References from the drop-down menu
'    A listbox of available references will be displayed
'    Tick the check-box next to
'    'Microsoft Scripting Runtime'
'    'Microsoft Visual Basic For Applications Extensibility'
'    'Microsoft XML, V6.0'
'    The full name and path of the scrrun.dll file will be displayed below the listbox
'    Click on the OK button.
' Source : https://stackoverflow.com/questions/3233203/how-do-i-use-filesystemobject-in-vba
' Source : https://www.excelforum.com/excel-programming-vba-macros/561236-dim-vbcomp-as-vbide-vbcomponent.html

'In the VBE Editor set a reference to "Microsoft Visual Basic For Applications Extensibility 5.3" and to "Microsoft Scripting Runtime" and then save the file.

'You also need to enable programmatic access to the VBA Project in Excel. In Excel 2003 and earlier,
'go the Tools>Macros>Security(in Excel), click on the Trusted Publishers tab and check the Trust access to the Visual Basic Project setting.
'In Excel 2007-2013, click the Developer tab and then click the Macro Security item.
'In that dialog, choose Macro Settings and check the Trust access to the VBA project object model.
'You can also try the shortcut ALT tms to go to this dialog.
'Source : https://www.rondebruin.nl/win/s9/win002.htm

Public Sub Main()

    Dim folder As String
    Dim urls() As Variant
    
    Debug.Print CreateFolder(GetPersonalPath() & "VBAProjectFiles")
    folder = Environ("Temp") & "\VBAProjectFiles-" & Replace(Replace(Replace(Now(), " ", "-"), ":", "-"), "/", "-") & "\"
    'ExportModules (folder)
    'DeleteVBAModulesAndUserForms "PERSONAL.XLSB", "ModuleImportExport"
    
    urls = Array( _
           "https://drive.google.com/file/d/1P_DmnMt32gWkMop66gdHXElEhLQhZ7zl/view?usp=sharing", _
           "https://drive.google.com/file/d/1PeZzCoY0WaeLz8pAJSXKSNibNqsFSU_w/view?usp=sharing", _
           "https://drive.google.com/file/d/1PW8_0cj6FqyA1r4Kp9Dv0byYR0NFR4Np/view?usp=sharing", _
           "https://drive.google.com/file/d/1Pg3EFkECQqlY2Lx7Z8RxPAepeQx68N_E/view?usp=sharing", _
           "https://drive.google.com/file/d/1PmBBovncODg8oBiqLDVZIUlqQrSWg33G/view?usp=sharing", _
           "https://drive.google.com/file/d/1PaFmzK-_8WW2R8UYOhDK-M78iHyQqkbu/view?usp=sharing", _
           "https://drive.google.com/file/d/1PkqSzsPJiLNUdyndH7wBx-WOs_TywgwO/view?usp=sharing")
    
    DownloadGoogleDriveWithFilename folder, urls

    'ImportModules folder, True, "PERSONAL.XLSB", "ModuleImportExport"
    
End Sub

Private Function CreateFolder(FolderPath As String) As String
    CreateFolder = "Error"
    If Not IsFolderExist(FolderPath) Then
        On Error Resume Next
        MkDir FolderPath
        On Error GoTo 0
    End If
    If IsFolderExist(FolderPath) Then
        CreateFolder = FolderPath
    End If
End Function

Private Function IsFolderExist(FolderPath As String) As Boolean
    IsFolderExist = False
    Dim FSO As Object
    Set FSO = CreateObject("scripting.filesystemobject")
    If Right(FolderPath, 1) <> "\" Then
        FolderPath = FolderPath & "\"
    End If
    IsFolderExist = FSO.FolderExists(FolderPath)
End Function

Private Function GetPersonalPath() As String
    Dim WshShell As Object
    Dim appData As String
    Set WshShell = CreateObject("WScript.Shell")
    appData = WshShell.expandEnvironmentStrings("%APPDATA%")
    GetPersonalPath = appData + "\Microsoft\Excel\XLSTART\"
End Function

Private Function DeleteVBAModulesAndUserForms(szSourceWorkbook As String, Optional Ignored As String = "")
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent

    Set VBProj = Application.Workbooks(szSourceWorkbook).VBProject

    For Each VBComp In VBProj.VBComponents
        If VBComp.Type = vbext_ct_Document Then
            'Thisworkbook or worksheet module
            'We do nothing
        ElseIf (Ignored = "" Or Not VBComp.Name Like ("*" + Ignored + "*")) Then
            'ignore modules which name contain something
            VBProj.VBComponents.Remove VBComp
        End If
    Next VBComp
End Function

Private Function DeleteAllFiles(szExportPath As String)
    Kill szExportPath & "*.*"
End Function

Private Sub ExportModules(ByVal ExportPath As String, Optional ByVal CreateSubFolder As Boolean = True, Optional ByVal szSourceWorkbook As String = "PERSONAL.XLSB")
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szFileName As String
    Dim szExportPath As String
    Dim cmpComponent As VBIDE.VBComponent
    Dim ExportFolderPath As String
    
    szExportPath = CreateFolder(ExportPath)
    
    If szExportPath = "Error" Then
        MsgBox "Export Folder does not exist"
        Exit Sub
    End If
    
    If CreateSubFolder Then
        szExportPath = CreateFolder(ExportPath + "VBAProjectFiles")
    
        If szExportPath = "Error" Then
            MsgBox "Export Folder does not exist"
            Exit Sub
        End If
    End If
        
    'szSourceWorkbook = ActiveWorkbook.Name
    Set wkbSource = Application.Workbooks(szSourceWorkbook)

    If wkbSource.VBProject.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected," & _
               "not possible to export the code"
        Exit Sub
    End If

    For Each cmpComponent In wkbSource.VBProject.VBComponents

        bExport = True
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
        Case vbext_ct_ClassModule
            szFileName = szFileName & ".cls"
        Case vbext_ct_MSForm
            szFileName = szFileName & ".frm"
        Case vbext_ct_StdModule
            szFileName = szFileName & ".bas"
        Case vbext_ct_Document
            ''' This is a worksheet or workbook object.
            ''' Don't try to export.
            bExport = False
        End Select
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            ''' remove it from the project if you want
            '''wkbSource.VBProject.VBComponents.Remove cmpComponent
        End If
    Next cmpComponent
    If MsgBox("Ouvrir le dossier ?", vbQuestion + vbYesNo + vbDefaultButton2, "Export OK !") = vbYes Then
        ''open folder path location to look at the files
        Call Shell("explorer.exe" & " " & szExportPath, vbNormalFocus)
    End If
End Sub

Public Sub ImportModules(ByVal ImportPath As String, Optional ByVal SubFolder As Boolean = True, Optional ByVal szTargetWorkbook As String = "PERSONAL.XLSB", Optional Ignored As String = "")
    Dim wkbTarget As Excel.Workbook
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim cmpComponents As VBIDE.VBComponents
    Dim szImportPath As String
    
    If Right(ImportPath, 1) <> "\" Then
        ImportPath = ImportPath & "\"
    End If

    If Not IsFolderExist(ImportPath) Then
        MsgBox "Import Folder does not exist"
        Exit Sub
    End If

    If SubFolder Then
        If Not IsFolderExist(ImportPath + "VBAProjectFiles") Then
            MsgBox "Import Sub Folder does not exist"
            Exit Sub
        End If
    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = ImportPath + "VBAProjectFiles"
    
    Set wkbTarget = Application.Workbooks(szTargetWorkbook)

    If wkbTarget.VBProject.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected," & _
               "not possible to Import the code"
        Exit Sub
    End If

    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
        MsgBox "There are no files to import"
        Exit Sub
    End If

    Set cmpComponents = wkbTarget.VBProject.VBComponents

    ''' Import all the code modules in the specified path to the Workbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files
        On Error Resume Next
        If objFile.Name Like ("*" + Ignored + "*") Then
            'To skip modules which name contain "ModuleImportExport"
        Else
            If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
                                                               (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
                                                               (objFSO.GetExtensionName(objFile.Name) = "bas") Then
                cmpComponents.Import objFile.path
            End If
        End If
    Next objFile

    MsgBox "Import OK !"
End Sub

Private Function FileExists(FilePath As String) As Boolean
    Dim TestStr As String
    TestStr = ""
    On Error Resume Next
    TestStr = Dir(FilePath)
    On Error GoTo 0
    If TestStr = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function

Private Function DownloadGoogleDriveWithFilename(ByVal DownloadPath As String, myOriginalURLs() As Variant, Optional ByVal CreateSubFolder As Boolean = True) As Boolean
    Dim myURL As String
    Dim FileID As String
    Dim xmlhttp As Object
    Dim FolderPath As String
    Dim FilePath As String
    Dim name0 As Variant
    Dim oStream As Object
    Dim user As String
    Dim mdp As String

    Dim myOriginalURL As Variant
    Dim xmlhttptemp As Variant
    Dim myXmlhttps() As Object
    Dim TimeOut As Single

    Dim objFSO As Scripting.FileSystemObject
    Dim message As String
    Dim result As Boolean
    Dim time As String

    DownloadGoogleDriveWithFilename = False
    Application.ScreenUpdating = False
    Debug.Print "Starting ..."
    'URL from share link or Google sheet URL or Google doc URL
    Dim i As Integer
    i = 0
    user = ""
    mdp = ""
    user = InputBox("adresse gmail")
    If user <> "" Then
        mdp = InputBox("adresse gmail")
    End If
        
    FolderPath = CreateFolder(DownloadPath)
    
    If FolderPath = "Error" Then
        MsgBox "Export Folder does not exist"
        Exit Function
    End If
    
    If CreateSubFolder Then
        FolderPath = CreateFolder(FolderPath + "VBAProjectFiles")
    
        If FolderPath = "Error" Then
            MsgBox "Download Folder does not exist"
            Exit Function
        End If
    End If
     
    For Each myOriginalURL In myOriginalURLs
    
        ReDim Preserve myXmlhttps(1 To i + 1) As Object
    
        FileID = Split(myOriginalURL, "/d/")(1)  ''split after "/d/"
        FileID = Split(FileID, "/")(0)           ''split before "/"
        'Const UrlLeft As String = "http://drive.google.com/u/0/uc?id="
        Const UrlLeft As String = "http://drive.google.com/uc?id="
        Const UrlRight As String = "&export=download&confirm=t"
        myURL = UrlLeft & FileID & UrlRight
        Debug.Print myURL

        'Set xmlhttp = CreateObject("winhttp.winhttprequest.5.1")
        Set myXmlhttps(i + 1) = CreateObject("Msxml2.ServerXMLHTTP.6.0") 'New MSXML2.ServerXMLHTTP60 '
        
        myXmlhttps(i + 1).Open "GET", myURL
        
        myXmlhttps(i + 1).setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        myXmlhttps(i + 1).setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:47.0) Gecko/20100101 Firefox/47.0"
        If user <> "" And mdp <> "" Then
            myXmlhttps(i + 1).setRequestHeader "Authorization", "Basic " + Base64Encode(user + ":" + mdp)
        End If
        myXmlhttps(i + 1).Send
        'Set myXmlhttps(i + 1) = xmlhttp
        i = i + 1
    Next myOriginalURL

    For Each xmlhttptemp In myXmlhttps
        TimeOut = Timer
   
        Do While xmlhttptemp.readyState <> 4     'And xmlhttptemp.Status <> 200
            If (Timer - TimeOut > 10) Then
                MsgBox "TimeOut : " & Timer - TimeOut & " ReadyState=" & xmlhttptemp.readyState
                Exit Function                    'Do
            End If
            DoEvents
            Application.Wait (Now + TimeValue("00:00:01"))
        Loop
        Debug.Print Timer - TimeOut & " ReadyState=" & xmlhttptemp.readyState & " Status=" & xmlhttptemp.Status

        If xmlhttptemp.Status = 200 Then
            name0 = xmlhttptemp.getResponseHeader("Content-Disposition")
            If name0 = "" Then
                MsgBox "file name not found"
                Exit Function
            End If
        
            Debug.Print name0
            name0 = Split(name0, "=""")(1) ''split after "=""
            name0 = Split(name0, """;")(0)  ''split before "";"
            '        name0 = Replace(name0, """", "") ' Remove double quotes
            Debug.Print name0
        
            FilePath = FolderPath & name0
            ''This part is equivalent to URLDownloadToFile(0, myURL, FolderPath & "\" & name0, 0, 0)
            ''just without having to write Windows API code for 32 bit and 64 bit.
        
            Set oStream = CreateObject("ADODB.Stream")
            oStream.Open
            oStream.Type = 1
            oStream.Write xmlhttptemp.responseBody
            oStream.SaveToFile FilePath, 2       ' 1 = no overwrite, 2 = overwrite
            oStream.Close
        End If
    Next xmlhttptemp

    Application.ScreenUpdating = True
  
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(FolderPath).Files.Count = i Then
        message = "Download OK !"
        DownloadGoogleDriveWithFilename = True
    Else
        message = "Download Failed !"
    End If

    If MsgBox("Ouvrir le dossier ?", vbQuestion + vbYesNo + vbDefaultButton2, message) = vbYes Then
        ''open folder path location to look at the files
        Call Shell("explorer.exe" & " " & FolderPath, vbNormalFocus)
    End If
    Debug.Print "-- End --"
End Function


Function Base64Encode(sText)
    Dim oXML, oNode
    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.createElement("base64")
    oNode.DataType = "bin.base64"
    oNode.nodeTypedValue = Stream_StringToBinary(sText)
    Base64Encode = oNode.Text
    Set oNode = Nothing
    Set oXML = Nothing
End Function


'Stream_StringToBinary Function
'2003 Antonin Foller, http://www.motobit.com
'Text - string parameter To convert To binary data
Function Stream_StringToBinary(Text)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.Charset = "us-ascii"

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.WriteText Text

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeBinary

  'Ignore first two bytes - sign of
  BinaryStream.Position = 0

  'Open the stream And get binary data from the object
  Stream_StringToBinary = BinaryStream.Read

  Set BinaryStream = Nothing
End Function

'Stream_BinaryToString Function
'2003 Antonin Foller, http://www.motobit.com
'Binary - VT_UI1 | VT_ARRAY data To convert To a string
Function Stream_BinaryToString(Binary)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeBinary

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.Write Binary

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.Charset = "us-ascii"

  'Open the stream And get binary data from the object
  Stream_BinaryToString = BinaryStream.ReadText
  Set BinaryStream = Nothing
End Function
