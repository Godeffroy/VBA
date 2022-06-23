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
    Debug.Print CreateFolder(GetPersonalPath() & "VBAProjectFiles")
    folder = Environ("Temp") & "\VBAProjectFiles-" & Replace(Replace(Replace(Now(), " ", "-"), ":", "-"), "/", "-") & "\"
    ExportModules (folder)
    DeleteVBAModulesAndUserForms "PERSONAL.XLSB", "ModuleImportExport"
    
    ImportModules folder, True, "PERSONAL.XLSB", "ModuleImportExport"
    
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
