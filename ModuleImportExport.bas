Attribute VB_Name = "ModuleImportExport"
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
