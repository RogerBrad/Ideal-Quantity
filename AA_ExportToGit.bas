Attribute VB_Name = "AA_ExportToGit"
Option Explicit
'Remember to add a reference to Microsoft Visual Basic for Applications Extensibility
'Exports all VBA project components containing code to a folder in the same directory as this spreadsheet.
'
' Destination directory: "C:\Users\rbradfield\Documents\GitHub\Promotions\Live6"
'
'

Public Sub ExportAllComponents()
    Dim VBComp As VBIDE.VBComponent
    Dim destDir As String, fName As String, ext As String
    'Create the directory where code will be created.
    'Alternatively, you could change this so that the user is prompted
    If ActiveWorkbook.path = "" Then
        MsgBox "You must first save this workbook somewhere so that it has a path.", , "Error"
        Exit Sub
    End If
    destDir = ActiveWorkbook.path & "\" & ActiveWorkbook.name & " Modules"
    destDir = "C:\Users\rbradfield\Documents\Test"
    destDir = "C:\Users\rbradfield\Documents\GitHub\Promotions\Live6"
    destDir = "C:\Users\rbradfield\Documents\OK\Services2\LivePromos\IdealQty\Git\Ideal-Quantity"
    
    If Dir(destDir, vbDirectory) = vbNullString Then MkDir destDir
    
    'Export all non-blank components to the directory
    For Each VBComp In ActiveWorkbook.VBProject.VBComponents
        If VBComp.codeModule.CountOfLines > 0 Then
            'Determine the standard extention of the exported file.
            'These can be anything, but for re-importing, should be the following:
            Select Case VBComp.Type
                Case vbext_ct_ClassModule: ext = ".cls"
                Case vbext_ct_Document: ext = ".cls"
                Case vbext_ct_StdModule: ext = ".bas"
                Case vbext_ct_MSForm: ext = ".frm"
                Case Else: ext = vbNullString
            End Select
            If ext <> vbNullString Then
                fName = destDir & "\" & VBComp.name & ext
                'Overwrite the existing file
                'Alternatively, you can prompt the user before killing the file.
                If Dir(fName, vbNormal) <> vbNullString Then kill (fName)
                VBComp.Export (fName)
            End If
        End If
    Next VBComp
End Sub


