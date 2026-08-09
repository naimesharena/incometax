Attribute VB_Name = "Initializer"
Option Explicit

Sub Intialize()
On Error Resume Next
'Intial level setting for Workbook--------------------------------------------------------Start---
Dim strMsgOnStart, strMsgOnStart1, strMsgOnStart2, vbMessgaeCaption As String
'If ThisWorkbook.ProtectStructure = False Then ThisWorkbook.Protect Password:=getmsgstate
    With Application
        .caption = ""
        
        'Cut
        .CommandBars("Worksheet Menu Bar").Controls("Cut").Controls("Rows").Enabled = False
        .CommandBars("Row").Controls("Cut").Enabled = False
        
        .CommandBars("Column").Controls("Cut").Enabled = False
        .CommandBars("Worksheet Menu Bar").Controls("Cut").Controls("Columns").Enabled = False
        
        .CommandBars("Standard").Controls.item("Cut").Enabled = False
        .CommandBars("Edit").Controls.item("Cut").Enabled = False
        .CommandBars("Cell").Controls.item("Cut").Enabled = False

        'Insert
        .CommandBars("Row").Controls("Insert").Enabled = False
        .CommandBars("Worksheet Menu Bar").Controls("Insert").Controls("Rows").Enabled = False
        .CommandBars("Worksheet Menu Bar").Controls("Insert...").Controls("Rows").Enabled = False
        
        .CommandBars("Column").Controls("Insert").Enabled = False
        .CommandBars("Worksheet Menu Bar").Controls("Insert...").Controls("Columns").Enabled = False
        .CommandBars("Worksheet Menu Bar").Controls("Insert").Controls("Columns").Enabled = False

        .CommandBars("Cell").Controls.item("Insert...").Enabled = False
        
        'Delete
        .CommandBars("Row").Controls("Delete").Enabled = False
        .CommandBars("Worksheet Menu Bar").Controls("Delete").Controls("Rows").Enabled = False
        .CommandBars("Worksheet Menu Bar").Controls("Delete...").Controls("Rows").Enabled = False

        .CommandBars("Column").Controls("Delete").Enabled = False
        .CommandBars("Worksheet Menu Bar").Controls("Delete").Controls("Columns").Enabled = False
        .CommandBars("Worksheet Menu Bar").Controls("Delete...").Controls("Columns").Enabled = False
        .CommandBars("Cell").Controls.item("Delete...").Enabled = False
         
        'Other Properties Set
        .OnKey "^x", "DisableCut"
        .WindowState = xlMaximized
        .DisplayFormulaBar = False
        .Calculation = xlAutomatic
        '.DefaultFilePath = "C:\My Documents\Excel"
        .CellDragAndDrop = False
        '.AutoCorrect = True
        '.AutoRecover = True
        .EnableEvents = True
        .MaxChange = 0.001
    End With
    UserForm2.Show

    Sheet33.EnableCalculation = False
    Sheet35.Activate
        If UCase(ActiveSheet.name) <> UCase("Home") Then
          Sheet35.Activate
        End If
    
'Intial level setting for Workbook---------------------------------------------------------End---


End Sub

Sub DisableCut()
fmsgboxStatus "You are not allowed to Cut a cell. instead use copy function." ', vbInformation, "Error"
End Sub





