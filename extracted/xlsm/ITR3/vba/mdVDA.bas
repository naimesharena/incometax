Attribute VB_Name = "mdVDA"
Option Explicit
Public msgValidateSheetVDA As String

Sub Cmd_Validate_VDA_Click()
ValidateScheduleVDA
fmsgboxStatus "Sheet VDA  is ok" ', vbOKOnly, "ITR 3"
End Sub

Sub ValidateScheduleVDA()
    subProcCaption = "Validating VDA"
    If Not ValidateScheduleVDA1 Then
        Sheet47.Activate
        fmsgboxsmall msgValidateSheetVDA ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
End Sub


Function ValidateScheduleVDA1() As Boolean
ValidateScheduleVDA1 = True

If Not ValidateMandatoryVDA() Then ValidateScheduleVDA1 = False

End Function
Function msgbox_VDA(strmsg As String) As String
     msgValidateSheetVDA = msgValidateSheetVDA & strmsg & Chr(13)
End Function


Function ValidateMandatoryVDA() As Boolean
ValidateMandatoryVDA = True
Dim i As Long

For i = 2 To Sheet47.Range("VDA_Date_of_Acquisition").count
    If (Sheet47.Range("VDA_Date_of_Acquisition").item(i).value <> "") Then
        If Sheet47.Range("VDA_Date_of_Acquisition").item(i - 1).value = "" Then
            msgbox_VDA ("* Please fill all the mandatory fields of the row " & i - 1 & " before filling the next row in Schedule VDA" & Chr(13))
            ValidateMandatoryVDA = False
            Exit Function
        End If
    End If
Next

For i = 1 To Sheet47.Range("VDA_Date_of_Acquisition").count

    If Sheet47.Range("VDA_Date_of_Acquisition").Cells(i, 1).value <> "" Or Sheet47.Range("VDA_Date_of_Transfer").Cells(i, 1).value <> "" Or _
        (Sheet47.Range("VDA_Head_Income").Cells(i, 1).value <> "" And Sheet47.Range("VDA_Head_Income").Cells(i, 1).value <> "(Select)") Or _
        Sheet47.Range("VDA_Cost_of_Acquisition").Cells(i, 1).value <> "" Or Sheet47.Range("VDA_Consideration_Received").Cells(i, 1).value <> "" Then
        'Some thing is filled in the row
        
        If Sheet47.Range("VDA_Date_of_Acquisition").Cells(i, 1).value = "" Then
            ValidateMandatoryVDA = False
            msgbox_VDA ("Please select the date of Acquisition at Sr. No. " & i)
        End If
        
        If Sheet47.Range("VDA_Date_of_Transfer").Cells(i, 1).value = "" Then
            ValidateMandatoryVDA = False
            msgbox_VDA ("Please select the date of Transfer at Sr. No. " & i)
        End If
        
        If Sheet47.Range("VDA_Head_Income").Cells(i, 1).value = "" Or Sheet47.Range("VDA_Head_Income").Cells(i, 1).value = "(Select)" Then
            ValidateMandatoryVDA = False
            msgbox_VDA ("Please select head of income at Sr. No. " & i)
        End If
        
        If Sheet47.Range("VDA_Cost_of_Acquisition").Cells(i, 1).value = "" Then
            ValidateMandatoryVDA = False
            msgbox_VDA ("Please enter Cost of Acquisition at Sr. No. " & i)
        End If
        
        If Sheet47.Range("VDA_Consideration_Received").Cells(i, 1).value = "" Then
            ValidateMandatoryVDA = False
            msgbox_VDA ("Please enter Consideration Received at Sr. No. " & i)
        End If
        
    End If
Next i

End Function

Sub AddRowsVDA()
    Sheets("VDA").Activate
    EfilingCommon.DefinedgridNameRange = "VDA_Date_of_Acquisition||VDA_Date_of_Transfer||VDA_Head_Income||VDA_Cost_of_Acquisition||VDA_Consideration_Received||VDA_Income_from_transfer"
    ActiveCellRange = EfilingCommon.searchLastRow("VDA_Date_of_Acquisition")
    EfilingCommon.insertRowUnderSectionWithFormula
End Sub


Sub AddRowsVDA_Import(DiffRows As Long)
    Sheets("VDA").Activate
    Dim rngname_VDA As Variant
'    rngname_VDA = "VDA_Date_of_Acquisition||VDA_Date_of_Transfer||VDA_Head_Income||VDA_Cost_of_Acquisition||VDA_Consideration_Received||VDA_Income_from_transfer"
    rngname_VDA = "VDA_Date_of_Acquisition;VDA_Date_of_Transfer;VDA_Head_Income;VDA_Cost_of_Acquisition;VDA_Consideration_Received;VDA_Income_from_transfer;"
    SelectLastRow ("VDA_Date_of_Acquisition")
    InsertDiffRowsAndFillFormulas (DiffRows)
    Call ExendRangeNameToTable(DiffRows, rngname_VDA)
End Sub

