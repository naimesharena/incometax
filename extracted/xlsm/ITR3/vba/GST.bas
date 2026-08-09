Attribute VB_Name = "GST"

Public MsgBox_GST As String
Public end_GSTIN, end_GSTAMOUNT, GSTINAMOUNT_BP, GSTIN_BP As Variant

Sub ValidateSheetGST_Click()
ValidateGSTNew
fmsgboxStatus "Sheet GST is ok" ', vbOKOnly, "ITR-3"
End Sub

Sub AddRowsGST_Click()
Dim vRows  As Long
Sheets("GST").Activate
EfilingCommon.DefinedgridNameRange = "GST_GSTRno||GST_GSTRAmount"
ActiveCellRange = EfilingCommon.searchLastRow("GST_GSTRno")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub


Sub ValidateGSTNew()
    If Not ValidateGST1 Then
        Sheet51.Activate
        fmsgboxsmall MsgBox_GST ', vbOKOnly, "Error(s)"
        CloseMsg
    End If
End Sub

Sub setTblinfo_GSTIN()
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet51.Range("GST_GSTRno").count
 Set rangecells = Sheet51.Range("GST_GSTRno").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_GSTIN = ccount
 
 End Sub
 
 
 Sub setTblinfo_GSTAMOUNT()
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet51.Range("GST_GSTRAmount").count
 Set rangecells = Sheet51.Range("GST_GSTRAmount").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_GSTAMOUNT = ccount
 
 End Sub
Function ValidateGST1()
ValidateGST1 = True
If Not ValidateGSTIN Then ValidateGST1 = False
End Function
 
 Function ValidateGSTIN() As Boolean
 
    ValidateGSTIN = True
    setTblinfo_GSTIN
    setTblinfo_GSTAMOUNT
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim i As Long
    Dim cellrange1 As String
    Set rangecells1 = Sheet51.Range("GST_GSTRno").Cells
    Set rangecells = Sheet51.Range("GST_GSTRAmount").Cells
    
    end_GSTIN = WorksheetFunction.Max(0, end_GSTIN, end_GSTAMOUNT)
    ReDim GSTINAMOUNT_BP(end_GSTIN)
    ReDim GSTIN_BP(end_GSTIN)
    For i = 1 To end_GSTIN
    
        cellrange1 = GetMergedAddressCell(rangecells1, i)
        GSTIN_BP(i) = Sheet51.Range(cellrange1).value
         GSTINAMOUNT_BP(i) = rangecells.item(i).value
       
        If Not Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Then
            If Not CheckGSTR(GSTIN_BP(i)) Then
            MsgBox_GST = MsgBox_GST + "* GSTIN format should be First 2 numeric, next 5 Alphabets, next 4 numeric , next 1 Alphabet, then next 3 alphanumeric.at Sr. No  " & i & " in schedule GST " & Chr(13)
                ValidateGSTIN = False
                Exit Function
            End If
        Else
            If Not CheckSpecialCharacter(GSTIN_BP(i)) Then
            MsgBox_GST = MsgBox_GST + "* Invaild GSTIN. GSTIN should be in alphanumeric digits at Sr. No  " & i & " in schedule GST " & Chr(13)
                ValidateGSTIN = False
                Exit Function
            End If
        End If
        
        If Not IsNumeric(GSTINAMOUNT_BP(i)) Then
        MsgBox_GST = MsgBox_GST + "*  Please enter the Amount of turnover/Gross receipt for the year corresponding to the GSTIN at Sr. No  " & i & "in schedule GST" & Chr(13)
            ValidateGSTIN = False
            Exit Function
        End If
        
        If end_GSTIN > 0 Then
         If Not chkCompulsory(GSTIN_BP(i)) Then
             MsgBox_GST = MsgBox_GST + "* Please enter the GSTIN  at Sr. No  " & i & " in schedule GST" & Chr(13)
            ValidateGSTIN = False
            Exit Function
        End If
        
        If Not chkCompulsory(GSTINAMOUNT_BP(i)) Then
             MsgBox_GST = MsgBox_GST + "* Please enter the Annual Value of Outward Supplies as per the GST return filed at Sr. No  " & i & "in schedule GST" & Chr(13)
            ValidateGSTIN = False
            Exit Function
        End If
        End If
        
        
        
     Next
End Function

Function CheckGSTR(GSTRNo As Variant) As Boolean
On Error Resume Next
'GSTR No : Consist of 10 characters
'GSTR No: First 2 numeric, 5 Alphabets, next 4 numeric , next 1 Alphabet, then next 3 alphanumeric.

    CheckGSTR = True
    If Len(GSTRNo) > 0 Then
        If Not IsNumeric(Mid(GSTRNo, 1, 1)) Then
            CheckGSTR = False
            Exit Function
        End If
        If Not IsNumeric(Mid(GSTRNo, 2, 1)) Then
            CheckGSTR = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(GSTRNo, 3, 1)) Then
            CheckGSTR = False
            Exit Function
        End If
                
       If Not ChkAlphabet(Mid(GSTRNo, 4, 1)) Then
            CheckGSTR = False
            Exit Function
        End If
        
        If Not ChkAlphabet(Mid(GSTRNo, 5, 1)) Then
            CheckGSTR = False
            Exit Function
        End If
        
        If Not ChkAlphabet(Mid(GSTRNo, 6, 1)) Then
            CheckGSTR = False
            Exit Function
        End If
        
        If Not ChkAlphabet(Mid(GSTRNo, 7, 1)) Then
            CheckGSTR = False
            Exit Function
        End If
         
        If Not IsNumeric(Mid(GSTRNo, 8, 4)) Then
            CheckGSTR = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(GSTRNo, 12, 1)) Then
            CheckGSTR = False
            Exit Function
        End If
        If (Not ChkAlphabet(Mid(GSTRNo, 13, 1))) Then
        If (Not IsNumeric(Mid(GSTRNo, 13, 1))) Then
            CheckGSTR = False
            Exit Function
        End If
        End If
        
        
        If (Not ChkAlphabet(Mid(GSTRNo, 14, 1))) Then
        If (Not IsNumeric(Mid(GSTRNo, 14, 1))) Then
            CheckGSTR = False
            Exit Function
        End If
        End If
        
        If (Not ChkAlphabet(Mid(GSTRNo, 15, 1))) Then
        If (Not IsNumeric(Mid(GSTRNo, 15, 1))) Then
            CheckGSTR = False
            Exit Function
        End If
        End If
       
    End If
End Function
