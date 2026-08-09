Attribute VB_Name = "SchRA"


Option Explicit

Public MsgBox_80GGA, MsgBox_80GGANew As String

Public RelevantClauseClaimed_80GGA, RelevantClauseClaimed_80GGANew As Variant
Public Name_of_Donee_80GGA, Name_of_Donee_80GGANew As Variant
Public Address_80GGA, Address_80GGANew As Variant
Public City_Town_District_80GGA, City_Town_District_80GGANew As Variant
Public State_Code_80GGA, State_Code_80GGANew As Variant
Public Pincode_80GGA, Pincode_80GGANew As Variant
Public PAN_of_donee_80GGA, PAN_of_donee_80GGANew As Variant
Public Donation_cash_80GGA, Donation_cash_80GGANew As Variant
Public Donation_other_80GGA, Donation_other_80GGANew As Variant
Public Donation_total_80GGA, Donation_total_80GGANew As Variant
Public Donation_Eligible_80GGA, Donation_Eligible_80GGANew As Variant

Public rngname_80GGA, rngname_80GGANew As Variant
Public Total_Donation_Eligible_80GGA, Total_Donation_Eligible_80GGANew, Total_Donation_Total_80GGANew, Total_Donation_OtherMode_80GGANew, Total_Donation_Cash_80GGANew As Variant

Public end80GGA, end80GGA1, end80GGA2, end80GGA3, end80GGA4, end80GGA5, end80GGA6, end80GGA7, end80GGA8, end80GGA9 As Long
Public end80GGAN, end80GGAN1, end80GGAN2, end80GGAN3, end80GGAN4, end80GGAN5, end80GGAN6, end80GGAN7, end80GGAN8, end80GGAN9 As Long

Sub ValidateSheet80GGANew_Click()
Validate80GGANew
fmsgboxStatus "Sheet 80GGA is ok" ', vbOKOnly, "ITR-3"
End Sub

Sub ValidateSheet80GGA_Click()
Validate80GGA
fmsgboxStatus "Sheet RA is ok" ', vbOKOnly, "ITR-3"
End Sub

Sub AddRows80GGA_Click()
Dim vRows  As Long
Sheets("RA").Activate
EfilingCommon.DefinedgridNameRange = "Name_of_Donee_80GGA||Address_80GGA||City_Town_District_80GGA||State_Code_80GGA||Pincode_80GGA||PAN_of_donee_80GGA||Donation_cash_80GGA||Donation_other_80GGA||Donation_total_80GGA||Donation_Eligible_80GGA"
ActiveCellRange = EfilingCommon.searchLastRow("Name_of_Donee_80GGA")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub AddRows80GGANew_Click()
Dim vRows  As Long
Sheets("80GGA").Activate
EfilingCommon.DefinedgridNameRange = "RelevantClauseClaimed_80GGANew||Name_of_Donee_80GGANew||Address_80GGANew||City_Town_District_80GGANew||State_Code_80GGANew||Pincode_80GGANew||PAN_of_donee_80GGANew||Donation_cash_80GGANew||Donation_other_80GGANew||Donation_total_80GGANew||Donation_Eligible_80GGANew"
ActiveCellRange = EfilingCommon.searchLastRow("Name_of_Donee_80GGANew")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub Validate80GGANew()

    If Not Validate80GGA_1New Then
        Sheet56.Activate
        fmsgboxsmall MsgBox_80GGANew ', vbOKOnly, "Error(s)"
        CloseMsg
    End If
End Sub


Sub Validate80GGA()
    If Not Validate80GGA_1 Then
        Sheet50.Activate
        fmsgboxsmall MsgBox_80GGA ', vbOKOnly, "Error(s)"
        CloseMsg
    End If
End Sub
Sub setTableInfo80GGANew()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet56.Range("RelevantClauseClaimed_80GGANew").Cells
    mIntCells = Sheet56.Range("RelevantClauseClaimed_80GGANew").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    end80GGAN = ccount
    rngname_80GGANew = "RelevantClauseClaimed_80GGANew;Name_of_Donee_80GGANew;Address_80GGANew;City_Town_District_80GGANew;State_Code_80GGANew;Pincode_80GGANew;PAN_of_donee_80GGANew;Donation_cash_80GGANew;Donation_other_80GGANew;Donation_total_80GGANew;Donation_Eligible_80GGANew;"
End Sub
Sub setTableInfo80GGA1()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet50.Range("Name_of_Donee_80GGA").Cells
    mIntCells = Sheet50.Range("Name_of_Donee_80GGA").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGA1 = ccount
    rngname_80GGA = "Name_of_Donee_80GGA;Address_80GGA;City_Town_District_80GGA;State_Code_80GGA;Pincode_80GGA;PAN_of_donee_80GGA;Donation_cash_80GGA;Donation_other_80GGA;Donation_total_80GGA;Donation_Eligible_80GGA;"
End Sub
Sub setTableInfo80GGA1New()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet56.Range("Name_of_Donee_80GGANew").Cells
    mIntCells = Sheet56.Range("Name_of_Donee_80GGANew").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGAN1 = ccount
End Sub


Sub setTableInfo80GGA2()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet50.Range("Address_80GGA").Cells
    mIntCells = Sheet50.Range("Address_80GGA").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGA2 = ccount
End Sub
Sub setTableInfo80GGA2New()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet56.Range("Address_80GGANew").Cells
    mIntCells = Sheet56.Range("Address_80GGANew").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGAN2 = ccount
End Sub
Sub setTableInfo80GGA3()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet50.Range("City_Town_District_80GGA").Cells
    mIntCells = Sheet50.Range("City_Town_District_80GGA").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGA3 = ccount
End Sub
Sub setTableInfo80GGA3New()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet56.Range("City_Town_District_80GGANew").Cells
    mIntCells = Sheet56.Range("City_Town_District_80GGANew").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGAN3 = ccount
End Sub

Sub setTableInfo80GGA4()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet50.Range("State_Code_80GGA").Cells
    mIntCells = Sheet50.Range("State_Code_80GGA").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    end80GGA4 = ccount
End Sub
Sub setTableInfo80GGA4New()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet56.Range("State_Code_80GGANew").Cells
    mIntCells = Sheet56.Range("State_Code_80GGANew").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    end80GGAN4 = ccount
End Sub
Sub setTableInfo80GGA5()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet50.Range("Pincode_80GGA").Cells
    mIntCells = Sheet50.Range("Pincode_80GGA").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGA5 = ccount
End Sub
Sub setTableInfo80GGA5New()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet56.Range("Pincode_80GGANew").Cells
    mIntCells = Sheet56.Range("Pincode_80GGANew").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGAN5 = ccount
End Sub
Sub setTableInfo80GGA6()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet50.Range("PAN_of_donee_80GGA").Cells
    mIntCells = Sheet50.Range("PAN_of_donee_80GGA").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGA6 = ccount
End Sub
Sub setTableInfo80GGA6New()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet56.Range("PAN_of_donee_80GGANew").Cells
    mIntCells = Sheet56.Range("PAN_of_donee_80GGANew").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGAN6 = ccount
End Sub
Sub setTableInfo80GGA7()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet50.Range("Donation_cash_80GGA").Cells
    mIntCells = Sheet50.Range("Donation_cash_80GGA").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGA7 = ccount
End Sub
Sub setTableInfo80GGA7New()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet56.Range("Donation_cash_80GGANew").Cells
    mIntCells = Sheet56.Range("Donation_cash_80GGANew").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGAN7 = ccount
End Sub
Sub setTableInfo80GGA8()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet50.Range("Donation_other_80GGA").Cells
    mIntCells = Sheet50.Range("Donation_other_80GGA").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGA8 = ccount
End Sub
Sub setTableInfo80GGA8New()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet56.Range("Donation_other_80GGANew").Cells
    mIntCells = Sheet56.Range("Donation_other_80GGANew").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGAN8 = ccount
End Sub
Sub setTableInfo80GGA9()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet50.Range("Donation_Eligible_80GGA").Cells
    mIntCells = Sheet50.Range("Donation_Eligible_80GGA").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGA9 = ccount
End Sub
Sub setTableInfo80GGA9New()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet56.Range("Donation_Eligible_80GGANew").Cells
    mIntCells = Sheet56.Range("Donation_Eligible_80GGANew").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGAN9 = ccount
End Sub
Function ValidateRelevantClauseClaimed_80GGANew() As Boolean
    ValidateRelevantClauseClaimed_80GGANew = True
    setTableInfo80GGANew
    Dim rangecells As Range
    Set rangecells = Sheet56.Range("RelevantClauseClaimed_80GGANew").Cells
    Dim i As Long
    ReDim RelevantClauseClaimed_80GGANew(end80GGAN)
    
    subProcCaption = "Validating 80GGA"
    noOfProcessSub = end80GGAN
    
    For i = 1 To end80GGAN
        RelevantClauseClaimed_80GGANew(i) = rangecells.item(i).value
        
'        If isdropdownblank(RelevantClauseClaimed_80GGANew(i)) Then
'
'
'            MsgBox_80GGANew = MsgBox_80GGANew + "* Please select relevant clause at Sr. No." & i & " from Drop-Down in Schedule 80GGA" & Chr(13)
'            ValidateRelevantClauseClaimed_80GGANew = False
'            Exit Function
'        End If
        UpdateProgressBar
    Next
End Function
Function ValidateName_of_Donee_80GGA() As Boolean
    ValidateName_of_Donee_80GGA = True
   ' setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet50.Range("Name_of_Donee_80GGA").Cells
    Dim i As Long
    ReDim Name_of_Donee_80GGA(end80GGA1)

    For i = 1 To end80GGA1
        Name_of_Donee_80GGA(i) = rangecells.item(i).value

        If Not chkMandatory(Name_of_Donee_80GGA(i)) Then
            MsgBox_80GGA = MsgBox_80GGA + "* Please enter name of donee at Sr. No " & i & "in Schedule RA" & Chr(13)
            ValidateName_of_Donee_80GGA = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function

Function ValidateName_of_Donee_80GGANew() As Boolean
    ValidateName_of_Donee_80GGANew = True
    setTableInfo80GGANew
    Dim rangecells As Range
    Set rangecells = Sheet56.Range("Name_of_Donee_80GGANew").Cells
    Dim i As Long
    ReDim Name_of_Donee_80GGANew(end80GGAN)

    For i = 1 To end80GGAN
        Name_of_Donee_80GGANew(i) = rangecells.item(i).value

'        If Not chkMandatory(Name_of_Donee_80GGANew(i)) Then
'
'            MsgBox_80GGANew = MsgBox_80GGANew + "* Please enter name of donee at Sr. No " & i & "in Schedule 80GGA" & Chr(13)
'            ValidateName_of_Donee_80GGANew = False
'            Exit Function
'        End If
        UpdateProgressBar
    Next
End Function

Function ValidateAddress_80GGA() As Boolean
    ValidateAddress_80GGA = True
    'setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet50.Range("Address_80GGA").Cells
    Dim i As Long
    ReDim Address_80GGA(end80GGA1)

    For i = 1 To end80GGA1
        Address_80GGA(i) = rangecells.item(i).value

        If Not chkMandatory(Address_80GGA(i)) Then
            MsgBox_80GGA = MsgBox_80GGA + "* Please enter address of donee at Sr. No " & i & "in Schedule RA" & Chr(13)
            ValidateAddress_80GGA = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function
Function ValidateAddress_80GGANew() As Boolean
    ValidateAddress_80GGANew = True
    setTableInfo80GGANew
    Dim rangecells As Range
    Set rangecells = Sheet56.Range("Address_80GGANew").Cells
    Dim i As Long
    ReDim Address_80GGANew(end80GGAN)

    For i = 1 To end80GGAN
        Address_80GGANew(i) = rangecells.item(i).value

'        If Not chkMandatory(Address_80GGANew(i)) Then
'            MsgBox_80GGANew = MsgBox_80GGANew + "* Please enter address of donee at Sr. No " & i & "in Schedule 80GGA" & Chr(13)
'            ValidateAddress_80GGANew = False
'            Exit Function
'        End If
        UpdateProgressBar
    Next
End Function
Function ValidateCity_Town_District_80GGA() As Boolean
    ValidateCity_Town_District_80GGA = True
    'setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet50.Range("City_Town_District_80GGA").Cells
    Dim i As Long
    ReDim City_Town_District_80GGA(end80GGA1)

    For i = 1 To end80GGA1
        City_Town_District_80GGA(i) = rangecells.item(i).value

        If Not chkMandatory(City_Town_District_80GGA(i)) Then
            MsgBox_80GGA = MsgBox_80GGA + "* Please enter city/town/district of donee at Sr. No " & i & "in Schedule RA" & Chr(13)
            ValidateCity_Town_District_80GGA = False
            Exit Function
        End If
        
        If Len(City_Town_District_80GGA(i)) > 50 Then
            MsgBox_80GGA = MsgBox_80GGA + "* City/town/district of donee  cannot be more than 50 characters at Sr. No " & i & "in Schedule RA" & Chr(13)
            ValidateCity_Town_District_80GGA = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function
Function ValidateCity_Town_District_80GGANew() As Boolean
    ValidateCity_Town_District_80GGANew = True
    setTableInfo80GGANew
    Dim rangecells As Range
    Set rangecells = Sheet56.Range("City_Town_District_80GGANew").Cells
    Dim i As Long
    ReDim City_Town_District_80GGANew(end80GGAN)

    For i = 1 To end80GGAN
        City_Town_District_80GGANew(i) = rangecells.item(i).value
'
'        If Not chkMandatory(City_Town_District_80GGANew(i)) Then
'            MsgBox_80GGANew = MsgBox_80GGANew + "* Please enter city/town/district of donee at Sr. No " & i & "in Schedule 80GGA" & Chr(13)
'            ValidateCity_Town_District_80GGANew = False
'            Exit Function
'        End If
        
        If Len(City_Town_District_80GGANew(i)) > 50 Then
            MsgBox_80GGANew = MsgBox_80GGANew + "* City/town/district of donee  cannot be more than 50 characters at Sr. No " & i & "in Schedule 80GGA" & Chr(13)
            ValidateCity_Town_District_80GGANew = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function

Function ValidateState_Code_80GGA() As Boolean
ValidateState_Code_80GGA = True
    'setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet50.Range("State_Code_80GGA").Cells
    Dim i As Long
    ReDim State_Code_80GGA(end80GGA1)
    For i = 1 To end80GGA1
        State_Code_80GGA(i) = rangecells.item(i).value
        
        
    'Changed By sai on 16/10/2024==================================================================
       If State_Code_80GGA(i) <> "" Then
       Sheets("RA").Activate
         Dim PIN_targetadd, state_targetadd As String
         
        Dim ws_1 As Worksheet
        Set ws_1 = Worksheets("RA")
         
      'state_targetadd = Target.address
         state_targetadd = Replace(rangecells.item(i).Address, "$", "")
         PIN_targetadd = Replace(state_targetadd, "H", "I")
      
            If Not state_Validation_Level(ws_1, PIN_targetadd, state_targetadd) Then
                     
                     Sheet50.Range(Replace(state_targetadd, "H", "I")).value = ""
                     
                   ValidateState_Code_80GGA = False
                 End If
            End If
        
    '===============================================================================================
        
        

        If ((State_Code_80GGA(i) = "(Select)") Or (State_Code_80GGA(i) = "")) Then
          MsgBox_80GGA = MsgBox_80GGA + "* Please select state of donee at Sr. No " & i & "in Schedule RA" & Chr(13)
          ValidateState_Code_80GGA = False
          Exit Function
        End If
    Next
End Function
Function ValidateState_Code_80GGANew() As Boolean
ValidateState_Code_80GGANew = True
    setTableInfo80GGANew
    Dim rangecells As Range
    Set rangecells = Sheet56.Range("State_Code_80GGANew").Cells
    Dim i As Long
    ReDim State_Code_80GGANew(end80GGAN)
    For i = 1 To end80GGAN
        State_Code_80GGANew(i) = rangecells.item(i).value

'        If ((State_Code_80GGANew(i) = "(Select)") Or (State_Code_80GGANew(i) = "")) Then
'          MsgBox_80GGANew = MsgBox_80GGANew + "* Please select state of donee at Sr. No " & i & "in Schedule 80GGA" & Chr(13)
'          ValidateState_Code_80GGANew = False
'          Exit Function
'        End If

    'Changed By sai on 16/10/2024==================================================================
       If State_Code_80GGANew(i) <> "" Then
       Sheets("80GGA").Activate
         Dim PIN_targetadd, state_targetadd As String
        Dim ws_2 As Worksheet
        Set ws_2 = Worksheets("80GGA")
         
      'state_targetadd = Target.address
     state_targetadd = Replace(rangecells.item(i).Address, "$", "")
      PIN_targetadd = Replace(state_targetadd, "I", "J")

            If Not state_Validation_Level(ws_2, PIN_targetadd, state_targetadd) Then
                     
                     Sheet56.Range(Replace(state_targetadd, "I", "J")).value = ""
                     
                   ValidateState_Code_80GGANew = False
                 End If
            End If

    '========================================================================================




    Next
End Function

Function ValidatePincode_80GGA() As Boolean
ValidatePincode_80GGA = True
   ' setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet50.Range("Pincode_80GGA").Cells
    Dim i, x  As Long
    ReDim Pincode_80GGA(end80GGA1)
    For i = 1 To end80GGA1
        Pincode_80GGA(i) = rangecells.item(i).value

        If Not chkMandatory(Pincode_80GGA(i)) Then
            MsgBox_80GGA = MsgBox_80GGA + "* Please enter pin code of donee at Sr. No " & i & "in Schedule RA" & Chr(13)
            ValidatePincode_80GGA = False
            Exit Function
        End If

        If Mid(Pincode_80GGA(i), 1, 1) = 0 Then
            MsgBox_80GGA = MsgBox_80GGA + "* Pin code of donee at Sr. No " & i & " must be 6 digits and cannot begin with '0'." & "in Schedule RA" & Chr(13)
            ValidatePincode_80GGA = False
            Exit Function
        End If

        For x = 1 To Len(Pincode_80GGA(i))
            If Not IsNumeric(Mid(Pincode_80GGA(i), x, 1)) Then
                  MsgBox_80GGA = MsgBox_80GGA + "* Pin code of donee at Sr. No " & i & " must contain only digits from 0 to 9." & "in Schedule RA" & Chr(13)
                ValidatePincode_80GGA = False
                Exit Function
            End If
        Next

        If Len(Pincode_80GGA(i)) > 6 Then
            MsgBox_80GGA = MsgBox_80GGA + "* Pin code of donee at Sr. No " & i & " must be 6 digits." & "in Schedule RA" & Chr(13)
            ValidatePincode_80GGA = False
            Exit Function
        End If

        If Not checkfieldspecialcharacter(Pincode_80GGA(i)) Then
             MsgBox_80GGA = MsgBox_80GGA + "* Pin code of donee at Sr. No  " & i & "cannot Contain Special Characters." & "in Schedule RA" & Chr(13)
             ValidatePincode_80GGA = False
             Exit Function
        End If
    Next
End Function
Function ValidatePincode_80GGANew() As Boolean
ValidatePincode_80GGANew = True
    setTableInfo80GGANew
    Dim rangecells As Range
    Set rangecells = Sheet56.Range("Pincode_80GGANew").Cells
    Dim i, x  As Long
    ReDim Pincode_80GGANew(end80GGAN)
    For i = 1 To end80GGAN
        Pincode_80GGANew(i) = rangecells.item(i).value
'
'        If Not chkMandatory(Pincode_80GGANew(i)) Then
'            MsgBox_80GGANew = MsgBox_80GGANew + "* Please enter pin code of donee at Sr. No " & i & "in Schedule 80GGA" & Chr(13)
'            ValidatePincode_80GGANew = False
'            Exit Function
'        End If

        If Mid(Pincode_80GGANew(i), 1, 1) = 0 Then
            MsgBox_80GGANew = MsgBox_80GGANew + "* Pin code of donee at Sr. No " & i & " must be 6 digits and cannot begin with '0'." & "in Schedule 80GGA" & Chr(13)
            ValidatePincode_80GGANew = False
            Exit Function
        End If

        For x = 1 To Len(Pincode_80GGANew(i))
            If Not IsNumeric(Mid(Pincode_80GGANew(i), x, 1)) Then
'                  MsgBox_80GGANew = MsgBox_80GGANew + "* Pin code of donee at Sr. No " & i & " must contain only digits from 0 to 9." & "in Schedule 80GGA" & Chr(13)
            'SIT-69327 changed by sai
                  MsgBox_80GGANew = MsgBox_80GGANew + "* ""Invalid Pin code"" at Sr. No " & i & " in Schedule 80GGA" & Chr(13)
                  
                ValidatePincode_80GGANew = False
                Exit Function
            End If
        Next

        If Len(Pincode_80GGANew(i)) > 6 Then
            MsgBox_80GGANew = MsgBox_80GGANew + "* Pin code of donee at Sr. No " & i & " must be 6 digits." & "in Schedule 80GGA" & Chr(13)
            ValidatePincode_80GGANew = False
            Exit Function
        End If

        If Not checkfieldspecialcharacter(Pincode_80GGANew(i)) Then
             MsgBox_80GGANew = MsgBox_80GGANew + "* Pin code of donee at Sr. No  " & i & "cannot Contain Special Characters." & "in Schedule 80GGA" & Chr(13)
             ValidatePincode_80GGANew = False
             Exit Function
        End If
    Next
End Function
Function ValidatePAN_of_donee_80GGA() As Boolean
ValidatePAN_of_donee_80GGA = True
    'setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet50.Range("PAN_of_donee_80GGA").Cells
    Dim i As Long
    ReDim PAN_of_donee_80GGA(end80GGA1)
    For i = 1 To end80GGA1
    PAN_of_donee_80GGA(i) = UCase(rangecells.item(i).value)
        If Not chkMandatory(PAN_of_donee_80GGA(i)) Then
            MsgBox_80GGA = MsgBox_80GGA + "* Please enter PAN of donee at Sr. No " & i & "in Schedule RA" & Chr(13)
            ValidatePAN_of_donee_80GGA = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter(PAN_of_donee_80GGA(i)) Then
             MsgBox_80GGA = MsgBox_80GGA + "* PAN of donee at Sr. No  " & i & " cannot Contain Special Characters." & "in Schedule RA" & Chr(13)
             ValidatePAN_of_donee_80GGA = False
             Exit Function
        End If

        If Not CheckDoneePAN(UCase(PAN_of_donee_80GGA(i))) Then
             MsgBox_80GGA = MsgBox_80GGA + "* PAN of donee at Sr. No  " & i & " is Invalid. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet." & "in Schedule RA" & Chr(13)
             ValidatePAN_of_donee_80GGA = False
             Exit Function
        End If

        If ((UCase(PAN_of_donee_80GGA(i)) = UCase(Sheet1.Range("sheet1.PAN").value)) Or (UCase(PAN_of_donee_80GGA(i)) = UCase(Sheet52.Range("sheet9.PAN").value))) Then
            MsgBox_80GGA = MsgBox_80GGA + "* PAN of donee at Sr. No  " & i & " is Invalid." & Chr(13) & " Donee PAN cannot be same as assesse PAN or verification PAN  " & "in Schedule RA" & Chr(13)
            ValidatePAN_of_donee_80GGA = False
            Exit Function
        End If
    Next
End Function
Function ValidatePAN_of_donee_80GGANew() As Boolean
ValidatePAN_of_donee_80GGANew = True
    setTableInfo80GGANew
    Dim rangecells As Range
    Set rangecells = Sheet56.Range("PAN_of_donee_80GGANew").Cells
    Dim i As Long
    ReDim PAN_of_donee_80GGANew(end80GGAN)
    For i = 1 To end80GGAN
    PAN_of_donee_80GGANew(i) = UCase(rangecells.item(i).value)
'        If Not chkMandatory(PAN_of_donee_80GGANew(i)) Then
'        'SIT-69327 changed by sai
'            MsgBox_80GGANew = MsgBox_80GGANew + "* ""Please enter PAN of donee."" at Sr. No " & i & "" & Chr(13)
'            ValidatePAN_of_donee_80GGANew = False
'            Exit Function
'        End If
        If Not checkfieldspecialcharacter(PAN_of_donee_80GGANew(i)) Then
             MsgBox_80GGANew = MsgBox_80GGANew + "* PAN of donee at Sr. No  " & i & " cannot Contain Special Characters." & "in Schedule 80GGA" & Chr(13)
             ValidatePAN_of_donee_80GGANew = False
             Exit Function
        End If

        If Not CheckDoneePAN(UCase(PAN_of_donee_80GGANew(i))) Then
'             MsgBox_80GGANew = MsgBox_80GGANew + "*Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet at Sr. No " & i & " in Schedule 80GGA." & Chr(13)
        'SIT-69327 changed by sai
             MsgBox_80GGANew = MsgBox_80GGANew + "*Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet in Schedule 80GGA at Sr. No " & i & "" & Chr(13)
             
             
             ValidatePAN_of_donee_80GGANew = False
             Exit Function
        End If

        If ((UCase(PAN_of_donee_80GGANew(i)) = UCase(Sheet1.Range("sheet1.PAN").value)) Or (UCase(PAN_of_donee_80GGANew(i)) = UCase(Sheet52.Range("sheet9.PAN").value))) Then
            MsgBox_80GGANew = MsgBox_80GGANew + "* PAN of donee at Sr. No  " & i & " is Invalid." & Chr(13) & " Donee PAN cannot be same as assesse PAN or verification PAN  " & "in Schedule 80GGA" & Chr(13)
            ValidatePAN_of_donee_80GGANew = False
            Exit Function
        End If
    Next
End Function
Function ValidateDonationAmt_80GGA() As Boolean
ValidateDonationAmt_80GGA = True
    'setTableInfo80GGA
    Dim rangecells As Range
    Dim rangecells1 As Range
    Set rangecells = Sheet50.Range("Donation_cash_80GGA").Cells
    Set rangecells1 = Sheet50.Range("Donation_other_80GGA").Cells
    Dim i As Long
    ReDim Donation_cash_80GGA(end80GGA1)
    ReDim Donation_other_80GGA(end80GGA1)

    For i = 1 To end80GGA1
        Donation_cash_80GGA(i) = rangecells.item(i).value
        Donation_other_80GGA(i) = rangecells1.item(i).value

        If Not chkMandatory(Donation_cash_80GGA(i)) And Not chkMandatory(Donation_other_80GGA(i)) Then
            MsgBox_80GGA = MsgBox_80GGA + "* Donation in Cash or Donation Other mode at Sr. No " & i & " is Mandatory." & "in Schedule RA" & Chr(13)
            
            
'            MsgBox_80GGA = MsgBox_80GGA + "* Donation in Cash or Donation Other mode at Sr. No " & i & " is Mandatory." & "in Schedule RA" & Chr(13)
            
            ValidateDonationAmt_80GGA = False
            Exit Function
        End If
    Next
End Function
Function ValidateDonationAmt_80GGANew() As Boolean
ValidateDonationAmt_80GGANew = True
    setTableInfo80GGANew
    Dim rangecells As Range
    Dim rangecells1 As Range
    Set rangecells = Sheet56.Range("Donation_cash_80GGANew").Cells
    Set rangecells1 = Sheet56.Range("Donation_other_80GGANew").Cells
    Dim i As Long
    ReDim Donation_cash_80GGANew(end80GGAN)
    ReDim Donation_other_80GGANew(end80GGAN)

    For i = 1 To end80GGAN
        Donation_cash_80GGANew(i) = rangecells.item(i).value
        Donation_other_80GGANew(i) = rangecells1.item(i).value

'        If Not chkMandatory(Donation_cash_80GGANew(i)) And Not chkMandatory(Donation_other_80GGANew(i)) Then
''            MsgBox_80GGANew = MsgBox_80GGANew + "* Donation in Cash or Donation Other mode at Sr. No " & i & " is Mandatory." & "in Schedule 80GGA" & Chr(13)
'
''              MsgBox_80GGANew = MsgBox_80GGANew + "*Enter the amount of donation either in field ""Donation in cash"" or ""Donation in other mode at Sr. No " & i & " in schedule 80GGA"" " & Chr(13)
'
'            'SIT-69327 changed by sai
'              MsgBox_80GGANew = MsgBox_80GGANew + "*Enter the amount of donation either in field ""Donation in cash"" or ""Donation in other mode in schedule 80GGA"" at Sr. No " & i & " " & Chr(13)
'
'            ValidateDonationAmt_80GGANew = False
'            Exit Function
'        End If
    Next
End Function
Function ValidateDonation_total_80GGA() As Boolean
ValidateDonation_total_80GGA = True
    'setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet50.Range("Donation_total_80GGA").Cells
    Dim i As Long
    ReDim Donation_total_80GGA(end80GGA1)

    For i = 1 To end80GGA1
        Donation_total_80GGA(i) = rangecells.item(i).value
        If Len(Donation_total_80GGA(i)) > 14 Then
            MsgBox_80GGA = MsgBox_80GGA + "* Total donation Amount at Sr. No " & i & " cannot exceed 14 digits." & "in Schedule RA" & Chr(13)
            ValidateDonation_total_80GGA = False
            Exit Function
        End If
    Next
End Function
Function ValidateDonation_total_80GGANew() As Boolean
ValidateDonation_total_80GGANew = True
    setTableInfo80GGANew
    Dim rangecells As Range
    Set rangecells = Sheet56.Range("Donation_total_80GGANew").Cells
    Dim i As Long
    ReDim Donation_total_80GGANew(end80GGAN)

    For i = 1 To end80GGAN
        Donation_total_80GGANew(i) = rangecells.item(i).value
        If Len(Donation_total_80GGANew(i)) > 14 Then
            MsgBox_80GGANew = MsgBox_80GGANew + "* Total donation Amount at Sr. No " & i & " cannot exceed 14 digits." & "in Schedule 80GGA" & Chr(13)
            ValidateDonation_total_80GGANew = False
            Exit Function
        End If
    Next
End Function
'
Function ValidateDonation_Eligible_80GGA() As Boolean
ValidateDonation_Eligible_80GGA = True
    'setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet50.Range("Donation_Eligible_80GGA").Cells
    Dim i As Long
    ReDim Donation_Eligible_80GGA(end80GGA1)

    For i = 1 To end80GGA1
        Donation_Eligible_80GGA(i) = rangecells.item(i).value
        If Len(Donation_Eligible_80GGA(i)) > 14 Then
            MsgBox_80GGA = MsgBox_80GGA + "* Eligible donation Amount at Sr. No " & i & " cannot exceed 14 digits." & "in Schedule RA" & Chr(13)
            ValidateDonation_Eligible_80GGA = False
            Exit Function
        End If
        
        If Donation_total_80GGA(i) > 0 Then
        If Not chkMandatory(Donation_Eligible_80GGA(i)) Then
            MsgBox_80GGA = MsgBox_80GGA + "* Eligible donation Amount at Sr. No " & i & " is mandatory" & "in Schedule RA" & Chr(13)
            ValidateDonation_Eligible_80GGA = False
            Exit Function
        End If
        End If
    Next
End Function
Function ValidateDonation_Eligible_80GGANew() As Boolean
ValidateDonation_Eligible_80GGANew = True
    setTableInfo80GGANew
    Dim rangecells As Range
    Set rangecells = Sheet56.Range("Donation_Eligible_80GGANew").Cells
    Dim i As Long
    ReDim Donation_Eligible_80GGANew(end80GGAN)

    For i = 1 To end80GGAN
        Donation_Eligible_80GGANew(i) = rangecells.item(i).value
        If Len(Donation_Eligible_80GGANew(i)) > 14 Then
            MsgBox_80GGANew = MsgBox_80GGANew + "* Eligible donation Amount at Sr. No " & i & " cannot exceed 14 digits." & "in Schedule 80GGA" & Chr(13)
            ValidateDonation_Eligible_80GGANew = False
            Exit Function
        End If
        
        If Donation_total_80GGANew(i) > 0 Then
        If Not chkMandatory(Donation_Eligible_80GGANew(i)) Then
            MsgBox_80GGANew = MsgBox_80GGANew + "* Eligible donation Amount at Sr. No " & i & " is mandatory" & "in Schedule 80GGA" & Chr(13)
            ValidateDonation_Eligible_80GGANew = False
            Exit Function
        End If
        End If
    Next
End Function
Function ValidateTotal_Donation_Eligible_80GGA() As Boolean
 ValidateTotal_Donation_Eligible_80GGA = True
 Total_Donation_Eligible_80GGA = Sheet50.Range("Total_Donation_Eligible_80GGA").value

 If Len(Total_Donation_Eligible_80GGA) > 14 Then
    MsgBox_80GGA = MsgBox_80GGA + "* Total donation amount cannot be greater than 14 digits." & "in Schedule RA" & Chr(13)
    ValidateTotal_Donation_Eligible_80GGA = False
    Exit Function
End If

End Function

Function ValidateTotal_Donation_Eligible_80GGANew() As Boolean
 ValidateTotal_Donation_Eligible_80GGANew = True
 Total_Donation_Eligible_80GGANew = Sheet56.Range("Total_Donation_Eligible_80GGANew").value

 If Len(Total_Donation_Eligible_80GGANew) > 14 Then
    MsgBox_80GGANew = MsgBox_80GGANew + "* Total donation amount cannot be greater than 14 digits." & "in Schedule 80GGA" & Chr(13)
    ValidateTotal_Donation_Eligible_80GGANew = False
    Exit Function
End If

End Function
Function ValidateTotal_Donation_Cash_80GGANew() As Boolean
 ValidateTotal_Donation_Cash_80GGANew = True
 Total_Donation_Cash_80GGANew = Sheet56.Range("Total_Donation_Cash_80GGANew").value

 If Len(Total_Donation_Cash_80GGANew) > 14 Then
    MsgBox_80GGANew = MsgBox_80GGANew + "* Total donation amount cannot be greater than 14 digits." & "in Schedule 80GGA" & Chr(13)
    ValidateTotal_Donation_Cash_80GGANew = False
    Exit Function
End If

End Function
Function ValidateTotal_Donation_OtherMode_80GGANew() As Boolean
 ValidateTotal_Donation_OtherMode_80GGANew = True
 Total_Donation_OtherMode_80GGANew = Sheet56.Range("Total_Donation_OtherMode_80GGANew").value

 If Len(Total_Donation_OtherMode_80GGANew) > 14 Then
    MsgBox_80GGANew = MsgBox_80GGANew + "* Total donation amount cannot be greater than 14 digits." & "in Schedule 80GGA" & Chr(13)
    ValidateTotal_Donation_OtherMode_80GGANew = False
    Exit Function
End If

End Function
Function ValidateTotal_Donation_Total_80GGANew() As Boolean
 ValidateTotal_Donation_Total_80GGANew = True
 Total_Donation_Total_80GGANew = Sheet56.Range("Total_Donation_Total_80GGANew").value

 If Len(Total_Donation_Total_80GGANew) > 14 Then
    MsgBox_80GGANew = MsgBox_80GGANew + "* Total donation amount cannot be greater than 14 digits." & "in Schedule 80GGA" & Chr(13)
    ValidateTotal_Donation_Total_80GGANew = False
    Exit Function
End If

End Function
Function Validate80GGA_1() As Boolean
    Validate80GGA_1 = True
    'setTableInfo80GGA
    setTableInfo80GGA1
    setTableInfo80GGA2
    setTableInfo80GGA3
    setTableInfo80GGA4
    setTableInfo80GGA5
    setTableInfo80GGA6
    setTableInfo80GGA7
    setTableInfo80GGA8
    setTableInfo80GGA9
    
    MsgBox_80GGA = "Schedule RA : " & Chr(10)

    end80GGA1 = WorksheetFunction.Max(0, end80GGA1, end80GGA2, end80GGA3, end80GGA4, end80GGA5, end80GGA6, end80GGA7, end80GGA8, end80GGA9)
    
'Change-08.05.2023.101.ESR -> RuntimeError
'    If Sheet12.Range("ESR1ii.ExcessAmtOverDebPL").value > 0 Or Sheet12.Range("ESR1iia.ExcessAmtOverDebPL").value Or Sheet12.Range("ESR1iii.ExcessAmtOverDebPL").value Or Sheet12.Range("ESR2AA.ExcessAmtOverDebPL").value Then
    If Sheet12.Range("ESR1ii.ExcessAmtOverDebPL").value > 0 Or Sheet12.Range("ESR1iia.ExcessAmtOverDebPL").value > 0 Or Sheet12.Range("ESR1iii.ExcessAmtOverDebPL").value > 0 Or Sheet12.Range("ESR2AA.ExcessAmtOverDebPL").value > 0 Then
    If Not end80GGA1 > 0 Then
    MsgBox_80GGA = MsgBox_80GGA + "* Schedule RA is mandatory as there is entry in Schedule ESR[deduction under sections 35(1)(ii) or 35(1)(iia) or 35(1)(iii) or 35(2AA)]" & Chr(13)
        Validate80GGA_1 = False
        Exit Function
    End If
    End If

    
    'If Not ValidateRelevantClauseClaimed_80GGA Then Validate80GGA_1 = False
    If Not ValidateName_of_Donee_80GGA Then Validate80GGA_1 = False
    If Not ValidateAddress_80GGA Then Validate80GGA_1 = False
    If Not ValidateCity_Town_District_80GGA Then Validate80GGA_1 = False
    If Not ValidateState_Code_80GGA Then Validate80GGA_1 = False
    If Not ValidatePincode_80GGA Then Validate80GGA_1 = False
    If Not ValidatePAN_of_donee_80GGA Then Validate80GGA_1 = False
    If Not ValidateDonationAmt_80GGA Then Validate80GGA_1 = False
    If Not ValidateDonation_total_80GGA Then Validate80GGA_1 = False
   If Not ValidateDonation_Eligible_80GGA Then Validate80GGA_1 = False
   If Not ValidateTotal_Donation_Eligible_80GGA Then Validate80GGA_1 = False


End Function

Function Validate80GGA_1New() As Boolean
    Validate80GGA_1New = True
    Dim i As Long
    
    setTableInfo80GGA1New
    setTableInfo80GGA2New
    setTableInfo80GGA3New
    setTableInfo80GGA4New
    setTableInfo80GGA5New
    setTableInfo80GGA6New
    setTableInfo80GGA7New
    setTableInfo80GGA8New
    setTableInfo80GGA9New
        
    MsgBox_80GGANew = "Schedule 80GGA : " & Chr(10)

    end80GGAN = WorksheetFunction.Max(end80GGAN, end80GGAN1, end80GGAN2, end80GGAN3, end80GGAN4, end80GGAN5, end80GGAN6, end80GGAN7, end80GGAN8, end80GGAN9)
   
    
         If Not ValidateRelevantClauseClaimed_80GGANew Then Validate80GGA_1New = False
         If Not ValidateName_of_Donee_80GGANew Then Validate80GGA_1New = False
         If Not ValidateAddress_80GGANew Then Validate80GGA_1New = False
         If Not ValidateCity_Town_District_80GGANew Then Validate80GGA_1New = False
         If Not ValidateState_Code_80GGANew Then Validate80GGA_1New = False
         If Not ValidatePincode_80GGANew Then Validate80GGA_1New = False
         If Not ValidatePAN_of_donee_80GGANew Then Validate80GGA_1New = False
         If Not ValidateDonationAmt_80GGANew Then Validate80GGA_1New = False
         If Not ValidateDonation_total_80GGANew Then Validate80GGA_1New = False
         If Not ValidateDonation_Eligible_80GGANew Then Validate80GGA_1New = False
         If Not ValidateTotal_Donation_Eligible_80GGANew Then Validate80GGA_1New = False
         If Not ValidateTotal_Donation_Cash_80GGANew Then Validate80GGA_1New = False
         If Not ValidateTotal_Donation_OtherMode_80GGANew Then Validate80GGA_1New = False
         If Not ValidateTotal_Donation_Total_80GGANew Then Validate80GGA_1New = False
    
    
    
'        If ((end80GGAN <> end80GGAN1) Or (end80GGAN <> end80GGAN2) Or (end80GGAN <> end80GGAN3) Or _
'            (end80GGAN <> end80GGAN4) Or (end80GGAN <> end80GGAN5) Or (end80GGAN <> end80GGAN6)) Then
'        MsgBox_80GGANew = MsgBox_80GGANew + "Enter All Mandatory Fields."
'        Validate80GGA_1New = False
'        End If


   
For i = 1 To Sheet56.Range("RelevantClauseClaimed_80GGANew").count
'SIT 90918
'    If Sheet56.Range("RelevantClauseClaimed_80GGANew").Cells(i, 1).value <> "(Select)" Or Sheet56.Range("Name_of_Donee_80GGANew").Cells(i, 1).value <> "" Or _
'        (Sheet56.Range("Address_80GGANew").Cells(i, 1).value <> "" And Sheet56.Range("City_Town_District_80GGANew").Cells(i, 1).value <> "") Or _
'        Sheet56.Range("State_Code_80GGANew").Cells(i, 1).value <> "(Select)" Or Sheet56.Range("Pincode_80GGANew").Cells(i, 1).value <> "" Or Sheet56.Range("PAN_of_donee_80GGANew").Cells(i, 1).value <> "" Or Sheet56.Range("Donation_cash_80GGANew").Cells(i, 1).value <> "" Or Sheet56.Range("Donation_other_80GGANew").Cells(i, 1).value <> "" Then

    If ((Sheet56.Range("RelevantClauseClaimed_80GGANew").Cells(i, 1).value <> "(Select)" And Sheet56.Range("RelevantClauseClaimed_80GGANew").Cells(i, 1).value <> "") Or _
         Sheet56.Range("Name_of_Donee_80GGANew").Cells(i, 1).value <> "") Or _
        (Sheet56.Range("Address_80GGANew").Cells(i, 1).value <> "" And Sheet56.Range("City_Town_District_80GGANew").Cells(i, 1).value <> "") Or _
        (Sheet56.Range("State_Code_80GGANew").Cells(i, 1).value <> "" And Sheet56.Range("State_Code_80GGANew").Cells(i, 1).value <> "(Select)") Or _
        Sheet56.Range("Pincode_80GGANew").Cells(i, 1).value <> "" Or Sheet56.Range("PAN_of_donee_80GGANew").Cells(i, 1).value <> "" Or Sheet56.Range("Donation_cash_80GGANew").Cells(i, 1).value <> "" Or Sheet56.Range("Donation_other_80GGANew").Cells(i, 1).value <> "" Then
        '--
        If Sheet56.Range("RelevantClauseClaimed_80GGANew").Cells(i, 1).value = "" Or Sheet56.Range("RelevantClauseClaimed_80GGANew").Cells(i, 1).value = "(Select)" Then
            Validate80GGA_1New = False
        MsgBox_80GGANew = MsgBox_80GGANew + "* Please select relevant clause from Drop-Down in Schedule 80GGA at Sr. No." & i & "" & Chr(13)
        End If
        
        If Sheet56.Range("Name_of_Donee_80GGANew").Cells(i, 1).value = "" Then
            Validate80GGA_1New = False
        MsgBox_80GGANew = MsgBox_80GGANew + "* Please enter name of donee in schedule 80GGA at Sr. No." & i & "." & Chr(13)
        End If

        If Sheet56.Range("Address_80GGANew").Cells(i, 1).value = "" Then
            Validate80GGA_1New = False
        MsgBox_80GGANew = MsgBox_80GGANew + "* Please enter address of donee in schedule 80GGA at Sr. No." & i & "." & Chr(13)

        End If

        If Sheet56.Range("City_Town_District_80GGANew").Cells(i, 1).value = "" Then
            Validate80GGA_1New = False
        MsgBox_80GGANew = MsgBox_80GGANew + "* Please enter city/town/district of donee in schedule 80GGA at Sr. No." & i & "" & Chr(13)
        End If

        If Sheet56.Range("State_Code_80GGANew").Cells(i, 1).value = "" Or Sheet56.Range("State_Code_80GGANew").Cells(i, 1).value = "(Select)" Then
            Validate80GGA_1New = False
        MsgBox_80GGANew = MsgBox_80GGANew + "* Please enter State of donee in schedule 80GGA at Sr. No." & i & "" & Chr(13)
        End If

        If Sheet56.Range("Pincode_80GGANew").Cells(i, 1).value = "" Then
            Validate80GGA_1New = False
        MsgBox_80GGANew = MsgBox_80GGANew + "* Please enter pin code of donee in schedule 80GGA at Sr. No." & i & "." & Chr(13)
        End If
        
        If Sheet56.Range("PAN_of_donee_80GGANew").Cells(i, 1).value = "" Then
            Validate80GGA_1New = False
        MsgBox_80GGANew = MsgBox_80GGANew + "* Please enter PAN of donee at Sr. No." & i & "." & Chr(13)

        End If
        
'Added by sudha as per SIT-91477

        If Sheet56.Range("Donation_cash_80GGANew").Cells(i, 1).value = "" Then
            'Validate80GGA_1New = False
            'MsgBox_80GGANew = MsgBox_80GGANew + "* Enter the amount of donation either in field ""Donation in cash"" or ""Donation in other mode"" in schedule 80GGA at Sr. No." & i & "" & Chr(13)
        'End If
        
        If Sheet56.Range("Donation_other_80GGANew").Cells(i, 1).value = "" Then
            Validate80GGA_1New = False
            MsgBox_80GGANew = MsgBox_80GGANew + "* Enter the amount of donation either in field ""Donation in cash"" or ""Donation in other mode"" at Sr. No." & i & "" & Chr(13)
        End If
End If

    End If
Next i


End Function




