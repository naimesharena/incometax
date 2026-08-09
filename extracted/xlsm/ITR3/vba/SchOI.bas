Attribute VB_Name = "SchOI"
Option Explicit

Dim MsgBx_MOF1 As String, MsgBx_MOF2 As String, MsgBx_Item4 As String

Sub CmdValidate_OI_Click()

    Validate_SchOI_All

fmsgboxStatus "Sheet Other Information (OI) is ok"
End Sub

Sub Validate_SchOI_All()

If Not ValidateItem4All Then
Sheet5.Activate
    fmsgboxsmall (MsgBx_Item4) ', vbOKOnly, "Error(s)!"
    CloseMsg
End If



End Sub

Function ValidateMetofAccounting() As Boolean
ValidateMetofAccounting = True
MsgBx_MOF1 = ""
If Sheet5.Range("sheet5.MethodOfAcct").value = "None" Or Sheet5.Range("sheet5.MethodOfAcct").value = "" Then
    ValidateMetofAccounting = False
     Sheet5.Range("sheet5.MethodOfAcct").Select
     MsgBx_MOF1 = MsgBx_MOF1 + "* Please select Method of accounting employed in the previous year" & Chr(13)
End If
End Function



Function ValidateChangeinMOA_2() As Boolean
ValidateChangeinMOA_2 = True
MsgBx_MOF2 = ""
If Sheet5.Range("sheet5.ChangeInAcctMethFlg").value = "(Select)" Or Sheet5.Range("sheet5.ChangeInAcctMethFlg").value = "" Then
    ValidateChangeinMOA_2 = False
     Sheet5.Range("sheet5.ChangeInAcctMethFlg").Select
     MsgBx_MOF2 = MsgBx_MOF2 + "* Please select whether change in method of accounting." & Chr(13)
End If
End Function


Function ValidateItem4All() As Boolean
ValidateItem4All = True
MsgBx_Item4 = ""
    'If Not Validate4a Then ValidateItem4All = False
    'If Not Validate4b Then ValidateItem4All = False
    If Not Validate4c Then ValidateItem4All = False
End Function


Function Validate4a() As Boolean
Validate4a = True

End Function


Function Validate4b() As Boolean
Validate4b = True
If Sheet5.Range("sheet5.ValFinishedGoods").value = "(Select)" Or Sheet5.Range("sheet5.ValFinishedGoods").value = "" Then
    Validate4b = False
     Sheet5.Range("sheet5.ValFinishedGoods").Select
     MsgBx_Item4 = MsgBx_Item4 + "* Item 4b Please select method of Valuation of closing Stock  Part A - OI" & Chr(13)
End If
End Function


Function Validate4c() As Boolean
Validate4c = True
If Sheet5.Range("sheet6.ChngStockValMetFlg").value = "(Select)" Or Sheet5.Range("sheet6.ChngStockValMetFlg").value = "" Then
    Validate4c = False
     Sheet5.Range("sheet6.ChngStockValMetFlg").Select
     MsgBx_Item4 = MsgBx_Item4 + "* Item 4cPlease select whether change in Stock valuation method  Part A - OI" & Chr(13)
End If
End Function
