Attribute VB_Name = "UserForm1"
Attribute VB_Base = "0{68659549-277A-473B-97EC-945D39821F29}{4343D60B-AD10-4590-9FB1-A20693A0CF54}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False



Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = vbFormControlMenu Then
    Cancel = True
End If
End Sub
