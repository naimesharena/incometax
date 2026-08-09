Attribute VB_Name = "UserForm3"
Attribute VB_Base = "0{CCF05AF3-1F63-47EB-9849-C74DA94BD76E}{45A1F906-37F5-4E54-BE78-571404EFA034}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub Frame1_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim tblRange As Range
    Dim r As Range
    Dim dataArr As Variant
    Dim i As Long, j As Long

leftpos = 0
toppos = 0
lastrow = Sheet62.Range("A" & Rows.count).End(xlUp).row
lastcol = 4

For tablerow = 1 To lastrow
        For tablEcol = 1 To lastcol
        
                    If tablerow = 1 Then
                    Set txt = UserForm3.Frame1.Controls.add("forms.label.1")
                    Else
                    Set txt = UserForm3.Frame1.Controls.add("forms.textbox.1")
                    End If
                    
                    
                    With txt
                    
                            If tablEcol = 1 Then .Width = 80
                            If tablEcol = 2 Then .Width = 350
                            If tablEcol = 3 Then .Width = 350
                            If tablEcol = 4 Then .Width = 80
                            .Top = toppos
                            .Left = leftpos
                            'leftpos = (.Left + 40) + .Width
                            leftpos = (.Left) + .Width
                            'leftpos = .Left + .Top
                            .BorderStyle = 1
                            
                            If tablerow = 1 Then
                                .caption = Sheet62.Cells(tablerow, tablEcol)
                                .BackColor = RGB(200, 200, 200)
                                .Height = 20
                                .TextAlign = 2
                            Else
                                .text = Sheet62.Cells(tablerow, tablEcol)
                                .Locked = True
                                .Multiline = True
                                .WordWrap = True
                                .ScrollBars = 2
                                .Height = 40
                            End If
                            'If tablEcol = 1 Then
                            '.AutoSize = True
                            'End If
                            
                            '        If tablEcol >= 3 Then
                            '        .Width = 200
                            '        End If
                            If tablerow = 1 Then .Font.Bold = True
                            If tablerow = 1 Then .TextAlign = 1
                  End With
        Next tablEcol
        
        toppos = txt.Top + txt.Height
        leftpos = 0
Next tablerow

    Me.Frame1.ScrollBars = fmScrollBarsBoth
    Me.Frame1.ScrollHeight = txt.Top + txt.Height
    Me.ScrollWidth = txt.Left + txt.Width
 
End Sub

Private Sub CommandButton1_Click()
    Unload Me
     
End Sub
Sub ShowTablePopup()
MsgBox UserForm3.Show, vbInformation
    'UserForm1.Show
End Sub



