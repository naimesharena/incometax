Attribute VB_Name = "mdProgressbar"
'Developed By Vishesh Jindal

Option Explicit
Public ProgressFrameCaption As Variant      'Set the Frame Caption
Public mainProcCaption As Variant           'Set the main procedure Caption
Public subProcCaption As Variant            'Set the Sub procedure Caption
Public mainProcPerc As Variant              'Set the main procedure completed percentage Caption
Public subProcPerc As Variant               'Set the sub procedure completed percentage Caption
Public noOfProcessMain As Variant           'Set the no of mainProcedures Caption
Public noOfProcessSub As Variant            'Set the no Of subProcedures Caption
Public processMainCounter As Variant        'count the no of MainProcedure completed
Public processSubCounter As Variant         'count the no of SubProcedure  completed
Public tempProgressFrameCaption As Variant
Public tempmainProcCaption As Variant
Public tempsubProcCaption As Variant



Sub UpdateProgressBar()

    If noOfProcessMain > 0 Then mainProcPerc = (100 / noOfProcessMain) * (processMainCounter) Else mainProcPerc = 0
    If noOfProcessSub > 0 Then subProcPerc = (100 / noOfProcessSub) * (processSubCounter) Else subProcPerc = 0
    If noOfProcessSub > 0 Then
        processSubCounter = processSubCounter + 1
        If processSubCounter = noOfProcessSub Then
            noOfProcessSub = 0
            processSubCounter = 0
        End If
    Else
    If Not mainProcPerc >= 100 Then processMainCounter = processMainCounter + 1
        
    End If
        'processMainCounter = processMainCounter + 1
        Dim widthForm As Variant
        widthForm = UserForm1.ProgressFrame.Width
        With UserForm1
        ChangeCaptions

    .lblmainProgPercDisp.caption = Round(mainProcPerc, 2) & "%"
    .lblmainProgPerc.Width = (mainProcPerc * (widthForm - 10) * 0.01)
    .lblsubProgPercDisp.caption = Round(subProcPerc, 2) & "%"
    .lblsubProgPerc.Width = (subProcPerc * (widthForm - 10) * 0.01)
   
    End With
    UserForm1.Repaint
'    DoEvents
End Sub
Sub ChangeCaptions()
    If UserForm1.ProgressFrame.caption = "" Then
        UserForm1.ProgressFrame.caption = ProgressFrameCaption
        ElseIf UserForm1.ProgressFrame.caption <> ProgressFrameCaption Then
        UserForm1.ProgressFrame.caption = ProgressFrameCaption
    End If
    If UserForm1.lblmainProgCaption.caption = "" Then
        UserForm1.lblmainProgCaption.caption = mainProcCaption
        ElseIf UserForm1.lblmainProgCaption.caption <> mainProcCaption Then
        UserForm1.lblmainProgCaption.caption = mainProcCaption
    End If
    If UserForm1.lblsubProgCaption.caption = "" Then
        UserForm1.lblsubProgCaption.caption = subProcCaption
        ElseIf UserForm1.lblsubProgCaption.caption <> subProcCaption Then
        UserForm1.lblsubProgCaption.caption = subProcCaption
    End If
End Sub
Sub ShowProgressBar()
UserForm1.Show
End Sub

Sub ProgressBarHide()
Unload UserForm1
InitProgBar
End Sub

Sub InitProgBar()
mainProcCaption = ""
subProcCaption = ""
mainProcPerc = 0
subProcPerc = 0
noOfProcessMain = 0
noOfProcessSub = 0
processMainCounter = 0
processSubCounter = 0
ProgressFrameCaption = ""
xmlLinesCount = 0
End Sub

'Vishesh Jindal

