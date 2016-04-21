VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} begin 
   Caption         =   "Commencer"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3330
   OleObjectBlob   =   "begin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "begin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_activate()
Fond_dégradé Me, , RGB(198, 209, 43)
trois_boutons Me
End Sub

Private Sub UserForm_Resize()
maForm_Resize Me
End Sub
Private Sub Buttons(Optional Ctrl As Control)
Dim i As Byte
    For i = 1 To 4
        Me.Controls("Label" & i).ForeColor = vbBlack
        Me.Controls("Label" & i).BackStyle = fmBackStyleTransparent
        Me.Controls("Label" & i).BackColor = &HE0E0E0
        Me.Controls("Label" & i).SpecialEffect = 0
        Me.Controls("Label" & i).BorderStyle = fmBorderStyleSingle
        Me.Controls("Label" & i).BorderColor = &H8000&
    Next i
    If Not Ctrl Is Nothing Then
        Ctrl.ForeColor = vbWhite
        Ctrl.BackStyle = fmBackStyleOpaque
        Ctrl.BackColor = RGB(1, 103, 53)
        Ctrl.SpecialEffect = fmSpecialEffectRaised
        Ctrl.BorderStyle = fmBorderStyleSingle
        Ctrl.BorderColor = vbWhite
    End If
End Sub

Private Sub Label1_Click()
Call generateNumTie
'Mess Label1.Caption
End Sub
Private Sub Label2_Click()
begin.Hide
usfadd.Show 0
'Mess Label2.Caption
End Sub
Private Sub Label3_Click()
Unload Me
UsfOpen.Show 0
'Mess Label3.Caption
End Sub
Private Sub Label4_Click()
Unload Me
End Sub


Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Buttons Label1
End Sub
Private Sub Label2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Buttons Label2
End Sub
Private Sub Label3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Buttons Label3
End Sub
Private Sub Label4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Buttons Label4
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Buttons
End Sub

