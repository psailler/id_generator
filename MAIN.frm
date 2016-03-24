VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MAIN 
   Caption         =   "Commencer"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4425
   OleObjectBlob   =   "MAIN.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
Call generateNumTie
End Sub

Private Sub CommandButton2_Click()
Call generateIdCdiscount
End Sub

Private Sub CommandButton3_Click()
MAIN.Hide
UsfAdd.Show 0
End Sub

Private Sub CommandButton4_Click()
Unload Me
End Sub

Private Sub CommandButton5_Click()
Call generate
End Sub

Private Sub CommandButton6_Click()
Unload Me
UsfOpen.Show 0
End Sub

Private Sub UserForm_Activate()
trois_boutons Me
End Sub

Private Sub UserForm_Resize()
maForm_Resize Me
End Sub
