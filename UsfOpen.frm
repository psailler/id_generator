VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UsfOpen 
   Caption         =   "            Cashback Generator"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "UsfOpen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UsfOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
UsfOpen.Hide
MAIN.Show 0
End Sub

Private Sub CommandButton2_Click()
UsfOpen.Hide
OpenDoc.Show 0
End Sub

Private Sub CommandButton3_Click()
Unload Me
End Sub

Private Sub UserForm_Activate()
trois_boutons Me
End Sub

Private Sub UserForm_Resize()
maForm_Resize Me
End Sub

