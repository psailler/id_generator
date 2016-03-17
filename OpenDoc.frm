VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OpenDoc 
   Caption         =   "Manuel d'utilisation"
   ClientHeight    =   10470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11805
   OleObjectBlob   =   "OpenDoc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OpenDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Unload Me
UsfOpen.Show 0
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub CommandButton3_Click()
If Frame2.Zoom > 381 Then 'le zoom est limité à 400 au max
MsgBox ("Zomm maximum atteint")
Else
Frame2.Zoom = Frame2.Zoom + 20 'on zoom de 20 en 20
End If
End Sub

Private Sub CommandButton4_Click()
If Frame2.Zoom < 30 Then 'le zoom min est 10
MsgBox ("Zoom minimum atteint")
Else
Frame2.Zoom = Frame2.Zoom - 20 'on dezoom de 20 en 20
End If
End Sub


Private Sub CommandButton5_Click()
Dim AppWrd As Object
Set AppWrd = CreateObject("Word.Application")
AppWrd.Documents.Open Filename:="U:\Pierrick\RétentionCashBack\Doc.docx"
AppWrd.Visible = True
End Sub



Private Sub CommandButton6_Click()
If Image1.Tag = "Doc" Then
    With Image1
    .Picture = LoadPicture("U:\Pierrick\RétentionCashBack\Doc1.jpg")
    .Tag = "Doc1"
    .BorderStyle = fmBorderStyleSingle
    .PictureSizeMode = fmPictureSizeModeZoom
    End With
    Label1.Caption = "Page 2/5"
ElseIf Image1.Tag = "Doc1" Then
    With Image1
    .Picture = LoadPicture("U:\Pierrick\RétentionCashBack\Doc2.jpg")
    .Tag = "Doc2"
    .BorderStyle = fmBorderStyleSingle
    .PictureSizeMode = fmPictureSizeModeZoom
    End With
    Label1.Caption = "Page 3/5"
ElseIf Image1.Tag = "Doc2" Then
    With Image1
    .Picture = LoadPicture("U:\Pierrick\RétentionCashBack\Doc3.jpg")
    .Tag = "Doc3"
    .BorderStyle = fmBorderStyleSingle
    .PictureSizeMode = fmPictureSizeModeZoom
    End With
    Label1.Caption = "Page 4/5"
ElseIf Image1.Tag = "Doc3" Then
    With Image1
    .Picture = LoadPicture("U:\Pierrick\RétentionCashBack\Doc4.jpg")
    .Tag = "Doc4"
    .BorderStyle = fmBorderStyleSingle
    .PictureSizeMode = fmPictureSizeModeZoom
    End With
    Label1.Caption = "Page 5/5"
Else
End If
End Sub
Private Sub CommandButton7_Click()
If Image1.Tag = "Doc4" Then
    With Image1
    .Picture = LoadPicture("U:\Pierrick\RétentionCashBack\Doc3.jpg")
    .Tag = "Doc3"
    .BorderStyle = fmBorderStyleSingle
    .PictureSizeMode = fmPictureSizeModeZoom
    End With
    Label1.Caption = "Page 4/5"
ElseIf Image1.Tag = "Doc3" Then
    With Image1
    .Picture = LoadPicture("U:\Pierrick\RétentionCashBack\Doc2.jpg")
    .Tag = "Doc2"
    .BorderStyle = fmBorderStyleSingle
    .PictureSizeMode = fmPictureSizeModeZoom
    End With
    Label1.Caption = "Page 3/5"
ElseIf Image1.Tag = "Doc2" Then
    With Image1
    .Picture = LoadPicture("U:\Pierrick\RétentionCashBack\Doc1.jpg")
    .Tag = "Doc1"
    .BorderStyle = fmBorderStyleSingle
    .PictureSizeMode = fmPictureSizeModeZoom
    End With
    Label1.Caption = "Page 2/5"
ElseIf Image1.Tag = "Doc1" Then
    With Image1
    .Picture = LoadPicture("U:\Pierrick\RétentionCashBack\Doc.jpg")
    .Tag = "Doc"
    .BorderStyle = fmBorderStyleSingle
    .PictureSizeMode = fmPictureSizeModeZoom
    End With
    Label1.Caption = "Page 1/5"
Else
End If
End Sub

Private Sub UserForm_Initialize()
With Frame2
        .ScrollBars = fmScrollBarsBoth

        '~~> Change the values of 2 as Per your requirements
        .ScrollHeight = .InsideHeight * 2
        .ScrollWidth = .InsideWidth * 9
    End With

    With Image1
        .Picture = LoadPicture("U:\Pierrick\RétentionCashBack\Doc.jpg")
        .Tag = "Doc"
        .BorderStyle = fmBorderStyleSingle
        .PictureSizeMode = fmPictureSizeModeZoom
    End With
    Label1.Caption = "Page 1/5"
End Sub
