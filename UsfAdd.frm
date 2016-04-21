VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UsfAdd 
   Caption         =   "Formulaire d'ajout client"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10110
   OleObjectBlob   =   "UsfAdd.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UsfAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ColorsLabel(Optional Ctrl As Control)
Dim j As Byte
    For j = 1 To 13
    Me.Controls("Label" & j).ForeColor = RGB(1, 103, 53)
    Next j
End Sub

Private Sub UserForm_activate()
    usfadd.BackColor = RGB(198, 209, 43)
    Call ColorsLabel
    trois_boutons Me
End Sub

Private Sub UserForm_Resize()
    maForm_Resize Me
End Sub

Private Sub Buttons(Optional Ctrl As Control)
Dim i As Byte
    For i = 14 To 17
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

Private Sub Label14_Click()

Dim L As Long
    
    If MsgBox("Etes-vous certain de vouloir enregistrer ce contact ?", vbYesNo, "Demande de confirmation") = vbYes Then
    L = Worksheets("ACC_CLIENT_PORTEUR").Range("A1048576").End(xlUp).Row + 1
    
    If TextBox1.BackColor = vbRed Then
        Exit Sub
    ElseIf TextBox4.BackColor = vbRed Then
        Exit Sub
    ElseIf TextBox6.BackColor = vbRed Then
        Exit Sub
    ElseIf TextBox9.BackColor = vbRed Then
        Exit Sub
    ElseIf TextBox10.BackColor = vbRed Then
        Exit Sub
    ElseIf TextBox11.BackColor = vbRed Then
        Exit Sub
    ElseIf TextBox12.BackColor = vbRed Then
        Exit Sub
    End If


'ID_CDISCOUNT
    Sheets("ACC_CLIENT_PORTEUR").Range("A" & L).Value = TextBox1

'CIVILITE
    If OptionButton1.Value = True Then
    Sheets("ACC_CLIENT_PORTEUR").Range("B" & L).Value = OptionButton1.Caption
    ElseIf OptionButton2.Value = True Then
    Sheets("ACC_CLIENT_PORTEUR").Range("B" & L).Value = OptionButton2.Caption
    ElseIf OptionButton3.Value = True Then
    Sheets("ACC_CLIENT_PORTEUR").Range("B" & L).Value = OptionButton3.Caption
    ElseIf OptionButton1.Value = False Or OptionButton2.Value = False Or OptionButton3.Value = False Then
    Sheets("ACC_CLIENT_PORTEUR").Range("B" & L).Value = ""
    End If


'NOM
    Sheets("ACC_CLIENT_PORTEUR").Range("C" & L).Value = TextBox2

'PRENOM
    Sheets("ACC_CLIENT_PORTEUR").Range("D" & L).Value = TextBox3

'DATE_NAISSANCE
    Sheets("ACC_CLIENT_PORTEUR").Range("E" & L).Value = TextBox4

'ADRESSE
    Sheets("ACC_CLIENT_PORTEUR").Range("F" & L).Value = TextBox5

'CP
    Sheets("ACC_CLIENT_PORTEUR").Range("G" & L).Value = TextBox6

'VILLE
    Sheets("ACC_CLIENT_PORTEUR").Range("H" & L).Value = TextBox7

'EMAIL
    Sheets("ACC_CLIENT_PORTEUR").Range("I" & L).Value = TextBox8

'RIB
    Sheets("ACC_CLIENT_PORTEUR").Range("J" & L).Value = TextBox9

'NUM_ISO
    Sheets("ACC_CLIENT_PORTEUR").Range("K" & L).Value = TextBox10

'NUM_TIE
    Sheets("ACC_CLIENT_PORTEUR").Range("L" & L).Value = TextBox11

'REF
    Sheets("ACC_CLIENT_PORTEUR").Range("M" & L).Value = TextBox12

End If
    
    OptionButton1 = Unchecked
    OptionButton2 = Unchecked
    OptionButton3 = Unchecked
    TextBox1 = ""
    TextBox1.BackColor = vbWhite
    TextBox2 = ""
    TextBox2.BackColor = vbWhite
    TextBox3 = ""
    TextBox3.BackColor = vbWhite
    TextBox4 = ""
    TextBox4.BackColor = vbWhite
    TextBox5 = ""
    TextBox5.BackColor = vbWhite
    TextBox6 = ""
    TextBox6.BackColor = vbWhite
    TextBox7 = ""
    TextBox7.BackColor = vbWhite
    TextBox8 = ""
    TextBox8.BackColor = vbWhite
    TextBox9 = ""
    TextBox9.BackColor = vbWhite
    TextBox10 = ""
    TextBox10.BackColor = vbWhite
    TextBox11 = ""
    TextBox11.BackColor = vbWhite
        With TextBox4
            .Enabled = True
            .BackStyle = fmBackStyleOpaque
        End With
    TextBox12 = ""
    TextBox12.BackColor = vbWhite
End Sub
Private Sub Label15_Click()
    OptionButton1 = Unchecked
    OptionButton2 = Unchecked
    OptionButton3 = Unchecked
    TextBox1 = ""
    TextBox1.BackColor = vbWhite
    TextBox2 = ""
    TextBox2.BackColor = vbWhite
    TextBox3 = ""
    TextBox3.BackColor = vbWhite
    TextBox4 = ""
    TextBox4.BackColor = vbWhite
    TextBox5 = ""
    TextBox5.BackColor = vbWhite
    TextBox6 = ""
    TextBox6.BackColor = vbWhite
    TextBox7 = ""
    TextBox7.BackColor = vbWhite
    TextBox8 = ""
    TextBox8.BackColor = vbWhite
    TextBox9 = ""
    TextBox9.BackColor = vbWhite
    TextBox10 = ""
    TextBox10.BackColor = vbWhite
    TextBox11 = ""
    TextBox11.BackColor = vbWhite
        With TextBox4
            .Enabled = True
            .BackStyle = fmBackStyleOpaque
        End With
    TextBox12 = ""
    TextBox12.BackColor = vbWhite
End Sub
Private Sub Label16_Click()
    Unload Me
    begin.Show 0
End Sub
Private Sub Label17_Click()
    Unload Me
End Sub

Private Sub Label14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Buttons Label14
End Sub
Private Sub Label15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Buttons Label15
End Sub
Private Sub Label16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Buttons Label16
End Sub
Private Sub Label17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Buttons Label17
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Buttons
End Sub

Private Sub TextBox1_Change()
    TextBox1.MaxLength = 12
    TextBox1.Text = UCase(TextBox1.Text)
End Sub

Private Sub TextBox1_AfterUpdate()
    TextBox1.MaxLength = 12
        If TextBox1.TextLength = 12 Then
            TextBox1.BackColor = vbWhite
        ElseIf TextBox1.TextLength = 11 Then
            MsgBox "L'identifiant que vous avez saisi n'est pas au bon format (11 caractères)" & vbNewLine & "Exemple : 000000001D3K ", vbOKOnly + vbCritical, "Format de l'identifiant"
            TextBox1.BackColor = vbRed
            Exit Sub
        ElseIf TextBox1.TextLength = 10 Then
            MsgBox "L'identifiant que vous avez saisi n'est pas au bon format (10 caractères)" & vbNewLine & "Exemple : 000000001D3K ", vbOKOnly + vbCritical, "Format de l'identifiant"
            TextBox1.BackColor = vbRed
            Exit Sub
        ElseIf TextBox1.TextLength = 9 Then
            MsgBox "L'identifiant que vous avez saisi n'est pas au bon format (9 caractères)" & vbNewLine & "Exemple : 000000001D3K ", vbOKOnly + vbCritical, "Format de l'identifiant"
            TextBox1.BackColor = vbRed
            Exit Sub
        ElseIf TextBox1.TextLength = 8 Then
            MsgBox "L'identifiant que vous avez saisi n'est pas au bon format (8 caractères)" & vbNewLine & "Exemple : 000000001D3K ", vbOKOnly + vbCritical, "Format de l'identifiant"
            TextBox10.BackColor = vbRed
            Exit Sub
        ElseIf TextBox1.TextLength = 7 Then
            MsgBox "L'identifiant que vous avez saisi n'est pas au bon format (7 caractères)" & vbNewLine & "Exemple : 000000001D3K ", vbOKOnly + vbCritical, "Format de l'identifiant"
            TextBox1.BackColor = vbRed
            Exit Sub
        ElseIf TextBox1.TextLength = 6 Then
            MsgBox "L'identifiant que vous avez saisi n'est pas au bon format (6 caractères)" & vbNewLine & "Exemple : 000000001D3K ", vbOKOnly + vbCritical, "Format de l'identifiant"
            TextBox1.BackColor = vbRed
            Exit Sub
        ElseIf TextBox1.TextLength = 5 Then
            MsgBox "L'identifiant que vous avez saisi n'est pas au bon format (5 caractères)" & vbNewLine & "Exemple : 000000001D3K ", vbOKOnly + vbCritical, "Format de l'identifiant"
            TextBox1.BackColor = vbRed
            Exit Sub
        ElseIf TextBox1.TextLength = 4 Then
            MsgBox "L'identifiant que vous avez saisi n'est pas au bon format (4 caractères)" & vbNewLine & "Exemple : 000000001D3K ", vbOKOnly + vbCritical, "Format de l'identifiant"
            TextBox1.BackColor = vbRed
            Exit Sub
        ElseIf TextBox1.TextLength = 3 Then
            MsgBox "L'identifiant que vous avez saisi n'est pas au bon format (3 caractères)" & vbNewLine & "Exemple : 000000001D3K ", vbOKOnly + vbCritical, "Format de l'identifiant"
            TextBox1.BackColor = vbRed
            Exit Sub
        ElseIf TextBox1.TextLength = 2 Then
            MsgBox "L'identifiant que vous avez saisi n'est pas au bon format (2 caractères)" & vbNewLine & "Exemple : 000000001D3K ", vbOKOnly + vbCritical, "Format de l'identifiant"
            TextBox1.BackColor = vbRed
            Exit Sub
        ElseIf TextBox1.TextLength = 1 Then
            MsgBox "L'identifiant que vous avez saisi n'est pas au bon format (1 caractère)" & vbNewLine & "Exemple : 000000001D3K ", vbOKOnly + vbCritical, "Format de l'identifiant"
            TextBox1.BackColor = vbRed
            Exit Sub
        End If
End Sub

Private Sub TextBox2_Change()
    TextBox2.Text = UCase(TextBox2.Text)
End Sub
Private Sub TextBox3_Change()
    TextBox3.Text = UCase(TextBox3.Text)
End Sub

Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Dim Valeur As Byte

    TextBox4.MaxLength = 10

    Valeur = Len(TextBox4)
        If Valeur = 2 Or Valeur = 5 Then TextBox4 = TextBox4 & "/"
    
        If InStr("0123456789/", Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
            MsgBox "Il faut saisir une date" & vbNewLine & "Exemple : 01/01/2016 ", vbOKOnly + vbCritical, "Format de la date"
        End If
        If Chr(KeyAscii) = "/" Then
            KeyAscii = 0
            MsgBox "Il n'est pas nécessaire de saisir '/'," & vbNewLine & "Il faut seulement saisir les chiffres" & vbNewLine & "Exemple : 01/01/2016 ", vbOKOnly + vbCritical, "Format de la date"
        End If
End Sub

Private Sub TextBox4_AfterUpdate()
    If TextBox4.TextLength = 10 Then
        TextBox4.BackColor = vbWhite
    ElseIf TextBox4.TextLength = 9 Then
        MsgBox "Le format que vous avez saisi est incorrect (9 caractères)" & vbNewLine & "Exemple : 01/01/2016 ", vbOKOnly + vbCritical, "Format de la date"
        TextBox4.BackColor = vbRed
    ElseIf TextBox4.TextLength = 8 Then
        MsgBox "Le format que vous avez saisi est incorrect (8 caractères)" & vbNewLine & "Exemple : 01/01/2016 ", vbOKOnly + vbCritical, "Format de la date"
        TextBox4.BackColor = vbRed
    ElseIf TextBox4.TextLength = 7 Then
        MsgBox "Le format que vous avez saisi est incorrect (7 caractères)" & vbNewLine & "Exemple : 01/01/2016 ", vbOKOnly + vbCritical, "Format de la date"
        TextBox4.BackColor = vbRed
    ElseIf TextBox4.TextLength = 6 Then
        MsgBox "Le format que vous avez saisi est incorrect (6 caractères)" & vbNewLine & "Exemple : 01/01/2016 ", vbOKOnly + vbCritical, "Format de la date"
        TextBox4.BackColor = vbRed
    ElseIf TextBox4.TextLength = 5 Then
        MsgBox "Le format que vous avez saisi est incorrect (5 caractères)" & vbNewLine & "Exemple : 01/01/2016 ", vbOKOnly + vbCritical, "Format de la date"
        TextBox4.BackColor = vbRed
    ElseIf TextBox4.TextLength = 4 Then
        MsgBox "Le format que vous avez saisi est incorrect (4 caractères)" & vbNewLine & "Exemple : 01/01/2016 ", vbOKOnly + vbCritical, "Format de la date"
        TextBox4.BackColor = vbRed
    ElseIf TextBox4.TextLength = 3 Then
        MsgBox "Le format que vous avez saisi est incorrect (3 caractères)" & vbNewLine & "Exemple : 01/01/2016 ", vbOKOnly + vbCritical, "Format de la date"
        TextBox4.BackColor = vbRed
    ElseIf TextBox4.TextLength = 2 Then
        MsgBox "Le format que vous avez saisi est incorrect (2 caractères)" & vbNewLine & "Exemple : 01/01/2016 ", vbOKOnly + vbCritical, "Format de la date"
        TextBox4.BackColor = vbRed
    ElseIf TextBox4.TextLength = 1 Then
        MsgBox "Le format que vous avez saisi est incorrect (1 caractère)" & vbNewLine & "Exemple : 01/01/2016 ", vbOKOnly + vbCritical, "Format de la date"
        TextBox4.BackColor = vbRed
    End If
End Sub

Private Sub TextBox5_Change()
    TextBox5.Text = UCase(TextBox5.Text)
End Sub

Private Sub TextBox6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

TextBox6.MaxLength = 5

    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
        MsgBox "Il faut saisir un Code Postal" & vbNewLine & "Exemple : 33 - 33000 ", vbOKOnly + vbCritical, "Format du Code Postal"
    End If
End Sub

Private Sub TextBox6_AfterUpdate()

TextBox6.MaxLength = 5
    If TextBox6.TextLength = 5 Then
        TextBox6.BackColor = vbWhite
    ElseIf TextBox6.TextLength = 4 Then
        MsgBox "Le Code Postal que vous avez saisi est incorrect (4 caractères)" & vbNewLine & "Exemple : 33 - 33000 ", vbOKOnly + vbCritical, "Code Postal incorrect"
        TextBox6.BackColor = vbRed
    ElseIf TextBox6.TextLength = 3 Then
        MsgBox "Le Code Postal que vous avez saisi est incorrect (3 caractères)" & vbNewLine & "Exemple : 33 - 33000 ", vbOKOnly + vbCritical, "Code Postal incorrect"
        TextBox6.BackColor = vbRed
    ElseIf TextBox6.TextLength = 2 Then
        TextBox6.BackColor = vbWhite
    ElseIf TextBox6.TextLength = 1 Then
        MsgBox "Le Code Postal que vous avez saisi est incorrect (1 caractère)" & vbNewLine & "Exemple : 33 - 33000 ", vbOKOnly + vbCritical, "Code Postal incorrect"
        TextBox6.BackColor = vbRed
    End If
End Sub

Private Sub TextBox7_Change()
    TextBox7.Text = UCase(TextBox7.Text)
End Sub

Private Sub TextBox8_Change()
    TextBox8.Text = UCase(TextBox8.Text)
End Sub

Private Sub TextBox9_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

TextBox9.MaxLength = 21
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
        MsgBox "Il faut saisir un RIB de 21 chiffres" & vbNewLine & "Exemple : 146289551400028000000 ", vbOKOnly + vbCritical, "Format du RIB"
    End If
End Sub

Private Sub TextBox9_AfterUpdate()
    TextBox9.MaxLength = 21
        If TextBox9.TextLength = 21 Then
            TextBox9.BackColor = vbWhite
        ElseIf TextBox9.TextLength = 20 Then
            MsgBox "Le format que vous avez saisi est incorrect (20 caractères)" & vbNewLine & "Exemple : 146289551400028000000 ", vbOKOnly + vbCritical, "Format de du RIB"
            TextBox9.BackColor = vbRed
        ElseIf TextBox9.TextLength = 19 Then
            MsgBox "Le format que vous avez saisi est incorrect (19 caractères)" & vbNewLine & "Exemple : 146289551400028000000 ", vbOKOnly + vbCritical, "Format de du RIB"
            TextBox9.BackColor = vbRed
        ElseIf TextBox9.TextLength = 18 Then
            MsgBox "Le format que vous avez saisi est incorrect (18 caractères)" & vbNewLine & "Exemple : 146289551400028000000 ", vbOKOnly + vbCritical, "Format de du RIB"
            TextBox9.BackColor = vbRed
        ElseIf TextBox9.TextLength = 17 Then
            MsgBox "Le format que vous avez saisi est incorrect (17 caractères)" & vbNewLine & "Exemple : 146289551400028000000 ", vbOKOnly + vbCritical, "Format de du RIB"
            TextBox9.BackColor = vbRed
        ElseIf TextBox9.TextLength = 16 Then
            MsgBox "Le format que vous avez saisi est incorrect (16 caractères)" & vbNewLine & "Exemple : 146289551400028000000 ", vbOKOnly + vbCritical, "Format de du RIB"
            TextBox9.BackColor = vbRed
        ElseIf TextBox9.TextLength = 15 Then
            MsgBox "Le format que vous avez saisi est incorrect (15 caractères)" & vbNewLine & "Exemple : 146289551400028000000 ", vbOKOnly + vbCritical, "Format de du RIB"
            TextBox9.BackColor = vbRed
        ElseIf TextBox9.TextLength = 14 Then
            MsgBox "Le format que vous avez saisi est incorrect (14 caractères)" & vbNewLine & "Exemple : 146289551400028000000 ", vbOKOnly + vbCritical, "Format de du RIB"
            TextBox9.BackColor = vbRed
        ElseIf TextBox9.TextLength = 13 Then
            MsgBox "Le format que vous avez saisi est incorrect (13 caractères)" & vbNewLine & "Exemple : 146289551400028000000 ", vbOKOnly + vbCritical, "Format de du RIB"
            TextBox9.BackColor = vbRed
        ElseIf TextBox9.TextLength = 12 Then
            MsgBox "Le format que vous avez saisi est incorrect (12 caractères)" & vbNewLine & "Exemple : 146289551400028000000 ", vbOKOnly + vbCritical, "Format de du RIB"
            TextBox9.BackColor = vbRed
        ElseIf TextBox9.TextLength = 11 Then
            MsgBox "Le format que vous avez saisi est incorrect (11 caractères)" & vbNewLine & "Exemple : 146289551400028000000 ", vbOKOnly + vbCritical, "Format de du RIB"
            TextBox9.BackColor = vbRed
        ElseIf TextBox9.TextLength = 10 Then
            MsgBox "Le format que vous avez saisi est incorrect (10 caractères)" & vbNewLine & "Exemple : 146289551400028000000 ", vbOKOnly + vbCritical, "Format de du RIB"
            TextBox9.BackColor = vbRed
        ElseIf TextBox9.TextLength = 9 Then
            MsgBox "Le format que vous avez saisi est incorrect (9 caractères)" & vbNewLine & "Exemple : 146289551400028000000 ", vbOKOnly + vbCritical, "Format de du RIB"
            TextBox9.BackColor = vbRed
        ElseIf TextBox9.TextLength = 8 Then
            MsgBox "Le format que vous avez saisi est incorrect (8 caractères)" & vbNewLine & "Exemple : 146289551400028000000 ", vbOKOnly + vbCritical, "Format de du RIB"
            TextBox9.BackColor = vbRed
        ElseIf TextBox9.TextLength = 7 Then
            MsgBox "Le format que vous avez saisi est incorrect (7 caractères)" & vbNewLine & "Exemple : 146289551400028000000 ", vbOKOnly + vbCritical, "Format de du RIB"
            TextBox9.BackColor = vbRed
        ElseIf TextBox9.TextLength = 6 Then
            MsgBox "Le format que vous avez saisi est incorrect (6 caractères)" & vbNewLine & "Exemple : 146289551400028000000 ", vbOKOnly + vbCritical, "Format de du RIB"
            TextBox9.BackColor = vbRed
        ElseIf TextBox9.TextLength = 5 Then
            MsgBox "Le format que vous avez saisi est incorrect (5 caractères)" & vbNewLine & "Exemple : 146289551400028000000 ", vbOKOnly + vbCritical, "Format de du RIB"
            TextBox9.BackColor = vbRed
        ElseIf TextBox9.TextLength = 4 Then
            MsgBox "Le format que vous avez saisi est incorrect (4 caractères)" & vbNewLine & "Exemple : 146289551400028000000 ", vbOKOnly + vbCritical, "Format de du RIB"
            TextBox9.BackColor = vbRed
        ElseIf TextBox9.TextLength = 3 Then
            MsgBox "Le format que vous avez saisi est incorrect (3 caractères)" & vbNewLine & "Exemple : 146289551400028000000 ", vbOKOnly + vbCritical, "Format de du RIB"
            TextBox9.BackColor = vbRed
        ElseIf TextBox9.TextLength = 2 Then
            MsgBox "Le format que vous avez saisi est incorrect (2 caractères)" & vbNewLine & "Exemple : 146289551400028000000 ", vbOKOnly + vbCritical, "Format de du RIB"
            TextBox9.BackColor = vbRed
        ElseIf TextBox9.TextLength = 1 Then
            MsgBox "Le format que vous avez saisi est incorrect (1 caractère)" & vbNewLine & "Exemple : 146289551400028000000 ", vbOKOnly + vbCritical, "Format de du RIB"
            TextBox9.BackColor = vbRed
        End If

If TextBox9.Value <> "" Then
TextBox12.Value = Right(TextBox9.Value, 11)
ElseIf TextBox9.Value = "" Then
TextBox12.Value = ""
End If
End Sub

Private Sub TextBox10_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

TextBox10.MaxLength = 16

If InStr("0123456789", Chr(KeyAscii)) = 0 Then
    KeyAscii = 0
    MsgBox "Il faut saisir un numéro de carte composé de 16 chiffres" & vbNewLine & "Exemple : 5399601010685240 ", vbOKOnly + vbCritical, "Format du numéro de carte"
End If
End Sub

Private Sub TextBox10_AfterUpdate()
TextBox10.MaxLength = 16
If TextBox10.TextLength = 16 Then
TextBox10.BackColor = vbWhite
ElseIf TextBox10.TextLength = 15 Then
MsgBox "Le format que vous avez saisi est incorrect" & vbNewLine & "Exemple : 5399601010685240 ", vbOKOnly + vbCritical, "Format de la date"
TextBox10.BackColor = vbRed
ElseIf TextBox10.TextLength = 14 Then
MsgBox "Le format que vous avez saisi est incorrect" & vbNewLine & "Exemple : 5399601010685240 ", vbOKOnly + vbCritical, "Format de la date"
TextBox10.BackColor = vbRed
ElseIf TextBox10.TextLength = 13 Then
MsgBox "Le format que vous avez saisi est incorrect" & vbNewLine & "Exemple : 5399601010685240 ", vbOKOnly + vbCritical, "Format de la date"
TextBox10.BackColor = vbRed
ElseIf TextBox10.TextLength = 12 Then
MsgBox "Le format que vous avez saisi est incorrect" & vbNewLine & "Exemple : 5399601010685240 ", vbOKOnly + vbCritical, "Format de la date"
TextBox10.BackColor = vbRed
ElseIf TextBox10.TextLength = 11 Then
MsgBox "Le format que vous avez saisi est incorrect" & vbNewLine & "Exemple : 5399601010685240 ", vbOKOnly + vbCritical, "Format de la date"
TextBox10.BackColor = vbRed
ElseIf TextBox10.TextLength = 10 Then
MsgBox "Le format que vous avez saisi est incorrect" & vbNewLine & "Exemple : 5399601010685240 ", vbOKOnly + vbCritical, "Format de la date"
TextBox10.BackColor = vbRed
ElseIf TextBox10.TextLength = 9 Then
MsgBox "Le format que vous avez saisi est incorrect" & vbNewLine & "Exemple : 5399601010685240 ", vbOKOnly + vbCritical, "Format de la date"
TextBox10.BackColor = vbRed
ElseIf TextBox10.TextLength = 8 Then
MsgBox "Le format que vous avez saisi est incorrect" & vbNewLine & "Exemple : 5399601010685240 ", vbOKOnly + vbCritical, "Format de la date"
TextBox10.BackColor = vbRed
ElseIf TextBox10.TextLength = 7 Then
MsgBox "Le format que vous avez saisi est incorrect" & vbNewLine & "Exemple : 5399601010685240 ", vbOKOnly + vbCritical, "Format de la date"
TextBox10.BackColor = vbRed
ElseIf TextBox10.TextLength = 6 Then
MsgBox "Le format que vous avez saisi est incorrect" & vbNewLine & "Exemple : 5399601010685240 ", vbOKOnly + vbCritical, "Format de la date"
TextBox10.BackColor = vbRed
ElseIf TextBox10.TextLength = 5 Then
MsgBox "Le format que vous avez saisi est incorrect" & vbNewLine & "Exemple : 5399601010685240 ", vbOKOnly + vbCritical, "Format de la date"
TextBox10.BackColor = vbRed
ElseIf TextBox10.TextLength = 4 Then
MsgBox "Le format que vous avez saisi est incorrect" & vbNewLine & "Exemple : 5399601010685240 ", vbOKOnly + vbCritical, "Format de la date"
TextBox10.BackColor = vbRed
ElseIf TextBox10.TextLength = 3 Then
MsgBox "Le format que vous avez saisi est incorrect" & vbNewLine & "Exemple : 5399601010685240 ", vbOKOnly + vbCritical, "Format de la date"
TextBox10.BackColor = vbRed
ElseIf TextBox10.TextLength = 2 Then
MsgBox "Le format que vous avez saisi est incorrect" & vbNewLine & "Exemple : 5399601010685240 ", vbOKOnly + vbCritical, "Format de la date"
TextBox10.BackColor = vbRed
ElseIf TextBox10.TextLength = 1 Then
MsgBox "Le format que vous avez saisi est incorrect" & vbNewLine & "Exemple : 5399601010685240 ", vbOKOnly + vbCritical, "Format de la date"
TextBox10.BackColor = vbRed
End If
End Sub

Private Sub TextBox11_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

TextBox11.MaxLength = 13

If InStr("0123456789", Chr(KeyAscii)) = 0 Then
    KeyAscii = 0
    MsgBox "Il faut saisir un numéro de tiers" & vbNewLine & "Exemple : 2000008007022 ", vbOKOnly + vbCritical, "Format du numéro de tiers"
End If
End Sub

Private Sub TextBox11_AfterUpdate()
TextBox11.MaxLength = 13
If TextBox11.TextLength = 13 Then
TextBox11.BackColor = vbWhite
ElseIf TextBox11.TextLength = 12 Then
MsgBox "Le format que vous avez saisi est incorrect (12 caractères)" & vbNewLine & "Exemple : 2000008007022 ", vbOKOnly + vbCritical, "Format du numéro de tiers"
TextBox11.BackColor = vbRed
ElseIf TextBox11.TextLength = 11 Then
MsgBox "Le format que vous avez saisi est incorrect (11 caractères)" & vbNewLine & "Exemple : 2000008007022 ", vbOKOnly + vbCritical, "Format du numéro de tiers"
TextBox11.BackColor = vbRed
ElseIf TextBox11.TextLength = 10 Then
MsgBox "Le format que vous avez saisi est incorrect (10 caractères)" & vbNewLine & "Exemple : 2000008007022 ", vbOKOnly + vbCritical, "Format du numéro de tiers"
TextBox11.BackColor = vbRed
ElseIf TextBox11.TextLength = 9 Then
MsgBox "Le format que vous avez saisi est incorrect (9 caractères)" & vbNewLine & "Exemple : 2000008007022 ", vbOKOnly + vbCritical, "Format du numéro de tiers"
TextBox11.BackColor = vbRed
ElseIf TextBox11.TextLength = 8 Then
MsgBox "Le format que vous avez saisi est incorrect (8 caractères)" & vbNewLine & "Exemple : 2000008007022 ", vbOKOnly + vbCritical, "Format du numéro de tiers"
TextBox11.BackColor = vbRed
ElseIf TextBox11.TextLength = 7 Then
MsgBox "Le format que vous avez saisi est incorrect (7 caractères)" & vbNewLine & "Exemple : 2000008007022 ", vbOKOnly + vbCritical, "Format du numéro de tiers"
TextBox11.BackColor = vbRed
ElseIf TextBox11.TextLength = 6 Then
MsgBox "Le format que vous avez saisi est incorrect (6 caractères)" & vbNewLine & "Exemple : 2000008007022 ", vbOKOnly + vbCritical, "Format du numéro de tiers"
TextBox11.BackColor = vbRed
ElseIf TextBox11.TextLength = 5 Then
MsgBox "Le format que vous avez saisi est incorrect (5 caractères)" & vbNewLine & "Exemple : 2000008007022 ", vbOKOnly + vbCritical, "Format du numéro de tiers"
TextBox11.BackColor = vbRed
ElseIf TextBox11.TextLength = 4 Then
MsgBox "Le format que vous avez saisi est incorrect (4 caractères)" & vbNewLine & "Exemple : 2000008007022 ", vbOKOnly + vbCritical, "Format du numéro de tiers"
TextBox11.BackColor = vbRed
ElseIf TextBox11.TextLength = 3 Then
MsgBox "Le format que vous avez saisi est incorrect (3 caractères)" & vbNewLine & "Exemple : 2000008007022 ", vbOKOnly + vbCritical, "Format du numéro de tiers"
TextBox11.BackColor = vbRed
ElseIf TextBox11.TextLength = 2 Then
MsgBox "Le format que vous avez saisi est incorrect (2 caractères)" & vbNewLine & "Exemple : 2000008007022 ", vbOKOnly + vbCritical, "Format du numéro de tiers"
TextBox11.BackColor = vbRed
ElseIf TextBox11.TextLength = 1 Then
MsgBox "Le format que vous avez saisi est incorrect (1 caractère)" & vbNewLine & "Exemple : 2000008007022 ", vbOKOnly + vbCritical, "Format du numéro de tiers"
TextBox11.BackColor = vbRed
End If
End Sub

Private Sub TextBox12_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

TextBox12.MaxLength = 12

If InStr("0123456789", Chr(KeyAscii)) = 0 Then
    KeyAscii = 0
    MsgBox "Il faut saisir de référence tiers" & vbNewLine & "Exemple : 00028000000 ", vbOKOnly + vbCritical, "Format de référence tiers"
End If
End Sub

Private Sub TextBox12_AfterUpdate()
TextBox12.MaxLength = 11
If TextBox12.TextLength = 11 Then
TextBox12.BackColor = vbWhite
ElseIf TextBox12.TextLength = 10 Then
MsgBox "Le format que vous avez saisi est incorrect (10 caractères)" & vbNewLine & "Exemple : 00028000000 ", vbOKOnly + vbCritical, "Format de référence tiers"
TextBox12.BackColor = vbRed
ElseIf TextBox12.TextLength = 9 Then
MsgBox "Le format que vous avez saisi est incorrect (9 caractères)" & vbNewLine & "Exemple : 00028000000 ", vbOKOnly + vbCritical, "Format de référence tiers"
TextBox12.BackColor = vbRed
ElseIf TextBox12.TextLength = 8 Then
MsgBox "Le format que vous avez saisi est incorrect (8 caractères)" & vbNewLine & "Exemple : 00028000000 ", vbOKOnly + vbCritical, "Format de référence tiers"
TextBox12.BackColor = vbRed
ElseIf TextBox12.TextLength = 7 Then
MsgBox "Le format que vous avez saisi est incorrect (7 caractères)" & vbNewLine & "Exemple : 00028000000 ", vbOKOnly + vbCritical, "Format de référence tiers"
TextBox12.BackColor = vbRed
ElseIf TextBox12.TextLength = 6 Then
MsgBox "Le format que vous avez saisi est incorrect (6 caractères)" & vbNewLine & "Exemple : 00028000000 ", vbOKOnly + vbCritical, "Format de référence tiers"
TextBox12.BackColor = vbRed
ElseIf TextBox12.TextLength = 5 Then
MsgBox "Le format que vous avez saisi est incorrect (5 caractères)" & vbNewLine & "Exemple : 00028000000 ", vbOKOnly + vbCritical, "Format de référence tiers"
TextBox12.BackColor = vbRed
ElseIf TextBox12.TextLength = 4 Then
MsgBox "Le format que vous avez saisi est incorrect (4 caractères)" & vbNewLine & "Exemple : 00028000000 ", vbOKOnly + vbCritical, "Format de référence tiers"
TextBox12.BackColor = vbRed
ElseIf TextBox12.TextLength = 3 Then
MsgBox "Le format que vous avez saisi est incorrect (3 caractères)" & vbNewLine & "Exemple : 00028000000 ", vbOKOnly + vbCritical, "Format de référence tiers"
TextBox12.BackColor = vbRed
ElseIf TextBox12.TextLength = 2 Then
MsgBox "Le format que vous avez saisi est incorrect (2 caractères)" & vbNewLine & "Exemple : 00028000000 ", vbOKOnly + vbCritical, "Format de référence tiers"
TextBox12.BackColor = vbRed
ElseIf TextBox12.TextLength = 1 Then
MsgBox "Le format que vous avez saisi est incorrect (1 caractère)" & vbNewLine & "Exemple : 00028000000 ", vbOKOnly + vbCritical, "Format de référence tiers"
TextBox12.BackColor = vbRed
End If
End Sub

