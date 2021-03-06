Attribute VB_Name = "NumTie"
Sub generateNumTie()

Dim targetI As Range
Dim numEmpty As Range
Dim mntEmpty As Range
Dim lignes As Range
Dim nomFic As String
Dim maDate As String

maDate = Format(Date, "yyyymmdd")
nomFic = "Cashback_CUP_" & maDate


If Sheets(Sheets.Count).Name = "CashbackGenerator" Then
    Set newF = Worksheets.Add(, Worksheets(Worksheets.Count))
    newF.Name = "Feuil1"
ElseIf Sheets(Sheets.Count).Name <> "CashbackGenerator" Then
    Application.DisplayAlerts = False
    Sheets(Sheets.Count).Delete
    Set newF = Worksheets.Add(, Worksheets(Worksheets.Count))
    newF.Name = "Feuil1"
End If

Sheets("CashbackGenerator").Select
Range("A2:A" & [A1048576].End(xlUp).Row).Select
For Each numEmpty In Selection
    If numEmpty.Value = Empty Then
        MsgBox "Il manque un ou plusieurs num�ro de tiers pour g�n�rer le cashback", vbCritical + vbOKOnly, "Erreur"
        Exit Sub
    End If
Next

Sheets("CashbackGenerator").Select
Range("B2:B" & [B1048576].End(xlUp).Row).Select
For Each mntEmpty In Selection
    If mntEmpty.Value = Empty Then
        MsgBox "Il manque un ou plusieurs montant de bon d'achat pour g�n�rer le cashback", vbCritical + vbOKOnly, "Erreur"
        Exit Sub
    End If
Next

Worksheets("CashbackGenerator").Range("C2:C1048576").ClearContents
For n = 2 To Sheets("CashbackGenerator").Range("A1048576").End(xlUp).Row
    Set C = Sheets("ACC_CLIENT_PORTEUR").Columns(12).Find(Sheets("CashbackGenerator").Range("A" & n), LookIn:=xlValues, lookat:=xlWhole)
    Set d = Sheets("ACC_CLIENT_PORTEUR").Columns(13).Find(Sheets("CashbackGenerator").Range("A" & n), LookIn:=xlValues, lookat:=xlWhole)
    Set e = Sheets("CashbackGenerator").Columns(1).Find(Sheets("CashbackGenerator").Range("A" & n), LookIn:=xlValues, lookat:=xlWhole)
    
        If Sheets("CashbackGenerator").Range("C" & n).Value = "Introuvable" Then
    
            If Not C Is Nothing Then
                Sheets("CashbackGenerator").Range("C" & n) = Sheets("ACC_CLIENT_PORTEUR").Range("A" & C.Row)
                Sheets("Feuil1").Range("A" & n) = Sheets("ACC_CLIENT_PORTEUR").Range("A" & C.Row) & ";" & Sheets("CashbackGenerator").Range("B" & e.Row) * 100 & ";" & DateSerial(Year(Date), Month(Date) + 4, 1) - 1 & " 00:00:00"
            ElseIf Not d Is Nothing Then
                Sheets("CashbackGenerator").Range("C" & n) = Sheets("ACC_CLIENT_PORTEUR").Range("A" & d.Row)
                Sheets("Feuil1").Range("A" & n) = Sheets("ACC_CLIENT_PORTEUR").Range("A" & d.Row) & ";" & Sheets("CashbackGenerator").Range("B" & e.Row) * 100 & ";" & DateSerial(Year(Date), Month(Date) + 4, 1) - 1 & " 00:00:00"
            Else
                Sheets("CashbackGenerator").Range("C" & n) = "Introuvable"
                Sheets("Feuil1").Range("A" & n) = "Introuvable"
            End If
        
        ElseIf Sheets("CashbackGenerator").Range("C" & n).Value <> "Introuvable" Then
     
            If Not C Is Nothing Then
                Sheets("CashbackGenerator").Range("C" & n) = Sheets("ACC_CLIENT_PORTEUR").Range("A" & C.Row)
                Sheets("Feuil1").Range("A" & n) = Sheets("ACC_CLIENT_PORTEUR").Range("A" & C.Row) & ";" & Sheets("CashbackGenerator").Range("B" & e.Row) * 100 & ";" & DateSerial(Year(Date), Month(Date) + 4, 1) - 1 & " 00:00:00"
            ElseIf Not d Is Nothing Then
                Sheets("CashbackGenerator").Range("C" & n) = Sheets("ACC_CLIENT_PORTEUR").Range("A" & d.Row)
                Sheets("Feuil1").Range("A" & n) = Sheets("ACC_CLIENT_PORTEUR").Range("A" & d.Row) & ";" & Sheets("CashbackGenerator").Range("B" & e.Row) * 100 & ";" & DateSerial(Year(Date), Month(Date) + 4, 1) - 1 & " 00:00:00"
            Else
                Sheets("CashbackGenerator").Range("C" & n) = "Introuvable"
                Sheets("Feuil1").Range("A" & n) = "Introuvable"
            End If
        
        Else
            Range("C" & n).Value = ""
        
        End If
Next n

Sheets("Feuil1").Select
Range("A1:A" & [A1048576].End(xlUp).Row).Select
Set targetI = Columns(1).Find(what:="Introuvable", LookIn:=xlValues, lookat:=xlWhole)
    If Not targetI Is Nothing Then
        Worksheets("CashbackGenerator").Range("A2:C1048576").Borders.LineStyle = xlNone
        Sheets("CashbackGenerator").Select
        Range("C2:C" & [C1048576].End(xlUp).Row).Select
        For Each a In Selection
        If a.Value <> "Introuvable" Then a.EntireRow.Hidden = True
        Next
        MsgBox "Certains identifiants n'ont pas trouv� de correspondance", vbCritical + vbOKOnly, "Erreur"
        Exit Sub
    Else
        If Dir("C:\Users\" & Environ("username") & "\Desktop\" & nomFic & ".txt") = "" Then
            ChDir "C:\Users\" & Environ("username") & "\Desktop"
            ActiveWorkbook.SaveAs ("C:\Users\" & Environ("username") & "\Desktop\" & nomFic & ".txt"), FileFormat:= _
            xlText, CreateBackup:=False
            MsgBox "Le fichier Cashback_CUP_" & maDate & ".txt vient d'�tre cr�e sur le Bureau", vbInformation + vbOKOnly, "Cr�ation du fichier"
        Else
            Msg = ("Le fichier Cashback_CUP_" & maDate & ".txt  existe d�j� !" & vbCrLf & vbCrLf & "Voulez-vous le remplacer ?")
            Style = vbYesNo + vbInformation
            Title = "Demande de confirmation"
            Response = MsgBox(Msg, Style, Title)
            If Response = vbYes Then
                Application.DisplayAlerts = False
                ActiveWorkbook.SaveAs ("C:\Users\" & Environ("username") & "\Desktop\" & nomFic & ".txt"), FileFormat:= _
                xlText, CreateBackup:=False
                Application.DisplayAlerts = True
                MsgBox "Le fichier Cashback_CUP_" & maDate & ".txt vient d'�tre cr�e sur le Bureau", vbInformation + vbOKOnly, "Cr�ation du fichier"
            Else
                MsgBox "Le fichier Cashback_CUP_" & maDate & ".txt n'a pas �t� cr�e", vbInformation + vbOKOnly, "Cr�ation du fichier"
            End If
        End If
    End If
    
Worksheets("CashbackGenerator").Range("A2:C1048576").ClearContents
Worksheets("CashbackGenerator").Range("A2:C1048576").Borders.LineStyle = xlNone
Application.DisplayAlerts = False
Sheets(Sheets.Count).Delete
Application.DisplayAlerts = True
    
If Dir("U:\Pierrick\R�tentionCashBack\CashbackGenerator.xlsm") = "" Then
    ActiveWorkbook.SaveAs ("U:\Pierrick\R�tentionCashBack\CashbackGenerator.xlsm"), FileFormat:= _
    xlText, CreateBackup:=False
    MsgBox "Le fichier CashbackGenerator.xlsm a bien �t� sauvegard�", vbInformation + vbOKOnly, "Sauvegarde du fichier"
Else
    Msg = ("Le fichier CashbackGenerator.xlsm existe d�j� !" & vbCrLf & vbCrLf & "Voulez-vous le remplacer ?")
    Style = vbYesNo + vbInformation
    Title = "Demande de confirmation"
    Response = MsgBox(Msg, Style, Title)
    If Response = vbYes Then
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs ("U:\Pierrick\R�tentionCashBack\CashbackGenerator.xlsm"), FileFormat:= _
        xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
        Application.DisplayAlerts = True
        MsgBox "Le fichier CashbackGenerator.xlsm a bien �t� sauvegard�", vbInformation + vbOKOnly, "Sauvegarde du fichier"
    Else
        MsgBox "Le fichier CashbackGenerator.xlsm n'a pas �t� sauvegard�", vbInformation + vbOKOnly, "Sauvegarde du fichier"
    End If
End If

'ChDir "\\uf96-001.cm-cic.fr\BCA_DPOI\02-SIT\Temporaire"
'ActiveWorkbook.SaveAs Filename:= _
'    "\\uf96-001.cm-cic.fr\BCA_DPOI\02-SIT\Temporaire\CashbackGenerator.xlsm", _
'    FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False

End Sub

