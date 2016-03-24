Attribute VB_Name = "Cashback"
Sub generate()

Dim targetI As Range
Dim idEmpty As Range
Dim lignes As Range
Dim nomFic As String
Dim maDate As String
Dim plage As Range
Dim plageA As Integer
Dim plageB As Integer
Dim plageC As Integer


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
Range("A1:C1048576").EntireRow.Hidden = False
Set plage = Union(Range("A1:A" & [A1048576].End(xlUp).Row), Range("B1:B" & [B1048576].End(xlUp).Row), Range("C1:C" & [C1048576].End(xlUp).Row))
plage.Select
For Each a In Selection
    If Application.CountA(a.EntireRow) = Empty Then
        MsgBox "Il y a une ou plusieurs lignes vides, impossible de générer le cashback", vbCritical + vbOKOnly, "Erreur"
        Exit Sub
    End If
Next

Sheets("CashbackGenerator").Select
plageA = Range("A1", [A1048576].End(xlUp)).Rows.Count
plageB = Range("B1", [B1048576].End(xlUp)).Rows.Count
plageC = Range("C1", [C1048576].End(xlUp)).Rows.Count
f = Sheets("CashbackGenerator").Range("B1048576").End(xlUp).Row
If ((plageA > plageB) And (plageA > plageC)) Then
    Range("B2:B" & [A1048576].End(xlUp).Row).Select
        For Each emptyCellB In Selection
            If emptyCellB.Value = Empty Then
                MsgBox "Il manque un ou plusieurs montant pour générer le cashback", vbCritical + vbOKOnly, "Erreur"
                Exit Sub
            End If
        Next
ElseIf ((plageB > plageA) And (plageB > plageC)) Then
    Range("A2:A" & [B1048576].End(xlUp).Row).Select
        For Each emptyCellA In Selection
            If emptyCellA.Value = Empty Then
                emptyCellA.Select
                If ActiveCell.Offset(0, 2).Value = Empty Then
                    MsgBox "Il manque un ou plusieurs numéros tiers pour générer le cashback", vbCritical + vbOKOnly, "Erreur"
                    Exit Sub
                ElseIf ActiveCell.Offset(0, 2).Value <> Empty Then
                    ActiveCell.Offset(0, 2).Select
                    e = ActiveCell.Address()
                    Range(e & ":C" & f).Select
                    For Each emptyCellC In Selection
                        If emptyCellC.Value = Empty Then
                            MsgBox "Il manque un ou plusieurs identifiants pour générer le cashback", vbCritical + vbOklonly, "Erreur"
                            Exit Sub
                        End If
                    Next
                End If
            End If
        Next
    Range("C2:C" & [B1048576].End(xlUp).Row).Select
        For Each emptyCellC In Selection
            If emptyCellC.Value = Empty Then
                MsgBox "Il manque un ou plusieurs identifiants pour générer le cashback", vbCritical + vbOklonly, "Erreur"
                Exit Sub
            End If
        Next
ElseIf ((plageC > plageA) And (plageC > plageB)) Then
    Range("B2:B" & [C1048576].End(xlUp).Row).Select
        For Each emptyCellC In Selection
            If emptyCellC.Value = Empty Then
                MsgBox "Il manque un ou plusieurs montants pour générer le cashback", vbCritical + vbOKOnly, "Erreur"
                Exit Sub
            End If
        Next
ElseIf ((plageA = plageB) Or (plageB = plageC)) Then
    If plageA > plageB Then
        Range("A2:A" & [A1048576].End(xlUp).Row).Select
            For Each emptyCellA In Selection
                If emptyCellA.Value = Empty Then
                    MsgBox "Il manque un ou plusieurs numéros tiers pour générer le cashback", vbCritical + vbOKOnly, "Erreur"
                    Exit Sub
                End If
            Next
        Range("B2:B" & [A1048576].End(xlUp).Row).Select
            For Each emptyCellB In Selection
                If emptyCellB.Value = Empty Then
                    MsgBox "Il manque un ou plusieurs montant pour générer le cashback", vbCritical + vbOKOnly, "Erreur"
                    Exit Sub
                End If
            Next
    ElseIf plageB > plageA Or plageC > plageA Then
        Range("A2:A" & [B1048576].End(xlUp).Row).Select
            For Each emptyCellA In Selection
                If emptyCellA.Value = Empty Then
                    emptyCellA.Select
                    If ActiveCell.Offset(0, 2).Value = Empty And ActiveCell.Offset(0, 1) <> Empty Then
                        MsgBox "Il manque un numéro tiers ou un identifiant pour générer le cashback", vbCritical + vbOKOnly, "Erreur"
                        Exit Sub
                    ElseIf ActiveCell.Offset(0, 2).Value <> Empty Then
                        ActiveCell.Offset(0, 2).Select
                        e = ActiveCell.Address()
                        Range(e & ":C" & f).Select
                        For Each emptyCellC In Selection
                            If emptyCellC.Value = Empty Then
                                MsgBox "Il manque un ou plusieurs identifiants pour générer le cashback", vbCritical + vbOklonly, "Erreur"
                                Exit Sub
                            End If
                        Next
                    End If
                End If
            Next
        Range("B2:B" & [C1048576].End(xlUp).Row).Select
            For Each emptyCellB In Selection
                If emptyCellB.Value = Empty Then
                    MsgBox "Il manque un ou plusieurs montant pour générer le cashback", vbCritical + vbOKOnly, "Erreur"
                    Exit Sub
                End If
            Next
    ElseIf plageC > plageA Or plageC > plageB Then
        Range("C2:C" & [C1048576].End(xlUp).Row).Select
            For Each emptyCellC In Selection
                If emptyCellC.Value = Empty Then
                    MsgBox "Il manque un ou plusieurs numéros tiers pour générer le cashback", vbCritical + vbOKOnly, "Erreur"
                    Exit Sub
                End If
            Next
        Range("B2:B" & [C1048576].End(xlUp).Row).Select
            For Each emptyCellB In Selection
                If emptyCellB.Value = Empty Then
                    MsgBox "Il manque un ou plusieurs montant pour générer le cashback", vbCritical + vbOKOnly, "Erreur"
                    Exit Sub
                End If
            Next
    ElseIf plageA = plageB And plageB = plageC Then
        Range("A2:A" & [A1048576].End(xlUp).Row).Select
            For Each emptyCellA In Selection
                If emptyCellA.Value = Empty Then
                    MsgBox "Il manque un ou plusieurs numéros tiers pour générer le cashback", vbCritical + vbOKOnly, "Erreur"
                    Exit Sub
                End If
            Next
        Range("B2:B" & [B1048576].End(xlUp).Row).Select
            For Each emptyCellB In Selection
                If emptyCellB.Value = Empty Then
                    MsgBox "Il manque un ou plusieurs montant pour générer le cashback", vbCritical + vbOKOnly, "Erreur"
                    Exit Sub
                End If
            Next
        Range("C2:C" & [C1048576].End(xlUp).Row).Select
            For Each emptyCellC In Selection
                If emptyCellC.Value = Empty Then
                    MsgBox "Il manque un ou plusieurs numéros tiers pour générer le cashback", vbCritical + vbOKOnly, "Erreur"
                    Exit Sub
                End If
            Next
    ElseIf plageA = plageB Or plageB = plageC Then
        If plageA = plageB Then
        Range("A2:A" & [A1048576].End(xlUp).Row).Select
            For Each emptyCellA In Selection
                If emptyCellA.Value = Empty Then
                    MsgBox "Il manque un ou plusieurs numéros tiers pour générer le cashback", vbCritical + vbOKOnly, "Erreur"
                    Exit Sub
                End If
            Next
        Range("B2:B" & [B1048576].End(xlUp).Row).Select
            For Each emptyCellB In Selection
                If emptyCellB.Value = Empty Then
                    MsgBox "Il manque un ou plusieurs montant pour générer le cashback", vbCritical + vbOKOnly, "Erreur"
                    Exit Sub
                End If
            Next
        ElseIf plageB = plageC Then
                Range("B2:B" & [B1048576].End(xlUp).Row).Select
            For Each emptyCellB In Selection
                If emptyCellB.Value = Empty Then
                    MsgBox "Il manque un ou plusieurs montant pour générer le cashback", vbCritical + vbOKOnly, "Erreur"
                    Exit Sub
                End If
            Next
        Range("C2:C" & [C1048576].End(xlUp).Row).Select
            For Each emptyCellC In Selection
                If emptyCellC.Value = Empty Then
                    MsgBox "Il manque un ou plusieurs numéros tiers pour générer le cashback", vbCritical + vbOKOnly, "Erreur"
                    Exit Sub
                End If
            Next
        End If
    End If
Else
End If

With Sheets("CashbackGenerator")
    If .AutoFilterMode Then
        .Cells.AutoFilter
    End If
End With

Sheets("CashbackGenerator").Select
Range("C2:C" & [C1048576].End(xlUp).Row).Select

For Each idEmpty In Selection

    If idEmpty.Value = Empty Or idEmpty.Value = "Introuvable" Then
    
        For o = 2 To Sheets("CashbackGenerator").Range("A1048576").End(xlUp).Row
        
            Set C = Sheets("ACC_CLIENT_PORTEUR").Columns(12).Find(Sheets("CashbackGenerator").Range("A" & o), LookIn:=xlValues, lookat:=xlWhole)
            Set d = Sheets("ACC_CLIENT_PORTEUR").Columns(13).Find(Sheets("CashbackGenerator").Range("A" & o), LookIn:=xlValues, lookat:=xlWhole)
            Set e = Sheets("CashbackGenerator").Columns(1).Find(Sheets("CashbackGenerator").Range("A" & o), LookIn:=xlValues, lookat:=xlWhole)
        
                If Sheets("CashbackGenerator").Range("C" & o).Value = "Introuvable" Or Sheets("CashbackGenerator").Range("C" & o).Value = Empty Then
    
                    If Not C Is Nothing Then
                        Sheets("CashbackGenerator").Range("C" & o) = Sheets("ACC_CLIENT_PORTEUR").Range("A" & C.Row)
                        Sheets("Feuil1").Range("A" & o) = Sheets("ACC_CLIENT_PORTEUR").Range("A" & C.Row) & ";" & Sheets("CashbackGenerator").Range("B" & e.Row) * 100 & ";" & DateSerial(Year(Date), Month(Date) + 4, 1) - 1 & " 00:00:00"
                    ElseIf Not d Is Nothing Then
                        Sheets("CashbackGenerator").Range("C" & o) = Sheets("ACC_CLIENT_PORTEUR").Range("A" & d.Row)
                        Sheets("Feuil1").Range("A" & o) = Sheets("ACC_CLIENT_PORTEUR").Range("A" & d.Row) & ";" & Sheets("CashbackGenerator").Range("B" & e.Row) * 100 & ";" & DateSerial(Year(Date), Month(Date) + 4, 1) - 1 & " 00:00:00"
                    Else
                        Sheets("CashbackGenerator").Range("C" & o) = "Introuvable"
                        Sheets("Feuil1").Range("A" & o) = "Introuvable"
                    End If
            
                ElseIf Sheets("CashbackGenerator").Range("C" & o).Value <> "Introuvable" Then
     
                    If Not C Is Nothing Then
                        Sheets("CashbackGenerator").Range("C" & o) = Sheets("ACC_CLIENT_PORTEUR").Range("A" & C.Row)
                        Sheets("Feuil1").Range("A" & o) = Sheets("ACC_CLIENT_PORTEUR").Range("A" & C.Row) & ";" & Sheets("CashbackGenerator").Range("B" & e.Row) * 100 & ";" & DateSerial(Year(Date), Month(Date) + 4, 1) - 1 & " 00:00:00"
                    ElseIf Not d Is Nothing Then
                        Sheets("CashbackGenerator").Range("C" & o) = Sheets("ACC_CLIENT_PORTEUR").Range("A" & d.Row)
                        Sheets("Feuil1").Range("A" & o) = Sheets("ACC_CLIENT_PORTEUR").Range("A" & d.Row) & ";" & Sheets("CashbackGenerator").Range("B" & e.Row) * 100 & ";" & DateSerial(Year(Date), Month(Date) + 4, 1) - 1 & " 00:00:00"
                    Else
                        Sheets("CashbackGenerator").Range("C" & o) = "Introuvable"
                        Sheets("Feuil1").Range("A" & o) = "Introuvable"
                    End If
                End If
        Next o
        
    ElseIf idEmpty.Value <> Empty Then
        For q = 2 To Sheets("CashbackGenerator").Range("C1048576").End(xlUp).Row
            If Sheets("CashbackGenerator").Range("C" & q).Value = "Introuvable" Then
                MsgBox "Certains identifiants sont introuvables", vbCritical + vbOKOnly, "Erreur"
                Exit Sub
            ElseIf Sheets("CashbackGenerator").Range("C" & q).Value <> "Introuvable" And Sheets("CashbackGenerator").Range("C" & q).Value <> "" Then
                Sheets("Feuil1").Range("A" & q) = Sheets("CashbackGenerator").Range("C" & q) & ";" & Sheets("CashbackGenerator").Range("B" & q) * 100 & ";" & DateSerial(Year(Date), Month(Date) + 4, 1) - 1 & " 00:00:00"
            Else
            End If
        Next q
    
    Else
    End If
Next

Sheets("Feuil1").Select
Range("A2:A" & [A1048576].End(xlUp).Row).Select
Set targetI = Columns(1).Find(what:="Introuvable", LookIn:=xlValues, lookat:=xlWhole)
    If Not targetI Is Nothing Then
        Worksheets("CashbackGenerator").Range("A2:C1048576").Borders.LineStyle = xlNone
        Sheets("CashbackGenerator").Select
        Range("C2:C" & [C1048576].End(xlUp).Row).Select
        For Each a In Selection
            If a.Value <> "Introuvable" Then a.EntireRow.Hidden = True
        Next
        MsgBox "Certains identifiants sont introuvables", vbCritical + vbOKOnly, "Erreur"
        Exit Sub
    Else
        If Dir("C:\Users\" & Environ("username") & "\Desktop\" & nomFic & ".txt") = "" Then
            ChDir "C:\Users\" & Environ("username") & "\Desktop"
            ActiveWorkbook.SaveAs ("C:\Users\" & Environ("username") & "\Desktop\" & nomFic & ".txt"), FileFormat:= _
            xlText, CreateBackup:=False
            MsgBox "Le fichier Cashback_CUP_" & maDate & ".txt vient d'être crée sur le Bureau", vbInformation + vbOKOnly, "Création du fichier"
            Application.DisplayAlerts = False
            Worksheets("CashbackGenerator").Range("A2:C1048576").ClearContents
            Worksheets("CashbackGenerator").Range("A2:C1048576").Borders.LineStyle = xlNone
            Sheets(Sheets.Count).Delete
            ActiveWorkbook.SaveAs Filename:= _
            "\\uf96-001.cm-cic.fr\BCA_DPOI\02-SIT\Temporaire\CashbackGenerator.xlsm", _
            FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
            Application.DisplayAlerts = True
            MsgBox "Le fichier CashbackGenerator.xlsm a été sauvegardé", vbInformation + vbOKOnly, "Sauvegarde du fichier"
            
        Else
            Msg = ("Le fichier Cashback_CUP_" & maDate & ".txt  existe déjà !" & vbCrLf & vbCrLf & "Voulez-vous le remplacer ?")
            Style = vbYesNo + vbInformation
            Title = "Demande de confirmation"
            Response = MsgBox(Msg, Style, Title)
            If Response = vbYes Then
                Application.DisplayAlerts = False
                ActiveWorkbook.SaveAs ("C:\Users\" & Environ("username") & "\Desktop\" & nomFic & ".txt"), FileFormat:= _
                xlText, CreateBackup:=False
                Worksheets("CashbackGenerator").Range("A2:C1048576").ClearContents
                Worksheets("CashbackGenerator").Range("A2:C1048576").Borders.LineStyle = xlNone
                Sheets(Sheets.Count).Delete
                ActiveWorkbook.SaveAs Filename:= _
                "\\uf96-001.cm-cic.fr\BCA_DPOI\02-SIT\Temporaire\CashbackGenerator.xlsm", _
                FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
                Application.DisplayAlerts = True
                MsgBox "Le fichier Cashback_CUP_" & maDate & ".txt vient d'être crée sur le Bureau", vbInformation + vbOKOnly, "Création du fichier"
                MsgBox "Le fichier CashbackGenerator.xlsm a été sauvegardé", vbInformation + vbOKOnly, "Sauvegarde du fichier"
            ElseIf Response = vbNo Then
                MsgBox "Le fichier Cashback_CUP_" & maDate & ".txt n'a pas été crée", vbInformation + vbOKOnly, "Création du fichier"
                If Dir("\\uf96-001.cm-cic.fr\BCA_DPOI\02-SIT\Temporaire\CashbackGenerator.xlsm") = "" Then
                    Worksheets("CashbackGenerator").Range("A2:C1048576").ClearContents
                    Worksheets("CashbackGenerator").Range("A2:C1048576").Borders.LineStyle = xlNone
                    Sheets(Sheets.Count).Delete
                    ActiveWorkbook.SaveAs Filename:= _
                    "\\uf96-001.cm-cic.fr\BCA_DPOI\02-SIT\Temporaire\CashbackGenerator.xlsm", _
                    FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
                    MsgBox "Le fichier CashbackGenerator.xlsm a bien été sauvegardé", vbInformation + vbOKOnly, "Sauvegarde du fichier"
                Else
                    Msg = ("Le fichier CashbackGenerator.xlsm existe déjà !" & vbCrLf & vbCrLf & "Voulez-vous le remplacer ?")
                    Style = vbYesNo + vbInformation
                    Title = "Demande de confirmation"
                    Response = MsgBox(Msg, Style, Title)
                    If Response = vbYes Then
                        Application.DisplayAlerts = False
                        Worksheets("CashbackGenerator").Range("A2:C1048576").ClearContents
                        Worksheets("CashbackGenerator").Range("A2:C1048576").Borders.LineStyle = xlNone
                        Sheets(Sheets.Count).Delete
                        ActiveWorkbook.SaveAs Filename:= _
                        "\\uf96-001.cm-cic.fr\BCA_DPOI\02-SIT\Temporaire\CashbackGenerator.xlsm", _
                        FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
                        Application.DisplayAlerts = True
                        MsgBox "Le fichier CashbackGenerator.xlsm a bien été sauvegardé", vbInformation + vbOKOnly, "Sauvegarde du fichier"
                        Exit Sub
                    Else
                        MsgBox "Le fichier CashbackGenerator.xlsm n'a pas été sauvegardé", vbInformation + vbOKOnly, "Sauvegarde du fichier"
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
    


End Sub



