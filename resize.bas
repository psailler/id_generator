Attribute VB_Name = "resize"
'**********************************************************************************************************************
'*                                     CREATEUR :Patrick toulon Alias chamalin1@msn.com                               *
'*                                                    DATE :23/09/2010                                                *
'*                                       UTILISATION D'UNE SEULE API LE "USER32.DLL"                                  *
'*                                    EXEMPLE DE USERFORM REDIMENTIONNABLE NOUVELLE VERSION                           *
'*                                      LES CONTROLS SONT REDIMENTIONNES EN MEME TEMPS                                *
'*                                               AINSI QUE LES FONT SIZE                                              *
'*                                                                                                                    *
'* REVISION:21:02:2013(Modification de la gestion du font.size)                                                       *
'                                                                                                                     *
'*le font size est géré control par control ,il peuvent donc avoir un fontsize différent                           *
'*                                                                                                                    *
'**********************************************************************************************************************
Option Explicit
Public Declare Function FWA Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SWH Lib "User32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SWLA Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GWLA Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public old_largeur As Long, handle As Long, old_hauteur As Long, newhauteur As Single, newlargeur As Single
Dim Ctl As Object
Dim ctrl As Object

Sub trois_boutons(uf As Object)    'on va ajouter les deux boutons manquants et l'élasticité a l'userform
'*******************************************************************
'*ici on memorise les dimention de depart de l'userform            *
    old_largeur = uf.InsideWidth: old_hauteur = uf.InsideHeight   '*
'*******************************************************************

'*******************************************************************************************************************
'nouvelle version                                                                                                  *
' ici on va memoriser l'operateur corespondant a l'userform/ par le font size de chaques control _                 *
'  sauf ce qui n'ont pas cette propriété                                                                           *
    For Each ctrl In uf.Controls                                                                                  '*
        If TypeName(ctrl) <> "ScrollBar" And TypeName(ctrl) <> "Image" And TypeName(ctrl) <> "SpinButton" Then    '*
            ctrl.Tag = uf.InsideWidth / ctrl.Font.Size                                                            '*
        End If                                                                                                    '*
    Next                                                                                                          '*
    '***************************************************************************************************************

    '***************************************************************************************************************
    ' ici on determine le handle par la classe de frame en testant la version de l'application ( DE EXCEL97 A 2010)*
    handle = FWA("Thunder" & IIf(Application.Version Like "8*", "0*", "D") & "Frame", uf.Caption)                 '*
    ' ici on applique les changement (&h70000= les trois bouton et l'elasticité)                                   *
    SWLA handle, -16, GWLA(handle, -16) Or &H70000                                                                '*
    '***************************************************************************************************************


End Sub
Sub plein_ecran()
' on affiche le userform en plein ecran avec l'api showwindowa de la user32.dll _
  bien moins lourd que mes versions precedente de maximisation de l'userform et plus rapide et plus propre
'1= mode normal
'3 =maximiser
'6 =minimiser
'le handle a été declaré en public au debut du module et  identifié dans la routine des trois boutons il n'est donc plus necessaire de l'identifier
    SWH handle, 3
End Sub

Sub maForm_Resize(usf As Object)
'ici on determine le multiplicateur qui differenci la dimention de base a celle actuelle de l'userform
    newlargeur = usf.InsideWidth / old_largeur: newhauteur = usf.InsideHeight / old_hauteur

    'ici on boucle sur tout les controls
    For Each Ctl In usf.Controls
        'et on applique le multiplicateur au controls pour la largeur et la hauteur en une seule ligne
        Ctl.Move Ctl.Left * newlargeur, Ctl.Top * newhauteur, Ctl.Width * newlargeur, Ctl.Height * newhauteur
        'tout les controls qui ont le multiplicateur enregistré dans leurs tags respectifs verront leur font size redimentionné en proportion
          If Ctl.Tag <> "" Then Ctl.Font.Size = Round(usf.InsideWidth / Ctl.Tag, 0) - 1
    Next
    'ici on indique que l'ancienne largeur devient la nouvelle largeur et pareil pour la hauteur indispensable pour un futur redimentionnement
    old_largeur = usf.InsideWidth: old_hauteur = usf.InsideHeight: usf.Repaint
End Sub






