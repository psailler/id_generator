Attribute VB_Name = "Degradé"
 
 
'methode pour le fill en degradé
'fill.OneColorGradient(Style, Variant,    Degree)
'"""""""""""""""""""""""""""MODULE POUR INSERER UNE PROGRESSBARRE PERSO""""""""""""""""""""""""""""""
'                                                                                                   "
'                                creation patricktoulon                                             "
'                                                                                                   "
'                               Theme : personalisation des applications userform couleur dégradé                           "
'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
 
Option Explicit
 
 
Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(8) As Byte
End Type
 
Public Type PICTDESC
    cbSize As Long
    picType As Long
    hImage As Long
End Type
Public Declare Function OpenClipboard& Lib "User32" (ByVal hwnd As Long)
Public Declare Function EmptyClipboard Lib "User32" () As Long
Public Declare Function GetClipboardData& Lib "User32" (ByVal wFormat%)
Public Declare Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function CloseClipboard& Lib "User32" ()
Public Declare Function CopyImage& Lib "User32" (ByVal handle&, ByVal un1&, ByVal n1&, ByVal n2&, ByVal un2&)
Public Declare Function IIDFromString Lib "ole32" (ByVal lpsz As String, ByRef lpiid As GUID) As Long
Public Declare Function OleCreatePictureIndirect Lib "olepro32" (pPictDesc As PICTDESC, ByRef riid As GUID, ByVal fOwn As Long, ByRef ppvObj As IPicture) As Long
Public maform As Object
Public iPic As IPicture
 
 
Function Fond_dégradé(usf As Object, Optional texture As Variant = "", Optional couleur As Variant = vbBlue)
    Dim image As IPicture
    With ActiveSheet.Shapes.AddShape(1, 10, 15, usf.Width, usf.Height)
        .Name = "fond_usf"
        .Line.Visible = msoFalse
        .Fill.Visible = msoTrue
        If texture <> "" Then .Fill.PresetTextured texture: GoTo suite 'on saute le gradient pour les texture
        .Fill.ForeColor.RGB = couleur
        .Fill.TwoColorGradient Style:=msoGradientDiagonalDown, Variant:=4
 
 
suite:
        'on copie la forme dans le clipboard
        .CopyPicture xlScreen, xlBitmap    'copie la selection dans le clipboard
        ActiveSheet.Shapes("fond_usf").Delete
        'prend l'image dans le cliboard
        Dim hCopy&: OpenClipboard 0&
        hCopy = CopyImage(GetClipboardData(2), 0, 0, 0, &H8)
        CloseClipboard    ' ferme le cliboard
        If hCopy = 0 Then Exit Function    'si il y a rien on sort de la fonction
        Const IPictureIID = "{7BF80981-BF32-101A-8BBB-00AA00300CAB}"
        Dim tIID As GUID, tPICTDEST As PICTDESC, Ret As Long
        Ret = IIDFromString(StrConv(IPictureIID, vbUnicode), tIID)
        If Ret Then Exit Function
        With tPICTDEST
            .cbSize = Len(tPICTDEST)
            .picType = 1
            .hImage = hCopy
        End With
        'on créé le itmap
        Ret = OleCreatePictureIndirect(tPICTDEST, tIID, 1, image)
        If Ret Then Exit Function
        On Error GoTo 0
 
        usf.Picture = image
 
        Set iPic = Nothing
    End With
End Function
