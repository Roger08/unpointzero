Attribute VB_Name = "modConsole"
'                       ##########################################
'                       ##############  FRoG Creator 1.0   #############
'                       ##  Module de gestion du texte affiché sur le serveur  ##
'                       ####### Dernière modification : JJ/MM/AAAA  #######
'                       ##########################################

' -- Affichage des différents types de messages sur le serveur --
Public Sub DispMessage(ByVal msg As String)
    With frmServer.txtReturn
        .SelStart = Len(.text)
        .SelColor = QBColor(Black)
        .SelText = (vbNewLine & ">" & msg)
        .SelStart = Len(.text) - 1
    End With
End Sub

Public Sub DispErreur(ByVal msg As String)
    With frmServer.txtReturn
        .SelStart = Len(.text)
        .SelColor = QBColor(Black)
        .SelText = (vbNewLine & "> ")
        .SelStart = Len(.text)
        .SelColor = QBColor(Red)
        .SelText = ("ERREUR : " & msg)
        .SelStart = Len(.text) - 1
    End With
End Sub

Public Sub DispInfo(ByVal msg As String)
    With frmServer.txtReturn
        .SelStart = Len(.text)
        .SelColor = QBColor(Black)
        .SelText = (vbNewLine & "> ")
        .SelStart = Len(.text)
        .SelColor = QBColor(Cyan)
        .SelText = (msg)
        .SelStart = Len(.text) - 1
    End With
End Sub

' -- Traitement des commandes --
