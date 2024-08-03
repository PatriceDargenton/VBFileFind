
Module modDepart

#If DEBUG Then
    Public Const bDebug As Boolean = True
    Public Const bRelease As Boolean = False
#Else
        Public Const bDebug As Boolean = False
        Public Const bRelease As Boolean = True
#End If

    Public ReadOnly sNomAppli$ = My.Application.Info.Title
    Public ReadOnly sTitreMsg$ = sNomAppli
    Private Const sDateVersionVBFileFind$ = "03/08/2024"
    Public Const sDateVersionAppli$ = sDateVersionVBFileFind

    Public ReadOnly sVersionAppli$ =
        My.Application.Info.Version.Major & "." &
        My.Application.Info.Version.Minor &
        My.Application.Info.Version.Build

    Public m_sTitreMsg$ = sNomAppli

    Public Sub DefinirTitreApplication(sTitreMsg As String)
        m_sTitreMsg = sTitreMsg
    End Sub

    Public Sub Main()

        ' S'il n'y a aucune gestion d'erreur, on peut déboguer dans l'IDE
        ' Sinon, ce n'est pas pratique de retrouver la ligne du bug :
        '  il faut cocher Thrown dans le menu Debug:Exception... pour les 2 lignes
        '  (dans ce cas, il peut y avoir beaucoup d'interruptions selon la logique 
        '   de programmation : mieux vaut prévenir les erreurs que de les traiter)
        ' C'était plus simple avec On Error Goto X, car on pouvait
        '  désactiver la gestion d'erreur avec une simple constante bTrapErr.
        If bDebug Then Depart() : Exit Sub

        ' Attention : En mode Release il faut un Try Catch ici  
        '  car sinon il n'y a pas de gestion d'erreur !
        ' (.Net renvoie un message d'erreur équivalent 
        '  à un plantage complet sans explication)
        Try
            Depart()
        Catch ex As Exception
            AfficherMsgErreur2(ex, "Main " & sTitreMsg)
        End Try

    End Sub

    Private Sub Depart()

        Dim sArg0$ = Microsoft.VisualBasic.Interaction.Command
        Dim asArgs$() = asArgLigneCmd(sArg0)

        Dim sCheminDossier$ = ""
        If (sArg0.Length <> 0) Then
            Dim iNbArguments% = UBound(asArgs) + 1
            If iNbArguments = 1 Then sCheminDossier = asArgs(0)
        End If

        Dim frm As New frmVBFileFind
        frm.m_sCheminDossier = sCheminDossier
        Application.Run(frm)

    End Sub

End Module