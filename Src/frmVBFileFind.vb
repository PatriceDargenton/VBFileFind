
' VBFileFind : Recherche de fichiers pour remplacer celle de Windows
' D'après la source :
' File Searcher in C#
' By Manfred Bittersam | 24 Apr 2009
' A freeware file searcher in C#
' https://www.codeproject.com/Articles/35044/File-Searcher-in-C

' Conventions de nommage des variables :
' ------------------------------------
' b pour Boolean (booléen vrai ou faux)
' i pour Integer : % (en VB .Net, l'entier a la capacité du VB6.Long)
' l pour Long : &
' r pour nombre Réel (Single!, Double# ou Decimal : D)
' s pour String : $
' c pour Char ou Byte
' d pour Date
' u pour Unsigned (non signé : entier positif)
' a pour Array (tableau) : ()
' m_ pour variable Membre de la classe ou de la feuille (Form)
'  (mais pas pour les constantes)
' frm pour Form
' cls pour Classe
' mod pour Module
' ...
' ------------------------------------

Imports System.Text ' Pour StringBuilder
Imports System.IO ' Pour DirectoryInfo

' https://www.nuget.org/packages/OSVersionExt
' https://github.com/pruggitorg/detect-windows-version
Imports OSVersionExtension ' Pour OSVersion.GetOSVersion

Public Class frmVBFileFind

    Private Const bUtiliserFiltreExclusion As Boolean = True

#Region "Interface"

    Public m_sCheminDossier$ = ""

    Public Enum enumColonnes
        iColChemin = 0
        iColTaille = 1
        iColTailleTxt = 2
        iColDate = 3
        iColDateAcces = 3
    End Enum

#End Region

#Region "Déclarations"

    Private m_bUtiliserBackgroundWorker As Boolean = False ' 25/09/2016

    Const sCmdLancer$ = "Lancer"
    Const sCmdPause$ = "Pause"
    Const sCmdPoursuivre$ = "Poursuivre"
    Const sCmdStop$ = "Stop"
    Const sTxtRechercheEnCours$ = "Recherche en cours..."
    Const sTxtRechercher$ = "Veuillez saisir un texte à rechercher."

    Private Shared ReadOnly m_oVerrou As New Object ' 25/11/2012

    ' Menu contextuel
    Private Const sMenuCtx_TypeDossier$ = "Directory"
    Private Const sMenuCtx_TypeLecteur$ = "Drive" ' 16/12/2012

    ' Il vaut mieux indiquer VBFileFind devant Rechercher pour rappeler quel logiciel ajoute cette clé
    Private Const sMenuCtx_CleCmdRechercher$ = "VBFileFind.Rechercher"
    Private Const sMenuCtx_CleCmdRechercherDescription$ = "Rechercher avec VBFileFind"

    Private Const dDateNulle As Date = #12:00:00 AM#

    Private m_dTpsDeb As DateTime = Now
    Private m_dTpsPrecedListeFichiers As DateTime = Now
    Private m_dTpsPrecedBarreMsg As DateTime = Now

    Private WithEvents m_msgDelegue As clsMsgDelegue = New clsMsgDelegue

    Private Delegate Sub GestEvAfficherMsg(msg As clsMsgEventArgs)
    Private m_gestAffichage As GestEvAfficherMsg

    Private m_llviQueue As New List(Of ListViewItem) ' Résultats à afficher

    Private m_oVBFF As New clsVBFileFind
    Private m_bRechEnCours As Boolean = False
    Private m_sRaccourciBlocNotesOccurrSuiv$ = ""
    Private m_b1ereOuvertureBlocNotes As Boolean = False

    Private m_iMemTailleRech% = 0
    Private m_bInit As Boolean = False ' frm déjà initialisé ?
    Private m_BlocNotesRechercheDesactivee As Boolean = False ' 22/06/2024

#End Region

#Region "Initialisations"

    ' Note : l'appel à InitialiserFenetre() se trouve dans la fonction New()
    ' cf. frmVBFileFind.Designer.vb
    Private Sub InitialiserFenetre()

        ' Reprendre la taille et la position précédente de la fenêtre

        ' Positionnement de la fenêtre par le code : mode manuel
        Me.StartPosition = FormStartPosition.Manual
        If bDebug Then Me.StartPosition = FormStartPosition.CenterScreen ' 25/10/2015

        ' Fixer la position et la taille de la feuille sauvées dans le fichier .exe.config
        Me.Location = My.Settings.frmPosition
        Me.Size = My.Settings.frmTaille

        Me.WindowState = DirectCast(My.Settings.frm_EtatFenetre, FormWindowState)

        ' 16/10/2016
        Me.lvResultats.Columns(enumColonnes.iColChemin).Width = My.Settings.TailleColChemin
        Me.lvResultats.Columns(enumColonnes.iColTaille).Width = 0 'My.Settings.TailleColTaille
        Me.lvResultats.Columns(enumColonnes.iColTailleTxt).Width = My.Settings.TailleColTailleTxt
        Me.lvResultats.Columns(enumColonnes.iColDate).Width = My.Settings.TailleColDate
        Me.lvResultats.Columns(enumColonnes.iColDateAcces).Width = My.Settings.TailleColDateAcces
        m_iMemTailleRech = Me.lvResultats.Width

    End Sub

    Private Sub frmVBFileFind_Load(sender As Object, e As EventArgs) _
        Handles MyBase.Load

        ' 04/05/2014 modUtilFichier peut maintenant être compilé dans une dll
        DefinirTitreApplication(sTitreMsg)

        ' 11/11/2012
        m_gestAffichage = New GestEvAfficherMsg(AddressOf AfficherMsgDirect)
        Me.lblInfo.Text = ""
        AfficherMsg("")

        ' 15/07/2012
        ' Disable automatic sorting to enable manual sorting.
        'Me.lvResultats.Sorting = SortOrder.None
        Me.lvResultats.Columns(enumColonnes.iColChemin).Tag = GetType(String)
        Me.lvResultats.Columns(enumColonnes.iColTaille).Tag = GetType(Long)
        ' Colonne masquée, c'est la colonne texte en octets qui est affichée
        Me.lvResultats.Columns(enumColonnes.iColTaille).Width = 0
        Me.lvResultats.Columns(enumColonnes.iColTailleTxt).Tag = GetType(String)
        Me.lvResultats.Columns(enumColonnes.iColDate).Tag = GetType(Date)
        Me.lvResultats.Columns(enumColonnes.iColDateAcces).Tag = GetType(Date)
        Me.lvResultats.m_iColTriSrc = enumColonnes.iColTailleTxt
        Me.lvResultats.m_iColTriDest = enumColonnes.iColTaille
        Me.lvResultats.DefinirMsgDelegue(m_msgDelegue)

        VerifierMenuCtx()

        Dim sVersion$ = " - V" & sVersionAppli & " (" & sDateVersionAppli & ")"
        Dim sDebug$ = " - Debug"
        Dim sTxt$ = Me.Text & sVersion
        If bDebug Then sTxt &= sDebug
        Me.Text = sTxt

        Me.tbCheminDossier.Text = My.Settings.CheminRecherche
        If Me.m_sCheminDossier.Length > 0 Then
            Me.tbCheminDossier.Text = Me.m_sCheminDossier
        End If

        Me.chkSousDossiers.Checked = My.Settings.bInclureSousDossiers
        Me.tbFiltresFichiers.Text = My.Settings.Filtre
        Me.tbFiltresFichiersExclus.Text = My.Settings.FiltreExclusion ' 25/10/2015

        Me.chkDateMin.Checked = My.Settings.bDateMin
        If My.Settings.DateMin <> dDateNulle Then _
            Me.dtpDateMin.Value = My.Settings.DateMin
        Me.chkDateMax.Checked = My.Settings.bDateMax
        If My.Settings.DateMax <> dDateNulle Then _
            Me.dtpDateMax.Value = My.Settings.DateMax

        ' 20/06/2022
        Me.chkTailleMin.Checked = My.Settings.bTailleMin
        If My.Settings.TailleMin.Length > 0 Then _
            Me.tbTailleMin.Text = My.Settings.TailleMin
        Me.chkTailleMax.Checked = My.Settings.bTailleMax
        If My.Settings.TailleMax.Length > 0 Then _
            Me.tbTailleMax.Text = My.Settings.TailleMax

        Me.chkTexteRech.Checked = My.Settings.bContient
        Me.chkCasse.Checked = My.Settings.bCasse
        Me.tbTexteRech.Text = My.Settings.MotARechercher
        Me.chkBlocNotes.Checked = My.Settings.bOuvrirBlocNotes

        ' 21/06/2024 Le bloc-notes ne peut plus recevoir des envois de touches
        Dim bWindows11 As Boolean = False
        ' Note : on peut mettre tout le code dans une seule classe, mais alors elle fait 31 Ko !
        Dim sVersionOS$ = OSVersion.GetOperatingSystem().ToString()
        If sVersionOS = "Windows11" Then bWindows11 = True
        If bWindows11 Then
            ' C'est seulement la recherche qui est désactivée (via l'envoi de touches clavier)
            m_BlocNotesRechercheDesactivee = True
            Me.toolTip1.SetToolTip(Me.chkBlocNotes,
                "Ouvrir le fichier avec le Bloc-notes (via un double-clic dans la liste trouvée), " &
                "sinon ouvrir avec l'application associée, ou bien ouvrir le dossier dans l'explo" &
                "rateur de fichiers. La recherche du mot trouvée ne marche plus à partir de Windows 11.")
        End If

        Me.pnlTexteRech.Enabled = Me.chkTexteRech.Checked

        Me.rbASCII.Checked = False
        Me.rbUnicode.Checked = False
        Me.rbDbleEncod.Checked = False
        If My.Settings.TypeEncodage = clsVBFileFind.TypeEncodage.ASCII_Ou_Unicode Then _
            Me.rbDbleEncod.Checked = True
        If My.Settings.TypeEncodage = clsVBFileFind.TypeEncodage.ASCII Then _
            Me.rbASCII.Checked = True
        If My.Settings.TypeEncodage = clsVBFileFind.TypeEncodage.Unicode Then _
            Me.rbUnicode.Checked = True
        If My.Settings.TypeEncodage = clsVBFileFind.TypeEncodage.Detecter Then _
            Me.rbDetect.Checked = True

        m_bUtiliserBackgroundWorker = My.Settings.bBackgroundWorker ' 24/09/2016

        ' 24/09/2016 Configuration automatique en anglais
        Dim ci As Globalization.CultureInfo = Globalization.CultureInfo.CurrentCulture
        Dim bAnglais As Boolean = False
        If ci.Name.StartsWith("en-") Then bAnglais = True
        m_sRaccourciBlocNotesOccurrSuiv = My.Settings.RaccourciBlocNotesOccurrSuivAlt_S
        If bAnglais Then
            m_sRaccourciBlocNotesOccurrSuiv = My.Settings.RaccourciBlocNotesOccurrSuivAnglaisAlt_F
        End If

        If bDebug Then
            Me.tbCheminDossier.Text = Application.StartupPath
            'Me.chkTailleMin.Checked = True
            'Me.chkTailleMax.Checked = True
            Me.tbTailleMin.Text = "3000000"
            Me.tbTailleMax.Text = "1000"
            Me.chkSousDossiers.Checked = False
            Me.tbFiltresFichiers.Text = "*.txt"
            'Me.tbFiltresFichiersExclus.Text = ""

            ' Test de recherche de ièce dans Pièces : ce n'est pas possible, car Pièces est encodé
            '  à la fois sur un et deux octets : è est encodé par C3 A8 en UTF-8,
            '  avec ou sans nomenclature (Byte Order Marks) :
            ' -P -i -- -è -c -e -s
            ' 50 69 C3 A8 63 65 73
            ' Note : pour ouvrir l'éditeur Hexadécimal dans Visual Studio Code :
            ' Bouton droit sur l'onglet du fichier ouvert : Reopen Editor With... : Hex Editor (à installer avant)
            Me.tbTexteRech.Text = "Pieces" ' Marche
            'Me.tbTexteRech.Text = "Pièces" ' Ne marche plus !
            'Me.rbDetect.Checked = True

            Me.chkTexteRech.Checked = True
            Me.pnlTexteRech.Enabled = True
            'Me.chkCasse.Checked = False
            'Me.chkDateMin.Checked = True
            'Me.chkDateMax.Checked = False
            'Me.dtpDateMin.Value = #1/1/2014#
        End If

        If Not bUtiliserFiltreExclusion Then
            Me.tbFiltresFichiersExclus.Enabled = False
        End If

        Activation(bActiver:=True) ' 17/09/2022

    End Sub

    Private Sub frmVBFileFind_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        m_bInit = True
    End Sub

    Private Function iLireTypeEncodage%()

        Dim iTypeEncodage% = clsVBFileFind.TypeEncodage.ASCII_Ou_Unicode
        If Me.rbASCII.Checked Then iTypeEncodage = clsVBFileFind.TypeEncodage.ASCII
        If Me.rbUnicode.Checked Then iTypeEncodage = clsVBFileFind.TypeEncodage.Unicode
        If Me.rbDetect.Checked Then iTypeEncodage = clsVBFileFind.TypeEncodage.Detecter
        iLireTypeEncodage = iTypeEncodage

    End Function

    Private Sub frmVBFileFind_FormClosing(sender As Object,
        e As FormClosingEventArgs) Handles MyBase.FormClosing

        Me.m_oVBFF.cmdStop()

        SauverConfig(Me.Location, Me.Size, Me.WindowState)

    End Sub

    Public Sub SauverConfig(
        positionFen As Point,
        tailleFen As Size,
        Optional fws As Windows.Forms.FormWindowState = FormWindowState.Normal)

        ' Sauver la configuration (emplacement de la fenêtre) dans le fichier .exe.config

        ' Le fichier sera sauvé ici :
        '\Documents and Settings\<utilisateur>\Local Settings\Application Data\
        ' ORS_Production\VBFileFind.exe_Url_xxx...xxx\1.0.x.xxxxx\user.config

        If fws = FormWindowState.Normal Then
            My.Settings.frmPosition = positionFen
            My.Settings.frmTaille = tailleFen
            My.Settings.frm_EtatFenetre = 0
        ElseIf fws = FormWindowState.Minimized Then
            My.Settings.frm_EtatFenetre = 0 ' Remetre normal 1
        ElseIf fws = FormWindowState.Maximized Then
            My.Settings.frm_EtatFenetre = 2
        End If

        My.Settings.CheminRecherche = Me.tbCheminDossier.Text
        My.Settings.bInclureSousDossiers = Me.chkSousDossiers.Checked
        My.Settings.Filtre = Me.tbFiltresFichiers.Text
        My.Settings.FiltreExclusion = Me.tbFiltresFichiersExclus.Text ' 25/10/2015

        My.Settings.bDateMin = Me.chkDateMin.Checked
        My.Settings.DateMin = Me.dtpDateMin.Value
        My.Settings.bDateMax = Me.chkDateMax.Checked
        My.Settings.DateMax = Me.dtpDateMax.Value

        ' 20/06/2022
        My.Settings.bTailleMin = Me.chkTailleMin.Checked
        My.Settings.TailleMin = Me.tbTailleMin.Text
        My.Settings.bTailleMax = Me.chkTailleMax.Checked
        My.Settings.TailleMax = Me.tbTailleMax.Text

        My.Settings.bContient = Me.chkTexteRech.Checked
        My.Settings.bCasse = Me.chkCasse.Checked
        My.Settings.MotARechercher = Me.tbTexteRech.Text
        My.Settings.bOuvrirBlocNotes = Me.chkBlocNotes.Checked

        My.Settings.TypeEncodage = iLireTypeEncodage()

        ' 16/10/2016
        My.Settings.TailleColChemin = Me.lvResultats.Columns(enumColonnes.iColChemin).Width
        My.Settings.TailleColTailleTxt = Me.lvResultats.Columns(enumColonnes.iColTailleTxt).Width
        My.Settings.TailleColDate = Me.lvResultats.Columns(enumColonnes.iColDate).Width
        My.Settings.TailleColDateAcces = Me.lvResultats.Columns(enumColonnes.iColDateAcces).Width

        ' Si l'infrastructure de l'appli. est activée, l'appel peut être automatique
        ' (simple case à cocher)
        My.Settings.Save()

    End Sub

    Private Sub frmVBFileFind_KeyPress(sender As Object,
        e As Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress

        ' Si on presse la touche Entrée, alors envoyer la touche Tabulation
        '  pour passer au controle suivant (frm.KeyPreview = True pour que ça marche)
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            SendKeys.Send("{TAB}")
            e.Handled = True
        End If

    End Sub

#End Region

#Region "Traitements"

    Private Sub chkTexteRech_Click(sender As Object, e As EventArgs) _
        Handles chkTexteRech.Click

        Dim bActif As Boolean = Me.chkTexteRech.Checked
        Me.pnlTexteRech.Enabled = bActif

    End Sub

    Private Sub cmdParcourir_Click(sender As Object, e As EventArgs) _
        Handles cmdParcourir.Click

        Dim dlg As New FolderBrowserDialog
        dlg.SelectedPath = Me.tbCheminDossier.Text
        If dlg.ShowDialog(Me) <> DialogResult.OK Then Exit Sub
        Me.tbCheminDossier.Text = dlg.SelectedPath

    End Sub

    Private Sub cmdLancer_Click(sender As Object, e As EventArgs) _
        Handles cmdLancer.Click

        If Me.m_bRechEnCours Then
            Me.m_oVBFF.bPause = Not Me.m_oVBFF.bPause
            If Me.m_oVBFF.bPause Then
                Me.cmdLancer.Text = sCmdPoursuivre
            Else
                Me.cmdLancer.Text = sCmdPause
            End If
            Me.lvResultats.m_bNePasTrier = Not Me.m_oVBFF.bPause ' 18/11/2012
            If Me.lvResultats.m_bNePasTrier Then Me.lvResultats.DesactiverTri()
            Exit Sub
        End If

        If Me.chkTexteRech.Checked AndAlso Me.tbTexteRech.Text = "" Then
            MessageBox.Show(sTxtRechercher, sTitreMsg, MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            Exit Sub
        End If

        Me.m_bRechEnCours = True ' 16/01/2011 Après le test précédant
        Me.lvResultats.m_bNePasTrier = True ' 25/11/2012
        Me.lvResultats.DesactiverTri()

        Me.m_dTpsDeb = Now()
        Me.m_dTpsPrecedListeFichiers = Now()
        Me.m_dTpsPrecedBarreMsg = Now()
        Dim sMsgREC$ = sTxtRechercheEnCours
        Me.lblInfo.Text = sMsgREC
        If Not My.Settings.bSignalerChaqueFichierBarreEtat Then AfficherMsg(sMsgREC)

        Me.lvResultats.Items.Clear()

        Dim fileNames As String() = Me.tbFiltresFichiers.Text.Split(New Char() {";"c})
        Dim fileNamesExcl As String() = Me.tbFiltresFichiersExclus.Text.Split(New Char() {";"c})
        Dim validFileNames As New List(Of String)
        For Each fileName As String In fileNames
            Dim trimmedFileName As String = fileName.Trim
            If trimmedFileName.Length > 0 Then validFileNames.Add(trimmedFileName)
        Next
        Dim validFileNamesExcl As New List(Of String)
        For Each fileName As String In fileNamesExcl
            Dim trimmedFileName As String = fileName.Trim
            If trimmedFileName.Length > 0 Then validFileNamesExcl.Add(trimmedFileName)
        Next

        Dim iTypeEncodage As clsVBFileFind.TypeEncodage =
            clsVBFileFind.TypeEncodage.ASCII_Ou_Unicode ' Les 2
        If Me.rbASCII.Checked Then iTypeEncodage = clsVBFileFind.TypeEncodage.ASCII
        If Me.rbUnicode.Checked Then iTypeEncodage = clsVBFileFind.TypeEncodage.Unicode
        If Me.rbDetect.Checked Then iTypeEncodage = clsVBFileFind.TypeEncodage.Detecter

        ' 20/06/2022
        Dim lTailleMin& = 0
        Dim lTailleMax& = 0
        If Me.tbTailleMin.Text.Length > 0 Then lTailleMin = lConv(Me.tbTailleMin.Text)
        If Me.tbTailleMax.Text.Length > 0 Then lTailleMax = lConv(Me.tbTailleMax.Text)

        Dim prm As New clsVBFileFind.clsPrm(
            Me.tbCheminDossier.Text, Me.chkSousDossiers.Checked,
            validFileNames, validFileNamesExcl,
            Me.chkDateMin.Checked, Me.dtpDateMin.Value,
            Me.chkDateMax.Checked, Me.dtpDateMax.Value,
            Me.chkTailleMin.Checked, lTailleMin,
            Me.chkTailleMax.Checked, lTailleMax,
            Me.chkTexteRech.Checked, Me.chkCasse.Checked, Me.tbTexteRech.Text,
            iTypeEncodage, m_bUtiliserBackgroundWorker)

        Activation(bActiver:=False)

        ' Initialisation
        Me.m_oVBFF.cmdStart(prm, m_msgDelegue)
        If m_bUtiliserBackgroundWorker Then Me.BackgroundWorker1.WorkerSupportsCancellation = True
        DepilerJob()

    End Sub

    Private Sub cmdStop_Click(sender As Object, e As EventArgs) _
        Handles cmdStop.Click

        Me.m_oVBFF.cmdStop()
        Terminer()

    End Sub

    Private Sub DepilerJob()

        ' Dépiler 1 job = une recherche dans un dossier

AutreJob:
        If Not Me.m_oVBFF.bResteJob() Then Terminer() : Exit Sub

        ' Il faut laisser ce test en cas de réaffichage
        If m_bUtiliserBackgroundWorker AndAlso Me.BackgroundWorker1.IsBusy Then Exit Sub

        Dim sDossier$ = Me.m_oVBFF.sDepilerJob()

        If m_bUtiliserBackgroundWorker Then
            Me.BackgroundWorker1.RunWorkerAsync(sDossier)
        Else
            DepilerJobInterne(sDossier)
            ' Un job est terminé, afficher puis dépiler le job suivant jusqu'à la fin
            AfficherResultats()
            'DepilerJob() ' Appel récursif ? Pas besoin :
            GoTo AutreJob
        End If

    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object,
        e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Dim arg As String = DirectCast(e.Argument, String)
        DepilerJobInterne(arg)
        e.Result = arg

    End Sub

    Private Sub DepilerJobInterne(sChemin$)

        Dim dirInfo As New DirectoryInfo(sChemin)
        ' 01/10/2016 Eviter de lancer une recherche si aucun filtre d'exclusion
        If bUtiliserFiltreExclusion AndAlso Me.tbFiltresFichiersExclus.Text.Length > 0 Then _
            Me.m_oVBFF.ChercherArbo(dirInfo, bExclusion:=True) ' 25/10/2015
        Me.m_oVBFF.ChercherArbo(dirInfo, bExclusion:=False)

    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object,
        e As System.ComponentModel.RunWorkerCompletedEventArgs) _
        Handles BackgroundWorker1.RunWorkerCompleted

        ' Un job est terminé, afficher puis dépiler le job suivant jusqu'à la fin

        AfficherResultats()
        'Dim arg As String = DirectCast(e.Result, String)
        DepilerJob()

    End Sub

    Private Sub Terminer()

        Me.m_oVBFF.CalculerTaillesDossiers()
        AfficherResultats(bFin:=True)
        MAJTailleDossiers()
        Activation(bActiver:=True)
        Me.m_bRechEnCours = False
        Me.lvResultats.m_bNePasTrier = False

    End Sub

    Private Sub AfficherInfo(Optional bFin As Boolean = False)

        Dim dTpsFin As DateTime = Now
        Dim tsDiffTps As TimeSpan = dTpsFin.Subtract(Me.m_dTpsDeb)

        Dim iNbFichiersOuDossiersTrouves% = 0, lNbOctetsLus& = 0, lNbOctetsCompares& = 0,
            lNbOctetsFichiers& = 0, lTailleMoyFichier& = 0, iNbFichiersOuDossiers% = 0,
            iNbDossiersParcourus% = 0, iNbDossiers% = 0
        Me.m_oVBFF.LireInfos(iNbFichiersOuDossiersTrouves, iNbFichiersOuDossiers,
            lNbOctetsLus, lNbOctetsCompares, lNbOctetsFichiers, lTailleMoyFichier,
            iNbDossiersParcourus, iNbDossiers)
        Dim rDebitLecture! = 0, rDebitCompare! = 0

        ' Fichiers ou dossiers : objets ou éléments
        Dim sTrouves$ = iNbFichiersOuDossiersTrouves & "/" & iNbFichiersOuDossiers &
            " élément(s) trouvé(s)"
        Dim sb As New StringBuilder(sTrouves)

        ' 16/10/2016 Inutile d'afficher la taille si on ne fait pas de recherche dans les fichiers
        If Me.chkTexteRech.Checked Then
            If lNbOctetsLus > 0 Then
                sb.Append(", " & sFormaterTailleOctets(lNbOctetsLus) & " lus")
                If tsDiffTps.TotalSeconds > 0 Then
                    rDebitLecture = CSng(lNbOctetsLus / tsDiffTps.TotalSeconds)
                    sb.Append(" (" & sFormaterTailleOctets(CLng(rDebitLecture)) & "/sec.)")
                End If
            End If
            If lNbOctetsCompares > 0 Then
                sb.Append(", " & sFormaterTailleOctets(lNbOctetsCompares) & " comparés")
                If tsDiffTps.TotalSeconds > 0 Then
                    rDebitCompare = CSng(lNbOctetsCompares / tsDiffTps.TotalSeconds)
                    sb.Append(" (" & sFormaterTailleOctets(CLng(rDebitCompare)) & "/sec.)")
                End If
            End If
            If lNbOctetsFichiers > 0 Then
                sb.Append(", " & sFormaterTailleOctets(lNbOctetsFichiers) &
                    " de fichiers parcourus")
                sb.Append(", taille moy.: " & sFormaterTailleOctets(lTailleMoyFichier))
                'If Me.chkTexteRech.Checked Then
                Dim rPC! = CSng(lNbOctetsLus / lNbOctetsFichiers)
                Dim rFreq! = 1 - rPC ' Fréquence du texte recherché
                sb.Append(" (fréq. texte : " & rFreq.ToString("0.0%") & ")")
                'End If
            End If
        End If

        If iNbDossiersParcourus > 0 Then ' 25/09/2016
            sb.Append(", " & iNbDossiersParcourus & " dossiers parcourus / " & iNbDossiers)
        End If
        If Me.m_oVBFF.m_lexErr.Count > 0 Then
            sb.Append(", " & Me.m_oVBFF.m_lexErr.Count & " erreur(s)")
        End If

        ' 02/09/2012 Correction du temps de recherche : + logique !
        Dim iNbHeures% = CInt(Math.Truncate(tsDiffTps.TotalHours))
        Dim iNbMinutes% = CInt(Math.Truncate(tsDiffTps.TotalMinutes - iNbHeures * 60))
        Dim iNbSecondes% = CInt(Math.Truncate(tsDiffTps.TotalSeconds _
            - iNbHeures * 3660 - iNbMinutes * 60))
        Dim rNbMilliSec# = tsDiffTps.TotalMilliseconds _
            - iNbHeures * 3660000 - iNbMinutes * 60000 - iNbSecondes * 1000
        Dim rTot# = rNbMilliSec + iNbSecondes * 1000 + iNbMinutes * 60000 + iNbHeures * 3660000
        Dim sTps$ = ", temps de recherche : " &
            iNbHeures.ToString() & "h " &
            iNbMinutes.ToString() & "' " &
            iNbSecondes.ToString() & "'' " &
            rNbMilliSec.ToString("0") & " msec."

        sb.Append(sTps)
        Dim sTxt$ = sb.ToString
        Me.lblInfo.Text = sTxt

        If Not bFin Then Exit Sub

        AfficherMsg(sTrouves)
        If Not My.Settings.bSignalerChaqueFichierBarreEtat Then AfficherMsg("Recherche terminée.")
        If Not My.Settings.bLogBilan Then Exit Sub

        Dim sCheminFichierLog$ = Application.StartupPath & "\VBFileFind.log"
        Dim sb0 As New StringBuilder
        sb0.AppendLine(Me.Text)
        sb0.AppendLine(Now & " : " & Me.tbCheminDossier.Text)
        sb0.AppendLine("Sous-dossiers = " & Me.chkSousDossiers.Checked)
        If Me.tbFiltresFichiersExclus.Text.Length > 0 Then
            sb0.AppendLine("Exclusions = " & Me.tbFiltresFichiersExclus.Text)
        End If
        If Me.chkTexteRech.Checked Then
            sb0.AppendLine("Recherche = " & Me.tbTexteRech.Text)
            sb0.AppendLine("Casse = " & Me.chkCasse.Checked &
                ", Type encodage = " & iLireTypeEncodage())
        End If
        sb0.Append(sTxt)

        sb0.AppendLine("")
        For Each ex As Exception In Me.m_oVBFF.m_lexErr
            sb0.AppendLine(ex.Message)
        Next

        If Me.tbFiltresFichiersExclus.Text.Length > 0 Then
            sb0.AppendLine("")
            sb0.AppendLine("Dossiers exclus : " & m_oVBFF.m_hsExclus.Count)
            For Each sChemin As String In m_oVBFF.m_hsExclus
                sb0.AppendLine(sChemin)
            Next
        End If

        sb0.AppendLine("")
        sb0.AppendLine("Dossiers : " & m_oVBFF.m_hsDossiers.Count)
        For Each sChemin As String In m_oVBFF.m_hsDossiers
            sb0.AppendLine(sChemin)
        Next

        sb0.AppendLine("")
        sb0.AppendLine("Dossiers trouvés : " & m_oVBFF.m_hsDossiersTrouves.Count)
        For Each sChemin As String In m_oVBFF.m_hsDossiersTrouves
            sb0.AppendLine(sChemin)
        Next

        sb0.AppendLine("")
        sb0.AppendLine("Eléments trouvés : " & m_oVBFF.m_hsElementsTrouves.Count)
        For Each sChemin As String In m_oVBFF.m_hsElementsTrouves
            sb0.AppendLine(sChemin)
        Next

        bEcrireFichier(sCheminFichierLog, sb0)

    End Sub

    Private Sub Activation(bActiver As Boolean)

        If Not bActiver Then Me.m_msgDelegue.m_bAnnuler = False ' 19/03/2017
        Me.cmdStop.Enabled = Not bActiver

        If bActiver Then
            Me.cmdLancer.Text = sCmdLancer
        Else
            Me.cmdLancer.Text = sCmdPause
        End If

        Me.tbCheminDossier.Enabled = bActiver
        Me.cmdParcourir.Enabled = bActiver
        Me.chkSousDossiers.Enabled = bActiver
        Me.tbFiltresFichiers.Enabled = bActiver

        If bUtiliserFiltreExclusion Then
            Me.tbFiltresFichiersExclus.Enabled = bActiver
        Else
            Me.tbFiltresFichiersExclus.Enabled = False
        End If

        Me.chkDateMin.Enabled = bActiver
        Me.dtpDateMin.Enabled = bActiver AndAlso Me.chkDateMin.Checked
        Me.chkDateMax.Enabled = bActiver
        Me.dtpDateMax.Enabled = bActiver AndAlso Me.chkDateMax.Checked

        ' 17/09/2022
        Me.chkTailleMin.Enabled = bActiver
        Me.tbTailleMin.Enabled = bActiver AndAlso Me.chkTailleMin.Checked
        Me.chkTailleMax.Enabled = bActiver
        Me.tbTailleMax.Enabled = bActiver AndAlso Me.chkTailleMax.Checked

        Me.chkTexteRech.Enabled = bActiver
        Me.tbTexteRech.Enabled = bActiver
        Me.rbASCII.Enabled = bActiver
        Me.rbUnicode.Enabled = bActiver
        Me.rbDbleEncod.Enabled = bActiver
        Me.chkCasse.Enabled = bActiver
        Me.chkBlocNotes.Enabled = bActiver

    End Sub

#End Region

#Region "Affichage des résultats"

    Private Sub AfficherFSIEv(sender As Object, e As clsFSIEventArgs) _
        Handles m_msgDelegue.EvAfficherFSIEnCours

        AjouterElement(e.fsi)

    End Sub

    Private Sub AfficherMessage(sender As Object, e As clsMsgEventArgs) _
        Handles m_msgDelegue.EvAfficherMessage
        Me.tsslblBarreMessage.Text = e.sMessage
    End Sub

    Private Sub GestSablier(sender As Object, e As clsSablierEventArgs) _
        Handles m_msgDelegue.EvSablier

        Sablier(e.bDesactiver)
        Me.Enabled = e.bDesactiver

        ' Test sur la possibilité d'annuler un tri : mais le pb c'est qu'on suspend 
        '  l'affichage pour aller plus vite justement
        'If Not e.bDesactiver Then
        '    Me.cmdLancer.Enabled = False
        '    Me.cmdStop.Enabled = True
        'Else
        '    Me.cmdLancer.Enabled = True
        '    Me.cmdStop.Enabled = False
        'End If

    End Sub

    Private Sub Sablier(Optional bDesactiver As Boolean = False)

        ' Me.Cursor : Curseur de la fenêtre
        ' Cursor.Current : Curseur de l'application

        If bDesactiver Then
            Me.Cursor = Cursors.Default
        Else
            Me.Cursor = Cursors.WaitCursor
        End If

        ' 19/03/2017 Ne rien faire de plus, sinon cela annule le sablier !

        ' Curseur de l'application : il est réinitialisé à chaque Application.DoEvents
        '  ou bien lorsque l'application ne fait rien
        '  du coup, il faut insister grave pour conserver le contrôle du curseur tout en 
        '  voulant afficher des messages de progression et vérifier les interruptions...
        'Dim ctrl As Control
        'For Each ctrl In Me.Controls
        '    ctrl.Cursor = Me.Cursor ' Curseur de chaque contrôle de la feuille
        'Next ctrl
        'Cursor.Current = Me.Cursor

        'TraiterMsgSysteme_DoEvents()

    End Sub

    Private Sub AfficherMsg(sTxt$)

        If Not m_bUtiliserBackgroundWorker Then
            Me.tsslblBarreMessage.Text = sTxt
        Else
            ' 11/12/2012 Faire un appel indirect pour éviter l'erreur d'appel depuis un autre thread
            ' (l'erreur survient lorsque l'on redimensionne la fenêtre pendant une recherche)
            Dim e As New clsMsgEventArgs(sTxt)
            Dim args() As Object = {e}
            MyBase.Invoke(m_gestAffichage, args)
        End If

    End Sub

    Private Sub AfficherMsgDirect(msg As clsMsgEventArgs)

        Me.tsslblBarreMessage.Text = msg.sMessage

    End Sub

    Private Sub AjouterElement(fsi As IO.FileSystemInfo)

        Dim lvi As New ListViewItem
        lvi.Text = fsi.FullName

        Dim lvsiTailleTxt As New ListViewItem.ListViewSubItem
        Dim lvsiTailleL As New ListViewItem.ListViewSubItem
        If TypeOf fsi Is IO.FileInfo Then
            Dim fi As IO.FileInfo = DirectCast(fsi, IO.FileInfo)
            Dim lLong& = fi.Length
            ' Colonne masquée avec la taille exacte de chaque fichier, pour le tri
            lvsiTailleTxt.Text = lLong.ToString
            lvsiTailleL.Text = sFormaterTailleKOctets(lLong, bSupprimerPt0:=True)
        Else
            lvsiTailleTxt.Text = (0L).ToString
            lvsiTailleL.Text = ""
        End If
        lvi.SubItems.Add(lvsiTailleTxt)
        lvi.SubItems.Add(lvsiTailleL)

        Dim lvsiDate As New ListViewItem.ListViewSubItem
        lvsiDate.Text = (fsi.LastWriteTime.ToShortDateString & " " &
            fsi.LastWriteTime.ToShortTimeString)
        lvi.SubItems.Add(lvsiDate)

        Dim lvsiDateAcces As New ListViewItem.ListViewSubItem
        lvsiDateAcces.Text = (fsi.LastAccessTime.ToShortDateString & " " &
            fsi.LastAccessTime.ToShortTimeString)
        lvi.SubItems.Add(lvsiDateAcces)

        lvi.ToolTipText = fsi.FullName

        ' Colorer les dossiers
        If TypeOf fsi Is IO.DirectoryInfo Then lvi.BackColor = Color.LightYellow

        ' Ajouter l'élément dans une file d'attente
        If m_bUtiliserBackgroundWorker Then
            SyncLock m_oVerrou ' 25/11/2012 On ne peut pas énumérer la liste pendant un ajout
                Me.m_llviQueue.Add(lvi)
            End SyncLock
        Else
            Me.m_llviQueue.Add(lvi)
        End If

        If Not My.Settings.bSignalerChaqueFichierBarreEtat Then Exit Sub
        ' Afficher les éléments de la file d'attente à chaque 10è de sec. écoulée
        Dim dTpsFin As DateTime = Now
        Dim tsDiffTpsPreced As TimeSpan = dTpsFin.Subtract(Me.m_dTpsPrecedBarreMsg)
        If tsDiffTpsPreced.TotalSeconds < My.Settings.DelaiAffichageBarreMsgSec Then Exit Sub
        Me.m_dTpsPrecedBarreMsg = dTpsFin
        'If My.Settings.bSignalerChaqueFichierBarreEtat Then AfficherMsg(lvi.Text)
        AfficherMsg(lvi.Text)

    End Sub

    Private Sub AfficherResultats(Optional bFin As Boolean = False)

        ' Traiter la file d'attente : afficher tous les résultats qu'elle contient

        ' Le plus svt possible afficher le fichier en cours de traitement
        '  pour donner l'impression de vitesse
        '  (mais par contre, pour la liste, pas besoin d'aller aussi vite)
        'For Each lvi0 As ListViewItem In Me.m_llviQueue
        '    AfficherMsg(lvi0.Text)
        '    Exit For
        'Next

        ' Afficher les éléments de la file d'attente à chaque 1/4 sec. écoulée
        If Not bFin Then
            Dim dTpsFin As DateTime = Now
            Dim tsDiffTpsPreced As TimeSpan = dTpsFin.Subtract(Me.m_dTpsPrecedListeFichiers)
            If tsDiffTpsPreced.TotalSeconds < My.Settings.DelaiAffichageListeFichiersSec Then Exit Sub
            Me.m_dTpsPrecedListeFichiers = dTpsFin
        End If

        Me.lvResultats.SuspendLayout()
        Me.SuspendLayout() ' En dernier
        Me.lvResultats.BeginUpdate() ' 21/07/2012

        If m_bUtiliserBackgroundWorker Then
            SyncLock m_oVerrou ' 25/11/2012 On ne peut pas énumérer la liste pendant un ajout
                Dim iNbItems% = Me.m_llviQueue.Count
                For Each lvi0 As ListViewItem In Me.m_llviQueue
                    ' Style possible pour chaque sous-item (pour colorier la colonne triée)
                    lvi0.UseItemStyleForSubItems = False
                    Me.lvResultats.Items.Add(lvi0)
                Next
            End SyncLock
        Else
            Dim iNbItems% = Me.m_llviQueue.Count
            For Each lvi0 As ListViewItem In Me.m_llviQueue
                ' Style possible pour chaque sous-item (pour colorier la colonne triée)
                lvi0.UseItemStyleForSubItems = False
                Me.lvResultats.Items.Add(lvi0)
            Next
        End If

        If bFin Then
            For Each ex As Exception In Me.m_oVBFF.m_lexErr
                Dim lvi As New ListViewItem
                lvi.UseItemStyleForSubItems = False
                lvi.Text = ex.Message
                lvi.BackColor = Color.Orange
                Me.lvResultats.Items.Add(lvi)
            Next
            If Not m_oVBFF.m_bSucces Then
                Dim lvi As New ListViewItem
                lvi.UseItemStyleForSubItems = False
                lvi.Text = m_oVBFF.m_sMsgErr
                lvi.BackColor = Color.Orange
                Me.lvResultats.Items.Add(lvi)
            End If
        End If

        Me.lvResultats.EndUpdate() ' 21/07/2012
        Me.lvResultats.ResumeLayout()

        AfficherInfo(bFin)

        Me.ResumeLayout() ' En dernier aussi

        Me.m_llviQueue = New List(Of ListViewItem)

    End Sub

    Private Sub MAJTailleDossiers()

        Me.lvResultats.SuspendLayout()
        Me.SuspendLayout() ' En dernier

        Me.lvResultats.BeginUpdate()

        If m_bUtiliserBackgroundWorker Then
            SyncLock m_oVerrou ' On ne peut pas énumérer la liste pendant un ajout
                For Each lvi0 As ListViewItem In Me.lvResultats.Items
                    Dim sCle$ = lvi0.Text
                    If Me.m_oVBFF.m_dicoTaillesDossiers.ContainsKey(sCle) Then
                        Dim lTailleSousDossier& = Me.m_oVBFF.m_dicoTaillesDossiers(sCle)
                        lvi0.SubItems(enumColonnes.iColTaille).Text = lTailleSousDossier.ToString
                        lvi0.SubItems(enumColonnes.iColTailleTxt).Text = sFormaterTailleKOctets(lTailleSousDossier, bSupprimerPt0:=True)
                    End If
                Next
            End SyncLock
        Else
            For Each lvi0 As ListViewItem In Me.lvResultats.Items
                Dim sCle$ = lvi0.Text
                If Me.m_oVBFF.m_dicoTaillesDossiers.ContainsKey(sCle) Then
                    Dim lTailleSousDossier& = Me.m_oVBFF.m_dicoTaillesDossiers(sCle)
                    lvi0.SubItems(enumColonnes.iColTaille).Text = lTailleSousDossier.ToString
                    lvi0.SubItems(enumColonnes.iColTailleTxt).Text = sFormaterTailleKOctets(lTailleSousDossier, bSupprimerPt0:=True)
                End If
            Next
        End If

        Me.lvResultats.EndUpdate()

        Me.lvResultats.ResumeLayout()
        Me.ResumeLayout() ' En dernier aussi

    End Sub

#End Region

#Region "Gestion des événements"

    Private Sub lvResultats_DoubleClick(sender As Object, e As EventArgs) _
        Handles lvResultats.DoubleClick

        ' Ouvrir le fichier en question, soit par l'application associée
        '  soit via le bloc-notes (avec recherche de l'occurrence via SendKeys)

        If Me.lvResultats.SelectedItems.Count = 0 Then Exit Sub

        Dim sChemin As String = Me.lvResultats.SelectedItems.Item(0).Text
        If IO.Directory.Exists(sChemin) Then
            ' C'est un dossier : l'ouvrir
            OuvrirDossier(sChemin)
            Exit Sub
        End If
        If Not IO.File.Exists(sChemin) Then Exit Sub

        If Not Me.chkBlocNotes.Checked Then OuvrirAppliAssociee(sChemin) : Exit Sub

        Dim sTxt$ = ""
        If Me.chkTexteRech.Checked Then sTxt = Me.tbTexteRech.Text
        Dim bUnicode As Boolean = False
        Dim sExt$ = IO.Path.GetExtension(sChemin)
        ' Si on choisi les 2 types d'encodage, on ne peut pas savoir à priori le type
        ' Si le fichier est txt, alors même en unicode, il s'affiche normalement dans le bloc-notes
        ' (pas besoin d'ajouter des espaces après chaque caractère)
        ' Si on consulte un fichier binaire avec une occurrence trouvée en unicode
        '  alors ajouter des espaces après chaque caractère
        If Me.rbUnicode.Checked And sExt.ToLower <> ".txt" Then bUnicode = True

        ' 18/08/2016 Configurer selon la langue, en anglais un seul paramètre est différent
        OuvrirBlocNotes(sChemin, sTxt, Me.chkCasse.Checked, bUnicode,
            My.Settings.RaccourciBlocNotesRechercherCtrl_f,
            My.Settings.RaccourciBlocNotesSensibleCasseAlt_c,
            m_sRaccourciBlocNotesOccurrSuiv,
            My.Settings.ExeBlocNotes,
            m_b1ereOuvertureBlocNotes,
            My.Settings.DelaiMSec1OuvertureBlocNotes,
            bOuvrirSeulement:=m_BlocNotesRechercheDesactivee)
        m_b1ereOuvertureBlocNotes = True

    End Sub

    Private Sub lvResultats_Resize(sender As Object, e As EventArgs) Handles lvResultats.Resize
        AjusterColonnesResultats()
    End Sub

    Private Sub AjusterColonnesResultats()

        ' 19/03/2017 Attendre que la fenêtre soit initialisée (sinon sa taille n'est pas connue)
        If Not m_bInit Then Exit Sub

        ' 01/04/2017
        If Me.WindowState = FormWindowState.Minimized Then Exit Sub

        ' Redimensionner la 1ère colonne pour pouvoir voir le chemin complet
        'Me.lvResultats.Columns.Item(0).Width = Me.lvResultats.Width - 240

        ' 16/10/2016
        ' Redimensionner la 1ère colonne avec l'agrandissement de la zone complète
        Me.lvResultats.Columns.Item(0).Width += Me.lvResultats.Width - m_iMemTailleRech
        m_iMemTailleRech = Me.lvResultats.Width

    End Sub

#End Region

#Region "Gestion des menus contextuels"

    Private Sub cmdAjouterMenuCtx_Click(sender As Object,
        e As EventArgs) Handles cmdAjouterMenuCtx.Click
        AjouterMenuCtx()
        VerifierMenuCtx()
    End Sub

    Private Sub cmdEnleverMenuCtx_Click(sender As Object,
        e As EventArgs) Handles cmdEnleverMenuCtx.Click
        EnleverMenuCtx()
        VerifierMenuCtx()
    End Sub

    Private Sub VerifierMenuCtx()

        Dim sCleDescriptionCmd$ = sMenuCtx_TypeDossier & "\" & sDossierShell & "\" &
            sMenuCtx_CleCmdRechercher
        If bCleRegistreCRExiste(sCleDescriptionCmd) Then
            Me.cmdAjouterMenuCtx.Enabled = False
            Me.cmdEnleverMenuCtx.Enabled = True
        Else
            Me.cmdAjouterMenuCtx.Enabled = True
            Me.cmdEnleverMenuCtx.Enabled = False
        End If

    End Sub

    Private Sub AjouterMenuCtx()

        If MsgBoxResult.Cancel = MsgBox("Ajouter le menu contextuel ?",
            MsgBoxStyle.OkCancel Or MsgBoxStyle.Question) Then Exit Sub

        Dim sCheminExe$ = Application.ExecutablePath
        Const bPrompt As Boolean = False
        Const sChemin$ = """%1"""
        If Not bAjouterMenuContextuel(sMenuCtx_TypeDossier, sMenuCtx_CleCmdRechercher,
            bPrompt, , sMenuCtx_CleCmdRechercherDescription, sCheminExe, sChemin) Then Exit Sub
        bAjouterMenuContextuel(sMenuCtx_TypeLecteur, sMenuCtx_CleCmdRechercher,
            bPrompt, , sMenuCtx_CleCmdRechercherDescription, sCheminExe, sChemin)

    End Sub

    Private Sub EnleverMenuCtx()

        If MsgBoxResult.Cancel = MsgBox("Enlever le menu contextuel ?",
            MsgBoxStyle.OkCancel Or MsgBoxStyle.Question) Then Exit Sub

        If Not bAjouterMenuContextuel(sMenuCtx_TypeDossier, sMenuCtx_CleCmdRechercher,
            bEnlever:=True, bPrompt:=False) Then Exit Sub
        bAjouterMenuContextuel(sMenuCtx_TypeLecteur, sMenuCtx_CleCmdRechercher,
            bEnlever:=True, bPrompt:=False)

    End Sub

    Private Sub tbTailleMin_TextChanged(sender As Object, e As EventArgs) _
        Handles tbTailleMin.TextChanged

        Dim lTailleMin& = 0
        If Me.tbTailleMin.Text.Length > 0 Then lTailleMin = lConv(Me.tbTailleMin.Text)
        Dim sTailleMin$ = sFormaterTailleOctets(lTailleMin)
        Me.toolTip1.SetToolTip(Me.tbTailleMin,
            "Taille min. des fichiers : " & sTailleMin)

    End Sub

    Private Sub tbTailleMax_TextChanged(sender As Object, e As EventArgs) _
        Handles tbTailleMax.TextChanged

        Dim lTailleMax& = 0
        If Me.tbTailleMax.Text.Length > 0 Then lTailleMax = lConv(Me.tbTailleMax.Text)
        Dim sTailleMax$ = sFormaterTailleOctets(lTailleMax)
        Me.toolTip1.SetToolTip(Me.tbTailleMax,
            "Taille max. des fichiers : " & sTailleMax)

    End Sub

    ' 17/09/2022
    Private Sub chkTailleMax_Click(sender As Object, e As EventArgs) Handles chkTailleMax.Click
        Me.tbTailleMax.Enabled = Me.chkTailleMax.Checked
    End Sub
    Private Sub chkTailleMin_Click(sender As Object, e As EventArgs) Handles chkTailleMin.Click
        Me.tbTailleMin.Enabled = Me.chkTailleMin.Checked
    End Sub
    Private Sub chkDateMax_Click(sender As Object, e As EventArgs) Handles chkDateMax.Click
        Me.dtpDateMax.Enabled = Me.chkDateMax.Checked
    End Sub
    Private Sub chkDateMin_Click(sender As Object, e As EventArgs) Handles chkDateMin.Click
        Me.dtpDateMin.Enabled = Me.chkDateMin.Checked
    End Sub

#End Region

End Class