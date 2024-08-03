
Imports System.IO
Imports System.Linq
Imports System.Text ' Pour Encoding

Public Class clsVBFileFind

    Private m_bDebugRecherche As Boolean = False
    Private m_bDebugTailleDossier As Boolean = False
    Private m_bDossierDebug As Boolean = False

#Region "Interface"

    Public m_sMsgErr$ = ""
    Public m_lexErr As List(Of Exception) = New List(Of Exception)
    Public m_bSucces As Boolean = True

    ' Hashset des fichiers ou dossiers déjà trouvés, au cas où les filtres soient redondants
    Public m_hs As New HashSet(Of String)
    Public m_hsDossiers As New HashSet(Of String) ' Pour log : dossiers restants (non exclus)
    Public m_hsExclus As New HashSet(Of String) ' 25/10/2015 Fichiers ou dossiers exclus
    Public m_hsElementsTrouves As New HashSet(Of String) ' Fichiers ou dossiers qui matchent
    Public m_hsDossiersTrouves As New HashSet(Of String) ' Dossiers qui matchent
    Private m_hsbSousDossiers As HashSet(Of String) ' Clé : SousDossier : tous les sous-dossiers
    'Private m_hsbSousDossiers2 As HashSet(Of String) ' Clé : SousDossier : tous les sous-dossiers

    ' Les dossiers exclus ne sont pas comptabilisés (car on compte la taille de chaque fichier
    '  de chaque dossier que l'on doit parcourir, pour mesurer la vitesse de recherche)
    Public m_dicoTaillesDossiers As Dictionary(Of String, Long) ' Clé : Dossier -> lTaille

    Public Enum TypeEncodage
        ASCII_Ou_Unicode = 0
        ASCII = 1
        Unicode = 2
        Detecter = 3 ' 27/07/2024
    End Enum

#End Region

#Region "Déclarations"

    ' Pour pouvoir détecter l'encodage, via l'option detectEncodingFromByteOrderMarks du StreamReader
    '  il faut lire le fichier depuis le StreamReader, et non plus directement depuis le FileStream
    Private Const bLectureDepuisFileStream As Boolean = False ' Ancienne technique

    Private m_msgDelegue As clsMsgDelegue

    'Private m_queue As New Queue(Of String) ' Queue des répertoires à parcourir
    ' 05/03/2017 Queue.Contains bcp + rapide :
    Private m_queue As New HashQueue(Of String) ' Queue des répertoires à parcourir

    Private m_abTxtRech As Byte() = Nothing
    Private m_abTxtRechMin As Byte() = Nothing ' Minuscules
    Private m_abTxtRechUC As Byte() = Nothing ' Unicode
    Private m_abTxtRechMinUC As Byte() = Nothing ' Minuscules
    Private m_prm As clsPrm = Nothing
    Private m_bStop As Boolean = False

    Private m_iNbDossiersEmpiles% ' Tous les dossiers empilés
    Private m_iNbDossiersParcourus% ' Tous les dossiers en train d'être parcourus ou terminés

    ' Rencontrés : avant l'examen des conditions
    Private m_iNbFichiersOuDossiers% ' Tous les fichiers ou dossiers rencontrés
    Private m_iNbFichiers% ' Tous les fichiers rencontrés
    Private m_iNbDossiers% ' Tous les dossiers rencontrés

    ' Trouvés : qui matchent les conditions
    Private m_iNbFichiersOuDossiersTrouves% ' Tous les fichiers ou dossiers qui matchent
    Private m_iNbDossiersTrouves% ' Tous les dossiers qui matchent

    Private m_lNbOctetsLus&
    Private m_lNbOctetsCompares&
    Private m_lNbOctetsFichiers&
    Private m_lTailleMoyFichier&

    Private m_bAlerte As Boolean = False

    Private m_bPause As Boolean = False
    Property bPause() As Boolean
        Get
            bPause = m_bPause
        End Get
        Set(bVal As Boolean)
            m_bPause = bVal
        End Set
    End Property

    Public Class clsPrm

        ' Fields
        Public m_containingChecked As Boolean
        Public m_casseChecked As Boolean
        Public m_containingText$
        Public m_iTypeEncodage As TypeEncodage ' 0, 1, 2 : 0 : les deux, 1 : ASCII, 2 : Unicode

        Public m_fileNames As List(Of String)
        Public m_fileNamesExcl As List(Of String)

        Public m_includeSubDirsChecked As Boolean

        Public m_newerThanChecked As Boolean
        Public m_newerThanDateTime As DateTime
        Public m_olderThanChecked As Boolean
        Public m_olderThanDateTime As DateTime

        ' 20/06/2022
        Public m_greaterThanChecked As Boolean
        Public m_greaterThanSize As Long
        Public m_smallerThanChecked As Boolean
        Public m_smallerThanSize As Long

        Public m_searchDir$
        Public m_bUseThread As Boolean

        ' Methods
        Public Sub New(searchDir$, includeSubDirsChecked As Boolean,
            fileNames As List(Of String),
            fileNamesExcl As List(Of String),
            newerThanChecked As Boolean, newerThanDateTime As DateTime,
            olderThanChecked As Boolean, olderThanDateTime As DateTime,
            greaterThanChecked As Boolean, greaterThanSize&,
            smallerThanChecked As Boolean, smallerThanSize&,
            containingChecked As Boolean,
            casseChecked As Boolean,
            containingText$,
            iTypeEncodage As clsVBFileFind.TypeEncodage,
            bUseThread As Boolean)

            Me.m_bUseThread = bUseThread
            Me.m_searchDir = searchDir
            Me.m_includeSubDirsChecked = includeSubDirsChecked

            Me.m_fileNames = fileNames
            Me.m_fileNamesExcl = fileNamesExcl

            Me.m_newerThanChecked = newerThanChecked
            Me.m_newerThanDateTime = newerThanDateTime
            Me.m_olderThanChecked = olderThanChecked
            Me.m_olderThanDateTime = olderThanDateTime

            ' 20/06/2022
            Me.m_greaterThanChecked = greaterThanChecked
            Me.m_greaterThanSize = greaterThanSize
            Me.m_smallerThanChecked = smallerThanChecked
            Me.m_smallerThanSize = smallerThanSize

            Me.m_containingChecked = containingChecked
            Me.m_casseChecked = casseChecked
            Me.m_containingText = containingText
            Me.m_iTypeEncodage = iTypeEncodage

        End Sub

    End Class

#End Region

#Region "Interface : procédure"

    Public Sub cmdStart(prm As clsPrm, msgDelegue As clsMsgDelegue)

        Me.ResetVariables()
        Me.m_prm = prm
        Me.m_msgDelegue = msgDelegue

        Depart() ' Empilage niveau 1

    End Sub

    Public Sub cmdStop()

        Me.m_bPause = False
        Me.m_bStop = True

    End Sub

    Private Sub Depart()

        If Me.m_prm.m_searchDir.Length < 3 OrElse
           Not Directory.Exists(Me.m_prm.m_searchDir) Then
            Me.m_bSucces = False
            Me.m_sMsgErr = "Impossible de trouver le dossier : " & Me.m_prm.m_searchDir
            GoTo Fin
        End If

        If Me.m_prm.m_containingChecked Then
            Try
                Dim sTxt$ = Me.m_prm.m_containingText
                Dim sTxtMin$ = sTxt.ToLower
                Me.m_abTxtRech = Encoding.ASCII.GetBytes(sTxt)
                Me.m_abTxtRechMin = Encoding.ASCII.GetBytes(sTxtMin)
                Me.m_abTxtRechUC = Encoding.Unicode.GetBytes(sTxt)
                Me.m_abTxtRechMinUC = Encoding.Unicode.GetBytes(sTxtMin)
            Catch ex As Exception
                Me.m_bSucces = False
                Me.m_sMsgErr = "Le texte ne peut pas être converti en octets : " &
                    Me.m_prm.m_containingText
            End Try
        End If
        If Not Me.m_bSucces Then GoTo Fin

        EmpilerJob(Me.m_prm.m_searchDir)

Fin:

    End Sub

    Private Sub EmpilerJob(sCheminDossier$)

        ' 02/10/2016 Si les filtres sont redondants, éviter d'empiler plusieurs fois les mêmes dossiers
        If Me.m_queue.Contains(sCheminDossier) Then Exit Sub

        Me.m_queue.Enqueue(sCheminDossier)
        Me.m_iNbDossiersEmpiles += 1 ' 25/09/2016

    End Sub

    Public Function bResteJob() As Boolean

        If Me.m_queue.Count = 0 Then Return False
        Return True

    End Function

    Public Function sDepilerJob$()

        ' Dépiler 1 job
        Dim sDossier$ = DirectCast(Me.m_queue.Dequeue, String)
        Return sDossier

    End Function

    Public Sub LireInfos(ByRef iNbFichiersOuDossiersTrouves%, ByRef iNbFichiersOuDossiers%,
        ByRef lNbOctetsLus&, ByRef lNbOctetsCompares&, ByRef lNbOctetsFichiers&,
        ByRef lTailleMoyFichier&, ByRef iNbDossiersParcourus%, ByRef iNbDossiers%)

        iNbFichiersOuDossiersTrouves = Me.m_iNbFichiersOuDossiersTrouves
        iNbFichiersOuDossiers = Me.m_iNbFichiersOuDossiers
        lNbOctetsLus = Me.m_lNbOctetsLus
        lNbOctetsCompares = Me.m_lNbOctetsCompares
        lNbOctetsFichiers = Me.m_lNbOctetsFichiers
        lTailleMoyFichier = Me.m_lTailleMoyFichier
        iNbDossiersParcourus = Me.m_iNbDossiersParcourus ' 25/09/2016
        'iNbDossiers = Me.m_iNbDossiers 
        iNbDossiers = Me.m_iNbDossiersEmpiles ' 25/09/2016

    End Sub

#End Region

#Region "Traitements"

    Private Sub ResetVariables()

        Me.m_bPause = False
        Me.m_bStop = False
        Me.m_bSucces = True
        Me.m_sMsgErr = ""
        Me.m_prm = Nothing
        Me.m_abTxtRech = Nothing
        Me.m_abTxtRechMin = Nothing
        Me.m_bAlerte = False
        Me.m_iNbFichiersOuDossiersTrouves = 0
        Me.m_lNbOctetsLus = 0
        Me.m_lNbOctetsCompares = 0
        Me.m_lNbOctetsFichiers = 0
        Me.m_lTailleMoyFichier = 0
        Me.m_iNbDossiersTrouves = 0
        Me.m_iNbFichiers = 0
        Me.m_iNbDossiers = 0
        Me.m_iNbDossiersEmpiles = 0
        Me.m_iNbDossiersParcourus = 0
        Me.m_iNbFichiersOuDossiers = 0
        Me.m_lexErr = New List(Of Exception)

        Me.m_hs = New HashSet(Of String)
        Me.m_hsDossiers = New HashSet(Of String)
        Me.m_hsExclus = New HashSet(Of String)
        Me.m_queue = New HashQueue(Of String)
        Me.m_dicoTaillesDossiers = New Dictionary(Of String, Long)
        Me.m_hsbSousDossiers = New HashSet(Of String)
        'Me.m_hsbSousDossiers2 = New HashSet(Of String)

        Me.m_hsElementsTrouves = New HashSet(Of String)
        Me.m_hsDossiersTrouves = New HashSet(Of String)

    End Sub

    Private Sub VerifierPause()

        Do While Me.m_bPause
            Attendre(100)
            TraiterMsgSysteme_DoEvents()
        Loop

    End Sub

    Public Sub ChercherArbo(dirInfo As DirectoryInfo, bExclusion As Boolean)

        If Me.m_bStop Then Return

        Dim iLongDossier% = dirInfo.FullName.Length
        If iLongDossier >= 248 Then Return ' 25/09/2016
        Const iLongMaxCheminComplet% = 260 - 1
        ' La longueur du chemin du dossier doit être < à 248 car. et les fichiers < 260 car.
        '  sinon on obtient l'exception suivante lorsque l'on fait dirInfo.GetFileSystemInfos :
        ' HResult=-2147024690
        ' Le chemin d'accès spécifié, le nom de fichier ou les deux sont trop longs.
        ' Le nom de fichier qualifié complet doit comprendre moins de 260 caractères
        '  et le nom du répertoire moins de 248 caractères.

        If Not bExclusion Then
            Me.m_iNbDossiersParcourus += 1
        End If

        Dim lTailleOctetsDossier& = 0

        Try
            Dim lstFiltre As List(Of String) = Me.m_prm.m_fileNames

            ' Note : Dans la fonction DepilerJobInterne, s'il y a une exclusion
            '  alors on appelle deux fois ChercherArbo, avec et sans exclusion
            ' En mode exclusion, on se contente de construire m_hsExclus 
            '  pour pouvoir filtrer ensuite le mode sans exclusion

            ' Whismeril a proposé en 2017 d'utiliser plutôt une requête linq pour les exclusions :
            ' Dim exclusions() As String = { "pas ca", "niCeluilà" }
            ' Dim resultats() As String = (
            '   From f In Directory.GetFiles("C:\Temp", "*a*.txt")
            '   Where (Not exclusions.Any(Function(ex) ExclureFichier(ex, f)))
            '   Select f).ToArray()
            ' https://codes-sources.commentcamarche.net/source/52496-vbfilefind-recherche-de-fichiers-pour-remplacer-celle-de-windows
            ' https://codes-sources.commentcamarche.net/profile/user/Whismeril
            ' Mais du coup, il faudrait réécrire en partie le code actuel, car il est
            '  basé sur dirInfo.GetFileSystemInfos et non pas sur Directory.GetFiles

            If bExclusion Then lstFiltre = Me.m_prm.m_fileNamesExcl
            For Each fileName As String In lstFiltre

                Dim iLongFichier% = fileName.Length
                Dim iLongTot% = iLongDossier + iLongFichier + 1 ' +1 pour le \
                If iLongTot > iLongMaxCheminComplet Then Continue For ' 25/09/2016

                Dim afsi As FileSystemInfo() = dirInfo.GetFileSystemInfos(fileName)
                For Each fsifileInfo As FileSystemInfo In afsi

                    ' S'il n'y a pas de thread on doit traiter les événements
                    If Not Me.m_prm.m_bUseThread Then TraiterMsgSysteme_DoEvents()
                    If Me.m_bStop Then Return
                    VerifierPause()

                    Dim sCheminComplet$ = fsifileInfo.FullName

                    ' 25/10/2015 D'abord on recherche les fichiers exclus du dossier
                    '  et ensuite on exclut ces fichiers du dossier
                    If bExclusion Then
                        If Not Me.m_hsExclus.Contains(sCheminComplet) Then _
                            Me.m_hsExclus.Add(sCheminComplet)
                        Continue For
                    Else
                        If Me.m_hsExclus.Contains(sCheminComplet) Then Continue For
                    End If

                    ' 06/07/2014 N'ajouter qu'une fois chaque élément, si les filtres sont redondants
                    If Me.m_hs.Contains(sCheminComplet) Then Continue For
                    Me.m_hs.Add(sCheminComplet)

                    Me.m_iNbFichiersOuDossiers += 1
                    Dim bFichier As Boolean = False
                    Dim lTailleFichier& = 0 ' 20/06/2022
                    If TypeOf (fsifileInfo) Is FileInfo Then ' Fichier (et non Dossier)
                        bFichier = True
                        Dim fi As FileInfo = DirectCast(fsifileInfo, FileInfo)
                        lTailleFichier = fi.Length
                        Me.m_iNbFichiers += 1
                        Me.m_lNbOctetsFichiers += lTailleFichier
                        Me.m_lTailleMoyFichier = Me.m_lNbOctetsFichiers \ Me.m_iNbFichiers

                        If m_bDebugTailleDossier AndAlso m_bDossierDebug Then
                            Debug.WriteLine("Fichier : " & sCheminComplet & " : " & lTailleFichier & " octets")
                            Debug.WriteLine("Dossier avant : " & lTailleOctetsDossier)
                        End If

                        lTailleOctetsDossier += lTailleFichier

                        If m_bDebugTailleDossier AndAlso m_bDossierDebug Then
                            Debug.WriteLine("Dossier après : " & lTailleOctetsDossier)
                        End If

                    Else
                        ' On a déjà vérifié les doublons avec m_hs
                        'If Not Me.m_hsDossiers.Contains(sCheminComplet) Then
                        Me.m_hsDossiers.Add(sCheminComplet)
                        Me.m_iNbDossiers += 1
                        'End If

                        ' 24/09/2022 Compter la taille des fichiers inclus dans le dossier
                        If Me.m_prm.m_greaterThanChecked OrElse Me.m_prm.m_smallerThanChecked Then
                            ' Cette fois la taille est juste et correspond à l'explorateur de fichier
                            Dim dirInfo2 As New DirectoryInfo(fsifileInfo.FullName)
                            lTailleFichier = dirInfo2.EnumerateFiles("*.*", SearchOption.AllDirectories).Sum(Function(file) file.Length)
                        End If
                    End If

                    ' On a déjà vérifié les doublons avec m_hs
                    'If Me.m_hsElementsTrouves.Contains(sCheminComplet) Then Continue For

                    If bElementCorrespondant(fsifileInfo, bFichier, lTailleFichier) Then
                        Me.m_iNbFichiersOuDossiersTrouves += 1
                        Me.m_hsElementsTrouves.Add(sCheminComplet)
                        If Not bFichier Then
                            ' Pour un dossier, on n'a vérifié que la date, le cas échéant
                            '  et la taille aussi maintenant 18/09/2022
                            Me.m_iNbDossiersTrouves += 1
                            Me.m_hsDossiersTrouves.Add(sCheminComplet)

                            m_bDossierDebug = False
                            If sCheminComplet = "" Then m_bDossierDebug = True

                            If m_bDebugTailleDossier AndAlso m_bDossierDebug Then
                                Debug.WriteLine("Dossier : " & sCheminComplet & " : " & lTailleFichier & " octets")
                                Debug.WriteLine("Dossier avant : " & lTailleOctetsDossier)
                            End If

                            lTailleOctetsDossier += lTailleFichier

                            If m_bDebugTailleDossier AndAlso m_bDossierDebug Then
                                Debug.WriteLine("Dossier après : " & lTailleOctetsDossier)
                            End If

                        End If
                        m_msgDelegue.AfficherFSIEnCours(fsifileInfo)
                    End If
                Next
            Next

            If Not bExclusion Then ' 01/10/2016 On n'a pas la taille dans ce cas
                ' Màj de la taille du dossier
                Dim sCle$ = dirInfo.FullName
                If Not m_dicoTaillesDossiers.ContainsKey(sCle) Then _
                    m_dicoTaillesDossiers.Add(sCle, lTailleOctetsDossier)
            End If

            If Me.m_prm.m_includeSubDirsChecked Then
                Dim subDirInfos As DirectoryInfo() = dirInfo.GetDirectories
                For Each subDirInfo As DirectoryInfo In subDirInfos

                    ' S'il n'y a pas de thread on doit traiter les événements
                    If Not Me.m_prm.m_bUseThread Then TraiterMsgSysteme_DoEvents()
                    If Me.m_bStop Then Return
                    VerifierPause()

                    ' Pour parcourir à la fin les sous-dossiers pour màj 
                    '  la taille de leur dossier parent
                    If Not Me.m_hsbSousDossiers.Contains(subDirInfo.FullName) Then _
                        Me.m_hsbSousDossiers.Add(subDirInfo.FullName)

                    ' 25/10/2015 Exclure les dossiers correspondants au filtre d'exclusion
                    If Me.m_hsExclus.Contains(subDirInfo.FullName) Then Continue For

                    EmpilerJob(subDirInfo.FullName) ' Appel récursif asynchrone
                Next
            End If

        Catch ex As Exception

            ' Exemple d'erreur : le dossier n'est pas accessible
            Me.m_lexErr.Add(ex)

        End Try

    End Sub

    Public Sub CalculerTaillesDossiers()

        ' Calculer la taille totale des dossiers avec leur sous-dossiers
        ' (pas seulement la taille des dossiers qui matchent)

        For Each sSousDossier As String In Me.m_hsbSousDossiers

            ' Les dossiers exclus ne sont pas comptabilisés
            If Not m_dicoTaillesDossiers.ContainsKey(sSousDossier) Then
                'Debug.WriteLine("Dossier non comptab.: " & sSousDossier)
                Continue For
            End If

            Dim lTailleSousDossier& = m_dicoTaillesDossiers(sSousDossier)
            If lTailleSousDossier = 0 Then
                'Debug.WriteLine("Sous-dossier vide : " & sSousDossier)
                Continue For
            End If

            Dim sDossierParent$ = IO.Path.GetDirectoryName(sSousDossier)
            Dim sCle$ = sDossierParent
            ' Le dossier parent racine peut être absent
            If Not m_dicoTaillesDossiers.ContainsKey(sCle) Then Continue For

            Dim bDebug1 As Boolean = False
            If m_bDebugTailleDossier Then
                Debug.WriteLine("Dossier : " & sSousDossier)
                bDebug1 = True
            End If

            Dim lTailleParent& = m_dicoTaillesDossiers(sCle)
            Dim lTailleParentNouv& = lTailleParent + lTailleSousDossier
            m_dicoTaillesDossiers.Remove(sCle)
            m_dicoTaillesDossiers.Add(sCle, lTailleParentNouv)
            If bDebug1 Then Debug.WriteLine("Cumul dossier : " & sDossierParent & " : " & lTailleParent &
                " + " & lTailleSousDossier & " -> " & lTailleParentNouv & " : " & sSousDossier)

        Next

    End Sub

    Private Function bElementCorrespondant(fsi As FileSystemInfo,
            bFichier As Boolean, lTailleFichier&) As Boolean

        Dim bOk As Boolean = True
        If bOk AndAlso Me.m_prm.m_newerThanChecked Then
            bOk = (fsi.LastWriteTime >= Me.m_prm.m_newerThanDateTime)
        End If
        If bOk AndAlso Me.m_prm.m_olderThanChecked Then
            bOk = (fsi.LastWriteTime <= Me.m_prm.m_olderThanDateTime)
        End If

        ' 20/06/2022 Faire un booléen pour filtrer aussi la taille des dossiers ?
        ' (le calcul de la taille du dossier risque de ralentir)
        ' Filtrer que les fichiers :
        If bOk AndAlso bFichier AndAlso Me.m_prm.m_greaterThanChecked Then
            bOk = (lTailleFichier >= Me.m_prm.m_greaterThanSize)
        End If
        If bOk AndAlso bFichier AndAlso Me.m_prm.m_smallerThanChecked Then
            bOk = (lTailleFichier <= Me.m_prm.m_smallerThanSize)
        End If
        ' Filtrer les fichiers et dossiers :
        If bOk AndAlso Not bFichier AndAlso Me.m_prm.m_greaterThanChecked Then
            bOk = (lTailleFichier >= Me.m_prm.m_greaterThanSize)
        End If
        If bOk AndAlso Not bFichier AndAlso Me.m_prm.m_smallerThanChecked Then
            bOk = (lTailleFichier <= Me.m_prm.m_smallerThanSize)
        End If

        If bOk AndAlso Me.m_prm.m_containingChecked Then

            bOk = False
            ' Si l'élément est un dossier, alors poursuivre
            If Not bFichier Then GoTo Suite

            ' C'est un fichier : lancer une recherche de son contenu

            Dim abTxtRech As Byte() = Nothing
            Dim abTxtRechMin As Byte() = Nothing
            Dim abTxtRechUc As Byte() = Nothing
            Dim abTxtRechMinUc As Byte() = Nothing

            ' Recherches sensibles à la casse (minuscules/majuscules)
            Dim bDetecterEncodage As Boolean = False
            If Me.m_prm.m_iTypeEncodage = TypeEncodage.ASCII Then
                abTxtRech = Me.m_abTxtRech
            ElseIf Me.m_prm.m_iTypeEncodage = TypeEncodage.Unicode Then
                abTxtRechUc = Me.m_abTxtRechUC
            ElseIf Me.m_prm.m_iTypeEncodage = TypeEncodage.ASCII_Ou_Unicode Then ' Les 2
                abTxtRech = Me.m_abTxtRech
                abTxtRechUc = Me.m_abTxtRechUC
            ElseIf Me.m_prm.m_iTypeEncodage = TypeEncodage.Detecter Then
                abTxtRech = Me.m_abTxtRech
                abTxtRechUc = Me.m_abTxtRechUC
                bDetecterEncodage = True ' 27/07/2024
            End If

            ' Recherches insensibles à la casse
            If Not Me.m_prm.m_casseChecked Then
                If Me.m_prm.m_iTypeEncodage = TypeEncodage.ASCII Then
                    abTxtRechMin = Me.m_abTxtRechMin
                ElseIf Me.m_prm.m_iTypeEncodage = TypeEncodage.Unicode Then
                    abTxtRechMinUc = Me.m_abTxtRechMinUC
                ElseIf Me.m_prm.m_iTypeEncodage = TypeEncodage.ASCII_Ou_Unicode Then ' Les 2
                    abTxtRechMin = Me.m_abTxtRechMin
                    abTxtRechMinUc = Me.m_abTxtRechMinUC
                ElseIf Me.m_prm.m_iTypeEncodage = TypeEncodage.Detecter Then
                    abTxtRech = Me.m_abTxtRechMin
                    abTxtRechUc = Me.m_abTxtRechMinUC
                    bDetecterEncodage = True ' 27/07/2024
                End If
            End If

            If bLectureDepuisFileStream Then
                ' Ancienne technique : pas d'option de détection de l'encodage
                bOk = bFichierContientOccFileStream(fsi.FullName,
                    abTxtRech, abTxtRechMin, abTxtRechUc, abTxtRechMinUc)
            Else
                ' 16/04/2023 Même code, mais en utilisant un StreamReader cette fois : pas de différence,
                '  mais on peut maintenant demander à détecter l'encodage, via l'option
                '  detectEncodingFromByteOrderMarks de StreamReader
                bOk = bFichierContientOccStreamReader(fsi.FullName,
                    abTxtRech, abTxtRechMin, abTxtRechUc, abTxtRechMinUc, bDetecterEncodage)
            End If

        End If

Suite:
        Return bOk

    End Function

    Private Function bFichierContientOccFileStream(sChemin$,
        abTxtRech As Byte(), abTxtRechMin As Byte(),
        abTxtRechUc As Byte(), abTxtRechMinUc As Byte()) As Boolean

        Dim bContient As Boolean = False
        Dim iTailleBloc% = 16384 ' 4096
        Dim iLongRech% = 0
        Dim iLongRechUc% = 0

        Dim bRech As Boolean = True
        If IsNothing(abTxtRech) Then bRech = False
        If bRech Then iLongRech = abTxtRech.Length
        Dim bRechUc As Boolean = True
        If IsNothing(abTxtRechUc) Then bRechUc = False
        If bRechUc Then iLongRechUc = abTxtRechUc.Length

        Dim bRechMin As Boolean = True
        If IsNothing(abTxtRechMin) Then bRechMin = False
        If bRechMin Then
            Dim iLongRechMin% = abTxtRechMin.Length
            If iLongRechMin <> iLongRech Then
                Debug.WriteLine("!")
                If bDebug Then Stop
            End If
        End If
        Dim bRechMinUc As Boolean = True
        If IsNothing(abTxtRechMinUc) Then bRechMinUc = False
        If bRechMinUc Then
            Dim iLongRechMinUc% = abTxtRechMinUc.Length
            If iLongRechMinUc <> iLongRechUc Then
                Debug.WriteLine("!")
                If bDebug Then Stop
            End If
        End If

        If Not bRech And bRechUc Then iLongRech = iLongRechUc

        ' Si le texte a chercher est vide, alors tous les fichiers correspondent
        If iLongRech = 0 Then GoTo Fin
        ' Si le nbre d'octets à rechercher est > à la taille du bloc, alors
        '  cela veut dire que l'on cherche un texte de très grande taille : idiot !
        If iLongRech > iTailleBloc Then
            If Not Me.m_bAlerte Then _
                MsgBox("La taille du tampon de recherche est insuffisante : " &
                    iTailleBloc & "<" & iLongRech & " (longueur de la chaîne à rechercher)",
                    MsgBoxStyle.Critical)
            Me.m_bAlerte = True
            GoTo Fin
        End If

        Dim iLongBloc% = iLongRech - 1 + iTailleBloc - 1
        Dim abBloc As Byte() = New Byte(iLongBloc) {}
        Dim acBloc As Char() = Nothing
        Dim sr As StreamReader = Nothing

        Try
            Using fs As New FileStream(sChemin, FileMode.Open, FileAccess.Read)

                If m_bDebugRecherche Then Debug.WriteLine("\nFichier : " & IO.Path.GetFileName(sChemin) & "\n")

                RechercheOcc(
                    fs, sr,
                    iLongRech, iTailleBloc,
                    bRech, bRechUc, bRechMin, bRechMinUc,
                    abTxtRech, abTxtRechMin,
                    abTxtRechUc, abTxtRechMinUc, acBloc, abBloc, bContient, bReadFileStream:=True)

            End Using 'fs.Close()
        Catch ex As Exception
            Me.m_lexErr.Add(ex)
        End Try

Fin:
        Return bContient

    End Function

    Private Function abConvMin(enc As Encoding, abBloc As Byte(), ByRef bEchec As Boolean) As Byte()

        ' Convertir un tableau d'octets en minuscules selon l'encodage demandé
        ' (en supposant que les octets forment un texte, ce qui n'est que pur hypothèse)

        ' On pourrait simplifier (sans vérification) en :
        ' enc.GetBytes(enc.GetString(abBloc).ToLower)

        bEchec = False
        Dim sChaine$ = enc.GetString(abBloc)
        Dim sChaineMin$ = sChaine.ToLower
        abConvMin = enc.GetBytes(sChaineMin)

        ' Vérification
        Dim iLenSrc% = abBloc.Length
Verifier:
        Dim iLenMin% = abConvMin.Length
        If iLenMin < iLenSrc Then
            Debug.WriteLine("!")
            If bDebug Then Stop
            bEchec = True
        End If
        If iLenMin > iLenSrc And Not bEchec Then
            ' Si la longueur en minuscule n'est plus la même, 
            '  on peut quand même trouver l'occurrence : poursuivre la recherche
            bEchec = True
            ReDim Preserve abConvMin(0 To iLenSrc - 1)
            GoTo Verifier
        End If

    End Function

    Private Function bContientTxt(iPosFin%, iLongRech%, abBloc As Byte(), abTxtRech As Byte(),
            bRechercheUnicode As Boolean) As Boolean

        For i As Integer = 0 To iPosFin - 1
            Dim j% = 0
            Do While j < iLongRech
                If abBloc(i + j) <> abTxtRech(j) Then Exit Do
                j += 1
            Loop
            Me.m_lNbOctetsCompares += j + 1
            If j = iLongRech Then Return True
        Next i
        Return False

    End Function

    ' Même code, mais en utilisant un streamreader cette fois : pas de différence, mais on peut détecter l'encodage !
    Private Function bFichierContientOccStreamReader(sChemin$,
            abTxtRech As Byte(), abTxtRechMin As Byte(),
            abTxtRechUc As Byte(), abTxtRechMinUc As Byte(), bDetecterEncodage As Boolean) As Boolean

        Dim bContient As Boolean = False
        Dim iTailleBloc% = 16384 ' 4096
        Dim iLongRech% = 0
        Dim iLongRechUc% = 0

        Dim bRech As Boolean = True
        If IsNothing(abTxtRech) Then bRech = False
        If bRech Then iLongRech = abTxtRech.Length
        Dim bRechUc As Boolean = True
        If IsNothing(abTxtRechUc) Then bRechUc = False
        If bRechUc Then iLongRechUc = abTxtRechUc.Length

        Dim bRechMin As Boolean = True
        If IsNothing(abTxtRechMin) Then bRechMin = False
        If bRechMin Then
            Dim iLongRechMin% = abTxtRechMin.Length
            If iLongRechMin <> iLongRech Then
                Debug.WriteLine("!")
                If bDebug Then Stop
            End If
        End If
        Dim bRechMinUc As Boolean = True
        If IsNothing(abTxtRechMinUc) Then bRechMinUc = False
        If bRechMinUc Then
            Dim iLongRechMinUc% = abTxtRechMinUc.Length
            If iLongRechMinUc <> iLongRechUc Then
                Debug.WriteLine("!")
                If bDebug Then Stop
            End If
        End If

        If Not bRech And bRechUc Then iLongRech = iLongRechUc

        ' Si le texte a chercher est vide, alors tous les fichiers correspondent
        If iLongRech = 0 Then GoTo Fin
        ' Si le nbre d'octets à rechercher est > à la taille du bloc, alors
        '  cela veut dire que l'on cherche un texte de très grande taille : idiot !
        If iLongRech > iTailleBloc Then
            If Not Me.m_bAlerte Then _
                MsgBox("La taille du tampon de recherche est insuffisante : " &
                    iTailleBloc & "<" & iLongRech & " (longueur de la chaîne à rechercher)",
                    MsgBoxStyle.Critical)
            Me.m_bAlerte = True
            GoTo Fin
        End If

        Dim iLongBloc% = iLongRech - 1 + iTailleBloc - 1
        Dim acBloc As Char() = New Char(iLongBloc) {}
        Dim abBloc As Byte() = Nothing

        Try
            Dim encodage As Encoding
            If bRechUc OrElse bRechMinUc Then
                encodage = Encoding.Unicode
            Else
                encodage = Encoding.ASCII
            End If
            Using fs As New FileStream(sChemin, FileMode.Open, FileAccess.Read)

                If bDetecterEncodage Then

                    If m_bDebugRecherche Then Debug.WriteLine("\nFichier : " & IO.Path.GetFileName(sChemin) & "\n")

                    Using sr As New StreamReader(fs, detectEncodingFromByteOrderMarks:=bDetecterEncodage)
                        RechercheOcc(
                            fs, sr,
                            iLongRech, iTailleBloc,
                            bRech, bRechUc, bRechMin, bRechMinUc,
                            abTxtRech, abTxtRechMin,
                            abTxtRechUc, abTxtRechMinUc, acBloc, abBloc, bContient, bReadFileStream:=False)
                    End Using

                Else

                    Using sr As New StreamReader(fs, encodage)
                        RechercheOcc(
                            fs, sr,
                            iLongRech, iTailleBloc,
                            bRech, bRechUc, bRechMin, bRechMinUc,
                            abTxtRech, abTxtRechMin,
                            abTxtRechUc, abTxtRechMinUc, acBloc, abBloc, bContient, bReadFileStream:=False)
                    End Using

                End If
            End Using
        Catch ex As Exception
            Me.m_lexErr.Add(ex)
        End Try

Fin:
        Return bContient

    End Function

    Private Sub RechercheOcc(
            fs As FileStream, sr As StreamReader,
            iLongRech%, iTailleBloc%,
            bRech As Boolean, bRechUc As Boolean, bRechMin As Boolean, bRechMinUc As Boolean,
            abTxtRech As Byte(), abTxtRechMin As Byte(),
            abTxtRechUc As Byte(), abTxtRechMinUc As Byte(), acBloc As Char(), abBloc As Byte(),
            ByRef bContient As Boolean, bReadFileStream As Boolean)

        Dim iNbOctetsLus%
        If bReadFileStream Then
            iNbOctetsLus = fs.Read(abBloc, 0, abBloc.Length)
        Else
            iNbOctetsLus = sr.Read(acBloc, 0, acBloc.Length)

            If m_bDebugRecherche Then
                Dim encodage As Encoding = sr.CurrentEncoding
                Debug.WriteLine("Encodage détecté : " & encodage.ToString)
            End If

        End If

        Me.m_lNbOctetsLus += iNbOctetsLus

Boucle:
        Dim iPosFin% = iNbOctetsLus - iLongRech + 1

        If bRech Then
            If bReadFileStream Then
                bContient = bContientTxt(iPosFin, iLongRech, abBloc, abTxtRech, bRechercheUnicode:=False)
            Else
                bContient = bContientTxt(iPosFin, iLongRech, acBloc, abTxtRech, bRechercheUnicode:=False)
            End If
            If bContient Then Exit Sub
        End If

        If bRechUc Then
            If bReadFileStream Then
                bContient = bContientTxt(iPosFin, iLongRech, abBloc, abTxtRechUc, bRechercheUnicode:=True)
            Else
                bContient = bContientTxt(iPosFin, iLongRech, acBloc, abTxtRechUc, bRechercheUnicode:=True)
            End If
            If bContient Then Exit Sub
        End If

        If bRechMin Then ' Rechercher aussi en minuscules
            ' Il faut aussi que abBloc soit ToLower, pour cela il faut l'encodage
            Dim bEchec As Boolean
            Dim abMin As Byte()
            If bReadFileStream Then
                abMin = abConvMin(Encoding.ASCII, abBloc, bEchec)
            Else
                abMin = abConvMin(Encoding.ASCII, acBloc, bEchec)
            End If
            bContient = bContientTxt(iPosFin, iLongRech, abMin, abTxtRechMin, bRechercheUnicode:=False)
            If bContient Then Exit Sub
        End If

        If bRechMinUc Then ' Rechercher aussi en minuscules
            Dim bEchec As Boolean
            Dim abMin As Byte()
            If bReadFileStream Then
                abMin = abConvMin(Encoding.Unicode, abBloc, bEchec)
            Else
                abMin = abConvMin(Encoding.Unicode, acBloc, bEchec)
            End If
            bContient = bContientTxt(iPosFin, iLongRech, abMin, abTxtRechMinUc, bRechercheUnicode:=True)
            ' Si la longueur en minuscule n'est plus la même (bEchec), 
            '  on peut quand même trouver l'occurrence : poursuivre la recherche
            If bContient Then Exit Sub
        End If

        If fs.Position >= fs.Length Then Exit Sub

        For i As Integer = 0 To iLongRech - 1 - 1
            acBloc(i) = acBloc(iTailleBloc + i)
        Next i
        Dim iNbOctetsLus0%
        If bReadFileStream Then
            iNbOctetsLus0 = fs.Read(abBloc, iLongRech - 1, iTailleBloc)
        Else
            iNbOctetsLus0 = sr.Read(acBloc, iLongRech - 1, iTailleBloc)
        End If
        Me.m_lNbOctetsLus += iNbOctetsLus0
        iNbOctetsLus = iLongRech - 1 + iNbOctetsLus0

        ' Si on enlève le thread alors traiter les événements
        TraiterMsgSysteme_DoEvents()
        VerifierPause()
        If Not Me.m_bStop Then GoTo Boucle

    End Sub

    Private Function abConvMin(enc As Encoding, acBloc As Char(),
        ByRef bEchec As Boolean) As Byte()

        ' Convertir un tableau d'octets en minuscules selon l'encodage demandé
        ' (en supposant que les octets forment un texte, ce qui n'est que pur hypothèse)

        ' On pourrait simplifier (sans vérification) en :
        ' enc.GetBytes(enc.GetString(abBloc).ToLower)

        bEchec = False
        Dim abBloc As Byte() = New Byte(acBloc.Length) {}
        For i As Integer = 0 To acBloc.Length - 1
            Dim c As Char = acBloc(i)
            Dim iAscW% = Microsoft.VisualBasic.AscW(c)
            If iAscW > Byte.MaxValue Then Continue For ' 16/04/2023
            Dim b As Byte = CByte(iAscW)
            abBloc(i) = b
        Next
        Dim sChaine$ = enc.GetString(abBloc)
        Dim sChaineMin$ = sChaine.ToLower
        abConvMin = enc.GetBytes(sChaineMin)

        ' Vérification
        Dim iLenSrc% = acBloc.Length
Verifier:
        Dim iLenMin% = abConvMin.Length
        If iLenMin < iLenSrc Then
            Debug.WriteLine("!")
            If bDebug Then Stop
            bEchec = True
        End If
        If iLenMin > iLenSrc And Not bEchec Then
            ' Si la longueur en minuscule n'est plus la même, 
            '  on peut quand même trouver l'occurrence : poursuivre la recherche
            bEchec = True
            ReDim Preserve abConvMin(0 To iLenSrc - 1)
            GoTo Verifier
        End If

    End Function

    Private Function bContientTxt(iPosFin%, iLongRech%, acBloc As Char(), abTxtRech As Byte(),
            bRechercheUnicode As Boolean) As Boolean

        If m_bDebugRecherche Then

            Debug.WriteLine("")
            If bRechercheUnicode Then
                Debug.WriteLine("Recherche Unicode : ")
            Else
                Debug.WriteLine("Recherche ASCII : ")
            End If

            Dim subArray As Char()
            ReDim subArray(iPosFin + 1)
            Array.Copy(acBloc, 0, subArray, 0, subArray.Length)
            Dim texteDetect$ = New String(subArray)
            Debug.WriteLine("Texte recherché : " & texteDetect & " : ")

            Dim texteRechercheUTF8$ = Encoding.UTF8.GetString(abTxtRech)
            Debug.WriteLine("Test UTF8 : " & texteRechercheUTF8 & " : ")
            Dim texteRechercheUc$ = Encoding.Unicode.GetString(abTxtRech)
            Debug.WriteLine("Test Unicode : " & texteRechercheUc & " : ")
            For i As Integer = 0 To abTxtRech.GetUpperBound(0)
                Dim c As Byte = abTxtRech(i)
                Debug.Write(c & ", ")
            Next
            Debug.WriteLine(" -> ")
        End If

        For i As Integer = 0 To iPosFin - 1
            Dim j% = 0
            Do While j < iLongRech
                Dim c As Char = acBloc(i + j)
                Dim b As Byte = abTxtRech(j)
                Dim ci% = Asc(c)
                If m_bDebugRecherche AndAlso i < 10 Then Debug.Write(ci & ", ")
                If ci <> b Then Exit Do
                j += 1
            Loop
            Me.m_lNbOctetsCompares += j + 1
            If j = iLongRech Then Return True
        Next i
        Return False

    End Function

#End Region

End Class