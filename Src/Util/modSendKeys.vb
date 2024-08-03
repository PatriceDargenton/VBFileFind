
Imports System.Runtime.InteropServices ' DllImport

Friend NotInheritable Class NativeMethods

    Private Sub New()
    End Sub

    '<DllImport("user32.dll", SetLastError:=True)> _
    'Public Shared Function ShowWindowAsync(hWnd As IntPtr, _
    '    <MarshalAs(UnmanagedType.I4)> nCmdShow As ShowWindowCommands) _
    '        As <MarshalAs(UnmanagedType.Bool)> Boolean
    'End Function

    '<DllImport("user32.dll")> _
    'Public Shared Function ShowWindow%(ByVal hwnd%, ByVal nCmdShow%)
    'End Function

    <DllImport("user32.dll")> _
    Public Shared Function ShowWindow%(hwnd As IntPtr, nCmdShow As Int32)
    End Function

End Class

Module modSendKeys

    'Private Declare Function apiShowWindow Lib "user32" Alias "ShowWindow" _
    '    (hWnd As IntPtr, nCmdShow As Int32) As Boolean

    Public Sub OuvrirBlocNotes(sChemin$,
        Optional sTxtRech$ = "", Optional bSensibleCasse As Boolean = False,
        Optional bUnicode As Boolean = False,
        Optional sRaccourciBlocNotesRechercherCtrl_f$ = "^(f)",
        Optional sRaccourciBlocNotesSensibleCasseAlt_c$ = "%(c)",
        Optional sRaccourciBlocNotesOccurrSuivAlt_S$ = "%(S)",
        Optional sExeBlocNotes$ = "Notepad.exe",
        Optional b1ereOuverture As Boolean = False,
        Optional iDelaiMSec1Ouverture% = 0,
        Optional bOuvrirSeulement As Boolean = False)

        Dim fi As New IO.FileInfo(sChemin)
        Dim iNbMo% = CInt(fi.Length \ (1024 * 1024))
        If iNbMo > 10 Then
            Dim sInfo$ = "Le fichier " & sChemin & vbLf &
                "a une taille de " & sFormaterTailleOctets(fi.Length) & " :" & vbLf &
                "Etes-vous sûr de vouloir l'ouvrir avec le bloc-notes ?"
            If MsgBoxResult.Cancel = MsgBox(sInfo,
                MsgBoxStyle.OkCancel Or MsgBoxStyle.Exclamation, sTitreMsg) Then Exit Sub
        End If

        If bOuvrirSeulement Then
            OuvrirAppliAssociee(sChemin)
            Exit Sub
        End If

        Try
            Dim startInfo As New ProcessStartInfo
            Dim sSysDir$ = Environment.GetFolderPath(Environment.SpecialFolder.System)
            'Dim sWinDir$ = IO.Path.GetDirectoryName(sSysDir)
            ' 24/09/2016 Ce dossier fonctionne dans tous les cas, y compris Windows server 2012
            startInfo.FileName = sSysDir & "\" & sExeBlocNotes
            'startInfo.FileName = sWinDir & "\" & sExeBlocNotes ' Ne fonctionne pas avec Windows server 2012
            If Not bFichierExiste(startInfo.FileName, bPrompt:=True) Then Exit Sub ' 18/08/2016
            startInfo.Arguments = sChemin
            Dim procNotePad As New Process With {.StartInfo = startInfo}
            procNotePad.Start()

            If sTxtRech.Length = 0 Then Exit Sub

            ' Need to wait for notepad to start
            procNotePad.WaitForInputIdle()
            Dim p As IntPtr = procNotePad.MainWindowHandle
            'Const SW_HIDE As Int32 = 0
            Const SW_SHOWNORMAL As Int32 = 1
            'Const SW_SHOWMINIMIZED As Int32 = 2
            'apiShowWindow(p, SW_SHOWNORMAL)
            Dim iRes% = NativeMethods.ShowWindow(p, SW_SHOWNORMAL)
            If iRes = 0 Then Exit Sub ' Pas de SendKeys si le ShowWindow a échoué

            ' 16/10/2016 Voir s'il faut attendre à nouveau ici, notamment lors de la 1ère ouverture
            If Not b1ereOuverture Then Attendre(iDelaiMSec1Ouverture)
            procNotePad.WaitForInputIdle()

            ' Alt : %, Shift : +, Ctrl : ^
            'SendKeys.SendWait("abcdef")
            'SendKeys.SendWait("^({HOME})^(f)d")

            ' Note : en anglais les raccourcis f, c et S seront différents
            '  Pour traiter ce cas, il faut lire la culture en cours
            '  et appliquer les raccourcis du bloc-notes dans la langue en cours 
            '  (ou sinon changer la configuration) :
            ' ^(f) : Ctrl+f : Menu Edition : Rechercher...
            ' %(c) : Alt +c : Case à cocher : sensible à la casse
            ' %(S) : Alt +S : Bouton Occurrence Suivante
            SendKeys.SendWait(sRaccourciBlocNotesRechercherCtrl_f)

            If bUnicode Then
                SendKeys.SendWait(sInsererEspacesTxt(sTxtRech))
            Else
                SendKeys.SendWait(sTxtRech)
            End If

            If bSensibleCasse Then
                ' Cocher sensible à la casse dans le bloc-notes
                SendKeys.SendWait(sRaccourciBlocNotesSensibleCasseAlt_c)
            End If
            SendKeys.SendWait(sRaccourciBlocNotesOccurrSuivAlt_S) ' Occurrence Suivante

        Catch ex As Exception
            AfficherMsgErreur2(ex, "OuvrirBlocNotes")
            'MsgBox(ex, MsgBoxStyle.Critical)
        End Try

    End Sub

    Private Function sInsererEspacesTxt$(sTxt$)

        Dim sb As New System.Text.StringBuilder
        For Each c As Char In sTxt
            sb.Append(c & " ")
        Next
        sInsererEspacesTxt = sb.ToString

    End Function

End Module
