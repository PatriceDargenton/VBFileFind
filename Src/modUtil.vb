
Imports System.Text ' Pour StringBuilder

Module modUtilitaires

    ' Module de fonctions utilitaires

    Public Sub AfficherMsgErreur2(ByRef Ex As Exception,
        Optional sTitreFct$ = "", Optional sInfo$ = "",
        Optional sDetailMsgErr$ = "",
        Optional bCopierMsgPressePapier As Boolean = True,
        Optional ByRef sMsgErrFinal$ = "")

        If Not Cursor.Current.Equals(Cursors.Default) Then _
            Cursor.Current = Cursors.Default
        Dim sMsg$ = ""
        If sTitreFct <> "" Then sMsg = "Fonction : " & sTitreFct
        If sInfo <> "" Then sMsg &= vbCrLf & sInfo
        If sDetailMsgErr <> "" Then sMsg &= vbCrLf & sDetailMsgErr
        If Ex.Message <> "" Then
            sMsg &= vbCrLf & Ex.Message.Trim
            If Not IsNothing(Ex.InnerException) Then _
                sMsg &= vbCrLf & Ex.InnerException.Message
        End If
        If bCopierMsgPressePapier Then CopierPressePapier(sMsg)
        sMsgErrFinal = sMsg
        MsgBox(sMsg, MsgBoxStyle.Critical)

    End Sub

    Public Sub CopierPressePapier(sInfo$)

        ' Copier des informations dans le presse-papier de Windows
        ' (elles resteront jusqu'à ce que l'application soit fermée)

        Try
            Dim dataObj As New DataObject
            dataObj.SetData(DataFormats.Text, sInfo)
            Clipboard.SetDataObject(dataObj)
        Catch ex As Exception
            ' Le presse-papier peut être indisponible
            AfficherMsgErreur2(ex, "CopierPressePapier",
                bCopierMsgPressePapier:=False)
        End Try

    End Sub

    Public Sub Attendre(Optional iMilliSec% = 200)
        Threading.Thread.Sleep(iMilliSec)
    End Sub

    Public Sub TraiterMsgSysteme_DoEvents()

        'Try
        Application.DoEvents() ' Peut planter avec OWC : Try Catch nécessaire
        'Catch
        'End Try

    End Sub

    Public Function iConv%(sVal$, Optional iValDef% = 0)

        If String.IsNullOrEmpty(sVal) Then Return iValDef

        Dim iVal%
        If Integer.TryParse(sVal, iVal) Then
            Return iVal
        Else
            Return iValDef
        End If

    End Function

    Public Function lConv&(sVal$, Optional lValDef& = 0)

        If String.IsNullOrEmpty(sVal) Then Return lValDef

        Dim lVal&
        If Long.TryParse(sVal, lVal) Then
            Return lVal
        Else
            Return lValDef
        End If

    End Function

End Module
