
Public Class SortableListView : Inherits ListView

Public Enum enumIdxImgTri
     iIdxImgTriDesactive = 0
     iIdxImgTriDescendant = 1
     iIdxImgTriMontant = 2
End Enum

' Fields
Private m_components As System.ComponentModel.IContainer = Nothing
Private m_iColTri As Integer = -1

' Possibilité d'attribuer un tri d'une colonne à une autre
' (ex.: une colonne affiche la taille en Ko en mode texte, mais on tri en octets)
' (actif si l'indice de la colonne est >=0)
Public m_iColTriSrc% = -1
Public m_iColTriDest% = -1
Public m_bNePasTrier As Boolean ' Ne pas trier pendant une recherche de fichiers par exemple

Private m_msgDelegue As clsMsgDelegue
Public m_bTriEnCours As Boolean = False

' Methods
Public Sub New()

    Me.InitializeComponent()
    MyBase.View = View.Details
    MyBase.AllowColumnReorder = True
    MyBase.FullRowSelect = True
    MyBase.ShowItemToolTips = True

End Sub

Public Sub DefinirMsgDelegue(msgDelegue As clsMsgDelegue)
    m_msgDelegue = msgDelegue
End Sub

Protected Overrides Sub Dispose(disposing As Boolean)

    If (disposing AndAlso (Not Me.m_components Is Nothing)) Then
        Me.m_components.Dispose()
    End If

    MyBase.Dispose(disposing)

End Sub

Private Sub InitializeComponent()
    Me.m_components = New System.ComponentModel.Container
End Sub

Public Sub DesactiverTri()

    If Not IsNothing(MyBase.ListViewItemSorter) Then

        If m_bTriEnCours Then
            While m_bTriEnCours
                Attendre()
                TraiterMsgSysteme_DoEvents()
            End While
        End If

        ' Enlever la précédente couleur de tri (en conservant les autres couleurs)
        Dim hsCoul As New HashSet(Of Color) ' Couleurs utilisées actuellement (sauf blanc et grisé)
        Dim coulTri As Color = Color.WhiteSmoke 'LightGray ' SystemColors.Info : non !
        Dim coulVide As Color = Color.White
        EnleverColoriageTri(hsCoul, coulTri, coulVide)
    End If

    MyBase.ListViewItemSorter = Nothing ' 16/12/2012 Désactiver le tri

End Sub

'Protected Overrides Sub OnColumnClick(e As ColumnClickEventArgs)
'    MyBase.OnColumnClick(e)
'End Sub
'Protected Overrides Sub OnColumnReordered(e As ColumnReorderedEventArgs)
'    ' Drag & drop des colonnes pour changer l'ordre d'affichage des colonnes
'    '  on récupère après coup l'événement
'    MyBase.OnColumnReordered(e)
'End Sub

Private Sub list_ColumnClick(sender As Object, e As ColumnClickEventArgs) _
    Handles MyBase.ColumnClick

    If m_bNePasTrier Then
        DesactiverTri()
        Exit Sub ' 18/11/2012
    End If

    m_bTriEnCours = True
    m_msgDelegue.AfficherMsg("Tri en cours...")
    m_msgDelegue.Sablier()

    Me.SuspendLayout()

    If e.Column <> m_iColTri Then
        m_iColTri = e.Column
        MyBase.Sorting = SortOrder.Ascending
    ElseIf MyBase.Sorting = SortOrder.Ascending Then
        MyBase.Sorting = SortOrder.Descending
    Else
        MyBase.Sorting = SortOrder.Ascending
    End If
    Dim iCol% = e.Column
    Dim iColTri0% = iCol
    ' Possibilité d'attribuer un tri d'une colonne à une autre
    ' (ex.: une colonne affiche la taille en Ko en mode texte, mais on tri en octets)
    ' (actif si l'indice de la colonne est >=0)
    If iCol = m_iColTriSrc Then iColTri0 = m_iColTriDest
    Dim typeDonneeColonne As Type = TryCast(MyBase.Columns.Item(iColTri0).Tag, Type)

    MyBase.ListViewItemSorter = New ListViewItemComparer(iColTri0, MyBase.Sorting, typeDonneeColonne)

    m_msgDelegue.AfficherMsg("Tri en cours 2...")
    ' Enlever la précédente couleur de tri (en conservant les autres couleurs)
    Dim hsCoul As New HashSet(Of Color) ' Couleurs utilisées actuellement (sauf blanc et grisé)
    Dim coulTri As Color = Color.WhiteSmoke 'LightGray ' SystemColors.Info : non !
    Dim coulVide As Color = Color.White
    EnleverColoriageTri(hsCoul, coulTri, coulVide)

    ' Colorier la colonne de tri (sauf les colonnes déjà coloriées au départ)

    ' Il y a 3 images : 0 : pas de tri, 1 : tri descendant, 2 : tri montant
    ' Initialisation : Me.lvResultats.SmallImageList = Me.imgLstLVH
    'If Not IsNothing(Me.SmallImageList) Then _
    Me.Columns.Item(iCol).ImageIndex = CInt(IIf(MyBase.Sorting = SortOrder.Descending, _
        enumIdxImgTri.iIdxImgTriDescendant, enumIdxImgTri.iIdxImgTriMontant))

    m_msgDelegue.AfficherMsg("Tri en cours 3...")
    'Dim iNbItems% = Me.Items.Count
    'Dim iNumItem% = 0
    For Each lvi As ListViewItem In Me.Items

        ' Comme on a mis le SuspendLayout, ça ne sert à rien d'envoyer des msg
        'iNumItem += 1
        'If iNumItem Mod 10000 = 0 Then
        '    Me.ResumeLayout()
        '    m_msgDelegue.AfficherMsg("Tri en cours 3 " & iNumItem & "/" & iNbItems & "...")
        '    TraiterMsgSysteme_DoEvents()
        '    If m_msgDelegue.m_bAnnuler Then Exit For ' 19/03/2017
        '    Me.SuspendLayout()
        'End If

        If iCol >= lvi.SubItems.Count Then Continue For ' Msg d'erreur : un seul item
        Dim coulSousItem As Color = lvi.SubItems(iCol).BackColor
        If Not hsCoul.Contains(coulSousItem) Then lvi.SubItems(iCol).BackColor = coulTri
    Next

    m_msgDelegue.AfficherMsg("Tri en cours 4...")
    Me.ResumeLayout()
    m_msgDelegue.AfficherMsg("Tri terminé.")
    m_msgDelegue.Sablier(bDesactiver:=True)
    m_bTriEnCours = False

End Sub

Private Sub EnleverColoriageTri(hsCoul As HashSet(Of Color), _
    coulTri As Color, coulVide As Color)

    For i As Integer = 0 To Me.Columns.Count - 1
        'If Not IsNothing(Me.SmallImageList) Then _
            Me.Columns.Item(i).ImageIndex = enumIdxImgTri.iIdxImgTriDesactive
        For Each lvi As ListViewItem In Me.Items
            If i >= lvi.SubItems.Count Then Continue For ' Msg d'erreur : un seul item
            Dim coulSousItem As Color = lvi.SubItems(i).BackColor
            If coulSousItem.ToArgb <> coulVide.ToArgb AndAlso _
               coulSousItem.ToArgb <> coulTri.ToArgb AndAlso _
               Not hsCoul.Contains(coulSousItem) Then hsCoul.Add(coulSousItem)
            If Not hsCoul.Contains(coulSousItem) Then lvi.SubItems(i).BackColor = coulVide
        Next
    Next

End Sub

#Region "ListViewItemComparer"

    ' Nested Types
    Public Class ListViewItemComparer : Implements Collections.IComparer

        ' Fields
        Private col As Integer
        Private columnType As Type
        Private order As SortOrder

        ' Methods
        Public Sub New()
            Me.col = 0
            Me.order = SortOrder.Ascending
            Me.columnType = Nothing
        End Sub

        Public Sub New(column As Integer, order As SortOrder, type As Type)
            Me.col = column
            Me.order = order
            Me.columnType = type
        End Sub

        Public Function Compare%(olviX As Object, olviY As Object) _
            Implements Collections.IComparer.Compare

            Dim iRetour% = -1
            Dim iCol% = Me.col
            Dim sTxtX$ = sLireColonne(olviX, iCol)
            Dim sTxtY$ = sLireColonne(olviY, iCol)

            If sTxtX.Length = 0 OrElse sTxtY.Length = 0 OrElse _
              (Me.columnType Is Nothing) OrElse _
              (Me.columnType Is GetType(String)) Then

                iRetour = String.Compare(sTxtX, sTxtY)

            ElseIf (Me.columnType Is GetType(DateTime)) Then

                iRetour = DateTime.Compare( _
                    DateTime.Parse(sTxtX), _
                    DateTime.Parse(sTxtY))

            ElseIf (Me.columnType Is GetType(Integer)) Then

                Dim i1% = Integer.Parse(sTxtX)
                Dim i2% = Integer.Parse(sTxtY)
                iRetour = (i1 - i2)

            ElseIf (Me.columnType Is GetType(Long)) Then

                Dim i1& = Long.Parse(sTxtX)
                Dim i2& = Long.Parse(sTxtY)
                iRetour = Math.Sign(i1 - i2)

            ElseIf (Me.columnType Is GetType(Single)) Then

                Dim s1! = Single.Parse(sTxtX)
                Dim s2! = Single.Parse(sTxtY)
                If s1 = s2 Then
                    iRetour = 0
                ElseIf s1 > s2 Then
                    iRetour = 1
                Else
                    iRetour = -1
                End If

            ElseIf (Me.columnType Is GetType(Double)) Then

                Dim d1 As Double = Double.Parse(sTxtX)
                Dim d2 As Double = Double.Parse(sTxtY)
                If d1 = d2 Then
                    iRetour = 0
                ElseIf d1 > d2 Then
                    iRetour = 1
                Else
                    iRetour = -1
                End If
            End If

            If Me.order = SortOrder.Descending Then iRetour = -iRetour

            Return iRetour

        End Function

        Private Function sLireColonne$(olvi As Object, iCol%)

            sLireColonne = ""
            Dim lviX As ListViewItem = DirectCast(olvi, ListViewItem)
            If iCol >= lviX.SubItems.Count Then Exit Function
            Dim lvsi As ListViewItem.ListViewSubItem = lviX.SubItems.Item(iCol)
            If String.IsNullOrEmpty(lvsi.Text) Then Exit Function
            Return lvsi.Text

        End Function

        'Private Function OKToCompare(X As Object, Y As Object) As Boolean
        '    If CompareOK(X) Then
        '        OKToCompare = Object.ReferenceEquals(X.GetType, Y.GetType)
        '    Else : OKToCompare = False
        '    End If
        'End Function
        'Private Function CompareOK(obj As Object) As Boolean
        '    CompareOK = False ' Assume not OK
        '    If obj Is Nothing Then Exit Function
        '    Dim IInfo() As Type = obj.GetType.GetInterfaces
        '    If IInfo Is Nothing Then Exit Function
        '    For Each Inter As Type In IInfo
        '        If Inter.Name.ToLower.StartsWith("icomparable") Then
        '            Return True
        '        End If
        '    Next
        'End Function

    End Class

#End Region

End Class


