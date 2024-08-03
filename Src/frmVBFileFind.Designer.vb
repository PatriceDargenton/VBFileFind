<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmVBFileFind : Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
    Public Sub New()
        MyBase.New()

        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()
        InitialiserFenetre()
    End Sub
    'La méthode substituée Dispose du formulaire pour nettoyer la liste des composants.
    Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
        If Disposing Then
            If Not components Is Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub
    'Requis par le Concepteur Windows Form
Private rbASCII As RadioButton
Private components As System.ComponentModel.IContainer = Nothing
Private WithEvents chkTexteRech As CheckBox
Private WithEvents tbTexteRech As TextBox
Private WithEvents tbFiltresFichiers As TextBox
Private groupBox1 As GroupBox
Private gpResultats As GroupBox
Private WithEvents chkSousDossiers As CheckBox
Private label1 As Label
Private label2 As Label
Private WithEvents chkDateMin As CheckBox
Private WithEvents dtpDateMin As DateTimePicker
Private WithEvents chkDateMax As CheckBox
Private WithEvents dtpDateMax As DateTimePicker
Private WithEvents tbCheminDossier As TextBox
Private WithEvents cmdParcourir As Button
Private WithEvents cmdLancer As Button
Private WithEvents cmdStop As Button
Private toolTip1 As ToolTip
Private WithEvents chkCasse As System.Windows.Forms.CheckBox
Private WithEvents rbDbleEncod As System.Windows.Forms.RadioButton
Friend WithEvents pnlTexteRech As System.Windows.Forms.Panel
Private WithEvents chkBlocNotes As System.Windows.Forms.CheckBox
Friend WithEvents cmdEnleverMenuCtx As System.Windows.Forms.Button
Friend WithEvents cmdAjouterMenuCtx As System.Windows.Forms.Button
Friend WithEvents lblInfo As System.Windows.Forms.Label
Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
Friend WithEvents tsslblBarreMessage As System.Windows.Forms.ToolStripStatusLabel
Private WithEvents rbUnicode As RadioButton
    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Il peut être modifié à l'aide du Concepteur Windows Form.
    'Ne pas le modifier à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmVBFileFind))
        Me.tbFiltresFichiers = New System.Windows.Forms.TextBox()
        Me.label1 = New System.Windows.Forms.Label()
        Me.label2 = New System.Windows.Forms.Label()
        Me.tbCheminDossier = New System.Windows.Forms.TextBox()
        Me.cmdParcourir = New System.Windows.Forms.Button()
        Me.chkSousDossiers = New System.Windows.Forms.CheckBox()
        Me.chkDateMin = New System.Windows.Forms.CheckBox()
        Me.dtpDateMin = New System.Windows.Forms.DateTimePicker()
        Me.dtpDateMax = New System.Windows.Forms.DateTimePicker()
        Me.chkDateMax = New System.Windows.Forms.CheckBox()
        Me.groupBox1 = New System.Windows.Forms.GroupBox()
        Me.tbTailleMax = New System.Windows.Forms.TextBox()
        Me.tbTailleMin = New System.Windows.Forms.TextBox()
        Me.chkTailleMax = New System.Windows.Forms.CheckBox()
        Me.chkTailleMin = New System.Windows.Forms.CheckBox()
        Me.pnlTexteRech = New System.Windows.Forms.Panel()
        Me.tbTexteRech = New System.Windows.Forms.TextBox()
        Me.chkCasse = New System.Windows.Forms.CheckBox()
        Me.rbDbleEncod = New System.Windows.Forms.RadioButton()
        Me.rbASCII = New System.Windows.Forms.RadioButton()
        Me.rbUnicode = New System.Windows.Forms.RadioButton()
        Me.chkTexteRech = New System.Windows.Forms.CheckBox()
        Me.gpResultats = New System.Windows.Forms.GroupBox()
        Me.lvResultats = New VBFileFind.SortableListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.imgLstLVH = New System.Windows.Forms.ImageList(Me.components)
        Me.cmdLancer = New System.Windows.Forms.Button()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.toolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdEnleverMenuCtx = New System.Windows.Forms.Button()
        Me.cmdAjouterMenuCtx = New System.Windows.Forms.Button()
        Me.chkBlocNotes = New System.Windows.Forms.CheckBox()
        Me.tbFiltresFichiersExclus = New System.Windows.Forms.TextBox()
        Me.lblInfo = New System.Windows.Forms.Label()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.tsslblBarreMessage = New System.Windows.Forms.ToolStripStatusLabel()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.rbDetect = New System.Windows.Forms.RadioButton()
        Me.groupBox1.SuspendLayout()
        Me.pnlTexteRech.SuspendLayout()
        Me.gpResultats.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'tbFiltresFichiers
        '
        Me.tbFiltresFichiers.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbFiltresFichiers.Location = New System.Drawing.Point(211, 39)
        Me.tbFiltresFichiers.Name = "tbFiltresFichiers"
        Me.tbFiltresFichiers.Size = New System.Drawing.Size(472, 20)
        Me.tbFiltresFichiers.TabIndex = 7
        Me.toolTip1.SetToolTip(Me.tbFiltresFichiers, "Types de fichiers ou dossiers à rechercher (*.* = tous), extension multiple possi" &
        "ble, exemple : Fich*.doc;*.txt")
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.Location = New System.Drawing.Point(156, 42)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(49, 13)
        Me.label1.TabIndex = 6
        Me.label1.Text = "Fichiers :"
        Me.label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'label2
        '
        Me.label2.AutoSize = True
        Me.label2.Location = New System.Drawing.Point(157, 16)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(48, 13)
        Me.label2.TabIndex = 3
        Me.label2.Text = "Chemin :"
        Me.label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'tbCheminDossier
        '
        Me.tbCheminDossier.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbCheminDossier.Location = New System.Drawing.Point(211, 12)
        Me.tbCheminDossier.Name = "tbCheminDossier"
        Me.tbCheminDossier.Size = New System.Drawing.Size(472, 20)
        Me.tbCheminDossier.TabIndex = 4
        Me.toolTip1.SetToolTip(Me.tbCheminDossier, "Chemin du dossier de recherche")
        '
        'cmdParcourir
        '
        Me.cmdParcourir.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdParcourir.Location = New System.Drawing.Point(689, 12)
        Me.cmdParcourir.Name = "cmdParcourir"
        Me.cmdParcourir.Size = New System.Drawing.Size(24, 21)
        Me.cmdParcourir.TabIndex = 5
        Me.cmdParcourir.Text = "..."
        Me.toolTip1.SetToolTip(Me.cmdParcourir, "Cliquer pour choisir un dossier de recherche")
        Me.cmdParcourir.UseVisualStyleBackColor = True
        '
        'chkSousDossiers
        '
        Me.chkSousDossiers.AutoSize = True
        Me.chkSousDossiers.Location = New System.Drawing.Point(12, 12)
        Me.chkSousDossiers.Name = "chkSousDossiers"
        Me.chkSousDossiers.Size = New System.Drawing.Size(124, 17)
        Me.chkSousDossiers.TabIndex = 2
        Me.chkSousDossiers.Text = "Inclure sous-dossiers"
        Me.toolTip1.SetToolTip(Me.chkSousDossiers, "Cocher pour parcourir tous les sous-dossiers, sinon limiter la recherche au seul " &
        "dossier indiqué")
        Me.chkSousDossiers.UseVisualStyleBackColor = True
        '
        'chkDateMin
        '
        Me.chkDateMin.AutoSize = True
        Me.chkDateMin.Location = New System.Drawing.Point(16, 22)
        Me.chkDateMin.Name = "chkDateMin"
        Me.chkDateMin.Size = New System.Drawing.Size(77, 17)
        Me.chkDateMin.TabIndex = 0
        Me.chkDateMin.Text = "Fichiers >="
        Me.toolTip1.SetToolTip(Me.chkDateMin, "Cocher pour limiter la recherche aux fichiers postérieurs à la date indiquée")
        Me.chkDateMin.UseVisualStyleBackColor = True
        '
        'dtpDateMin
        '
        Me.dtpDateMin.CustomFormat = "dd/MM/yyyy HH:mm"
        Me.dtpDateMin.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDateMin.Location = New System.Drawing.Point(99, 19)
        Me.dtpDateMin.Name = "dtpDateMin"
        Me.dtpDateMin.Size = New System.Drawing.Size(117, 20)
        Me.dtpDateMin.TabIndex = 1
        '
        'dtpDateMax
        '
        Me.dtpDateMax.CustomFormat = "dd/MM/yyyy HH:mm"
        Me.dtpDateMax.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDateMax.Location = New System.Drawing.Point(99, 44)
        Me.dtpDateMax.Name = "dtpDateMax"
        Me.dtpDateMax.Size = New System.Drawing.Size(117, 20)
        Me.dtpDateMax.TabIndex = 3
        '
        'chkDateMax
        '
        Me.chkDateMax.AutoSize = True
        Me.chkDateMax.Location = New System.Drawing.Point(16, 47)
        Me.chkDateMax.Name = "chkDateMax"
        Me.chkDateMax.Size = New System.Drawing.Size(77, 17)
        Me.chkDateMax.TabIndex = 2
        Me.chkDateMax.Text = "Fichiers <="
        Me.toolTip1.SetToolTip(Me.chkDateMax, "Cocher pour limiter la recherche aux fichiers antérieurs à la date indiquée")
        Me.chkDateMax.UseVisualStyleBackColor = True
        '
        'groupBox1
        '
        Me.groupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.groupBox1.Controls.Add(Me.tbTailleMax)
        Me.groupBox1.Controls.Add(Me.tbTailleMin)
        Me.groupBox1.Controls.Add(Me.chkTailleMax)
        Me.groupBox1.Controls.Add(Me.chkTailleMin)
        Me.groupBox1.Controls.Add(Me.pnlTexteRech)
        Me.groupBox1.Controls.Add(Me.chkTexteRech)
        Me.groupBox1.Controls.Add(Me.dtpDateMax)
        Me.groupBox1.Controls.Add(Me.chkDateMin)
        Me.groupBox1.Controls.Add(Me.chkDateMax)
        Me.groupBox1.Controls.Add(Me.dtpDateMin)
        Me.groupBox1.Location = New System.Drawing.Point(12, 103)
        Me.groupBox1.Name = "groupBox1"
        Me.groupBox1.Size = New System.Drawing.Size(701, 125)
        Me.groupBox1.TabIndex = 10
        Me.groupBox1.TabStop = False
        Me.groupBox1.Text = "Restrictions"
        '
        'tbTailleMax
        '
        Me.tbTailleMax.Location = New System.Drawing.Point(99, 97)
        Me.tbTailleMax.Name = "tbTailleMax"
        Me.tbTailleMax.Size = New System.Drawing.Size(117, 20)
        Me.tbTailleMax.TabIndex = 7
        Me.toolTip1.SetToolTip(Me.tbTailleMax, "Taille max. des fichiers en octets")
        '
        'tbTailleMin
        '
        Me.tbTailleMin.Location = New System.Drawing.Point(99, 74)
        Me.tbTailleMin.Name = "tbTailleMin"
        Me.tbTailleMin.Size = New System.Drawing.Size(117, 20)
        Me.tbTailleMin.TabIndex = 5
        Me.toolTip1.SetToolTip(Me.tbTailleMin, "Taille min. des fichiers en octets")
        '
        'chkTailleMax
        '
        Me.chkTailleMax.AutoSize = True
        Me.chkTailleMax.Location = New System.Drawing.Point(16, 99)
        Me.chkTailleMax.Name = "chkTailleMax"
        Me.chkTailleMax.Size = New System.Drawing.Size(77, 17)
        Me.chkTailleMax.TabIndex = 6
        Me.chkTailleMax.Text = "Fichiers <="
        Me.toolTip1.SetToolTip(Me.chkTailleMax, "Cocher pour limiter la recherche aux fichiers de taille inférieure à la taille in" &
        "diquée")
        Me.chkTailleMax.UseVisualStyleBackColor = True
        '
        'chkTailleMin
        '
        Me.chkTailleMin.AutoSize = True
        Me.chkTailleMin.Location = New System.Drawing.Point(16, 76)
        Me.chkTailleMin.Name = "chkTailleMin"
        Me.chkTailleMin.Size = New System.Drawing.Size(77, 17)
        Me.chkTailleMin.TabIndex = 4
        Me.chkTailleMin.Text = "Fichiers >="
        Me.toolTip1.SetToolTip(Me.chkTailleMin, "Cocher pour limiter la recherche aux fichiers de taille supérieure à la taille in" &
        "diquée")
        Me.chkTailleMin.UseVisualStyleBackColor = True
        '
        'pnlTexteRech
        '
        Me.pnlTexteRech.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlTexteRech.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlTexteRech.Controls.Add(Me.rbDetect)
        Me.pnlTexteRech.Controls.Add(Me.tbTexteRech)
        Me.pnlTexteRech.Controls.Add(Me.chkCasse)
        Me.pnlTexteRech.Controls.Add(Me.rbDbleEncod)
        Me.pnlTexteRech.Controls.Add(Me.rbASCII)
        Me.pnlTexteRech.Controls.Add(Me.rbUnicode)
        Me.pnlTexteRech.Location = New System.Drawing.Point(245, 34)
        Me.pnlTexteRech.Name = "pnlTexteRech"
        Me.pnlTexteRech.Size = New System.Drawing.Size(439, 65)
        Me.pnlTexteRech.TabIndex = 9
        '
        'tbTexteRech
        '
        Me.tbTexteRech.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbTexteRech.Location = New System.Drawing.Point(73, 9)
        Me.tbTexteRech.Name = "tbTexteRech"
        Me.tbTexteRech.Size = New System.Drawing.Size(354, 20)
        Me.tbTexteRech.TabIndex = 1
        Me.toolTip1.SetToolTip(Me.tbTexteRech, "Mot ou expression recherchée")
        '
        'chkCasse
        '
        Me.chkCasse.AutoSize = True
        Me.chkCasse.Location = New System.Drawing.Point(12, 11)
        Me.chkCasse.Name = "chkCasse"
        Me.chkCasse.Size = New System.Drawing.Size(55, 17)
        Me.chkCasse.TabIndex = 0
        Me.chkCasse.Text = "Casse"
        Me.toolTip1.SetToolTip(Me.chkCasse, "Cocher pour tenir compte de la casse (minuscules/majuscules)")
        Me.chkCasse.UseVisualStyleBackColor = True
        '
        'rbDbleEncod
        '
        Me.rbDbleEncod.AutoSize = True
        Me.rbDbleEncod.Location = New System.Drawing.Point(275, 35)
        Me.rbDbleEncod.Name = "rbDbleEncod"
        Me.rbDbleEncod.Size = New System.Drawing.Size(51, 17)
        Me.rbDbleEncod.TabIndex = 4
        Me.rbDbleEncod.TabStop = True
        Me.rbDbleEncod.Text = "Les 2"
        Me.toolTip1.SetToolTip(Me.rbDbleEncod, "Rechercher le mot selon les deux types d'encodage")
        Me.rbDbleEncod.UseVisualStyleBackColor = True
        '
        'rbASCII
        '
        Me.rbASCII.AutoSize = True
        Me.rbASCII.Location = New System.Drawing.Point(73, 35)
        Me.rbASCII.Name = "rbASCII"
        Me.rbASCII.Size = New System.Drawing.Size(82, 17)
        Me.rbASCII.TabIndex = 2
        Me.rbASCII.TabStop = True
        Me.rbASCII.Text = "ASCII/ANSI"
        Me.toolTip1.SetToolTip(Me.rbASCII, "Encodage simple des caractères sur un octet, qui correspond au codage ANSI indiqu" &
        "é par exemple dans le Bloc-notes de Windows (encodage par défaut du Bloc-notes)")
        Me.rbASCII.UseVisualStyleBackColor = True
        '
        'rbUnicode
        '
        Me.rbUnicode.AutoSize = True
        Me.rbUnicode.Location = New System.Drawing.Point(183, 35)
        Me.rbUnicode.Name = "rbUnicode"
        Me.rbUnicode.Size = New System.Drawing.Size(65, 17)
        Me.rbUnicode.TabIndex = 3
        Me.rbUnicode.TabStop = True
        Me.rbUnicode.Text = "Unicode"
        Me.toolTip1.SetToolTip(Me.rbUnicode, "Encodage des caractères sur deux octets (pour pouvoir tenir compte des caractères" &
        " étendus, tels que les lettres grecques, ...)")
        Me.rbUnicode.UseVisualStyleBackColor = True
        '
        'chkTexteRech
        '
        Me.chkTexteRech.AutoSize = True
        Me.chkTexteRech.Location = New System.Drawing.Point(245, 14)
        Me.chkTexteRech.Name = "chkTexteRech"
        Me.chkTexteRech.Size = New System.Drawing.Size(150, 17)
        Me.chkTexteRech.TabIndex = 8
        Me.chkTexteRech.Text = "Fichiers contenant le mot :"
        Me.toolTip1.SetToolTip(Me.chkTexteRech, "Cocher pour limiter la recherche aux fichiers contenant le mot indiqué (avec les " &
        "options précisées)")
        Me.chkTexteRech.UseVisualStyleBackColor = True
        '
        'gpResultats
        '
        Me.gpResultats.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gpResultats.Controls.Add(Me.lvResultats)
        Me.gpResultats.Location = New System.Drawing.Point(12, 302)
        Me.gpResultats.Name = "gpResultats"
        Me.gpResultats.Size = New System.Drawing.Size(701, 138)
        Me.gpResultats.TabIndex = 15
        Me.gpResultats.TabStop = False
        Me.gpResultats.Text = "Résultats"
        '
        'lvResultats
        '
        Me.lvResultats.AllowColumnReorder = True
        Me.lvResultats.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvResultats.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5})
        Me.lvResultats.FullRowSelect = True
        Me.lvResultats.HideSelection = False
        Me.lvResultats.Location = New System.Drawing.Point(6, 19)
        Me.lvResultats.Name = "lvResultats"
        Me.lvResultats.ShowItemToolTips = True
        Me.lvResultats.Size = New System.Drawing.Size(689, 113)
        Me.lvResultats.SmallImageList = Me.imgLstLVH
        Me.lvResultats.TabIndex = 0
        Me.lvResultats.UseCompatibleStateImageBehavior = False
        Me.lvResultats.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Chemin"
        Me.ColumnHeader1.Width = 212
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Taille octets"
        Me.ColumnHeader2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader2.Width = 100
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Taille"
        Me.ColumnHeader3.Width = 100
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Date de modification"
        Me.ColumnHeader4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ColumnHeader4.Width = 120
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Date d'accès"
        Me.ColumnHeader5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ColumnHeader5.Width = 120
        '
        'imgLstLVH
        '
        Me.imgLstLVH.ImageStream = CType(resources.GetObject("imgLstLVH.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imgLstLVH.TransparentColor = System.Drawing.Color.Transparent
        Me.imgLstLVH.Images.SetKeyName(0, "NonTri.bmp")
        Me.imgLstLVH.Images.SetKeyName(1, "FlecheBas.bmp")
        Me.imgLstLVH.Images.SetKeyName(2, "FlecheHaut.bmp")
        '
        'cmdLancer
        '
        Me.cmdLancer.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdLancer.Location = New System.Drawing.Point(638, 269)
        Me.cmdLancer.Name = "cmdLancer"
        Me.cmdLancer.Size = New System.Drawing.Size(75, 23)
        Me.cmdLancer.TabIndex = 0
        Me.cmdLancer.Text = "Lancer"
        Me.toolTip1.SetToolTip(Me.cmdLancer, "Lancer la recherche (Lancer/Pause/Poursuivre la recherche)")
        Me.cmdLancer.UseVisualStyleBackColor = True
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdStop.Enabled = False
        Me.cmdStop.Location = New System.Drawing.Point(557, 269)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(75, 23)
        Me.cmdStop.TabIndex = 1
        Me.cmdStop.Text = "Stop"
        Me.toolTip1.SetToolTip(Me.cmdStop, "Interrompre la recherche en cours")
        Me.cmdStop.UseVisualStyleBackColor = True
        '
        'cmdEnleverMenuCtx
        '
        Me.cmdEnleverMenuCtx.BackColor = System.Drawing.SystemColors.Control
        Me.cmdEnleverMenuCtx.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdEnleverMenuCtx.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdEnleverMenuCtx.Location = New System.Drawing.Point(271, 267)
        Me.cmdEnleverMenuCtx.Name = "cmdEnleverMenuCtx"
        Me.cmdEnleverMenuCtx.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdEnleverMenuCtx.Size = New System.Drawing.Size(25, 25)
        Me.cmdEnleverMenuCtx.TabIndex = 14
        Me.cmdEnleverMenuCtx.Text = "-"
        Me.toolTip1.SetToolTip(Me.cmdEnleverMenuCtx, "Enlever le menu contextuel (depuis Windows Vista, il faut au préalable lancer l'a" &
        "pplication en tant qu'admin.)")
        Me.cmdEnleverMenuCtx.UseVisualStyleBackColor = False
        '
        'cmdAjouterMenuCtx
        '
        Me.cmdAjouterMenuCtx.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAjouterMenuCtx.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAjouterMenuCtx.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAjouterMenuCtx.Location = New System.Drawing.Point(240, 267)
        Me.cmdAjouterMenuCtx.Name = "cmdAjouterMenuCtx"
        Me.cmdAjouterMenuCtx.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAjouterMenuCtx.Size = New System.Drawing.Size(25, 25)
        Me.cmdAjouterMenuCtx.TabIndex = 13
        Me.cmdAjouterMenuCtx.Text = "+"
        Me.toolTip1.SetToolTip(Me.cmdAjouterMenuCtx, "Ajouter un menu contextuel pour rechercher directement dans un dossier depuis l'e" &
        "xplorateur de fichiers (depuis Windows Vista, il faut au préalable lancer l'appl" &
        "ication en tant qu'admin.)")
        Me.cmdAjouterMenuCtx.UseVisualStyleBackColor = False
        '
        'chkBlocNotes
        '
        Me.chkBlocNotes.AutoSize = True
        Me.chkBlocNotes.Location = New System.Drawing.Point(12, 273)
        Me.chkBlocNotes.Name = "chkBlocNotes"
        Me.chkBlocNotes.Size = New System.Drawing.Size(145, 17)
        Me.chkBlocNotes.TabIndex = 12
        Me.chkBlocNotes.Text = "Ouvrir avec le Bloc-notes"
        Me.toolTip1.SetToolTip(Me.chkBlocNotes, "Ouvrir le fichier avec le Bloc-notes (via un double-clic dans la liste trouvée), " &
        "sinon ouvrir avec l'application associée, ou bien ouvrir le dossier dans l'explo" &
        "rateur de fichiers.")
        Me.chkBlocNotes.UseVisualStyleBackColor = True
        '
        'tbFiltresFichiersExclus
        '
        Me.tbFiltresFichiersExclus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbFiltresFichiersExclus.Location = New System.Drawing.Point(211, 65)
        Me.tbFiltresFichiersExclus.Name = "tbFiltresFichiersExclus"
        Me.tbFiltresFichiersExclus.Size = New System.Drawing.Size(472, 20)
        Me.tbFiltresFichiersExclus.TabIndex = 9
        Me.toolTip1.SetToolTip(Me.tbFiltresFichiersExclus, "Types de fichiers ou dossiers à exclure, extension multiple possible, exemple : *" &
        ".edmx")
        '
        'lblInfo
        '
        Me.lblInfo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblInfo.BackColor = System.Drawing.SystemColors.Control
        Me.lblInfo.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblInfo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInfo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInfo.Location = New System.Drawing.Point(12, 231)
        Me.lblInfo.Name = "lblInfo"
        Me.lblInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblInfo.Size = New System.Drawing.Size(700, 31)
        Me.lblInfo.TabIndex = 11
        Me.lblInfo.Text = "Informations :"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsslblBarreMessage})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 443)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(724, 22)
        Me.StatusStrip1.TabIndex = 16
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tsslblBarreMessage
        '
        Me.tsslblBarreMessage.Name = "tsslblBarreMessage"
        Me.tsslblBarreMessage.Size = New System.Drawing.Size(107, 17)
        Me.tsslblBarreMessage.Text = "tsslblBarreMessage"
        Me.tsslblBarreMessage.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'BackgroundWorker1
        '
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(161, 68)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(44, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Exclus :"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'rbDetect
        '
        Me.rbDetect.AutoSize = True
        Me.rbDetect.Location = New System.Drawing.Point(342, 35)
        Me.rbDetect.Name = "rbDetect"
        Me.rbDetect.Size = New System.Drawing.Size(66, 17)
        Me.rbDetect.TabIndex = 5
        Me.rbDetect.TabStop = True
        Me.rbDetect.Text = "Détecter"
        Me.toolTip1.SetToolTip(Me.rbDetect, "Détecter l'encodage (en utilisant l'option dédiée standard pour analyser les prem" &
        "iers octets du fichier)")
        Me.rbDetect.UseVisualStyleBackColor = True
        '
        'frmVBFileFind
        '
        Me.ClientSize = New System.Drawing.Size(724, 465)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.tbFiltresFichiersExclus)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.lblInfo)
        Me.Controls.Add(Me.cmdEnleverMenuCtx)
        Me.Controls.Add(Me.cmdAjouterMenuCtx)
        Me.Controls.Add(Me.chkBlocNotes)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.cmdLancer)
        Me.Controls.Add(Me.gpResultats)
        Me.Controls.Add(Me.groupBox1)
        Me.Controls.Add(Me.chkSousDossiers)
        Me.Controls.Add(Me.cmdParcourir)
        Me.Controls.Add(Me.label2)
        Me.Controls.Add(Me.tbCheminDossier)
        Me.Controls.Add(Me.label1)
        Me.Controls.Add(Me.tbFiltresFichiers)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MinimumSize = New System.Drawing.Size(485, 490)
        Me.Name = "frmVBFileFind"
        Me.Text = "VBFileFind"
        Me.groupBox1.ResumeLayout(False)
        Me.groupBox1.PerformLayout()
        Me.pnlTexteRech.ResumeLayout(False)
        Me.pnlTexteRech.PerformLayout()
        Me.gpResultats.ResumeLayout(False)
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lvResultats As VBFileFind.SortableListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents imgLstLVH As System.Windows.Forms.ImageList
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Private WithEvents Label3 As System.Windows.Forms.Label
    Private WithEvents tbFiltresFichiersExclus As System.Windows.Forms.TextBox
    Private WithEvents tbTailleMax As TextBox
    Private WithEvents tbTailleMin As TextBox
    Private WithEvents chkTailleMax As CheckBox
    Private WithEvents chkTailleMin As CheckBox
    Private WithEvents rbDetect As RadioButton
#End Region
End Class