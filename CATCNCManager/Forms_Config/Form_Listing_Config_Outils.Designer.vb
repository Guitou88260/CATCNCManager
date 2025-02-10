<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Listing_Config_Outils
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form_Listing_Config_Outils))
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Column12 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column13 = New System.Windows.Forms.DataGridViewButtonColumn()
        Me.Column14 = New System.Windows.Forms.DataGridViewButtonColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column17 = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.Column1 = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.Column5 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.Column19 = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.Column4 = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.Column27 = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.Column20 = New System.Windows.Forms.DataGridViewButtonColumn()
        Me.Column16 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column29 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column15 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column30 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton2 = New System.Windows.Forms.ToolStripButton()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AllowUserToResizeRows = False
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.BackgroundColor = System.Drawing.SystemColors.ActiveCaption
        Me.DataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.DataGridView1.CausesValidation = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.ColumnHeadersHeight = 50
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column12, Me.Column13, Me.Column14, Me.Column2, Me.Column17, Me.Column3, Me.Column1, Me.Column5, Me.Column19, Me.Column4, Me.Column27, Me.Column20, Me.Column16, Me.Column29, Me.Column15, Me.Column30})
        Me.DataGridView1.EnableHeadersVisualStyles = False
        Me.DataGridView1.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersWidth = 45
        Me.DataGridView1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DataGridView1.ShowCellErrors = False
        Me.DataGridView1.ShowEditingIcon = False
        Me.DataGridView1.ShowRowErrors = False
        Me.DataGridView1.Size = New System.Drawing.Size(762, 472)
        Me.DataGridView1.TabIndex = 0
        '
        'Column12
        '
        Me.Column12.HeaderText = "Numéro enregistrement"
        Me.Column12.MinimumWidth = 6
        Me.Column12.Name = "Column12"
        Me.Column12.ReadOnly = True
        Me.Column12.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column12.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column12.Width = 110
        '
        'Column13
        '
        Me.Column13.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Column13.HeaderText = "Enregistrer ligne"
        Me.Column13.MinimumWidth = 6
        Me.Column13.Name = "Column13"
        Me.Column13.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column13.Text = "Enregistrer"
        Me.Column13.UseColumnTextForButtonValue = True
        Me.Column13.Width = 110
        '
        'Column14
        '
        Me.Column14.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Column14.HeaderText = "Supprimer ligne"
        Me.Column14.MinimumWidth = 6
        Me.Column14.Name = "Column14"
        Me.Column14.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column14.Text = "Supprimer"
        Me.Column14.UseColumnTextForButtonValue = True
        Me.Column14.Width = 110
        '
        'Column2
        '
        Me.Column2.HeaderText = "Outil"
        Me.Column2.MinimumWidth = 6
        Me.Column2.Name = "Column2"
        Me.Column2.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column2.ToolTipText = "(2) Longueur de chaîne égale à ""0"" caractère(s) interdite, (7) Caractère(s) autre" &
    " que ""1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_"" interdit(" &
    "s)"
        Me.Column2.Width = 110
        '
        'Column17
        '
        Me.Column17.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Column17.HeaderText = "Type outil fraisage principal"
        Me.Column17.Items.AddRange(New Object() {"", "Foret", "Foret à pointer", "Foret à centrer", "Fraise à chanfreiner", "Taraud", "Foret à étager", "Outil à chanfreiner 2 côtés", "Outil à percer-chanfreiner", "Barre à aléser", "Alésoir", "Fraise 3 tailles", "Fraise à fileter", "Fraise à lamer", "Fraise à surfacer", "Fraise 2 tailles", "Fraise conique", "Barrel Mill", "Outil stylet à bille", "Outil stylet cylindrique"})
        Me.Column17.MinimumWidth = 6
        Me.Column17.Name = "Column17"
        Me.Column17.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column17.ToolTipText = "(2) Longueur de chaîne égale à ""0"" caractère(s) interdite"
        Me.Column17.Width = 110
        '
        'Column3
        '
        Me.Column3.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Column3.HeaderText = "Type outil tournage principal"
        Me.Column3.Items.AddRange(New Object() {"", "Porte-plaquette extérieur", "Porte-plaquette intérieur", "Porte-plaquette à gorge extérieur", "Porte-plaquette à gorge intérieur", "Porte-plaquette à gorge frontal", "Porte-plaquette à fileter extérieur", "Porte-plaquette à fileter intérieur"})
        Me.Column3.MinimumWidth = 6
        Me.Column3.Name = "Column3"
        Me.Column3.ToolTipText = "(2) Longueur de chaîne égale à ""0"" caractère(s) interdite"
        Me.Column3.Width = 110
        '
        'Column1
        '
        Me.Column1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Column1.HeaderText = "Type plaquette outil principal"
        Me.Column1.Items.AddRange(New Object() {"", "Plaquette ronde", "Plaquette rhombique", "Plaquette carrée", "Plaquette triangulaire", "Plaquette trigone", "Plaquette de gorge", "Plaquette de filetage"})
        Me.Column1.MinimumWidth = 6
        Me.Column1.Name = "Column1"
        Me.Column1.ToolTipText = "(2) Longueur de chaîne égale à ""0"" caractère(s) interdite"
        Me.Column1.Width = 110
        '
        'Column5
        '
        Me.Column5.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Column5.HeaderText = "Conversion outil"
        Me.Column5.MinimumWidth = 6
        Me.Column5.Name = "Column5"
        Me.Column5.Width = 110
        '
        'Column19
        '
        Me.Column19.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Column19.HeaderText = "Type outil fraisage secondaire"
        Me.Column19.Items.AddRange(New Object() {"", "Foret", "Foret à pointer", "Foret à centrer", "Fraise à chanfreiner", "Taraud", "Foret à étager", "Outil à chanfreiner 2 côtés", "Outil à percer-chanfreiner", "Barre à aléser", "Alésoir", "Fraise 3 tailles", "Fraise à fileter", "Fraise à lamer", "Fraise à surfacer", "Fraise 2 tailles", "Fraise conique", "Barrel Mill", "Outil stylet à bille", "Outil stylet cylindrique"})
        Me.Column19.MinimumWidth = 6
        Me.Column19.Name = "Column19"
        Me.Column19.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column19.Width = 110
        '
        'Column4
        '
        Me.Column4.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Column4.HeaderText = "Type outil tournage secondaire"
        Me.Column4.Items.AddRange(New Object() {"", "Porte-plaquette extérieur", "Porte-plaquette intérieur", "Porte-plaquette à gorge extérieur", "Porte-plaquette à gorge intérieur", "Porte-plaquette à gorge frontal", "Porte-plaquette à fileter extérieur", "Porte-plaquette à fileter intérieur"})
        Me.Column4.MinimumWidth = 6
        Me.Column4.Name = "Column4"
        Me.Column4.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column4.Width = 110
        '
        'Column27
        '
        Me.Column27.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Column27.HeaderText = "Type plaquette outil secondaire"
        Me.Column27.Items.AddRange(New Object() {"", "Plaquette ronde", "Plaquette rhombique", "Plaquette carrée", "Plaquette triangulaire", "Plaquette trigone", "Plaquette de gorge", "Plaquette de filetage"})
        Me.Column27.MinimumWidth = 6
        Me.Column27.Name = "Column27"
        Me.Column27.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column27.Width = 110
        '
        'Column20
        '
        Me.Column20.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Column20.HeaderText = "Configuration conversion outil"
        Me.Column20.MinimumWidth = 6
        Me.Column20.Name = "Column20"
        Me.Column20.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column20.Text = "Configurer"
        Me.Column20.UseColumnTextForButtonValue = True
        Me.Column20.Width = 110
        '
        'Column16
        '
        Me.Column16.HeaderText = "Date modification"
        Me.Column16.MinimumWidth = 6
        Me.Column16.Name = "Column16"
        Me.Column16.ReadOnly = True
        Me.Column16.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column16.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column16.Width = 110
        '
        'Column29
        '
        Me.Column29.HeaderText = "Auteur modification"
        Me.Column29.MinimumWidth = 6
        Me.Column29.Name = "Column29"
        Me.Column29.ReadOnly = True
        Me.Column29.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column29.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column29.Width = 110
        '
        'Column15
        '
        Me.Column15.HeaderText = "Date création"
        Me.Column15.MinimumWidth = 6
        Me.Column15.Name = "Column15"
        Me.Column15.ReadOnly = True
        Me.Column15.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column15.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column15.Width = 110
        '
        'Column30
        '
        Me.Column30.HeaderText = "Auteur création"
        Me.Column30.MinimumWidth = 6
        Me.Column30.Name = "Column30"
        Me.Column30.ReadOnly = True
        Me.Column30.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column30.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column30.Width = 110
        '
        'ToolStrip1
        '
        Me.ToolStrip1.AllowMerge = False
        Me.ToolStrip1.AutoSize = False
        Me.ToolStrip1.CanOverflow = False
        Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.Right
        Me.ToolStrip1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.ImageScalingSize = New System.Drawing.Size(22, 22)
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButton1, Me.ToolStripButton2})
        Me.ToolStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.VerticalStackWithOverflow
        Me.ToolStrip1.Location = New System.Drawing.Point(762, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Padding = New System.Windows.Forms.Padding(0)
        Me.ToolStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
        Me.ToolStrip1.ShowItemToolTips = False
        Me.ToolStrip1.Size = New System.Drawing.Size(32, 472)
        Me.ToolStrip1.Stretch = True
        Me.ToolStrip1.TabIndex = 1
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.AutoToolTip = False
        Me.ToolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton1.Image = Global.CATCNCManager.My.Resources.Resources.Filtrer
        Me.ToolStripButton1.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Margin = New System.Windows.Forms.Padding(2)
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(27, 26)
        Me.ToolStripButton1.Text = "ToolStripButton1"
        '
        'ToolStripButton2
        '
        Me.ToolStripButton2.AutoToolTip = False
        Me.ToolStripButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton2.Image = Global.CATCNCManager.My.Resources.Resources.Supprimer_Filtre_DGV
        Me.ToolStripButton2.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.ToolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton2.Margin = New System.Windows.Forms.Padding(2)
        Me.ToolStripButton2.Name = "ToolStripButton2"
        Me.ToolStripButton2.Size = New System.Drawing.Size(27, 26)
        Me.ToolStripButton2.Text = "ToolStripButton1"
        '
        'Form_Listing_Config_Outils
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.CausesValidation = False
        Me.ClientSize = New System.Drawing.Size(794, 472)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form_Listing_Config_Outils"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Listing des configurations d'outils"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripButton1 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButton2 As System.Windows.Forms.ToolStripButton
    Friend WithEvents Column12 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column13 As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents Column14 As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column17 As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Column5 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Column19 As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Column4 As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Column27 As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Column20 As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents Column16 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column29 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column15 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column30 As System.Windows.Forms.DataGridViewTextBoxColumn

End Class
