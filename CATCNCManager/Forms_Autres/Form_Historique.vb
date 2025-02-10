Public Class Form_Historique

    'Variables
    Private Tab_Chaines(0), Nom_Maitre_Local, Type_Historique_Local, Tab_Filtre_DGV_Local(,) As String
    Private Num_Enreg_Actuel, Num_Enreg_Maxi_CSV As Integer
    Private Derniere_Ligne_DGV_Auto As Boolean = False
    Private DGV_Maitresse_Local As DataGridView
    Private Form_Filtre_Locale As Form_Filtre

    '------------------------------------------------- Action sur événement ---------------------------------------------------

    'Action sur création nouvelle Form

    Public Sub New(Optional ByRef DGV_Maitresse As DataGridView = Nothing, _
                   Optional ByVal Nom_Maitre As String = Nothing, _
                   Optional ByVal Type_Historique As String = Nothing)

        'Initialisation composant
        InitializeComponent()

        'Paramètrage des colonnes DGV vérrouillées
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Numéro enregistrement maître")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Numéro enregistrement")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Date modification")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Auteur modification")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Date création")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Auteur création")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)

        'Si New Form créée sans les attributs (si créée juste pour récupérer des infos)
        If Not DGV_Maitresse Is Nothing And Not Nom_Maitre Is Nothing And Not Type_Historique Is Nothing Then

            'Valuation variables
            DGV_Maitresse_Local = DGV_Maitresse
            Nom_Maitre_Local = Nom_Maitre
            Type_Historique_Local = Type_Historique
            Tab_Chaines(0) = Dossier_Reseau & "\" & Type_Historique & "\Historique_" & _
                Nom_Maitre_Local & ".csv" 'Chemin fichier

            'Ouverture Form et remplissage DGV et si erreur, sortie de Sub
            If Fonctions_Form.Ouverture_Form_Avec_DGV _
            (Me, Me.DataGridView1, Tab_Chaines, Num_Enreg_Actuel, _
                Num_Enreg_Maxi_CSV, Derniere_Ligne_DGV_Auto, , , ) = False Then Exit Sub

            'Ajoute filtre par défaut dans Tab et relecture de tout le filtre
            Fonctions_Filtres_DGV.Ajout_Filtre_Et_Refiltrage_DGV _
                (Tab_Filtre_DGV_Local, "Numéro enregistrement maître", 1, _
                 DGV_Maitresse_Local.CurrentRow.Cells.Item _
                 (Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                  (DGV_Maitresse_Local, _
                   "Numéro enregistrement")).Value, True, Me.DataGridView1)

            'Ajout nom du Listing dans Form
            Me.Text = Me.Text & ": " & Nom_Maitre
        End If
    End Sub

    'Action sur fermeture Form

    Private Sub Form_Historique_Prog_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) _
        Handles MyBase.FormClosed

        'Fermeture Form
        Me.Dispose()
    End Sub

    'Gestion fenêtre si click DGV

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As DataGridViewCellEventArgs) _
        Handles DataGridView1.CellContentClick

        'Si ce n'est pas la dernière ligne et si pas en mode visualisation
        If Fonctions_DGV.Ctrl_Ligne_DGV_Niv_Ouverture(Me.DataGridView1) = 1 Then

            'Valeur de condition à trouver
            Select Case e.ColumnIndex

                'Si colonne Enregistrer
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Enregistrer ligne")

                    'Enregistrement ligne et si erreur, sortie de Sub
                    If Fonctions_DGV.Enregistrer_Ligne_DGV_Et_CSV(Me.DataGridView1, Tab_Chaines(0), _
                                                    Num_Enreg_Actuel, Num_Enreg_Maxi_CSV, ) = False Then Exit Sub

                    'Si colonne Supprimer
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Supprimer ligne")

                    'Suppression ligne et si erreur, sortie de Sub
                    If Fonctions_DGV.Sup_Ligne_DGV_Et_CSV(Me.DataGridView1, Tab_Chaines(0), _
                                                    Num_Enreg_Actuel, Num_Enreg_Maxi_CSV, ) = False Then Exit Sub
            End Select

            'Si click bouton dernière ligne et si pas en mode visualisation
        ElseIf Fonctions_DGV.Ctrl_Ligne_DGV_Niv_Ouverture(Me.DataGridView1) = 2 Then

            'Ajout de ligne
            Me.DataGridView1.Rows.Add()
        End If
    End Sub

    'Validation des données DGV

    Private Sub DataGridView1_CellValidating(ByVal sender As System.Object, ByVal e As DataGridViewCellValidatingEventArgs) _
        Handles DataGridView1.CellValidating

        'Si ce n'est pas la dernière ligne et si pas en mode visualisation
        If Fonctions_DGV.Ctrl_Ligne_DGV_Niv_Ouverture(Me.DataGridView1) = 1 Then

            'Si résultat des contrôles = False, refuse la validation
            If Fonctions_DGV.Ctrl_Chaine_DGV _
                (Me.DataGridView1, e.FormattedValue.ToString, e.ColumnIndex, True) = False Then e.Cancel = True
        End If
    End Sub

    'Ecriture dernière ligne DGV

    Private Sub DataGridView1_RowsAdded(ByVal sender As System.Object, ByVal e As DataGridViewRowsAddedEventArgs) _
        Handles DataGridView1.RowsAdded

        'Si dernière ligne DGV auto = True
        If Derniere_Ligne_DGV_Auto = True Then

            'Déclaration variable
            Dim Tab_Chaines_Supplementaires(0, 1) As String

            'Valuation variable
            Tab_Chaines_Supplementaires(0, 0) = Tab_Filtre_DGV_Local(0, 0)
            Tab_Chaines_Supplementaires(0, 1) = Tab_Filtre_DGV_Local(0, 2)

            'Ecriture dernière ligne dans DGV
            Fonctions_DGV.Ecriture_Derniere_Ligne_DGV _
            (Me.DataGridView1, Num_Enreg_Actuel, , , Tab_Chaines_Supplementaires, )
        End If
    End Sub

    '---------------------------------------------------- Filtres DGV -----------------------------------------------------

    'Si click bouton filtre

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click

        'Si Form Filtre = Nothing
        If Form_Filtre_Locale Is Nothing Then

            'Création nouvelle Form Filtre
            Form_Filtre_Locale = New Form_Filtre(Me.DataGridView1, Tab_Filtre_DGV_Local)

            'Ouverture Form filtre
            Form_Filtre_Locale.ShowDialog()

            'Sinon
        Else

            'Ouverture Form filtre existante
            Form_Filtre_Locale.ShowDialog()
        End If
    End Sub

    'Si click bouton suppression filtre

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click

        'Vidage variables
        Tab_Filtre_DGV_Local = Nothing
        Form_Filtre_Locale = Nothing

        'Ajoute filtre par défaut dans Tab et relecture de tout le filtre
        Fonctions_Filtres_DGV.Ajout_Filtre_Et_Refiltrage_DGV _
            (Tab_Filtre_DGV_Local, "Numéro enregistrement maître", 1, _
             DGV_Maitresse_Local.CurrentRow.Cells.Item _
             (Fonctions_DGV.Renvoi_Num_Colonne_DGV _
              (DGV_Maitresse_Local, _
               "Numéro enregistrement")).Value, True, Me.DataGridView1)
    End Sub
End Class