Public Class Form_Listing_Outils_Fraisage

    'Variables
    Private Num_Enreg_Actuel, Num_Enreg_Maxi_CSV As Integer
    Private Derniere_Ligne_DGV_Auto As Boolean = False
    Private Tab_Chaines(6), Type_Outil_Fraisage_Principal_Local As String
    Private Bouton_Ouverture_Local As System.Windows.Forms.Button
    Private ComboBox_Selection_Local As System.Windows.Forms.ComboBox
    Private Form_Filtre_Locale As Form_Filtre

    '------------------------------------------------- Action sur événement ---------------------------------------------------

    'Action sur création nouvelle Form

    Public Sub New(Optional ByRef Bouton_Ouverture As System.Windows.Forms.Button = Nothing, _
                   Optional ByRef ComboBox_Selection As System.Windows.Forms.ComboBox = Nothing, _
                   Optional ByRef Type_Outil_Fraisage_Principal As String = Nothing)

        'Initialisation composant
        InitializeComponent()

        'Paramètrage des colonnes DGV vérrouillées
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Numéro enregistrement")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Date modification")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Auteur modification")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Date création")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Auteur création")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        'Paramètrage colonnes param DGV
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Nombre d'outils neufs en stock")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvStock)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Nombre d'outils utilisés en stock")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvStock)

        'Si New Form créée sans les attributs (si créée juste pour récupérer des infos)
        If Not Bouton_Ouverture Is Nothing And Not ComboBox_Selection Is Nothing Then

            'Valuation variables
            Bouton_Ouverture_Local = Bouton_Ouverture
            ComboBox_Selection_Local = ComboBox_Selection
            Type_Outil_Fraisage_Principal_Local = Type_Outil_Fraisage_Principal
            Tab_Chaines(0) = Dossier_Reseau & "\Listings_Outils\Listing_" & ComboBox_Selection_Local.Text & ".csv" 'Chemin fichier Listing sélectionné
            Tab_Chaines(1) = ComboBox_Selection_Local.Text 'Nom du fichier Listing sélectionné
            Tab_Chaines(2) = Dossier_Reseau & "\Config_Outils\Fichier_Config_Outils_Fraisage_Stds.csv" 'Chemin fichier de config DGV
            Tab_Chaines(3) = "Type outil standard" 'Nom colonne maîtresse fichier de config DGV
            Tab_Chaines(4) = Type_Outil_Fraisage_Principal_Local 'Chaîne à rechercher dans colonne maîtresse fichier de config DGV
            Tab_Chaines(5) = "Outil à bout sphérique" 'Nom première colonne à configurer
            Tab_Chaines(6) = "Nombre de lèvres" 'Nom dernière colonne à configurer

            'Ouverture Form et remplissage DGV et si erreur, sortie de Sub
            If Fonctions_Form.Ouverture_Form_Avec_DGV(Me, Me.DataGridView1, Tab_Chaines, Num_Enreg_Actuel, _
                                                      Num_Enreg_Maxi_CSV, Derniere_Ligne_DGV_Auto, , , ) = False Then Exit Sub

            'Gestion ComboBox et Button de sélection
            Fonctions_Form.Gestion_Bouton_Et_ComboBox_Form _
                (Bouton_Ouverture_Local, 1, ComboBox_Selection_Local, Tab_Chaines(1))
        End If
    End Sub

    'Action sur fermeture Form

    Private Sub Form_Listing_Outils_Fraisage_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) _
        Handles MyBase.FormClosed

        'Traitement fermeture Form
        Fonctions_Form.Fermeture_Form_Enfant(Me, False, Derniere_Ligne_DGV_Auto)

        'Gestion ComboBox et Button de sélection
        Fonctions_Form.Gestion_Bouton_Et_ComboBox_Form _
            (Bouton_Ouverture_Local, 2, ComboBox_Selection_Local, Tab_Chaines(1))
    End Sub

    'Gestion fenêtre si click dans DGV

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
                                                    Num_Enreg_Actuel, Num_Enreg_Maxi_CSV, ) = _
                                                False Then Exit Sub

                    'Si colonne Supprimer
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Supprimer ligne")

                    'Suppression ligne et si erreur, sortie de Sub
                    If Fonctions_DGV.Sup_Ligne_DGV_Et_CSV(Me.DataGridView1, Tab_Chaines(0), _
                                                    Num_Enreg_Actuel, Num_Enreg_Maxi_CSV, ) = _
                                                False Then Exit Sub

                    'Si la celule active se trouve dans la colonne Envoyer vers CATIA
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Envoyer vers CATIA")

                    'Contrôle de chaîne de toutes les colonnes DGV
                    If Fonctions_DGV.Ctrl_Validation_Ligne_DGV(Me.DataGridView1, True) = False Then Exit Sub

                    'Création nouvelle Form
                    Dim New_Form As Form = New Form_Listing_Outils_Vers_CATIA _
                                           (Me.DataGridView1, Tab_Chaines)

                    'Visualisation Form
                    New_Form.ShowDialog()

                    'Si la celule active se trouve dans la colonne Recevoir de CATIA
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Recevoir de CATIA")

                    'Création nouvelle Form
                    Dim New_Form As Form = New Form_CATIA_Vers_Listing_Outils _
                                           (Me.DataGridView1, Tab_Chaines)

                    'Visualisation Form
                    New_Form.ShowDialog()

                    'Si colonne numéro de programme
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Ouvrir historique")

                    'Création nouvelle Form
                    Dim New_Form As Form = New Form_Historique(Me.DataGridView1, Tab_Chaines(1), "Historique_Outils")

                    'Visualisation Form
                    New_Form.ShowDialog()
            End Select

            'Si click bouton dernière ligne et si pas en mode visualisation
        ElseIf Fonctions_DGV.Ctrl_Ligne_DGV_Niv_Ouverture(Me.DataGridView1) = 2 Then

            'Ajout de ligne
            Me.DataGridView1.Rows.Add()
        End If
    End Sub

    'Entrée en édition dans case DGV

    Private Sub DataGridView1_CellEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) _
        Handles DataGridView1.CellEnter

        'Entrée en édition
        Fonctions_DGV.Entree_Edition_Case_DGV(Me.DataGridView1, e)
    End Sub

    'Sortie d'édition dans case DGV

    Private Sub DataGridView1_CellValidated(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) _
        Handles DataGridView1.CellValidated

        'Entrée en édition
        Fonctions_DGV.Sortie_Edition_Case_DGV(Me.DataGridView1, e)
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

        'Si en mode modification
        If Derniere_Ligne_DGV_Auto = True Then

            'Ecriture dernière ligne dans DGV
            Fonctions_DGV.Ecriture_Derniere_Ligne_DGV(Me.DataGridView1, Num_Enreg_Actuel, , , , )
        End If
    End Sub

    '---------------------------------------------------- Filtres DGV -----------------------------------------------------

    'Si click bouton filtre

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click

        'Si Form Filtre = Nothing
        If Form_Filtre_Locale Is Nothing Then

            'Création nouvelle Form Filtre
            Form_Filtre_Locale = New Form_Filtre(Me.DataGridView1, )

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

        'Vidage variable
        Form_Filtre_Locale = Nothing

        'Relecture de tous les filtres
        Fonctions_Filtres_DGV.Relecture_Tous_Filtres_DGV(Me.DataGridView1, , , )
    End Sub
End Class