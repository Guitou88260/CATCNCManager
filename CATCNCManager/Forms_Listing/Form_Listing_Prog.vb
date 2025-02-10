Public Class Form_Listing_Prog

    'Variables
    Private Num_Prog_Actuel, Index_Machine, Tab_Chaines(6), Tab_Config_Colonnes(1) As String
    Private Num_Enreg_Actuel, Num_Enreg_Maxi_CSV As Integer
    Private Derniere_Ligne_DGV_Auto As Boolean = False
    Private Bouton_Ouverture_Local As System.Windows.Forms.Button
    Private ComboBox_Selection_Local As System.Windows.Forms.ComboBox
    Private Form_Filtre_Locale As Form_Filtre

    '------------------------------------------------- Action sur événement ---------------------------------------------------

    'Action sur création nouvelle Form
    Public Sub New(Optional ByRef Bouton_Ouverture As System.Windows.Forms.Button = Nothing,
                   Optional ByRef ComboBox_Selection As System.Windows.Forms.ComboBox = Nothing)

        'Initialisation composant
        InitializeComponent()

        'Paramètrage des colonnes DGV vérrouillées
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Numéro enregistrement")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Date modification")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Auteur modification")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Date création")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Auteur création")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        'Paramètrage colonnes réf DGV
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "RefA")).HeaderText = ParamsApp.ColRefA
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "RefB")).HeaderText = ParamsApp.ColRefB
        'Paramètrage colonnes param DGV
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "ParamA")).ToolTipText = ParamsApp.VerifColParamA
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "ParamA")).HeaderText = ParamsApp.ColParamA
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "ParamB")).ToolTipText = ParamsApp.VerifColParamB
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "ParamB")).HeaderText = ParamsApp.ColParamB
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "ParamC")).ToolTipText = ParamsApp.VerifColParamC
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "ParamC")).HeaderText = ParamsApp.ColParamC
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "ParamD")).ToolTipText = ParamsApp.VerifColParamD
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "ParamD")).HeaderText = ParamsApp.ColParamD
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "ParamE")).ToolTipText = ParamsApp.VerifColParamE
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "ParamE")).HeaderText = ParamsApp.ColParamE
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "ParamF")).ToolTipText = ParamsApp.VerifColParamF
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "ParamF")).HeaderText = ParamsApp.ColParamF
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "ParamG")).ToolTipText = ParamsApp.VerifColParamG
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "ParamG")).HeaderText = ParamsApp.ColParamG
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "ParamH")).ToolTipText = ParamsApp.VerifColParamH
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "ParamH")).HeaderText = ParamsApp.ColParamH

        'Si New Form créée sans les attributs (si créée juste pour récupérer des infos)
        If Bouton_Ouverture IsNot Nothing And ComboBox_Selection IsNot Nothing Then

            'Valuation variables
            Bouton_Ouverture_Local = Bouton_Ouverture
            ComboBox_Selection_Local = ComboBox_Selection
            Tab_Chaines(0) = Dossier_Reseau & "\Listings_Prog\Listing_" & ComboBox_Selection_Local.Text & ".csv" 'Chemin fichier Listing sélectionné
            Tab_Chaines(1) = ComboBox_Selection_Local.Text 'Nom du fichier Listing sélectionné
            Tab_Chaines(2) = Dossier_Reseau & "\Config_Machines\Fichier_Config_Machines.csv" 'Chemin fichier de config DGV
            Tab_Chaines(3) = "Machine" 'Nom colonne maîtresse fichier de config DGV
            Tab_Chaines(4) = ComboBox_Selection_Local.Text 'Chaîne à rechercher dans colonne maîtresse fichier de config DGV
            Tab_Chaines(5) = ParamsApp.ColParamA 'Nom première colonne à configurer
            Tab_Chaines(6) = ParamsApp.ColParamH 'Nom dernière colonne à configurer

            'Déclaration variable
            Dim Colonnes_A_Desactiver(0, 2) As String

            'Valuation variable
            Colonnes_A_Desactiver(0, 0) = "Envoyer vers CATIA" 'Colonne DGV à désactiver
            Colonnes_A_Desactiver(0, 1) = "Machine" 'Colonne Tab dans laquelle chercher
            Colonnes_A_Desactiver(0, 2) = "Non-programmée sous CATIA" 'Colonne Tab chaîne à renvoyer

            'Ouverture Form et remplissage DGV et si erreur, sortie de Sub
            If Fonctions_Form.Ouverture_Form_Avec_DGV _
                (Me, Me.DataGridView1, Tab_Chaines, Num_Enreg_Actuel,
                 Num_Enreg_Maxi_CSV, Derniere_Ligne_DGV_Auto, Num_Prog_Actuel,
                 Index_Machine, Colonnes_A_Desactiver) = False Then Exit Sub

            ' Blocage des cellules si check
            Call GestionCellLocal()

            'Gestion ComboBox et Button de sélection
            Fonctions_Form.Gestion_Bouton_Et_ComboBox_Form _
                (Bouton_Ouverture_Local, 1, ComboBox_Selection_Local, Tab_Chaines(1))
        End If
    End Sub

    'Action sur fermeture Form

    Private Sub Form_Listing_Prog_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) _
        Handles MyBase.FormClosed

        'Traitement fermeture Form
        Fonctions_Form.Fermeture_Form_Enfant(Me, False, Derniere_Ligne_DGV_Auto)

        'Gestion ComboBox et Button de sélection
        Fonctions_Form.Gestion_Bouton_Et_ComboBox_Form _
            (Bouton_Ouverture_Local, 2, ComboBox_Selection_Local, Tab_Chaines(1))
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
                    If Fonctions_DGV.Enregistrer_Ligne_DGV_Et_CSV(Me.DataGridView1, Tab_Chaines(0),
                                                    Num_Enreg_Actuel, Num_Enreg_Maxi_CSV,
                                                    Num_Prog_Actuel) = False Then Exit Sub

                    'Si colonne Supprimer
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Supprimer ligne")

                    'Suppression ligne et si erreur, sortie de Sub
                    If Fonctions_DGV.Sup_Ligne_DGV_Et_CSV(Me.DataGridView1, Tab_Chaines(0),
                                                    Num_Enreg_Actuel, Num_Enreg_Maxi_CSV,
                                                    Num_Prog_Actuel) = False Then Exit Sub

                    'Si Envoyer vers CATIA
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Envoyer vers CATIA")

                    'Contrôle de chaîne de toutes les colonnes DGV
                    If Fonctions_DGV.Ctrl_Validation_Ligne_DGV(Me.DataGridView1, True) = False Then Exit Sub

                    'Création nouvelle Form
                    Dim New_Form As Form = New Form_Listing_Prog_Vers_CATIA(Me.DataGridView1, Tab_Chaines(1))

                    'Visualisation Form
                    New_Form.ShowDialog()

                    'Si colonne numéro de programme
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Numéro programme")

                    'Création nouvelle Form
                    Dim New_Form As Form = New Form_Choix_Num_Prog_Listing(Me.DataGridView1)

                    'Visualisation Form
                    New_Form.ShowDialog()

                    'Recherche numéro de programme le plus grand dans DGV et valuation variable
                    Num_Prog_Actuel = Fonctions_DGV.Renvoi_Valeur_Maxi_Colonne_DGV(Me.DataGridView1, "Numéro programme", True)

                    'Si colonne check pièce dans PLM
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Pièce dans PLM")

                    'Gestion CheckBox et colonnes concernées
                    Fonctions_DGV.Gestion_Dynamique_CheckBox _
                        (Me.DataGridView1, "Pièce dans PLM", Tab_Config_Colonnes, True, 1,
                         True, Me.DataGridView1.CurrentRow.Index)

                    'Si colonne numéro de programme
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Ouvrir historique")

                    'Création nouvelle Form
                    Dim New_Form As Form = New Form_Historique(Me.DataGridView1, Tab_Chaines(1), "Historique_Prog")

                    'Visualisation Form
                    New_Form.ShowDialog()
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

        'Si en mode modification
        If Derniere_Ligne_DGV_Auto = True Then

            'Déclaration et valuation variable
            Dim Tab_Noms_Colonnes_Check(0) As String
            Tab_Noms_Colonnes_Check(0) = "Pièce dans PLM"

            'Ecriture dernière ligne dans DGV
            Fonctions_DGV.Ecriture_Derniere_Ligne_DGV _
                (Me.DataGridView1, Num_Enreg_Actuel, Num_Prog_Actuel, Index_Machine, , Tab_Noms_Colonnes_Check)

            ' Blocage des cellules si check
            Call GestionCellLocal()
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


    ' Fonction gestion des CheckBox avec vérrouillage cellules
    Private Sub GestionCellLocal()

        'Valuation variable
        Tab_Config_Colonnes(0) = Tab_Chaines(5)
        Tab_Config_Colonnes(1) = Tab_Chaines(6)

        'Vidage et blocage des colonnes si Check
        Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
            (Me.DataGridView1, "Pièce dans PLM", Tab_Config_Colonnes, False, 1, True, )
    End Sub
End Class