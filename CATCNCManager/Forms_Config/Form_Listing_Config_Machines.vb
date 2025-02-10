Public Class Form_Listing_Config_Machines

    'Variables
    Private Num_Enreg_Actuel, Num_Enreg_Maxi_CSV As Integer
    Private Derniere_Ligne_DGV_Auto As Boolean = False
    Private Tab_Chaines(0), tabConfgMachTyp(1), confgFOCAS(1) As String
    Private Bouton_Ouverture_Local As System.Windows.Forms.Button
    Private Form_Filtre_Locale As Form_Filtre

    '------------------------------------------------- Action sur événement ---------------------------------------------------

    'Action sur création nouvelle Form

    Public Sub New(Optional ByRef Bouton_Ouverture As System.Windows.Forms.Button = Nothing)

        'Initialisation composant
        InitializeComponent()

        'Paramètrage des colonnes DGV vérrouillées
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Numéro enregistrement")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Date modification")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Auteur modification")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Date création")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Auteur création")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        'Paramètrage colonnes param DGV
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "ParamA")).HeaderText = ParamsApp.ColParamA
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "ParamB")).HeaderText = ParamsApp.ColParamB
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "ParamC")).HeaderText = ParamsApp.ColParamC
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "ParamD")).HeaderText = ParamsApp.ColParamD
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "ParamE")).HeaderText = ParamsApp.ColParamE
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "ParamF")).HeaderText = ParamsApp.ColParamF
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "ParamG")).HeaderText = ParamsApp.ColParamG
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "ParamH")).HeaderText = ParamsApp.ColParamH

        'Valuation variable
        Tab_Chaines(0) = Dossier_Reseau & "\Config_Machines\Fichier_Config_Machines.csv" 'Chemin fichier

        'Ouverture Form et remplissage DGV et si erreur, sortie de Sub
        If Fonctions_Form.Ouverture_Form_Avec_DGV(Me, Me.DataGridView1, Tab_Chaines, Num_Enreg_Actuel,
                                                  Num_Enreg_Maxi_CSV, Derniere_Ligne_DGV_Auto, , , ) = False Then Exit Sub

        ' Blocage des cellules si check
        Call GestionCellLocal()

        'Si New Form créée sans les attributs (si créée juste pour récupérer des infos)
        If Bouton_Ouverture IsNot Nothing Then

            'Valuation variables
            Bouton_Ouverture_Local = Bouton_Ouverture

            'Gestion Button de sélection
            Fonctions_Form.Gestion_Bouton_Et_ComboBox_Form(Bouton_Ouverture_Local, 3, , )
        End If
    End Sub

    'Action sur fermeture Form

    Private Sub Form_Listing_Config_Machines_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) _
        Handles MyBase.FormClosed

        'Traitement fermeture DGV
        Fonctions_Form.Fermeture_Form_Enfant(Me, True, Derniere_Ligne_DGV_Auto)

        'Gestion Button de sélection
        Fonctions_Form.Gestion_Bouton_Et_ComboBox_Form(Bouton_Ouverture_Local, 3, , )
    End Sub

    'Gestion fenêtre si click dans DGV

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As DataGridViewCellEventArgs) _
        Handles DataGridView1.CellContentClick

        'Si ce n'est pas la dernière ligne et si pas en mode visualisation
        If Fonctions_DGV.Ctrl_Ligne_DGV_Niv_Ouverture(Me.DataGridView1) = 1 Then

            'Valeur de condition à trouver
            Select Case e.ColumnIndex

                'Si la celule active se trouve dans la colonne Enregistrer
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Enregistrer ligne")

                    'Si en mode visualisation
                    If Niveau_Ouverture = 2 Then

                        'Ouverture Form MDP
                        Form_MDP_Admin.ShowDialog()
                    End If

                    'Si en mode administration
                    If Niveau_Ouverture = 3 Then

                        'Enregistrement ligne et si erreur, sortie de Sub
                        If Fonctions_DGV.Enregistrer_Ligne_DGV_Et_CSV _
                            (Me.DataGridView1, Tab_Chaines(0),
                             Num_Enreg_Actuel, Num_Enreg_Maxi_CSV, ) = False Then Exit Sub

                        'Déclaration et valuation variable
                        Dim DGV_Temp As DataGridView = New Form_Listing_Prog(, ).DataGridView1

                        'Création des dossiers et fichiers Listing et Historique si besoin et si erreur, renvoi False
                        If Fonctions_Diverses.Ctrl_Arbor_Ou_Creation_Fichier _
                        (Dossier_Reseau & "\Listings_Prog", False, "Listing_" & Me.DataGridView1.CurrentRow.Cells.Item _
                         (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Machine")).Value & ".csv", ,
                            DGV_Temp) = False Then Exit Sub

                        'Valuation variable
                        DGV_Temp = New Form_Historique().DataGridView1

                        If Fonctions_Diverses.Ctrl_Arbor_Ou_Creation_Fichier _
                            (Dossier_Reseau & "\Historique_Prog", False, "Historique_" & Me.DataGridView1.CurrentRow.Cells.Item _
                             (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Machine")).Value & ".csv", ,
                             DGV_Temp) = False Then Exit Sub
                    End If

                    'Si la celule active se trouve dans la colonne Supprimer
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Supprimer ligne")

                    'Si en mode visualisation
                    If Niveau_Ouverture = 2 Then

                        'Ouverture Form MDP
                        Form_MDP_Admin.ShowDialog()
                    End If

                    'Si en mode administration
                    If Niveau_Ouverture = 3 Then

                        'Suppression ligne et si erreur, sortie de Sub
                        If Fonctions_DGV.Sup_Ligne_DGV_Et_CSV(Me.DataGridView1, Tab_Chaines(0),
                                                        Num_Enreg_Actuel, Num_Enreg_Maxi_CSV, ) = False Then Exit Sub
                    End If

                    'Si la celule active se trouve dans la colonne Non-programmée sous CATIA
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Non-programmée sous CATIA")

                    'Gestion CheckBox et colonnes concernées
                    Fonctions_DGV.Gestion_Dynamique_CheckBox _
                        (Me.DataGridView1, "Non-programmée sous CATIA", tabConfgMachTyp, True, 1,
                         True, Me.DataGridView1.CurrentRow.Index)

                       'Si la celule active se trouve dans la colonne FOCAS disponible
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "FOCAS disponible")

                    'Gestion CheckBox et colonnes concernées
                    Fonctions_DGV.Gestion_Dynamique_CheckBox _
                        (Me.DataGridView1, "FOCAS disponible", confgFOCAS, True, 2,
                         True, Me.DataGridView1.CurrentRow.Index)
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
            If Fonctions_DGV.Ctrl_Chaine_DGV(Me.DataGridView1, e.FormattedValue.ToString, e.ColumnIndex, True) = False Then e.Cancel = True
        End If
    End Sub

    'Ecriture dernière ligne DGV

    Private Sub DataGridView1_RowsAdded(ByVal sender As System.Object, ByVal e As DataGridViewRowsAddedEventArgs) _
        Handles DataGridView1.RowsAdded

        'Si en mode modification
        If Derniere_Ligne_DGV_Auto = True Then

            'Déclaration et valuation variable
            Dim Tab_Noms_Colonnes_Check(1) As String
            Tab_Noms_Colonnes_Check(0) = "Non-programmée sous CATIA"
            Tab_Noms_Colonnes_Check(1) = "FOCAS disponible"

            'Ecriture dernière ligne dans DGV
            Fonctions_DGV.Ecriture_Derniere_Ligne_DGV(Me.DataGridView1, Num_Enreg_Actuel, , , , Tab_Noms_Colonnes_Check)

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

        ' Blocage des colonnes si Check
        tabConfgMachTyp(0) = "Type machine"
        tabConfgMachTyp(1) = "Mouvement axial/radial"
        Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
            (Me.DataGridView1, "Non-programmée sous CATIA", tabConfgMachTyp, False, 1, True, )

        confgFOCAS(0) = "Adresse IP"
        confgFOCAS(1) = "Nb outils"
        Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
                (Me.DataGridView1, "FOCAS disponible", confgFOCAS, False, 2, True, )
    End Sub
End Class