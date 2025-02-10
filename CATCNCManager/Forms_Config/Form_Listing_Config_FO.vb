Public Class Form_Listing_Config_FO

    'Variables
    Private Num_Enreg_Actuel, Num_Enreg_Maxi_CSV As Integer
    Private Derniere_Ligne_DGV_Auto As Boolean = False
    Private Tab_Chaines(0), traitEntet(1), traitOP(1), blocOP(1), traitOutils(1), traitCorr(1), sousProg(1), traitRep(1) As String

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

        'Valuation variable
        Tab_Chaines(0) = Dossier_Reseau & "\Config_FO\Fichier_Config_FO.csv" 'Chemin fichier

        'Ouverture Form et remplissage DGV et si erreur, sortie de Sub
        If Fonctions_Form.Ouverture_Form_Avec_DGV(Me, Me.DataGridView1, Tab_Chaines, Num_Enreg_Actuel,
                                                  Num_Enreg_Maxi_CSV, Derniere_Ligne_DGV_Auto, , , ) = False Then Exit Sub

        ' Blocage des cellules si check
        Call GestionCellLocal()

        'Si New Form créée sans les attributs (si créée juste pour récupérer des infos)
        If Not Bouton_Ouverture Is Nothing Then

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

                    'Si la celule active se trouve dans la colonne "Traiter entête"
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Traiter entête")

                    'Gestion CheckBox et colonnes concernées
                    Fonctions_DGV.Gestion_Dynamique_CheckBox _
                        (Me.DataGridView1, "Traiter entête", traitEntet, True, 2,
                         True, Me.DataGridView1.CurrentRow.Index)

                       'Si la celule active se trouve dans la colonne "Traiter opérations"
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Traiter opérations")

                    'Gestion CheckBox et colonnes concernées
                    Fonctions_DGV.Gestion_Dynamique_CheckBox _
                        (Me.DataGridView1, "Traiter opérations", traitOP, True, 2,
                         True, Me.DataGridView1.CurrentRow.Index)

                       'Si la celule active se trouve dans la colonne "Blocs par opération"
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Blocs par opération")

                    'Gestion CheckBox et colonnes concernées
                    Fonctions_DGV.Gestion_Dynamique_CheckBox _
                        (Me.DataGridView1, "Blocs par opération", blocOP, True, 2,
                         True, Me.DataGridView1.CurrentRow.Index)

                    'Si la celule active se trouve dans la colonne "Traiter outils"
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Traiter changements outil")

                    'Gestion CheckBox et colonnes concernées
                    Fonctions_DGV.Gestion_Dynamique_CheckBox _
                        (Me.DataGridView1, "Traiter changements outil", traitOutils, True, 2,
                         True, Me.DataGridView1.CurrentRow.Index)

                'Si la celule active se trouve dans la colonne "Traiter outils"
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Correcteurs type BUMOTEC")

                    'Gestion CheckBox et colonnes concernées
                    Fonctions_DGV.Gestion_Dynamique_CheckBox _
                        (Me.DataGridView1, "Correcteurs type BUMOTEC", traitCorr, True, 2,
                         True, Me.DataGridView1.CurrentRow.Index)

                       'Si la celule active se trouve dans la colonne "Sous-programme par outil"
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Sous-programmes par outil")

                    'Gestion CheckBox et colonnes concernées
                    Fonctions_DGV.Gestion_Dynamique_CheckBox _
                        (Me.DataGridView1, "Sous-programmes par outil", sousProg, True, 2,
                         True, Me.DataGridView1.CurrentRow.Index)

                       'Si la celule active se trouve dans la colonne "Traiter repères usinqge"
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Traiter repères usinage")

                    'Gestion CheckBox et colonnes concernées
                    Fonctions_DGV.Gestion_Dynamique_CheckBox _
                        (Me.DataGridView1, "Traiter repères usinage", traitRep, True, 2,
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
            Dim Tab_Noms_Colonnes_Check(6) As String
            Tab_Noms_Colonnes_Check(0) = "Traiter entête"
            Tab_Noms_Colonnes_Check(1) = "Traiter opérations"
            Tab_Noms_Colonnes_Check(2) = "Blocs par opération"
            Tab_Noms_Colonnes_Check(3) = "Traiter changements outil"
            Tab_Noms_Colonnes_Check(4) = "Correcteurs type BUMOTEC"
            Tab_Noms_Colonnes_Check(5) = "Sous-programmes par outil"
            Tab_Noms_Colonnes_Check(6) = "Traiter repères usinage"

            'Ecriture dernière ligne dans DGV
            Fonctions_DGV.Ecriture_Derniere_Ligne_DGV(Me.DataGridView1, Num_Enreg_Actuel, , , , Tab_Noms_Colonnes_Check)

            ' Gestion des CheckBox et vérrouillage cellules
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
        traitEntet(0) = "Cellule nom programme"
        traitEntet(1) = "Couleur RGB fond entête"
        Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
            (Me.DataGridView1, "Traiter entête", traitEntet, False, 2, True, )

        traitOP(0) = "Colonne nom opération"
        traitOP(1) = "Texte précédent bloc"
        Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
            (Me.DataGridView1, "Traiter opérations", traitOP, False, 2, True,  )

        blocOP(0) = "Colonne bloc"
        blocOP(1) = "Texte précédent bloc"
        Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
            (Me.DataGridView1, "Blocs par opération", blocOP, False, 2, True, )

        traitOutils(0) = "Colonne nom outil"
        traitOutils(1) = "Texte précédent sous-programme"
        Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
            (Me.DataGridView1, "Traiter changements outil", traitOutils, False, 2, True, )

        traitCorr(0) = "Colonne N° correcteur"
        traitCorr(1) = "Texte précédent N° correcteur"
        Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
            (Me.DataGridView1, "Correcteurs type BUMOTEC", traitCorr, False, 2, True, )

        sousProg(0) = "Colonne sous-programme"
        sousProg(1) = "Texte précédent sous-programme"
        Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
            (Me.DataGridView1, "Sous-programmes par outil", sousProg, False, 2, True, )

        traitRep(0) = "Colonne repère"
        traitRep(1) = "Couleur RGB fond repère"
        Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
            (Me.DataGridView1, "Traiter repères usinage", traitRep, False, 2, True, )
    End Sub
End Class