Public Class Form_Listing_Config_Conversion_Outils

    'Variables
    Private Tab_Chaines(0), Nom_Outil_Selectionne_Local, Type_Outil_Principal_Local, _
        Type_Plaquette_Principale_Local, Type_Outil_Secondaire_Local, Type_Plaquette_Secondaire_Local As String
    Private Num_Enreg_Actuel, Num_Enreg_Maxi_CSV As Integer
    Private Derniere_Ligne_DGV_Auto As Boolean = False
    Private Form_Filtre_Locale As Form_Filtre

    '------------------------------------------------- Action sur événement ---------------------------------------------------

    'Action sur création nouvelle Form

    Public Sub New(Optional ByVal Type_Outil_Principal As String = Nothing, _
                   Optional ByVal Type_Plaquette_Principale As String = Nothing, _
                   Optional ByVal Type_Outil_Secondaire As String = Nothing, _
                   Optional ByVal Type_Plaquette_Secondaire As String = Nothing, _
                   Optional ByVal Nom_Outil_Selectionne As String = Nothing)

        'Initialisation composant
        InitializeComponent()

        'Paramètrage des colonnes DGV vérrouillées
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Numéro enregistrement")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Date modification")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Auteur modification")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Date création")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Auteur création")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)

        'Si New Form créée sans les attributs (si créée juste pour récupérer des infos)
        If Type_Outil_Principal <> Nothing And Type_Outil_Secondaire <> Nothing And _
            Type_Plaquette_Secondaire <> Nothing Or Type_Outil_Principal <> Nothing And _
            Type_Plaquette_Principale <> Nothing And Type_Outil_Secondaire <> Nothing And _
            Nom_Outil_Selectionne <> Nothing Then

            'Valuation variables
            Type_Outil_Principal_Local = Type_Outil_Principal
            Type_Plaquette_Principale_Local = Type_Plaquette_Principale
            Type_Outil_Secondaire_Local = Type_Outil_Secondaire
            Type_Plaquette_Secondaire_Local = Type_Plaquette_Secondaire
            Nom_Outil_Selectionne_Local = Nom_Outil_Selectionne
            Tab_Chaines(0) = Dossier_Reseau & "\Config_Outils\Fichier_Config_Conversion_" & Nom_Outil_Selectionne_Local & ".csv" 'Chemin fichier

            'Si outil principal de fraisage et secondaire de tournage
            If Type_Plaquette_Principale_Local = Nothing

                'Déclaration variables
                Dim DGV_Outils_Principaux As DataGridView = New Form_Listing_Outils_Fraisage().DataGridView1
                Dim DGV_Outils_Secondaires As DataGridView = New Form_Listing_Outils_Tournage().DataGridView1

                'Remplissage des ComboBox DGV
                Fonctions_DGV.Remp_ComboBox_Avec_Entetes_DGV _
                    (DGV_Outils_Principaux, Me.DataGridView1.Columns.Item _
                     (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Paramètre CATIA outil principal")), _
                     "Outil à bout sphérique", "Nombre de lèvres", False, True, True)
                Fonctions_DGV.Remp_ComboBox_Avec_Entetes_DGV _
                     (DGV_Outils_Secondaires, Me.DataGridView1.Columns.Item _
                      (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Paramètre CATIA outil secondaire")), _
                       "Style d'orientation", "Largeur de plaquette (la)", False, True, True)
                Fonctions_DGV.Remp_ComboBox_Avec_Entetes_DGV _
                    (DGV_Outils_Secondaires, Me.DataGridView1.Columns.Item _
                     (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Paramètre CATIA plaquette secondaire")), _
                     "Type", "Profil de filetage", False, True, True)

                'Coloration/blocage fond cellule DGV
                Fonctions_DGV.Coloriage_Blocage_Celulle_DGV _
                    (Me.DataGridView1, "Paramètre CATIA plaquette principal", 1, , )

                'Si outil principal de tournage et secondaire de fraisage
            ElseIf Type_Plaquette_Secondaire_Local = Nothing Then

                'Déclaration et valuation variables
                Dim DGV_Outils_Principaux As DataGridView = New Form_Listing_Outils_Tournage().DataGridView1
                Dim DGV_Outils_Secondaires As DataGridView = New Form_Listing_Outils_Fraisage().DataGridView1

                'Remplissage des ComboBox DGV
                Fonctions_DGV.Remp_ComboBox_Avec_Entetes_DGV _
                     (DGV_Outils_Principaux, Me.DataGridView1.Columns.Item _
                      (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Paramètre CATIA outil principal")), _
                      "Style d'orientation", "Largeur de plaquette (la)", False, True, True)
                Fonctions_DGV.Remp_ComboBox_Avec_Entetes_DGV _
                    (DGV_Outils_Principaux, Me.DataGridView1.Columns.Item _
                     (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Paramètre CATIA plaquette principal")), _
                     "Type", "Profil de filetage", False, True, True)
                Fonctions_DGV.Remp_ComboBox_Avec_Entetes_DGV _
                    (DGV_Outils_Secondaires, Me.DataGridView1.Columns.Item _
                     (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Paramètre CATIA outil secondaire")), _
                     "Outil à bout sphérique", "Nombre de lèvres", False, True, True)

                'Coloration/blocage fond cellule DGV
                Fonctions_DGV.Coloriage_Blocage_Celulle_DGV _
                    (Me.DataGridView1, "Paramètre CATIA plaquette secondaire", 1, , )
            End If

            'Ouverture Form et remplissage DGV et si erreur, sortie de Sub
            If Fonctions_Form.Ouverture_Form_Avec_DGV(Me, Me.DataGridView1, Tab_Chaines, Num_Enreg_Actuel, _
                                       Num_Enreg_Maxi_CSV, Derniere_Ligne_DGV_Auto, , , ) = False Then Exit Sub

            'Déclaration variable
            Dim Tab_Colonne_Temp(0) As String

            'Valuation variable
            Tab_Colonne_Temp(0) = "Paramètre CATIA plaquette secondaire"
           
            'Vidage et blocage des colonnes si autre colonne vide
            Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
                (Me.DataGridView1, "Paramètre CATIA outil secondaire", Tab_Colonne_Temp, False, 1, False, )

            'Valuation variable
            Tab_Colonne_Temp(0) = "Paramètre CATIA outil secondaire"

            'Vidage et blocage des colonnes si autre colonne vide
            Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
                (Me.DataGridView1, "Paramètre CATIA plaquette secondaire", Tab_Colonne_Temp, False, 1, False, )
        End If
    End Sub

    'Action sur fermeture Form

    Private Sub Form_Listing_Config_Conversion_Outils_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) _
        Handles MyBase.FormClosed

        'Fermeture Form
        Me.Dispose()
    End Sub

    'Gestion fenêtre si click DGV

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As DataGridViewCellEventArgs) _
        Handles DataGridView1.CellContentClick

        'Valeur de condition à trouver
        Select Case e.ColumnIndex

            'Si colonne Enregistrer
            Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Enregistrer ligne")

                'Si en mode visualisation
                If Niveau_Ouverture = 2 Then

                    'Ouverture Form MDP
                    Form_MDP_Admin.ShowDialog()
                End If

                'Si en mode administration
                If Niveau_Ouverture = 3 Then

                    'Vidage ComboBox d'import des paramètres dans formule
                    Me.DataGridView1.CurrentRow.Cells.Item _
                        (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Paramètre CATIA outil principal")).Value = Nothing
                    Me.DataGridView1.CurrentRow.Cells.Item _
                        (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Paramètre CATIA plaquette principal")).Value = Nothing

                    'Enregistrement ligne et si erreur, sortie de Sub
                    If Fonctions_DGV.Enregistrer_Ligne_DGV_Et_CSV(Me.DataGridView1, Tab_Chaines(0), _
                                                    Num_Enreg_Actuel, Num_Enreg_Maxi_CSV, ) = False Then Exit Sub
                End If

                'Si colonne Supprimer
            Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Supprimer ligne")

                'Suppression ligne et si erreur, sortie de Sub
                If Fonctions_DGV.Sup_Ligne_DGV_Et_CSV(Me.DataGridView1, Tab_Chaines(0), _
                                                Num_Enreg_Actuel, Num_Enreg_Maxi_CSV, ) = False Then Exit Sub

                'Si colonne Ajouter paramètre CATIA
            Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Ajouter paramètre CATIA")

                'Si cellule DGV outil principal non-vide
                If Me.DataGridView1.CurrentRow.Cells.Item _
                        (Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                         (Me.DataGridView1, "Paramètre CATIA outil principal")).Value <> Nothing Then

                    'Envoi du paramètre dans cellule formule
                    Me.DataGridView1.CurrentRow.Cells.Item _
                        (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Formule conversion")).Value = _
                        Me.DataGridView1.CurrentRow.Cells.Item _
                        (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Formule conversion")).Value & _
                        "'" & Me.DataGridView1.CurrentRow.Cells.Item _
                        (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Paramètre CATIA outil principal")).Value & "'"

                    'Si cellule DGV plaquette principale non-vide
                ElseIf Me.DataGridView1.CurrentRow.Cells.Item _
                        (Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                         (Me.DataGridView1, "Paramètre CATIA plaquette principal")).Value <> Nothing Then

                    'Envoi du paramètre dans cellule formule
                    Me.DataGridView1.CurrentRow.Cells.Item _
                        (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Formule conversion")).Value = _
                        Me.DataGridView1.CurrentRow.Cells.Item _
                        (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Formule conversion")).Value & _
                        """" & Me.DataGridView1.CurrentRow.Cells.Item _
                        (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Paramètre CATIA plaquette principal")).Value & """"
                End If
        End Select
    End Sub

    'Gestion fenêtre en sortie d'édition

    Private Sub DataGridView1_CellEndEdit(ByVal sender As System.Object, ByVal e As DataGridViewCellEventArgs) _
        Handles DataGridView1.CellEndEdit

        'Déclaration variable
        Dim Premier_Passage As Boolean = True

        'Si 1er ouverture Form OK
        If Derniere_Ligne_DGV_Auto = True And Premier_Passage = True Then

            'Valeur de condition à trouver
            Select Case e.ColumnIndex

                'Outil principal
                'Si colonne outil principal est de fraisage
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Paramètre CATIA outil secondaire")

                    'Déclaration variable
                    Dim Tab_Colonne_Temp(0) As String

                    'Valuation variable
                    Tab_Colonne_Temp(0) = "Paramètre CATIA plaquette secondaire"

                    'Vidage et blocage des colonnes si autre colonne vide
                    Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
                        (Me.DataGridView1, "Paramètre CATIA outil secondaire", Tab_Colonne_Temp, True, 1, False, _
                         Me.DataGridView1.CurrentRow.Index)

                    'Valuation variable
                    Premier_Passage = False

                    'Si colonne outil principal est de tournage 
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Paramètre CATIA plaquette secondaire")

                    'Déclaration variable
                    Dim Tab_Colonne_Temp(0) As String

                    'Valuation variable
                    Tab_Colonne_Temp(0) = "Paramètre CATIA outil secondaire"

                    'Vidage et blocage des colonnes si autre colonne vide
                    Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
                        (Me.DataGridView1, "Paramètre CATIA plaquette secondaire", Tab_Colonne_Temp, True, 1, False, _
                         Me.DataGridView1.CurrentRow.Index)

                    'Valuation variable
                    Premier_Passage = False
            End Select
        End If
    End Sub

    'Ecriture dernière ligne DGV

    Private Sub DataGridView1_RowsAdded(ByVal sender As System.Object, ByVal e As DataGridViewRowsAddedEventArgs) _
        Handles DataGridView1.RowsAdded

        'Si en mode modification
        If Derniere_Ligne_DGV_Auto = True Then

            'Ecriture dernière ligne dans DGV
            Fonctions_DGV.Ecriture_Derniere_Ligne_DGV _
                (Me.DataGridView1, Num_Enreg_Actuel, , , , )
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