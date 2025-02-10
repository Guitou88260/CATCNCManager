Public Class Form_Listing_Config_Outils

    'Variables
    Private Num_Enreg_Actuel, Num_Enreg_Maxi_CSV As Integer
    Private Derniere_Ligne_DGV_Auto As Boolean = False
    Private Bouton_Ouverture As System.Windows.Forms.Button
    Private Tab_Chaines(0) As String
    Private Bouton_Ouverture_Local As System.Windows.Forms.Button
    Private Form_Filtre_Locale As Form_Filtre
    Private Vidage_Cellule_Sur_Check As Boolean = False

    '------------------------------------------------- Action sur événement ---------------------------------------------------

    'Action sur création nouvelle Form

    Public Sub New(ByRef Bouton_Ouverture As System.Windows.Forms.Button)

        'Initialisation composant
        InitializeComponent()

        'Paramètrage des colonnes DGV vérrouillées
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Numéro enregistrement")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Date modification")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Auteur modification")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Date création")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)
        Me.DataGridView1.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Auteur création")).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvBloq)

        'Valuation variable
        Tab_Chaines(0) = Dossier_Reseau & "\Config_Outils\Fichier_Config_Outils.csv" 'Chemin fichier

        'Ouverture Form et remplissage DGV et si erreur, sortie de Sub
        If Fonctions_Form.Ouverture_Form_Avec_DGV _
            (Me, Me.DataGridView1, Tab_Chaines, Num_Enreg_Actuel, _
             Num_Enreg_Maxi_CSV, Derniere_Ligne_DGV_Auto, , , ) = False Then Exit Sub

        'Outil principal
        'Déclaration variable
        Dim Tab_Colonne_Temp(1) As String

        'Valuation variable
        Tab_Colonne_Temp(0) = "Type outil tournage principal"
        Tab_Colonne_Temp(1) = "Type plaquette outil principal"

        'Vidage et blocage des colonnes si autre colonne non-vide
        Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
            (Me.DataGridView1, "Type outil fraisage principal", Tab_Colonne_Temp, False, 1, False, )

        'Redimensionnement Tab
        ReDim Tab_Colonne_Temp(0)

        'Valuation variable
        Tab_Colonne_Temp(0) = "Type outil fraisage principal"

        'Vidage et blocage des colonnes si autre colonne non-vide
        Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
            (Me.DataGridView1, "Type outil tournage principal", Tab_Colonne_Temp, False, 1, False, )

        'Outil secondaire
        'Redimensionnement Tab
        ReDim Tab_Colonne_Temp(1)

        'Valuation variable
        Tab_Colonne_Temp(0) = "Type outil tournage secondaire"
        Tab_Colonne_Temp(1) = "Type plaquette outil secondaire"

        'Vidage et blocage des colonnes si autre colonne non-vide
        Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
            (Me.DataGridView1, "Type outil fraisage secondaire", Tab_Colonne_Temp, False, 1, False, )

        'Redimensionnement Tab
        ReDim Tab_Colonne_Temp(0)

        'Valuation variable
        Tab_Colonne_Temp(0) = "Type outil fraisage secondaire"

        'Vidage et blocage des colonnes si autre colonne non-vide
        Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
            (Me.DataGridView1, "Type outil tournage secondaire", Tab_Colonne_Temp, False, 1, False, )

        'Conversion outil secondaire
        'Redimensionnement Tab
        ReDim Tab_Colonne_Temp(1)

        'Valuation variable
        Tab_Colonne_Temp(0) = "Type outil fraisage secondaire"
        Tab_Colonne_Temp(1) = "Type plaquette outil secondaire"

        'Vidage et blocage des colonnes si Check
        Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
            (Me.DataGridView1, "Conversion outil", Tab_Colonne_Temp, False, 2, True, )

        'Valuation variables
        Bouton_Ouverture_Local = Bouton_Ouverture

        'Gestion Button de sélection
        Fonctions_Form.Gestion_Bouton_Et_ComboBox_Form(Bouton_Ouverture_Local, 3, , )
    End Sub

    'Action sur fermeture Form

    Private Sub Form_Listing_Config_Conversion_Outils_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) _
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

                    'Si mode administration = True
                    If Niveau_Ouverture = 2 Then

                        'Ouverture Form MDP
                        Form_MDP_Admin.ShowDialog()
                    End If

                    'Si mode administration = True
                    If Niveau_Ouverture = 3 Then

                        'Enregistrement ligne et si erreur, sortie de Sub
                        If Fonctions_DGV.Enregistrer_Ligne_DGV_Et_CSV _
                        (Me.DataGridView1, Tab_Chaines(0), _
                            Num_Enreg_Actuel, Num_Enreg_Maxi_CSV, ) = False Then Exit Sub

                        'Si outil de fraisage
                        If Me.DataGridView1.CurrentRow.Cells.Item _
                         (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Type outil fraisage principal")).Value <> Nothing Then

                            'Création des dossiers et fichiers Listing si besoin et si erreur, renvoi False
                            If Fonctions_Diverses.Ctrl_Arbor_Ou_Creation_Fichier _
                           (Dossier_Reseau & "\Listings_Outils", False, "Listing_" & Me.DataGridView1.CurrentRow.Cells.Item _
                            (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Outil")).Value & ".csv", , _
                            New Form_Listing_Outils_Fraisage().DataGridView1) = False Then Exit Sub

                            'Si outil de tournage
                        ElseIf Me.DataGridView1.CurrentRow.Cells.Item _
                            (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Type outil tournage principal")).Value <> Nothing And _
                            Me.DataGridView1.CurrentRow.Cells.Item _
                            (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Type plaquette outil principal")).Value <> Nothing Then

                            'Création des dossiers et fichiers Listing si besoin et si erreur, renvoi False
                            If Fonctions_Diverses.Ctrl_Arbor_Ou_Creation_Fichier _
                           (Dossier_Reseau & "\Listings_Outils", False, "Listing_" & Me.DataGridView1.CurrentRow.Cells.Item _
                            (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Outil")).Value & ".csv", , _
                            New Form_Listing_Outils_Tournage().DataGridView1) = False Then Exit Sub
                        End If

                        'Création des dossiers et fichiers Config si besoin et si erreur, renvoi False
                        If Fonctions_Diverses.Ctrl_Arbor_Ou_Creation_Fichier _
                        (Dossier_Reseau & "\Config_Outils", False, "Fichier_Config_Conversion_" & Me.DataGridView1.CurrentRow.Cells.Item _
                         (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Outil")).Value & ".csv", , _
                         New Form_Listing_Config_Conversion_Outils().DataGridView1) = False Then Exit Sub

                        'Déclaration et valuation variable
                        Dim DGV_Temp As DataGridView = New Form_Historique().DataGridView1

                        If Fonctions_Diverses.Ctrl_Arbor_Ou_Creation_Fichier _
                            (Dossier_Reseau & "\Historique_Outils", False, "Historique_" & Me.DataGridView1.CurrentRow.Cells.Item _
                             (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Outil")).Value & ".csv", , _
                             DGV_Temp) = False Then Exit Sub
                    End If

                    'Si la celule active se trouve dans la colonne Supprimer
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Supprimer ligne")

                    'Si mode administration = True
                    If Niveau_Ouverture = 2 Then

                        'Ouverture Form MDP
                        Form_MDP_Admin.ShowDialog()
                    End If

                    'Si mode administration = True
                    If Niveau_Ouverture = 3 Then

                        'Suppression ligne et si erreur, sortie de Sub
                        If Fonctions_DGV.Sup_Ligne_DGV_Et_CSV(Me.DataGridView1, Tab_Chaines(0), _
                                                        Num_Enreg_Actuel, Num_Enreg_Maxi_CSV, ) = False Then Exit Sub
                    End If

                    'Si la celule active se trouve dans la colonne Configurer
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Configuration conversion outil")

                    'Contrôle de chaîne de toutes les colonnes DGV
                    If Fonctions_DGV.Ctrl_Validation_Ligne_DGV(Me.DataGridView1, True) = False Then Exit Sub

                    'Déclaration variables
                    Dim Type_Outil_Principal As String = Nothing
                    Dim Type_Plaquette_Principale As String = Nothing
                    Dim Type_Outil_Secondaire As String = Nothing
                    Dim Type_Plaquette_Secondaire As String = Nothing
                    Dim Nom_Outil_Selectionne As String = _
                        Me.DataGridView1.CurrentRow.Cells.Item _
                        (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Outil")).Value
                    Dim New_Form As Form

                    'Si outil principal de fraisage et secondaire de tournage
                    If Me.DataGridView1.CurrentRow.Cells.Item _
                        (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Type outil fraisage principal")).Value <> Nothing And _
                        Me.DataGridView1.CurrentRow.Cells.Item _
                        (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Type outil tournage secondaire")).Value <> Nothing And _
                        Me.DataGridView1.CurrentRow.Cells.Item _
                        (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Type plaquette outil secondaire")).Value <> Nothing Then

                        'Valuation variables
                        Type_Outil_Principal = Me.DataGridView1.CurrentRow.Cells.Item _
                        (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Type outil fraisage principal")).Value
                        Type_Outil_Secondaire = Me.DataGridView1.CurrentRow.Cells.Item _
                        (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Type outil tournage secondaire")).Value
                        Type_Plaquette_Secondaire = Me.DataGridView1.CurrentRow.Cells.Item _
                        (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Type plaquette outil secondaire")).Value

                        'Valuation nouvelle Form
                        New_Form = New Form_Listing_Config_Conversion_Outils _
                            (Type_Outil_Principal, , Type_Outil_Secondaire, _
                             Type_Plaquette_Secondaire, Nom_Outil_Selectionne)

                        'Si outil principal de tournage et secondaire de fraisage
                    ElseIf Me.DataGridView1.CurrentRow.Cells.Item _
                        (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Type outil tournage principal")).Value <> Nothing And _
                        Me.DataGridView1.CurrentRow.Cells.Item _
                        (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Type plaquette outil principal")).Value <> Nothing And _
                        Me.DataGridView1.CurrentRow.Cells.Item _
                        (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Type outil fraisage secondaire")).Value <> Nothing Then

                        'Valuation variables
                        Type_Outil_Principal = Me.DataGridView1.CurrentRow.Cells.Item _
                        (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Type outil tournage principal")).Value
                        Type_Plaquette_Principale = Me.DataGridView1.CurrentRow.Cells.Item _
                        (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Type plaquette outil principal")).Value
                        Type_Outil_Secondaire = Me.DataGridView1.CurrentRow.Cells.Item _
                        (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Type outil fraisage secondaire")).Value

                        'Valuation nouvelle Form
                        New_Form = New Form_Listing_Config_Conversion_Outils _
                            (Type_Outil_Principal, Type_Plaquette_Principale, _
                             Type_Outil_Secondaire, , Nom_Outil_Selectionne)
                            
                        'Sinon, message
                    Else

                        'Message et sortie de Sub
                        Fonctions_Messages.Appel_Msg(43, 3, , )
                        Exit Sub
                    End If

                    'Visualisation Form
                    New_Form.ShowDialog()

                    'Si colonne check conversion outil
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Conversion outil")

                    'Valuation variable pour ne pas déclencher l'événement CellValueChanged
                    Vidage_Cellule_Sur_Check = True

                    'Déclaration variable
                    Dim Tab_Nom_Colonne(1) As String

                    'Valuation variable
                    Tab_Nom_Colonne(0) = "Type outil fraisage secondaire"
                    Tab_Nom_Colonne(1) = "Type plaquette outil secondaire"

                    'Gestion CheckBox et colonnes concernées
                    Fonctions_DGV.Gestion_Dynamique_CheckBox _
                        (Me.DataGridView1, "Conversion outil", Tab_Nom_Colonne, _
                         True, 2, True, Me.DataGridView1.CurrentRow.Index)

                    'Valuation variable
                    Vidage_Cellule_Sur_Check = False
            End Select

            'Si click bouton dernière ligne et si pas en mode visualisation
        ElseIf Fonctions_DGV.Ctrl_Ligne_DGV_Niv_Ouverture(Me.DataGridView1) = 2 Then

            'Ajout de ligne
            Me.DataGridView1.Rows.Add()
        End If
    End Sub

    'Gestion fenêtre en sortie d'édition

    Private Sub DataGridView1_CellEndEdit(ByVal sender As System.Object, ByVal e As DataGridViewCellEventArgs) _
        Handles DataGridView1.CellEndEdit

        'Si 1er ouverture Form OK
        If Derniere_Ligne_DGV_Auto = True And Vidage_Cellule_Sur_Check = False Then

            'Valeur de condition à trouver
            Select Case e.ColumnIndex

                'Si colonne outil de fraisage principal
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Type outil fraisage principal")

                    'Déclaration variable
                    Dim Tab_Colonne_Temp(1) As String

                    'Valuation variable
                    Tab_Colonne_Temp(0) = "Type outil tournage principal"
                    Tab_Colonne_Temp(1) = "Type plaquette outil principal"

                    'Vidage et blocage des colonnes si autre colonne vide
                    Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
                        (Me.DataGridView1, "Type outil fraisage principal", Tab_Colonne_Temp, True, 1, False, _
                         Me.DataGridView1.CurrentRow.Index)

                    'Si colonne outil de tournage principal
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Type outil tournage principal")

                    'Déclaration variable
                    Dim Tab_Colonne_Temp(0) As String

                    'Valuation variable
                    Tab_Colonne_Temp(0) = "Type outil fraisage principal"

                    'Vidage et blocage des colonnes si autre colonne vide
                    Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
                        (Me.DataGridView1, "Type outil tournage principal", Tab_Colonne_Temp, True, 1, False, _
                         Me.DataGridView1.CurrentRow.Index)

                    'Si colonne outil de tournage principal
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Type plaquette outil principal")

                    'Déclaration variable
                    Dim Tab_Colonne_Temp(0) As String

                    'Valuation variable
                    Tab_Colonne_Temp(0) = "Type outil fraisage principal"

                    'Vidage et blocage des colonnes si autre colonne vide
                    Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
                        (Me.DataGridView1, "Type plaquette outil principal", Tab_Colonne_Temp, True, 1, False, _
                         Me.DataGridView1.CurrentRow.Index)

                    'Outil secondaire
                    'Si colonne outil de fraisage principal
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Type outil fraisage secondaire")

                    'Déclaration variable
                    Dim Tab_Colonne_Temp(1) As String

                    'Valuation variable
                    Tab_Colonne_Temp(0) = "Type outil tournage secondaire"
                    Tab_Colonne_Temp(1) = "Type plaquette outil secondaire"

                    'Vidage et blocage des colonnes si autre colonne vide
                    Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
                        (Me.DataGridView1, "Type outil fraisage secondaire", Tab_Colonne_Temp, True, 1, False, _
                         Me.DataGridView1.CurrentRow.Index)

                    'Si colonne outil de tournage principal
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Type outil tournage secondaire")

                    'Déclaration variable
                    Dim Tab_Colonne_Temp(0) As String

                    'Valuation variable
                    Tab_Colonne_Temp(0) = "Type outil fraisage secondaire"

                    'Vidage et blocage des colonnes si autre colonne vide
                    Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
                        (Me.DataGridView1, "Type outil tournage secondaire", Tab_Colonne_Temp, True, 1, False, _
                         Me.DataGridView1.CurrentRow.Index)

                    'Si colonne outil de tournage principal
                Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Type plaquette outil secondaire")

                    'Déclaration variable
                    Dim Tab_Colonne_Temp(0) As String

                    'Valuation variable
                    Tab_Colonne_Temp(0) = "Type outil fraisage secondaire"

                    'Vidage et blocage des colonnes si autre colonne vide
                    Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
                        (Me.DataGridView1, "Type plaquette outil secondaire", Tab_Colonne_Temp, True, 1, False, _
                         Me.DataGridView1.CurrentRow.Index)
            End Select
        End If
    End Sub

    'Ecriture dernière ligne DGV

    Private Sub DataGridView1_RowsAdded(ByVal sender As System.Object, ByVal e As DataGridViewRowsAddedEventArgs) _
        Handles DataGridView1.RowsAdded

        'Si en mode modification
        If Derniere_Ligne_DGV_Auto = True Then

            'Déclaration variable
            Dim Tab_Noms_Colonnes_Check(0) As String

            'Valuation variable
            Tab_Noms_Colonnes_Check(0) = "Conversion outil"

            'Ecriture dernière ligne dans DGV
            Fonctions_DGV.Ecriture_Derniere_Ligne_DGV _
                (Me.DataGridView1, Num_Enreg_Actuel, , , , Tab_Noms_Colonnes_Check)

            'Valuation variable pour ne pas déclencher l'événement CellValueChanged
            Vidage_Cellule_Sur_Check = True

            'Déclaration variable
            Dim Tab_Colonne_Temp(1) As String

            'Valuation variable
            Tab_Colonne_Temp(0) = "Type outil fraisage secondaire"
            Tab_Colonne_Temp(1) = "Type plaquette outil secondaire"

            'Vidage et blocage des colonnes si Check
            Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
            (Me.DataGridView1, Tab_Noms_Colonnes_Check(0), Tab_Colonne_Temp, _
              False, 2, True, e.RowIndex - 1)

            'Vidage des cellules car ne marche dans la fonction précédente
            'Car chargement de la valeur de la cellule aprés la création de la ligne
            Me.DataGridView1.Rows.Item(e.RowIndex - 1).Cells.Item _
                (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Type outil fraisage secondaire")).Value = Nothing
            Me.DataGridView1.Rows.Item(e.RowIndex - 1).Cells.Item _
                (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Type outil tournage secondaire")).Value = Nothing
            Me.DataGridView1.Rows.Item(e.RowIndex - 1).Cells.Item _
                (Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Type plaquette outil secondaire")).Value = Nothing

            'Valuation variable
            Vidage_Cellule_Sur_Check = False
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