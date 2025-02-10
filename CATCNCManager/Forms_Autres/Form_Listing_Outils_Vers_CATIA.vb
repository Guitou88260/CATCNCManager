Public Class Form_Listing_Outils_Vers_CATIA

    'Variables
    Private Doc_CATIA_Courant, Process_Courant, Prog_CATIA_Courant, _
        Assemblage_Outil_Selectionne, Outil_Selectionne, Plaq_Selectionnee As Object
    Private DGV_Maitresse_Local As DataGridView
    Private Tab_Chaine_Local(), Nom_Listing_Outil, Type_Outil_Principal, _
        Type_Plaq_Principal, Type_Outil_Secondaire, _
        Type_Plaq_Secondaire, Tab_Config_Outil(,) As String
    Private Conversion_Outil_Dans_Config, Conversion_Outil As Boolean

    'Action sur création nouvelle Form

    Public Sub New(ByVal DGV_Maitresse As DataGridView, ByVal Tab_Chaines() As String)

        'Initialisation composant
        InitializeComponent()

        'Valuation variables
        DGV_Maitresse_Local = DGV_Maitresse
        Tab_Chaine_Local = Tab_Chaines
        Nom_Listing_Outil = Tab_Chaine_Local(1)
        Type_Outil_Principal = Tab_Chaine_Local(4)

        'Si plaquette valuée
        If UBound(Tab_Chaine_Local, 1) > 6 Then

            Type_Plaq_Principal = Tab_Chaine_Local(9)
        End If

        'Lecture CSV et si erreur, sortie de Sub
        If Fonctions_CSV.Lecture_CSV _
            (Dossier_Reseau & "\Config_Outils\Fichier_Config_Outils.csv", _
             Tab_Config_Outil, 5) = False Then Exit Sub

        'Valuation conversion outil
        Conversion_Outil_Dans_Config = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab _
            (Tab_Config_Outil, Nom_Listing_Outil, "Outil", "Conversion outil", 1)

        'Si conversion d'outil = True
        If Conversion_Outil_Dans_Config = True Then

            'Si type plaquette outil principal valué
            If Type_Plaq_Principal <> Nothing Then

                'Valuation variable
                Type_Outil_Secondaire = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab _
                  (Tab_Config_Outil, Nom_Listing_Outil, "Outil", "Type outil fraisage secondaire", 1)

                'Sinon
            Else

                'Valuation variables
                Type_Outil_Secondaire = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab _
                      (Tab_Config_Outil, Nom_Listing_Outil, "Outil", "Type outil tournage secondaire", 1)
                Type_Plaq_Secondaire = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab _
                      (Tab_Config_Outil, Nom_Listing_Outil, "Outil", "Type plaquette outil secondaire", 1)
            End If
        End If

        'Variable
        Dim CATIA As Object = Nothing

        'Si pas de CATIA ouvert et CATProcess inactif, sortie de Sub
        If Fonctions_CATIA.Ctrl_Recherche_Donnees_CATIA(CATIA, Doc_CATIA_Courant, Process_Courant) = _
            False Then Exit Sub
    End Sub

    'Action sur fermeture Form

    Private Sub Form_Listing_Outils_Vers_CATIA_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) _
        Handles MyBase.FormClosed

        'Fermeture Form
        Me.Dispose()
    End Sub

    'Click bouton sélection dans CATIA

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'Variables
        Dim Activite_Select As Object = Nothing
        Dim Result_Test_Type As Boolean

        'Début boucle
        Do

            'Demande de sélection dans CATIA
            If Fonctions_CATIA.Selection_Element_CATIA _
                (Doc_CATIA_Courant, "ManufacturingActivity", _
                 "Sélectionner un changement d'outil parmi les activités de l'arbre", _
                 Activite_Select) = False Then Exit Sub

            'Contrôle type de la sélection et retourne faux si problème de lecture fichier CSV
            If Fonctions_CATIA.Controle_Selection_ToolChange _
                (Activite_Select, Result_Test_Type, Type_Outil_Principal, Type_Plaq_Principal, _
                 Type_Outil_Secondaire, Type_Plaq_Secondaire, Conversion_Outil) = False Then Exit Sub

            'Fin de boucle et test résultat
        Loop While Result_Test_Type = False

        'Si outil fraisage
        If Activite_Select.Type = "ToolChange" Then

            'Valuation variables
            Outil_Selectionne = Activite_Select.Tool

            'Modification état GroupBox
            Me.GroupBox2.Enabled = True
            Me.GroupBox10.Enabled = False

            'Si outil tournage
        ElseIf Activite_Select.Type = "ToolChangeLathe" Then

            'Valuation variables
            Assemblage_Outil_Selectionne = Activite_Select.ToolAssembly
            Outil_Selectionne = Activite_Select.Tool
            Plaq_Selectionnee = Assemblage_Outil_Selectionne.Insert

            'Modification état GroupBox
            Me.GroupBox2.Enabled = False
            Me.GroupBox10.Enabled = True
        End If

        'Valuation variables
        Prog_CATIA_Courant = Activite_Select.Parent

        'Remplissage TextBox et activation bouton d'envoi
        Me.TextBox1.Text = Outil_Selectionne.Name
        Me.Button2.Enabled = True
    End Sub

    'Click bouton envoyer dans CATIA

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        'Variables
        Dim Valeur_Bris_Outil As String = Nothing
        Dim Inversion_Outil_Tournage As Boolean

        'Si bris d'outil checké
        If CheckBox1.Checked = True Then

            'Valuation variable
            Valeur_Bris_Outil = Fonctions_Diverses.Renvoi_Valeur_Sans_Unite _
                (Me.TextBox2.Text, Me.ToolTip1.GetToolTip(Me.TextBox2))
        End If

        'Si outil sélectionné est de type tournage
        If Not Plaq_Selectionnee Is Nothing Then

            'Valeur de condition à trouver
            Select Case True

                'Type outil de tournage intérieur ST1 et extérieur ST2
                Case RadioButton2.Checked, RadioButton5.Checked

                    'Valuation variable
                    Inversion_Outil_Tournage = True

                    'Type outil de tournage extérieur ST1 et intérieur ST2
                Case RadioButton3.Checked, RadioButton4.Checked

                    'Valuation variable
                    Inversion_Outil_Tournage = False
            End Select
        End If

        'Si conversion d'outil = True
        If Conversion_Outil = True Then

            'Déclaration variable
            Dim Tab_Config_Conversion_Outil(,) As String = Nothing

            'Si type plaquette outil principal valué
            If Type_Plaq_Principal <> Nothing Then

                'Déclaration variables
                Dim DGV_Outil_Secondaire As DataGridView = New Form_Listing_Outils_Fraisage().DataGridView1

                'Lecture CSV et si erreur, sortie de Sub
                If Fonctions_CSV.Lecture_CSV _
                    (Dossier_Reseau & "\Config_Outils\Fichier_Config_Conversion_" & _
                     Nom_Listing_Outil & ".csv", Tab_Config_Conversion_Outil, 5) = False Then Exit Sub

                'Valuation des paramètres outils et si erreur, sortie de sub
                If Fonctions_CATIA.Valuation_Param_Outil _
                    (Doc_CATIA_Courant, Prog_CATIA_Courant, DGV_Maitresse_Local,
                     Outil_Selectionne, Assemblage_Outil_Selectionne, Plaq_Selectionnee,
                     Valeur_Bris_Outil, Inversion_Outil_Tournage,
                     Me.CheckBox3.Checked, Tab_Chaine_Local, Tab_Config_Conversion_Outil,
                     DGV_Outil_Secondaire) =
                 False Then Exit Sub

                'Sinon
            Else

                'Variables
                Dim DGV_Outil_Secondaire As DataGridView = New Form_Listing_Outils_Tournage().DataGridView1

                'Lecture CSV et si erreur, sortie de Sub
                If Fonctions_CSV.Lecture_CSV _
                    (Dossier_Reseau & "\Config_Outils\Fichier_Config_Conversion_" & _
                     Nom_Listing_Outil & ".csv", Tab_Config_Conversion_Outil, 5) = False Then Exit Sub

                'Valuation des paramètres outils et si erreur, sortie de sub
                If Fonctions_CATIA.Valuation_Param_Outil _
                    (Doc_CATIA_Courant, Prog_CATIA_Courant, DGV_Maitresse_Local,
                     Outil_Selectionne, Assemblage_Outil_Selectionne, Plaq_Selectionnee,
                     Valeur_Bris_Outil, Inversion_Outil_Tournage,
                     Me.CheckBox3.Checked, Tab_Chaine_Local, Tab_Config_Conversion_Outil,
                     DGV_Outil_Secondaire) =
                 False Then Exit Sub
            End If

            'Sinon
        Else

            'Valuation des paramètres outils et si erreur, sortie de sub
            If Fonctions_CATIA.Valuation_Param_Outil _
                (Doc_CATIA_Courant, Prog_CATIA_Courant, DGV_Maitresse_Local,
                 Outil_Selectionne, Assemblage_Outil_Selectionne, Plaq_Selectionnee,
                 Valeur_Bris_Outil, Inversion_Outil_Tournage,
                 Me.CheckBox3.Checked, Tab_Chaine_Local, , ) =
             False Then Exit Sub
        End If

        'Message de confirmation
        Fonctions_Messages.Appel_Msg(39, 1, , )

        'Fermeture Form
        Me.Dispose()
    End Sub

    'Click sur CheckBox bris d'outil

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles CheckBox1.CheckedChanged

        'Si CheckBox bris d'outil coché
        If Me.CheckBox1.Checked = True Then

            'Activation TextBox
            Me.TextBox2.Enabled = True

            'Sinon
        Else

            'Désactivation TextBox
            Me.TextBox2.Enabled = False
        End If
    End Sub
End Class