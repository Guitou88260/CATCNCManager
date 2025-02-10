Public Class Form_Listing_Prog_Vers_CATIA

    'Variables
    Private Document_CATIA_Courant, Process_Courant, _
        Phase_Courante, Prog_Selectionne As Object
    Private DGV_Maitresse_Local As DataGridView
    Private Nom_Machine_Locale As String

    'Action sur création nouvelle Form

    Public Sub New(ByVal DGV_Maitresse As DataGridView, _
                   ByVal Nom_Machine As String)

        'Initialisation composant
        InitializeComponent()

        'Valuation variables
        DGV_Maitresse_Local = DGV_Maitresse
        Nom_Machine_Locale = Nom_Machine

        'Variables
        Dim CATIA As Object = Nothing

        'Si pas de CATIA ouvert et CATProcess inactif, sortie de Sub
        If Fonctions_CATIA.Ctrl_Recherche_Donnees_CATIA(CATIA, Document_CATIA_Courant, Process_Courant) = _
            False Then Exit Sub
    End Sub

    'Action sur fermeture Form

    Private Sub Form_Listing_Prog_Vers_CATIA_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) _
        Handles MyBase.FormClosed

        'Fermeture Form
        Me.Dispose()
    End Sub

    'Click bouton sélection dans CATIA
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'Demande de sélection dans CATIA
        If Fonctions_CATIA.Selection_Element_CATIA _
            (Document_CATIA_Courant, "ManufacturingProgram", _
             "Sélectionner un programme d'usinage dans l'arbre", Prog_Selectionne) = False Then Exit Sub

        'Recherche de phase, programme courants, et si erreur, Message et sortie de Sub
        If Fonctions_CATIA.Recherche_Phase_Courante_Process _
            (Process_Courant, Prog_Selectionne, Phase_Courante) = False Then Exit Sub

        'Remplissage TextBox
        TextBox1.Text = Prog_Selectionne.Name

        'Activation bouton d'envoi
        Me.Button2.Enabled = True
    End Sub

    'Click bouton envoyer dans CATIA
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        'Gestion des exceptions
        Try

            'Si au moins 1 check dans les types de traitement
            If (Me.CheckBox1.Checked = True Or Me.CheckBox2.Checked = True Or Me.CheckBox3.Checked = True) Then

                'Si case nom de phase cochée
                If Me.CheckBox1.Checked = True Then

                    'Remplacement nom et commentaire de phase (avec passage en majuscule)
                    Phase_Courante.Name = UCase _
                        (DGV_Maitresse_Local.CurrentRow.Cells.Item _
                         (Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                          (DGV_Maitresse_Local, "Description programme")).Value)

                    Phase_Courante.Description = UCase _
                        (DGV_Maitresse_Local.CurrentRow.Cells.Item _
                         (Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                          (DGV_Maitresse_Local, "Description pièce")).Value)
                End If

                'Si case nom de programme cochée
                If Me.CheckBox2.Checked = True Then

                    'Renommage programme (avec passage en majuscule)
                    Prog_Selectionne.Name = UCase _
                        (DGV_Maitresse_Local.CurrentRow.Cells.Item _
                         (Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                          (DGV_Maitresse_Local, "Numéro programme")).Value)
                End If

                'Si case nom et type de machine cochée
                If Me.CheckBox3.Checked = True Then

                    'Variable
                    Dim Tab_Temp_Config(,) As String = Nothing

                    'Lecture CSV et si erreur, désactivation commandes Form et sortie de Sub
                    If Fonctions_CSV.Lecture_CSV(Dossier_Reseau & "\Config_Machines\Fichier_Config_Machines.csv", _
                                                 Tab_Temp_Config, 5) = False Then Exit Sub

                    'Création machine et si erreur, sortie de sub (avec passage en majuscule)
                    If Fonctions_CATIA.Creation_Machine_Dans_CATIA _
                        (Document_CATIA_Courant, Phase_Courante, Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab _
                         (Tab_Temp_Config, Nom_Machine_Locale, "Machine", "Type machine", 1), _
                         UCase(Nom_Machine_Locale)) = False Then Exit Sub

                    'Déclaration et valuation variable
                    Dim DGV_Temp As DataGridView = New Form_Listing_Config_Machines().DataGridView1

                    'Ecriture des pramètres dans CATIA
                    If Fonctions_CATIA.Modif_Variable_Selon_ToolTips _
                        (DGV_Temp, Phase_Courante.Machine, "Table PP", "Mouvement axial/radial", _
                         Tab_Temp_Config, Nom_Machine_Locale, "Machine") = False Then Exit Sub
                End If

                'Message de confirmation
                Fonctions_Messages.Appel_Msg(39, 1, , )

                'Fermeture Form
                Me.Dispose()

                'Sinon...
            Else

                'Message
                Fonctions_Messages.Appel_Msg(41, 3, , )
            End If

            'Pour toutes les erreurs
        Catch

            'Message et sortie de Sub
            Fonctions_Messages.Appel_Msg(45, 5, , )
            Exit Sub
        End Try
    End Sub
End Class