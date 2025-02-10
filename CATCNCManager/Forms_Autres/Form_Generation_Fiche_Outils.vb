Public Class Form_Generation_Fiche_Outils

    'Variables
    Private CATIA, Document_CATIA_Courant, Process_Courant,
        Phase_Courante, Prog_Selectionne As Object
    Private Bouton_Ouverture_Local As System.Windows.Forms.Button

    'Action sur création nouvelle Form

    Public Sub New(ByRef Bouton_Ouverture As System.Windows.Forms.Button)

        'Initialisation composant
        InitializeComponent()


        'Modif du dossier de reception FO
        'Remplissage TextBox
        Me.TextBox2.Text = Dossier_Temp '& "\" & Environment.UserName & "\Fiches_Outils"


        'Valuation variables
        Bouton_Ouverture_Local = Bouton_Ouverture

        'Valeur de condition à trouver
        Select Case Niveau_Ouverture

            'Si mode visualisation
            Case 1

                'Désactivation commandes Form
                Fonctions_Form.Desactivation_Cmd_Form(Me)

                'Si mode modification
            Case Is >= 2

                'Si erreur pendant remplissage ComboBox Fiche Outils, sortie de Sub
                If Fonctions_Form.Remp_ComboBox_Avec_Colonne_CSV _
                    (Dossier_Reseau & "\Config_FO\Fichier_Config_FO.csv", "Machine",
                      Me.ComboBox2, 5, ) = False Then Exit Sub

                'Si pas de CATIA ouvert et CATProcess inactif, sortie de Sub
                If Fonctions_CATIA.Ctrl_Recherche_Donnees_CATIA _
                    (CATIA, Document_CATIA_Courant, Process_Courant) = False Then Exit Sub
        End Select

        'Gestion Button de sélection
        Fonctions_Form.Gestion_Bouton_Et_ComboBox_Form(Bouton_Ouverture_Local, 3, , )
    End Sub

    'Action sur fermeture Form

    Private Sub Form_Listing_Prog_Vers_CATIA_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) _
        Handles MyBase.FormClosed

        'Traitement fermeture DGV
        Fonctions_Form.Fermeture_Form_Enfant(Me, False, )

        'Gestion Button de sélection
        Fonctions_Form.Gestion_Bouton_Et_ComboBox_Form(Bouton_Ouverture_Local, 3, , )
    End Sub

    'Click bouton chemin dossier fiche outils

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        'Remplissage TextBox avec chemin
        Fonctions_Form.Remplissage_TextBox_Chemin(Me.TextBox2, Me.FolderBrowserDialog1)
    End Sub

    'Click bouton ouvrir à l'emplacement

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        'Variable
        Dim Chemin_Temp As String = Me.TextBox2.Text & "\"

        'Ouverture explorer Windows et passage arguments
        Shell("explorer.exe /e,/root,""" & Chemin_Temp & """", AppWinStyle.MaximizedFocus)
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

        'Remplissage TextBox nom programme
        TextBox1.Text = Prog_Selectionne.Name

        'Pré-remplissage ComboBox machine
        If Phase_Courante.Machine IsNot Nothing Then
            Dim Machine_Courante As Object
            Machine_Courante = Phase_Courante.Machine
            Dim nameMachTmp() As String
            nameMachTmp = Split(Machine_Courante.Name, " ", 2)
            ComboBox2.Text = nameMachTmp(0)
        End If

        'Activation bouton d'envoi
        Me.Button2.Enabled = True
    End Sub

    'Click bouton générer fiche outils

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        'Si ComboBox non-vide
        If Me.ComboBox2.Text <> Nothing Then

            'Si erreur, renvoi False et Exit Sub
            If Fonctions_CATIA.Generation_Fiche_Outils _
                (CATIA, Document_CATIA_Courant, Phase_Courante, Prog_Selectionne, Me.ComboBox2.Text, Me.TextBox2.Text) =
                False Then Exit Sub

            'Sinon...
        Else

            'Message
            Fonctions_Messages.Appel_Msg(1, 3, , )
        End If
    End Sub
End Class