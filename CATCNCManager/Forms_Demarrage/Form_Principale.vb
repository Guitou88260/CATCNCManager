
Public Class Form_Principale


    '------------------------------------------------- Action sur événement ---------------------------------------------------

    'Réaction à l'ouverture de la fenêtre principale

    Private Sub Form_Principale_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Si besoin de regénérer l'XML
        'FctXMLParamApp.SerialXML()

        'Lecture config XML
        FctXMLParamApp.DeserialXML()

        'Val constantes des chemin des data
        If Environment.UserName = "Guitou" Then

            'Constantes perso
            Dossier_Reseau = ParamsApp.DataAppTest
            Dossier_Temp = ParamsApp.FoTmpAppTest

            'Sinon
        Else

            'Constantes perso
            Dossier_Reseau = ParamsApp.DataApp
            Dossier_Temp = ParamsApp.FoTmpApp
        End If

        'Ajout nom de fenêtre
        Me.Text = ParamsApp.NameApp



        'TODO Vérif version supprimée. A voir si pose problème

        'Variable
        'Dim Tab_Temp_Version_Appli(,) As String = Nothing

        'Lecture CSV et si erreur, End
        'If Fonctions_CSV.Lecture_CSV(Dossier_Reseau & "\Version_Appli\Fichier_Version_Appli.txt",
        'Tab_Temp_Version_Appli, 6) = False Then End

        'Contrôle version aplication
        'If Fonctions_Diverses.Ctrl_Chaine_Seul(Tab_Temp_Version_Appli(0, 0), 4, False, , , ,
        'My.Application.Info.Version.ToString) = False Then

        'MsgBox et End
        'Fonctions_Messages.Appel_Msg(31, 6, , )
        'End
        'End If



        'Ouverture Form MDP
        Form_MDP_Demarrage_Application.ShowDialog()

        'Ecriture mode ouverture dans entête Form principale
        Fonctions_Form.Ecriture_Mode_Ouverture(Me)

        'Si erreur pendant remplissage des ComboBox, Exit Sub
        If Fonctions_Form.Remp_ComboBox_Avec_Colonne_CSV(Dossier_Reseau & "\Config_Machines\Fichier_Config_Machines.csv",
                                                        "Machine", Me.ComboBox1, 5, ) = False Then Exit Sub
        If Fonctions_Form.Remp_ComboBox_Avec_Colonne_CSV(Dossier_Reseau & "\Config_Outils\Fichier_Config_Outils.csv",
                                                         "Outil", Me.ComboBox2, 5, ) = False Then Exit Sub
        If Fonctions_Form.Remp_ComboBox_Avec_Colonne_CSV(Dossier_Reseau & "\Config_Machines\Fichier_Config_Machines.csv",
                                                        "Machine", Me.ComboBox3, 5, "FOCAS disponible") = False Then Exit Sub
    End Sub

    'Click bouton config machines

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        'Création nouvelle Form
        Dim New_Form As Form = New Form_Listing_Config_Machines(Me.Button4)

        'Passage Form actuelle en enfant et visualisation Form
        New_Form.MdiParent = Me
        New_Form.Show()
    End Sub

    'Click bouton config outils de fraisage standards

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click

        'Création nouvelle Form
        Dim New_Form As Form = New Form_Listing_Config_Outils_Fraisage_Stds(Me.Button7)

        'Passage Form actuelle en enfant et visualisation Form
        New_Form.MdiParent = Me
        New_Form.Show()
    End Sub

    'Click bouton config outils tournage standards

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        'Création nouvelle Form
        Dim New_Form As Form = New Form_Listing_Config_Outils_Tournage_Stds(Me.Button3)

        'Passage Form actuelle en enfant et visualisation Form
        New_Form.MdiParent = Me
        New_Form.Show()
    End Sub

    'Click bouton config plaquettes tournage standards

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        'Création nouvelle Form
        Dim New_Form As Form = New Form_Listing_Config_Plaq_Tournage_Stds(Me.Button5)

        'Passage Form actuelle en enfant et visualisation Form
        New_Form.MdiParent = Me
        New_Form.Show()
    End Sub

    'Click bouton config outils et conversion outils (fraisage/tournage)

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click

        'Création nouvelle Form
        Dim New_Form As Form = New Form_Listing_Config_Outils(Me.Button6)

        'Passage Form actuelle en enfant et visualisation Form
        New_Form.MdiParent = Me
        New_Form.Show()
    End Sub

    'Click bouton config FO

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click

        'Création nouvelle Form
        Dim New_Form As Form = New Form_Listing_Config_FO(Me.Button14)

        'Passage Form actuelle en enfant et visualisation Form
        New_Form.MdiParent = Me
        New_Form.Show()
    End Sub

    'Click bouton ouverture listing programme

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'Si ComboBox non-vide
        If Me.ComboBox1.Text <> Nothing Then

            'Création nouvelle Form
            Dim New_Form As Form = New Form_Listing_Prog(Me.Button1, Me.ComboBox1)

            'Passage Form actuelle en enfant et visualisation Form
            New_Form.MdiParent = Me
            New_Form.Show()

            'Sinon...
        Else

            'Message
            Fonctions_Messages.Appel_Msg(20, 3, , )
        End If
    End Sub

    'Click bouton ouverture listing outil

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click

        'Si ComboBox non-vide
        If Me.ComboBox2.Text <> Nothing Then

            'Variable
            Dim Tab_Temp_Config(,) As String = Nothing

            'Lecture CSV et si erreur, désactivation commandes Form et sortie de Sub
            If Fonctions_CSV.Lecture_CSV(Dossier_Reseau & "\Config_Outils\Fichier_Config_Outils.csv",
                                         Tab_Temp_Config, 5) = False Then Exit Sub

            'Variable
            Dim New_Form As Form

            'Si colonne outil de fraisage valué
            If Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab _
                (Tab_Temp_Config, Me.ComboBox2.Text, "Outil", "Type outil fraisage principal", 1) <> Nothing Then

                'Création nouvelle Form
                New_Form = New Form_Listing_Outils_Fraisage _
                    (Me.Button11, Me.ComboBox2, Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab _
                     (Tab_Temp_Config, Me.ComboBox2.Text, "Outil", "Type outil fraisage principal", 1))

                'Passage Form actuelle en enfant et visualisation Form
                New_Form.MdiParent = Me
                New_Form.Show()

                'Si colonne outil de tournage valué
            ElseIf Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab _
                (Tab_Temp_Config, Me.ComboBox2.Text, "Outil", "Type outil tournage principal", 1) <> Nothing And
                Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab _
                (Tab_Temp_Config, Me.ComboBox2.Text, "Outil", "Type plaquette outil principal", 1) <> Nothing Then

                'Création nouvelle Form
                New_Form = New Form_Listing_Outils_Tournage _
                    (Me.Button11, Me.ComboBox2, Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab _
                     (Tab_Temp_Config, Me.ComboBox2.Text, "Outil", "Type outil tournage principal", 1),
                     Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab _
                     (Tab_Temp_Config, Me.ComboBox2.Text, "Outil", "Type plaquette outil principal", 1))

                'Passage Form actuelle en enfant et visualisation Form
                New_Form.MdiParent = Me
                New_Form.Show()

                'Sinon, message et sortie de Sub
            Else

                'Message
                Fonctions_Messages.Appel_Msg(52, 3, , )

                'Sortie de Sub
                Exit Sub
            End If

            'Sinon...
        Else

            'Message
            Fonctions_Messages.Appel_Msg(20, 3, , )
        End If
    End Sub

    'Click bouton générer fiche outils

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        'Création nouvelle Form
        Dim New_Form As Form = New Form_Generation_Fiche_Outils(Me.Button2)

        'Passage Form actuelle en enfant et visualisation Form
        New_Form.MdiParent = Me
        New_Form.Show()
    End Sub

    'Click bouton ouverture analyse des temps de cycle programme CATIA

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click

        'Création nouvelle Form
        Dim New_Form As Form = New Form_Analyse_Tps_Cycle(Me.Button8)

        'Passage Form actuelle en enfant et visualisation Form
        New_Form.MdiParent = Me
        New_Form.Show()
    End Sub

    'Click bouton ouverture lecture tableau outils machine
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        'Si ComboBox non-vide
        If Me.ComboBox3.Text <> Nothing Then

            'Création nouvelle Form
            Dim New_Form As Form = New Form_Tab_Outils(Me.Button9, Me.ComboBox3)

            'Passage Form actuelle en enfant et visualisation Form
            New_Form.MdiParent = Me
            New_Form.Show()

            'Sinon...
        Else

            'Message
            Fonctions_Messages.Appel_Msg(20, 3, , )
        End If
    End Sub


    '------------------------------------------------------------------- Ouverture documentation ---------------------------------------------------------------------

    'Click ouverture documentation PDF

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click

        'Ouverture PDF et si erreur, Exit Sub
        If Fonctions_Diverses.Ouverture_Fichier_PDF(Dossier_Reseau & "\Doc_Appli\Doc_Appli_CATIA_COMADUR.pdf") = False Then Exit Sub
    End Sub

    '------------------------------------------------- Gestion des Forms ---------------------------------------------------

    'Click bouton disposition Forms en mode cascade

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click

        'Passage Form en mode cascade
        Me.LayoutMdi(System.Windows.Forms.MdiLayout.Cascade)
    End Sub

    'Click bouton disposition Forms en mode vertical

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click

        'Passage Form en mode vertical
        Me.LayoutMdi(System.Windows.Forms.MdiLayout.TileVertical)
    End Sub

    'Click bouton disposition Forms en mode horizontal

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click

        'Passage Form en mode horizontal
        Me.LayoutMdi(System.Windows.Forms.MdiLayout.TileHorizontal)
    End Sub

    'Click bouton pour affichage/masquage menu

    Private Sub ToolStripButton5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.CheckedChanged

        'Gestion masquage et affichage du composant
        Fonctions_Form.Masquage_Affichage_Composants(Me.TabControl1, ToolStripButton5)
    End Sub
End Class
