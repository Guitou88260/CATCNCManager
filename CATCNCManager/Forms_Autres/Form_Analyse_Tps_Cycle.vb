Public Class Form_Analyse_Tps_Cycle

    'Variables CATIA
    Private CATIA_App, Doc_CATIA_Courant, Process_Courant, Phase_Courante, Prog_Courant As Object
    Private Bouton_Ouverture_Local As System.Windows.Forms.Button
    'Variables temps de cycle
    Private Tab_Tps_Cycle(,) As String = Nothing
    Private Coeff_Image_Tps As Double = Nothing
    Private Taille_Inc_PictureBox As Integer
    'Couleur visu temps de cycle
    Public Couleurs_OP(1) As Color

    'Action sur création nouvelle Form

    Public Sub New(ByRef Bouton_Ouverture As System.Windows.Forms.Button)

        'Initialisation composant
        InitializeComponent()

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

                'Si pas de CATIA ouvert et CATProcess inactif, sortie de Sub
                If Fonctions_CATIA.Ctrl_Recherche_Donnees_CATIA _
                    (CATIA_App, Doc_CATIA_Courant, Process_Courant) = False Then Exit Sub
        End Select

        'Valuation taille PictureBox
        Taille_Inc_PictureBox = (Me.PictureBox1.Width * 10) / 100

        'Valuation variable couleur
        Couleurs_OP(0) = Color.Red
        Couleurs_OP(1) = Color.Yellow

        'Couleur chgt outil/OP
        Me.Button6.BackColor = Couleurs_OP(0)
        Me.Button7.BackColor = Couleurs_OP(1)

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

    'Click bouton sélection dans CATIA

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'Demande de sélection dans CATIA
        If Fonctions_CATIA.Selection_Element_CATIA _
            (Doc_CATIA_Courant, "ManufacturingProgram", _
             "Sélectionner un programme d'usinage dans l'arbre", Prog_Courant) = False Then Exit Sub

        'Recherche de phase, programme courants, et si erreur, Message et sortie de Sub
        If Fonctions_CATIA.Recherche_Phase_Courante_Process _
            (Process_Courant, Prog_Courant, Phase_Courante) = False Then Exit Sub

        'Remplissage TextBox
        TextBox1.Text = Prog_Courant.Name

        'Activation bouton d'envoi
        Me.Button2.Enabled = True
    End Sub

    'Click bouton calcul du temps de cycle

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        'Si ComboBox non-vide
        If Me.ComboBox2.Text <> Nothing Then

            'Si erreur, renvoi False et Exit Sub
            If Fonctions_CATIA.Extract_Tab_Tps_Cycle_Prog _
            (Prog_Courant, Val(Me.ComboBox2.Text), Tab_Tps_Cycle) = False Then Exit Sub

            'Variable
            Dim Total_Tps As Double = Nothing

            'Boucle cumul des temps
            For i = 0 To UBound(Tab_Tps_Cycle)

                'Cumul des temps
                Total_Tps = Total_Tps + Tab_Tps_Cycle(i, 2)
            Next

            'Variable
            Dim Text_Result As String = Nothing

            'Composition texte
            Text_Result = "La durée totale du programme """ & Prog_Courant.Name & """ est de: " & _
                Conv_Tps_Cycle(Total_Tps) & "." & vbCrLf & vbCrLf & "Détail opération par opération:" & vbCrLf

            'Boucle détail des temps
            For i = 0 To UBound(Tab_Tps_Cycle)

                'Si pas la dernière ligne
                If i <> UBound(Tab_Tps_Cycle) Then

                    'Ecriture ligne
                    Text_Result = Text_Result & """" & Tab_Tps_Cycle(i, 1) & """: " & Conv_Tps_Cycle(Tab_Tps_Cycle(i, 2)) & "," & vbCrLf

                Else

                    'Ecriture fin de ligne
                    Text_Result = Text_Result & """" & Tab_Tps_Cycle(i, 1) & """: " & Conv_Tps_Cycle(Tab_Tps_Cycle(i, 2)) & "."
                End If
            Next

            'Remplissage TextBox
            Me.TextBox2.Text = Text_Result

            'Taille PictureBox 100%
            Me.PictureBox1.Width = Taille_Inc_PictureBox * 10

            'Valuation image
            Me.PictureBox1.Image = Fonctions_CATIA.Remp_Barre_Temps(Tab_Tps_Cycle, Me.PictureBox1, Couleurs_OP, Coeff_Image_Tps)

            'Changement %
            Me.Label4.Text = "100%"

            'Activation des contrôles
            Me.Button3.Enabled = True
            Me.Button4.Enabled = True
            Me.Button5.Enabled = True

            'Sinon...
        Else

            'Message
            Fonctions_Messages.Appel_Msg(47, 3, , )
        End If
    End Sub

    'Click bouton +

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        'Augmentation taille PictureBox
        Me.PictureBox1.Width = Me.PictureBox1.Width + Taille_Inc_PictureBox

        'Valuation image
        Me.PictureBox1.Image = Fonctions_CATIA.Remp_Barre_Temps(Tab_Tps_Cycle, Me.PictureBox1, Couleurs_OP, Coeff_Image_Tps)

        'Changement %
        Me.Label4.Text = (Val(Me.Label4.Text) + 10) & "%"
    End Sub

    'Click bouton -

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        'Si > 10%
        If Val(Me.Label4.Text) > 10 Then

            'Diminution taille PictureBox
            Me.PictureBox1.Width = Me.PictureBox1.Width - Taille_Inc_PictureBox

            'Valuation image
            Me.PictureBox1.Image = Fonctions_CATIA.Remp_Barre_Temps(Tab_Tps_Cycle, Me.PictureBox1, Couleurs_OP, Coeff_Image_Tps)

            'Changement %
            Me.Label4.Text = (Val(Me.Label4.Text) - 10) & "%"
        End If
    End Sub

    'Click bouton 100%

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        'Taille PictureBox 100%
        Me.PictureBox1.Width = Taille_Inc_PictureBox * 10

        'Valuation image
        Me.PictureBox1.Image = Fonctions_CATIA.Remp_Barre_Temps(Tab_Tps_Cycle, Me.PictureBox1, Couleurs_OP, Coeff_Image_Tps)

        'Changement %
        Me.Label4.Text = "100%"
    End Sub

    'Lors du passage avec le curseur sur la PictureBox...

    Private Sub PictureBox1_MouseMove(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.MouseMove

        'Si Picture <> Nothing
        If Not Me.PictureBox1.Image Is Nothing Then

            'Variables
            Dim Pt_Pos_Curseur, Pt_Debut_PictureBox As Integer
            Dim Cumul_Tps As Double = Nothing

            'Valuation variables
            Pt_Debut_PictureBox = Me.GroupBox4.Location.X + Me.GroupBox3.Location.X + _
                Me.Panel2.Location.X + Me.PictureBox1.Location.X + 1
            Pt_Pos_Curseur = PointToClient(Cursor.Position).X - Pt_Debut_PictureBox

            'Boucle lecture Tab temps
            For i = 0 To UBound(Tab_Tps_Cycle)

                'Cumul temps
                Cumul_Tps = Cumul_Tps + Int(Tab_Tps_Cycle(i, 2) * Coeff_Image_Tps)

                'Si position du curseur < temps cumulé
                If Pt_Pos_Curseur < Cumul_Tps Then

                    'Valaution ToolTip
                    ToolTip1.SetToolTip(Me.PictureBox1, "Nom de l'opération: " & Tab_Tps_Cycle(i, 1))

                    'Sortie de For
                    Exit For
                End If
            Next
        End If
    End Sub
End Class