Public Module Fonctions_Form

    'Contrôle d'un mot de passe

    Public Sub Ctrl_MDP(ByVal Form_MDP As Form, ByVal TextBox_MDP As System.Windows.Forms.TextBox, _
                     ByVal MDP_A_Trouver As String, ByRef Resultat_MDP As Boolean)

        'Si MDP OK
        If Fonctions_Diverses.Ctrl_Chaine_Seul(MDP_A_Trouver, 4, False, , , , TextBox_MDP.Text) = True Then

            'Valuation variable
            Resultat_MDP = True

            'Fermeture Form MDP
            Form_MDP.Dispose()

            'Sinon
        Else

            'Message
            Fonctions_Messages.Appel_Msg(28, 3, , )

            'Valuation variable
            Resultat_MDP = False

            'Vide et redonne le focus au TextBox
            TextBox_MDP.Text = Nothing
            TextBox_MDP.Focus()
        End If
    End Sub

    'Ouverture Form avec remplissage DGV et si erreur, retourne False

    Public Function Ouverture_Form_Avec_DGV _
        (ByRef Form_Concernee As Form, ByRef DGV_A_Remplir As DataGridView, _
         ByRef Tab_Chaines() As String, ByRef Num_Enreg_Actuel As Integer, _
         ByRef Num_Enreg_Maxi_CSV As Integer, ByRef Derniere_Ligne_DGV_Auto As Boolean, _
         Optional ByRef Num_Prog_Actuel As String = Nothing, _
         Optional ByRef Index_Machine As String = Nothing, _
         Optional ByVal Colonnes_A_Desactiver(,) As String = Nothing) As Boolean

        'Si fichier de config valué ou si noms de colonnes à désactiver différent Nothing
        If (UBound(Tab_Chaines, 1) = 6 Or UBound(Tab_Chaines, 1) = 11) Or Not Colonnes_A_Desactiver Is Nothing Then

            'Déclaration variable
            Dim Tab_Temp_Config_1(,) As String = Nothing

            'Lecture CSV et si erreur, désactivation commandes Form et retourne False
            If Fonctions_CSV.Lecture_CSV(Tab_Chaines(2), Tab_Temp_Config_1, 5) = False Then

                'Désactivation commandes Form et retourne False
                Fonctions_Form.Desactivation_Cmd_Form(Form_Concernee)
                Return False
            End If

            'Si fichier de config valué
            If (UBound(Tab_Chaines, 1) = 6 Or UBound(Tab_Chaines, 1) = 11) Then

                'Ajout nom du Listing dans Form
                Form_Concernee.Text = Form_Concernee.Text & ": " & Tab_Chaines(1)

                'Configuration DGV selon fichier de config
                Fonctions_DGV.Config_Colonnes_DGV(Tab_Temp_Config_1, Tab_Chaines(5), Tab_Chaines(6), _
                                                  Tab_Chaines(3), Tab_Chaines(4), DGV_A_Remplir)

                'Si 2ème fichier de config valué
                If UBound(Tab_Chaines, 1) = 11 Then

                    'Déclaration variable
                    Dim Tab_Temp_Config_2(,) As String = Nothing

                    'Lecture CSV et si erreur, désactivation commandes Form et retourne False
                    If Fonctions_CSV.Lecture_CSV(Tab_Chaines(7), Tab_Temp_Config_2, 5) = False Then

                        'Désactivation commandes Form et retourne False
                        Fonctions_Form.Desactivation_Cmd_Form(Form_Concernee)
                        Return False
                    End If

                    'Configuration DGV selon fichier de config
                    Fonctions_DGV.Config_Colonnes_DGV(Tab_Temp_Config_2, Tab_Chaines(10), Tab_Chaines(11), _
                                                      Tab_Chaines(8), Tab_Chaines(9), DGV_A_Remplir)
                End If

                'Si index programme trouvé dans tab config
                If Fonctions_Tableau.Renvoi_Num_Colonne_Tab(Tab_Temp_Config_1, "Index programme") <> -1 Then

                    'Valuation de l'index machine
                    Index_Machine = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab _
                        (Tab_Temp_Config_1, Tab_Chaines(1), Tab_Chaines(3), "Index programme", 1)
                End If
            End If

            'Si nom de colonnes à désactiver différent Nothing
            If Not Colonnes_A_Desactiver Is Nothing Then

                'Boucle relecture tab
                For Boucle = LBound(Colonnes_A_Desactiver, 1) To UBound(Colonnes_A_Desactiver, 1)

                    'Déclaration variable
                    Dim Result_Temp As Boolean = _
                        Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab _
                        (Tab_Temp_Config_1, Tab_Chaines(1), Colonnes_A_Desactiver(Boucle, 1), _
                         Colonnes_A_Desactiver(Boucle, 2), 1)

                    'Masquage colonne
                    DGV_A_Remplir.Columns(Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                    (DGV_A_Remplir, Colonnes_A_Desactiver(Boucle, 0))).Visible = Not Result_Temp
                Next
            End If
        End If

        'Déclaration variables
        Dim Tab_Temp_Listing(,) As String = Nothing

        'Lecture CSV et si erreur, désactivation commandes Form et retourne False
        If Fonctions_CSV.Lecture_CSV(Tab_Chaines(0), Tab_Temp_Listing, 5) = False Then

            'Désactivation commandes Form et retourne False
            Fonctions_Form.Desactivation_Cmd_Form(Form_Concernee)
            Return False
        End If

        'Remplissage DGV
        Fonctions_DGV.Remplissage_DGV_Complete(DGV_A_Remplir, Tab_Temp_Listing, Num_Enreg_Actuel, Num_Enreg_Maxi_CSV, _
                                               Derniere_Ligne_DGV_Auto, Num_Prog_Actuel, Index_Machine)

        'Contrôle mode visualisation
        Fonctions_DGV.Ctrl_Acces_Modif_DGV(DGV_A_Remplir)

        'Retourne True
        Return True
    End Function

    'Ecriture mode ouverture dans entête Form principale

    Public Sub Ecriture_Mode_Ouverture(ByVal Form_Concernee As Form)

        'Valeur de condition à trouver
        Select Case Niveau_Ouverture

            'Si mode visualisation
            Case 1

                'Ajout "Activée en mode visualisation" dans nom de fenêtre
                Form_Concernee.Text = ParamsApp.NameApp & ": Activée en mode visualisation"

                'Si mode modification
            Case 2

                'Ajout "Activée en mode modification" dans nom de fenêtre
                Form_Concernee.Text = ParamsApp.NameApp & ": Activée en mode modification"

                'Si mode administrateur
            Case 3

                'Ajout "Activée en mode administrateur" dans nom de fenêtre
                Form_Concernee.Text = ParamsApp.NameApp & ": Activée en mode administrateur"
        End Select
    End Sub

    'Remplissage ComboBox avec colonne CSV

    Public Function Remp_ComboBox_Avec_Colonne_CSV _
        (ByVal Chemin_CSV_A_Lire As String, ByVal Nom_Colonne_A_Lire As String, _
         ByVal ComboBox_A_Remplir As ComboBox, ByVal Type_MsgBox As Integer, _
         Optional ByVal Nom_Colonne_Filtre As String = Nothing) As Boolean

        'Variable
        Dim Tab_Temp(,) As String = Nothing

        'Lecture CSV et si erreur, renvoi False
        If Fonctions_CSV.Lecture_CSV(Chemin_CSV_A_Lire, Tab_Temp, Type_MsgBox) = False Then Return False

        'Vidage ComboBox
        ComboBox_A_Remplir.Items.Clear()

        'Affichage des infos avec colonne filtre
        If Nom_Colonne_Filtre <> Nothing Then

            'Boucle remplissage tab
            For Boucle = LBound(Tab_Temp, 1) + 1 To UBound(Tab_Temp, 1)

                'Si case fiche outils checkée
                If Tab_Temp(Boucle, Fonctions_Tableau.Renvoi_Num_Colonne_Tab(Tab_Temp, Nom_Colonne_Filtre)) = True Then

                    'Remplissage ComboBox
                    ComboBox_A_Remplir.Items.Add(Tab_Temp(Boucle, Fonctions_Tableau.Renvoi_Num_Colonne_Tab(Tab_Temp, Nom_Colonne_A_Lire)))
                End If
            Next

            'Sinon
        Else

            'Boucle remplissage tab
            For Boucle = LBound(Tab_Temp, 1) + 1 To UBound(Tab_Temp, 1)

                'Remplissage ComboBox
                ComboBox_A_Remplir.Items.Add(Tab_Temp(Boucle, Fonctions_Tableau.Renvoi_Num_Colonne_Tab(Tab_Temp, Nom_Colonne_A_Lire)))
            Next
        End If

        'Retourne True
        Return True
    End Function

    'Désactivation de toutes commandes d'une Form

    Public Sub Desactivation_Cmd_Form(ByVal Form_Concernee As Form)

        'Variable
        Dim Nb_Composant As Integer

        'Compte le nombre de composants
        Nb_Composant = Form_Concernee.Controls.Count

        'Boucle relecture de tous les composants
        For Boucle_Composant = 1 To Nb_Composant

            'Désactivation composants
            Form_Concernee.Controls(Boucle_Composant - 1).Enabled = False
        Next
    End Sub

    'Fermeture Form enfant

    Public Sub Fermeture_Form_Enfant _
        (ByVal Form_A_Fermer As Form, ByVal Form_Config As Boolean, _
         Optional ByRef Derniere_Ligne_DGV_Auto As Boolean = Nothing)

        'Si Form config et si en mode modif ou admin
        If Form_Config = True And Niveau_Ouverture >= 2 Then

            'Passage en mode modif
            Niveau_Ouverture = 2

            'Ecriture mode ouverture dans entête Form principale
            Fonctions_Form.Ecriture_Mode_Ouverture(Form_Principale)
        End If

        'Si DGV dans Form
        If Derniere_Ligne_DGV_Auto <> Nothing Then

            'Désactivation remplissage auto dernière ligne DGV
            Derniere_Ligne_DGV_Auto = False
        End If

        'Fermeture Form
        Form_A_Fermer.Dispose()
    End Sub

    'Remplissage TextBox avec sélection de chemin (FolderBrowserDialog)

    Public Sub Remplissage_TextBox_Chemin _
        (ByVal TextBox_Chemin As System.Windows.Forms.TextBox, _
         ByVal Dossier_Dialog As FolderBrowserDialog)

        'Si clic boutton sélection dossier
        If Dossier_Dialog.ShowDialog() = DialogResult.OK Then

            'Chargement du lien dans TextBox
            TextBox_Chemin.Text = Dossier_Dialog.SelectedPath
        End If
    End Sub

    'Remplissage TextBox point pivot, sélection aperçu et repositionnement TextBox

    Public Sub Calcul_Point_Pivot _
        (ByVal TextBox_Point_Pivot As System.Windows.Forms.TextBox, _
         ByVal TextBox_HT_MecaTool As System.Windows.Forms.TextBox, _
         ByVal TextBox_HT_Pince As System.Windows.Forms.TextBox, _
         ByRef TextBox_HT_Origine As System.Windows.Forms.TextBox, _
         ByRef Label_TextBox_HT_Origine As System.Windows.Forms.Label, _
         ByRef Image_Origine As PictureBox, ByVal Element_A_Focuser As Object)

        'Donne le focus à un autre élément que la TextBox (et évite la sélection de force)
        Element_A_Focuser.select()

        'Calcul hauteur origine
        TextBox_HT_Origine.Text = System.Math.Round((Val(TextBox_Point_Pivot.Text) - (Val(TextBox_HT_MecaTool.Text) + Val(TextBox_HT_Pince.Text))), 3) & "mm"

        'Si hauteur point de pivot > posage MecaTool...
        If Val(TextBox_HT_Origine.Text) >= 0 Then

            'Valuation apercu
            Image_Origine.Image = My.Resources.Apercu_Origine_Positive

            'Repositionnement TextBox
            TextBox_Point_Pivot.Location = New System.Drawing.Point(288, 265)
            TextBox_HT_Origine.Location = New System.Drawing.Point(255, 78)
            Label_TextBox_HT_Origine.Location = New System.Drawing.Point(235, 78)

            'Sinon
        ElseIf Val(TextBox_HT_Origine.Text) < 0 Then

            'Valuation apercu
            Image_Origine.Image = My.Resources.Apercu_Origine_Negative

            'Repositionnement TextBox
            TextBox_Point_Pivot.Location = New System.Drawing.Point(299, 363)
            TextBox_HT_Origine.Location = New System.Drawing.Point(278, 173)
            Label_TextBox_HT_Origine.Location = New System.Drawing.Point(258, 173)
        End If
    End Sub

    'Masquage/affichage de composants dans Form

    Public Sub Masquage_Affichage_Composants _
        (ByVal Composant_A_Masquer As Object, ByVal Bouton_Commande_Masquage As Object)

        'Si bouton commande checké
        If Bouton_Commande_Masquage.Checked = False Then

            'Masquage contrôle
            Composant_A_Masquer.Visible = False

            'Sinon
        Else

            'Affichage contrôle
            Composant_A_Masquer.Visible = True
        End If
    End Sub

    'Gestion ComboBox et Button suivant ouverture de Form dans Form principale

    Public Sub Gestion_Bouton_Et_ComboBox_Form _
        (ByRef Bouton_Ouverture As System.Windows.Forms.Button, ByVal Type_Traitement As Integer, _
         Optional ByRef ComboBox_Selection As System.Windows.Forms.ComboBox = Nothing, _
         Optional ByVal Chaine_Concernee As String = Nothing)

        'Si type de traitement = 1, retrait d'un item dans ComboBox
        If Type_Traitement = 1 Then

            'Si chaîne présente dans ComboBox
            If ComboBox_Selection.Items.Contains(Chaine_Concernee) = True Then

                'Suppression de la chaîne
                ComboBox_Selection.Items.RemoveAt(ComboBox_Selection.Items.IndexOf(Chaine_Concernee))

                'Si ComboBox vide
                If ComboBox_Selection.Items.Count = 0 Then

                    'Désactivation ComboBox et bouton
                    ComboBox_Selection.Enabled = False
                    Bouton_Ouverture.Enabled = False
                End If
            End If

            'Si type de traitement = 2, ajouter d'un item dans ComboBox
        ElseIf Type_Traitement = 2 Then

            'Si chaîne non-présente dans ComboBox
            If ComboBox_Selection.Items.Contains(Chaine_Concernee) = False Then

                'Ajout de la chaîne
                ComboBox_Selection.Items.Add(Chaine_Concernee)

                'Si ComboBox non-vide
                If ComboBox_Selection.Items.Count > 0 Then

                    'Activation ComboBox et bouton
                    ComboBox_Selection.Enabled = True
                    Bouton_Ouverture.Enabled = True
                End If
            End If

            'Si type de traitement = 3, changer l'état du bouton uniquement
        ElseIf Type_Traitement = 3 Then

            'Si bouton actif
            If Bouton_Ouverture.Enabled = True Then

                'Désactivation bouton
                Bouton_Ouverture.Enabled = False

                'Sinon
            Else

                'Activation bouton
                Bouton_Ouverture.Enabled = True
            End If
        End If
    End Sub
End Module