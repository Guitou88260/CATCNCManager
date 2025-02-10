Public Module Fonctions_DGV

    Public Function Renvoi_Valeur_Maxi_Colonne_DGV _
        (ByVal DGV_A_Lire As DataGridView, ByVal Nom_Colonne_A_Lire As String, _
         ByVal Traitement_Colonne_Prog As Boolean) As Integer

        'Valuation variable
        Renvoi_Valeur_Maxi_Colonne_DGV = 0

        'Boucle lecture ligne
        For Boucle_Ligne = 1 To DGV_A_Lire.Rows.Count - 1

            'Si colonne numéro de programme
            If Traitement_Colonne_Prog = True Then

                'Extraction des 4 chiffres uniquement et si valeur dans la DGV supérieure à la valeur du numéro
                If Mid(DGV_A_Lire.Rows.Item(Boucle_Ligne - 1).Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, Nom_Colonne_A_Lire)).Value, 3) > _
                    Renvoi_Valeur_Maxi_Colonne_DGV Then Renvoi_Valeur_Maxi_Colonne_DGV = _
                    Mid(DGV_A_Lire.Rows.Item(Boucle_Ligne - 1).Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, Nom_Colonne_A_Lire)).Value, 3)

                'Sinon
            Else

                'Si valeur dans la DGV supérieure à la valeur du numéro
                If DGV_A_Lire.Rows.Item(Boucle_Ligne - 1).Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, Nom_Colonne_A_Lire)).Value > _
                    Renvoi_Valeur_Maxi_Colonne_DGV Then Renvoi_Valeur_Maxi_Colonne_DGV = _
                    DGV_A_Lire.Rows.Item(Boucle_Ligne - 1).Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, Nom_Colonne_A_Lire)).Value
            End If
        Next
    End Function

    'Contrôle d'une chaîne contenue dans case DGV suivant ToolTip

    Public Function Ctrl_Chaine_DGV _
        (ByVal DGV_A_Lire As DataGridView, ByVal Chaine_A_Controler As String, _
         ByVal Colonne_DGV_A_Lire As Integer, ByVal Generation_Msg As Boolean) As Boolean

        'Valuation variable
        Ctrl_Chaine_DGV = True

        'Si colonne est de type TextBox ou ComboBox, colonne visible et cellule active
        If (DGV_A_Lire.Columns.Item(Colonne_DGV_A_Lire).GetType.ToString = "System.Windows.Forms.DataGridViewTextBoxColumn" Or _
            DGV_A_Lire.Columns.Item(Colonne_DGV_A_Lire).GetType.ToString = "System.Windows.Forms.DataGridViewComboBoxColumn") And _
        DGV_A_Lire.Columns.Item(Colonne_DGV_A_Lire).Visible = True And _
        DGV_A_Lire.CurrentRow.Cells.Item(Colonne_DGV_A_Lire).ReadOnly = False Then

            'Contrôle de la chaine et valuation variable
            Ctrl_Chaine_DGV = Fonctions_Diverses.Ctrl_Chaine_Suivant_Param _
                (Chaine_A_Controler, DGV_A_Lire.Columns.Item(Colonne_DGV_A_Lire).ToolTipText, True)
        End If
    End Function

    'Contrôle les chaînes d'une ligne courante d'une DGV (sans contrôle des colonnes cachées)

    Public Function Ctrl_Validation_Ligne_DGV _
        (ByVal DGV_A_Lire As DataGridView, ByVal Generation_Msg As Boolean) As Boolean

        'Valuation variable
        Ctrl_Validation_Ligne_DGV = True

        'Boucle lecture colonne
        For Boucle_Colonne = 1 To DGV_A_Lire.Columns.Count

            'Contrôle de la chaine et valuation variable
            If Fonctions_DGV.Ctrl_Chaine_DGV(DGV_A_Lire, DGV_A_Lire.CurrentRow.Cells.Item(Boucle_Colonne - 1).Value, _
                                              Boucle_Colonne - 1, True) = False Then Ctrl_Validation_Ligne_DGV = False
        Next
    End Function

    'Dimensionnement Tab depuis colonne DGV

    Public Function Dim_Tab_Depuis_DGV(ByVal DGV_A_Lire As DataGridView) As Integer

        'Valuation variable
        Dim_Tab_Depuis_DGV = 0

        'Boucle lecture colonne
        For Boucle_Colonne = 1 To DGV_A_Lire.Columns.Count

            'Si colonne différente d'un bouton ou si présence colonne programme dans DGV
            If DGV_A_Lire.Columns.Item(Boucle_Colonne - 1).GetType.ToString <> "System.Windows.Forms.DataGridViewButtonColumn" Or _
                Boucle_Colonne - 1 = Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, "Numéro programme") Then

                'Incrément compteur
                Dim_Tab_Depuis_DGV = Dim_Tab_Depuis_DGV + 1
            End If
        Next

        'Valuation variable
        Dim_Tab_Depuis_DGV = Dim_Tab_Depuis_DGV - 1
    End Function

    'Ecriture des entêtes DGV dans CSV

    Public Function Ecriture_Entete_DGV_Dans_CSV _
        (ByVal DGV_A_Lire As DataGridView, ByVal Lien_CSV As String) As Boolean

        'Gestion des exceptions
        Try

            'Déclaration variables
            Dim Tab_Temp(0, Fonctions_DGV.Dim_Tab_Depuis_DGV(DGV_A_Lire)) As String
            Dim Compteur As Integer = 0

            'Boucle lecture colonne
            For Boucle_Colonne = 1 To DGV_A_Lire.Columns.Count

                'Si colonne différente d'un bouton ou si présence colonne programme dans DGV
                If DGV_A_Lire.Columns.Item(Boucle_Colonne - 1).GetType.ToString <> "System.Windows.Forms.DataGridViewButtonColumn" Or
                    Boucle_Colonne - 1 = Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, "Numéro programme") Then

                    'Remplissage Tab
                    Tab_Temp(0, Compteur) = DGV_A_Lire.Columns.Item(Boucle_Colonne - 1).HeaderText

                    'Incrément compteur
                    Compteur = Compteur + 1
                End If
            Next

            'Ecriture texte complet
            System.IO.File.WriteAllText(Lien_CSV, Fonctions_CSV.Mise_En_Forme_Tab_Pour_Ecriture_CSV(Tab_Temp),
                                        System.Text.Encoding.Default)

            'Vidage variable
            Tab_Temp = Nothing

            'Retourne True
            Return True

            'Si chemin dossier inexistant
        Catch Excep As DirectoryNotFoundException

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(0, 5, , Excep.Message)
            Return False

            'Si chemin de fichier inexistant = 0
        Catch Excep As ArgumentException

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(0, 5, , Excep.Message)
            Return False

            'Si chemin de fichier trop long
        Catch Excep As PathTooLongException

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(0, 5, , Excep.Message)
            Return False

            'Pour toutes les erreurs
        Catch

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(5, 5, , )
            Return False
        End Try
    End Function

    'Recherche nom d'entête colonne et renvoi son numéro, et si non présente retourne -1

    Public Function Renvoi_Num_Colonne_DGV _
        (ByVal DGV_A_Lire As DataGridView,
         ByVal Entete_Colonne_A_Trouver As String) As Integer

        'Valuation variable
        Renvoi_Num_Colonne_DGV = Nothing

        'Boucle colonne
        For Boucle_Colonne = 1 To DGV_A_Lire.Columns.Count

            'Si chaîne trouvée
            If DGV_A_Lire.Columns.Item(Boucle_Colonne - 1).HeaderText = Entete_Colonne_A_Trouver Then

                'Valuation variable
                Renvoi_Num_Colonne_DGV = Boucle_Colonne - 1

                'Sortie de Fonction
                Exit Function
            End If
        Next

        'Valuation variable
        Renvoi_Num_Colonne_DGV = -1
    End Function

    'Recherche chaîne dans une colonne et renvoi le numéro de la ligne, et si non présente retourne -1

    Public Function Renvoi_Num_Ligne_DGV _
        (ByVal DGV_A_Lire As DataGridView, ByVal Nom_Colonne_A_Lire As String,
         ByVal Chaine_A_Trouver As String) As Integer

        'Valuation variable
        Renvoi_Num_Ligne_DGV = Nothing

        'Boucle colonne
        For Boucle_Ligne = 1 To DGV_A_Lire.Rows.Count

            'Sortie de boucle si ajout dernière ligne possible (ligne vide)
            If DGV_A_Lire.AllowUserToAddRows = True And Boucle_Ligne = DGV_A_Lire.Rows.Count Then Exit For

            'Si chaîne trouvée
            If DGV_A_Lire.Rows.Item(Boucle_Ligne - 1).Cells.Item _
                (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, Nom_Colonne_A_Lire)).Value = Chaine_A_Trouver Then

                'Valuation variable
                Renvoi_Num_Ligne_DGV = Boucle_Ligne - 1

                'Sortie de Fonction
                Exit Function
            End If
        Next

        'Valuation variable
        Renvoi_Num_Ligne_DGV = -1
    End Function

    'Envoi dans Tab des éléments présents dans ToolTips d'une DGV

    Public Sub Envoi_Element_ToolTip_DGV_Dans_Tab _
        (ByVal DGV_A_Lire As DataGridView, ByVal Nom_Colonne_Debut As String,
         ByVal Nom_Colonne_Fin As String, ByRef Tab_A_Valuer(,) As String,
         ByVal Num_Colonne_Tab_A_Lire As Integer, ByVal Num_Colonne_Tab_A_Valuer As Integer,
         ByVal Chaine_Debut_ToolTip As String, ByVal Chaine_Fin_ToolTip As String)

        'Boucle ligne Tab
        For Boucle_Ligne_Tab = LBound(Tab_A_Valuer, 1) To UBound(Tab_A_Valuer, 1)

            'Boucle colonne DGV
            For Boucle_Colonne_DGV = Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, Nom_Colonne_Debut) To _
                Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, Nom_Colonne_Fin)

                'Si colonne visible
                If DGV_A_Lire.Columns.Item(Boucle_Colonne_DGV).Visible = True And
                    DGV_A_Lire.Columns.Item(Boucle_Colonne_DGV).HeaderText =
                    Tab_A_Valuer(Boucle_Ligne_Tab, Num_Colonne_Tab_A_Lire) Then

                    'Valuation variable
                    Tab_A_Valuer(Boucle_Ligne_Tab, Num_Colonne_Tab_A_Valuer) =
                        Fonctions_Diverses.Renvoi_Chaine_Encadree _
                        (DGV_A_Lire.Columns.Item(Boucle_Colonne_DGV).ToolTipText,
                         Chaine_Debut_ToolTip, Chaine_Fin_ToolTip)
                End If
            Next
        Next
    End Sub

    'Remplace les paramètres contenus dans chaîne par valeur d'une DGV sans unité

    Public Sub Remplacement_Param_Par_Valeur_Dans_Chaine _
        (ByVal DGV_A_Lire As DataGridView, ByVal Nom_Colonne_Debut As String,
         ByVal Nom_Colonne_Fin As String, ByRef Tab_A_Valuer(,) As String,
         ByVal Num_Colonne_Tab_A_Lire As Integer, ByVal Num_Colonne_Tab_A_Valuer As Integer,
         ByVal Chaine_Encadrante As String)

        'Boucle ligne Tab
        For Boucle_Ligne_Tab = LBound(Tab_A_Valuer, 1) To UBound(Tab_A_Valuer, 1)

            'Transfert valeur
            Tab_A_Valuer(Boucle_Ligne_Tab, Num_Colonne_Tab_A_Valuer) =
                Tab_A_Valuer(Boucle_Ligne_Tab, Num_Colonne_Tab_A_Lire)

            'Boucle tant qu'un caractère encadrant est présent dans la chaîne à modifier
            While InStr(Tab_A_Valuer(Boucle_Ligne_Tab, Num_Colonne_Tab_A_Valuer), Chaine_Encadrante) <> 0

                'Boucle colonne DGV
                For Boucle_Colonne_DGV = Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, Nom_Colonne_Debut) To _
                    Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, Nom_Colonne_Fin)


                    'Si paramètre trouvé dans entête DGV
                    If DGV_A_Lire.Columns.Item(Boucle_Colonne_DGV).HeaderText =
                       Fonctions_Diverses.Renvoi_Chaine_Encadree _
                       (Tab_A_Valuer(Boucle_Ligne_Tab, Num_Colonne_Tab_A_Valuer),
                        Chaine_Encadrante, Chaine_Encadrante) Then

                        'Remplacement chaîne
                        Tab_A_Valuer(Boucle_Ligne_Tab, Num_Colonne_Tab_A_Valuer) =
                            Replace(Tab_A_Valuer(Boucle_Ligne_Tab, Num_Colonne_Tab_A_Valuer),
                                    Chaine_Encadrante & Fonctions_Diverses.Renvoi_Chaine_Encadree _
                                    (Tab_A_Valuer(Boucle_Ligne_Tab, Num_Colonne_Tab_A_Valuer), Chaine_Encadrante, Chaine_Encadrante) &
                                    Chaine_Encadrante, Fonctions_Diverses.Renvoi_Valeur_Sans_Unite _
                                    (DGV_A_Lire.CurrentRow.Cells.Item(Boucle_Colonne_DGV).Value, DGV_A_Lire.Columns.Item _
                                     (Boucle_Colonne_DGV).ToolTipText))
                    End If
                Next
            End While
        Next
    End Sub

    'Remplissage ComboBox ou DataGridViewComboBoxColumn avec colonne entière DGV

    Public Sub Remp_ComboBox_Avec_Colonne_DGV _
        (ByVal DGV_A_Lire As DataGridView, ByVal Nom_Colonne_A_Lire As String,
         ByRef ComboBox_A_Remplir As Object)

        'Vidage ComboBox
        ComboBox_A_Remplir.Items.Clear()

        'Boucle lecture ligne
        For Boucle_Ligne = 1 To DGV_A_Lire.Rows.Count - 1

            'Si case DGV non-vide et visible
            If DGV_A_Lire.Rows.Item(Boucle_Ligne - 1).Cells.Item _
                (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, Nom_Colonne_A_Lire)).Value <> Nothing And
                DGV_A_Lire.Rows.Item(Boucle_Ligne - 1).Visible = True Then

                'Si chaine non-présente dans ComboBox
                If ComboBox_A_Remplir.Items.Contains(DGV_A_Lire.Rows.Item(Boucle_Ligne - 1).Cells.Item _
                                                     (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, Nom_Colonne_A_Lire)).Value) = False Then

                    'Remplissage ComboBox
                    ComboBox_A_Remplir.Items.Add(DGV_A_Lire.Rows.Item(Boucle_Ligne - 1).Cells.Item _
                                                 (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, Nom_Colonne_A_Lire)).Value)
                End If
            End If
        Next
    End Sub

    'Désactivation édition DGV et suppression ligne cachées si mode visualisation

    Public Sub Ctrl_Acces_Modif_DGV(ByVal DGV_A_Lire As DataGridView)

        'Si mode visualisation
        If Niveau_Ouverture = 1 Then

            'Passage mode visu DGV Listing
            DGV_A_Lire.ReadOnly = True

            'Masquage des colonnes bouton sauf numéro de programme
            Fonctions_DGV.Masq_Colonne_Bouton_DGV(DGV_A_Lire)

            'Si colonne masquage ligne valuée
            If Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, "Masquer ligne") <> -1 Then

                'Suppression des lignes cachées dans DGV
                Fonctions_DGV.Sup_Lignes_Cachees_DGV(DGV_A_Lire)
            End If

            'Sinon
        Else

            'Passage mode modif DGV Listing
            DGV_A_Lire.ReadOnly = False
        End If
    End Sub

    'Masquage de toutes les colonnes type bouton DGV, sauf numéro de programme (si valué)

    Public Sub Masq_Colonne_Bouton_DGV(ByVal DGV_A_Lire As DataGridView)

        'Boucle lecture colonne
        For Boucle_Colonne = 1 To DGV_A_Lire.Columns.Count

            'Si colonne est de type Button et différent de colonne numéro de programme
            If DGV_A_Lire.Columns.Item(Boucle_Colonne - 1).GetType.ToString = "System.Windows.Forms.DataGridViewButtonColumn" And
                Boucle_Colonne - 1 <> Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, "Numéro programme") Then

                'Remplissage Tab
                DGV_A_Lire.Columns.Item(Boucle_Colonne - 1).Visible = False
            End If
        Next
    End Sub

    'Lecture tableau et écriture dans DGV complète

    Public Sub Remplissage_DGV_Complete _
        (ByRef DGV_A_Remplir As DataGridView, ByVal Tab_A_Lire As String(,),
         ByRef Num_Enreg_Actuel As Integer, ByRef Num_Enreg_Maxi_CSV As Integer,
         ByRef Derniere_Ligne_DGV_Auto As Boolean,
         Optional ByRef Num_Prog_Actuel As String = Nothing,
         Optional ByRef Index_Machine As String = Nothing)

        'Variables
        Num_Enreg_Maxi_CSV = Nothing
        Dim Compteur As Integer

        'Vidage DGV
        DGV_A_Remplir.Rows.Clear()

        'Boucle lecture ligne
        For Boucle_Ligne = LBound(Tab_A_Lire, 1) + 1 To UBound(Tab_A_Lire, 1)

            'Remise à 0 compteur
            Compteur = 0

            'Ajout ligne DGV
            DGV_A_Remplir.Rows.Add()

            'Boucle lecture colonne
            For Boucle_Colonne = 1 To DGV_A_Remplir.Columns.Count

                'Si colonne n'est pas de type Button ou égale à la colonne numéro de programme
                If DGV_A_Remplir.Columns.Item(Boucle_Colonne - 1).GetType.ToString <> "System.Windows.Forms.DataGridViewButtonColumn" Or
                    Boucle_Colonne - 1 = Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Remplir, "Numéro programme") Then

                    'Si colonne numéro d'enregistrement
                    If Boucle_Colonne - 1 = Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Remplir, "Numéro enregistrement") Then

                        'Variable
                        Num_Enreg_Actuel = Nothing

                        'Remplissage DGV
                        DGV_A_Remplir.Rows.Item(DGV_A_Remplir.NewRowIndex - 1).Cells.Item(Boucle_Colonne - 1).Value =
                            Tab_A_Lire(Boucle_Ligne, Compteur)

                        'Si trouve numéro d'enregistrement supérieur au dernier lu
                        If DGV_A_Remplir.Rows.Item(DGV_A_Remplir.NewRowIndex - 1).Cells.Item(Boucle_Colonne - 1).Value > Num_Enreg_Actuel Then _
                            Num_Enreg_Actuel = DGV_A_Remplir.Rows.Item(DGV_A_Remplir.NewRowIndex - 1).Cells.Item(Boucle_Colonne - 1).Value

                        'Si colonne numéro de programme
                    ElseIf Boucle_Colonne - 1 = Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Remplir, "Numéro programme") Then

                        'Variables
                        Num_Prog_Actuel = Nothing
                        Dim Dern_Num_Prog_Temp As Integer = Nothing

                        'Valuation variable
                        Dern_Num_Prog_Temp = Tab_A_Lire(Boucle_Ligne, Compteur)

                        'Si trouve numéro de programme supérieur au dernier lu
                        If Dern_Num_Prog_Temp > Num_Prog_Actuel Then Num_Prog_Actuel = Dern_Num_Prog_Temp

                        'Remplissage DGV
                        DGV_A_Remplir.Rows.Item(DGV_A_Remplir.NewRowIndex - 1).Cells.Item(Boucle_Colonne - 1).Value =
                            Index_Machine & Tab_A_Lire(Boucle_Ligne, Compteur)

                        'Sinon
                    Else

                        'Remplissage DGV
                        DGV_A_Remplir.Rows.Item(DGV_A_Remplir.NewRowIndex - 1).Cells.Item(Boucle_Colonne - 1).Value =
                            Tab_A_Lire(Boucle_Ligne, Compteur)
                    End If

                    'Incrément compteur
                    Compteur = Compteur + 1
                End If
            Next
        Next

        'Ajustement des largeurs de colonne
        DGV_A_Remplir.AutoResizeColumns()

        'Valuation variables
        Derniere_Ligne_DGV_Auto = True
        Num_Enreg_Maxi_CSV = Num_Enreg_Actuel
    End Sub

    'Suppression des lignes cachées (pour le mode visualisation)

    Public Sub Sup_Lignes_Cachees_DGV(ByVal DGV_A_Lire As DataGridView)

        'Variable
        Dim Compteur As Integer = 0

        'Boucle lecture ligne
        For Boucle_Ligne = 1 To DGV_A_Lire.Rows.Count - 1

            'Valuation variable
            Compteur = Compteur + 1

            'Si colonne ligne masquée checkée
            If DGV_A_Lire.Rows.Item(Compteur - 1).Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                                                             (DGV_A_Lire, "Masquer ligne")).Value = True Then

                'Suppression ligne
                DGV_A_Lire.Rows.Remove(DGV_A_Lire.Rows.Item(Compteur - 1))

                'Valuation variable (Recul d'une ligne suppression)
                Compteur = Compteur - 1
            End If
        Next
    End Sub

    'Lecture ligne courante DGV et écriture dans tableau 2D d'une ligne ou plus

    Public Sub Ecriture_Ligne_DGV_Dans_Tab _
        (ByVal DGV_A_Lire As DataGridView, ByRef Tab_A_Remplir(,) As String,
         ByVal Num_Ligne_Tab_A_Remplir As Integer)

        'Variable
        Dim Compteur As Integer = 0

        'Boucle lecture colonne
        For Boucle_Colonne = 1 To DGV_A_Lire.Columns.Count

            'Si colonne différente d'un bouton ou si présence colonne programme dans DGV
            If DGV_A_Lire.Columns.Item(Boucle_Colonne - 1).GetType.ToString <> "System.Windows.Forms.DataGridViewButtonColumn" Or
                Boucle_Colonne - 1 = Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, "Numéro programme") Then

                'Si colonne numéro de programme
                If Boucle_Colonne - 1 = Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, "Numéro programme") Then

                    'Remplissage Tab
                    Tab_A_Remplir(Num_Ligne_Tab_A_Remplir, Compteur) = Mid(DGV_A_Lire.CurrentRow.Cells.Item(Boucle_Colonne - 1).Value, 3)

                    'Si case à cocher
                ElseIf DGV_A_Lire.Columns.Item(Boucle_Colonne - 1).GetType.ToString = "System.Windows.Forms.DataGridViewCheckBoxColumn" Then

                    'Remplissage Tab
                    Tab_A_Remplir(Num_Ligne_Tab_A_Remplir, Compteur) = Fonctions_CSV.Ecriture_Result_CheckBox(DGV_A_Lire, Boucle_Colonne - 1)

                    'Sinon
                Else

                    'Remplissage Tab
                    Tab_A_Remplir(Num_Ligne_Tab_A_Remplir, Compteur) = DGV_A_Lire.CurrentRow.Cells.Item(Boucle_Colonne - 1).Value
                End If

                'Incrément compteur
                Compteur = Compteur + 1
            End If
        Next
    End Sub

    'Enregistrement ligne courante DGV dans CSV

    Public Function Enregistrer_Ligne_DGV_Et_CSV _
        (ByVal DGV_A_Lire As DataGridView, ByVal Lien_CSV As String,
         ByRef Num_Enreg_Actuel As Integer, ByRef Num_Enreg_Maxi_CSV As Integer,
         Optional ByRef Num_Prog_Actuel As String = Nothing) As Boolean

        'Gestion des exceptions
        Try

            'Contrôle de chaîne de toutes les colonnes DGV
            If Fonctions_DGV.Ctrl_Validation_Ligne_DGV(DGV_A_Lire, True) = False Then Return False

            'Variable
            Dim Tab_Listing_Initial(,) As String = Nothing

            'Lecture CSV et si erreur, retourne False
            If Fonctions_CSV.Lecture_CSV(Lien_CSV, Tab_Listing_Initial, 5) = False Then Return False

            'Ecriture date de modification et auteur
            DGV_A_Lire.CurrentRow.Cells.Item _
                (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, "Date modification")).Value =
                Fonctions_Diverses.Renvoi_Maintenant_Et_Auteur(1)
            DGV_A_Lire.CurrentRow.Cells.Item _
                (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, "Auteur modification")).Value =
                Fonctions_Diverses.Renvoi_Maintenant_Et_Auteur(2)

            'Si numéro d'enregistrement DGV > au plus grand numéro d'enregistrement du CSV + 1, Msg et Return False
            If DGV_A_Lire.CurrentRow.Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                                                (DGV_A_Lire, "Numéro enregistrement")).Value > Num_Enreg_Maxi_CSV + 1 Then

                'Message
                Fonctions_Messages.Appel_Msg(37, 3, , Num_Enreg_Maxi_CSV + 1)

                'Retourne False
                Return False
            End If

            'Variable
            Dim Num_Enreg_Maxi_CSV_Temp As Integer = Nothing

            'Recherche du numéro d'enregistrement le plus grand dans CSV et si erreur, Return False
            If Fonctions_CSV.Renvoi_Valeur_Maxi_CSV(Lien_CSV, "Numéro enregistrement",
                                      Num_Enreg_Maxi_CSV, 5) = False Then Return False

            'Si numéro d'enregistrement lu >= que l'initial (veut dire qu'une ligne à été créée par quelqu'un d'autre)
            If Num_Enreg_Maxi_CSV_Temp > Num_Enreg_Maxi_CSV Then

                'Message
                Fonctions_Messages.Appel_Msg(38, 3, , )

                'Retourne False
                Return False
            End If

            'Variable
            Dim Result_MsgBox As MsgBoxResult = Fonctions_Messages.Appel_Msg(11, 2, , )

            'Si OK
            If Result_MsgBox = MsgBoxResult.Yes Then

                'Variable
                Dim Nouveau_Tab_Listing(,) As String = Nothing

                'Si numéro d'enregistrement trouvé dans tableau
                If Fonctions_Tableau.Presence_Chaine_Dans_Tab(Tab_Listing_Initial, DGV_A_Lire.CurrentRow.Cells.Item _
                                                          (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, "Numéro enregistrement")).Value,
                                                          Fonctions_Tableau.Renvoi_Num_Colonne_Tab _
                                                          (Tab_Listing_Initial, "Numéro enregistrement"), , ) = True Then

                    'Variable
                    ReDim Nouveau_Tab_Listing(UBound(Tab_Listing_Initial, 1), UBound(Tab_Listing_Initial, 2))

                    'Boucle lecture lignes
                    For Boucle_Ligne = LBound(Tab_Listing_Initial, 1) To UBound(Tab_Listing_Initial, 1)

                        'Si différent du numéro d'enregistrement, écrire dans le tableau temp
                        If Tab_Listing_Initial(Boucle_Ligne, Fonctions_Tableau.Renvoi_Num_Colonne_Tab _
                                               (Tab_Listing_Initial, "Numéro enregistrement")).ToString <> DGV_A_Lire.CurrentRow.Cells.Item _
                                           (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, "Numéro enregistrement")).Value.ToString Then

                            'Boucle lecture colonnes
                            For Boucle_Colonne = LBound(Tab_Listing_Initial, 2) To UBound(Tab_Listing_Initial, 2)

                                'Remplissage tableau
                                Nouveau_Tab_Listing(Boucle_Ligne, Boucle_Colonne) = Tab_Listing_Initial(Boucle_Ligne, Boucle_Colonne)
                            Next

                            'Sinon
                        Else

                            'Lecture ligne courante DGV et écriture dans tableau temp
                            Fonctions_DGV.Ecriture_Ligne_DGV_Dans_Tab(DGV_A_Lire, Nouveau_Tab_Listing, Boucle_Ligne)
                        End If
                    Next

                    'Sinon
                Else

                    'Variable
                    ReDim Nouveau_Tab_Listing(UBound(Tab_Listing_Initial, 1) + 1, UBound(Tab_Listing_Initial, 2))

                    'Boucle lecture lignes
                    For Boucle_Ligne = LBound(Tab_Listing_Initial, 1) To UBound(Tab_Listing_Initial, 1)

                        'Boucle lecture colonnes
                        For Boucle_Colonne = LBound(Tab_Listing_Initial, 2) To UBound(Tab_Listing_Initial, 2)

                            'Remplissage tableau
                            Nouveau_Tab_Listing(Boucle_Ligne, Boucle_Colonne) = Tab_Listing_Initial(Boucle_Ligne, Boucle_Colonne)
                        Next
                    Next

                    'Vidage variable
                    Tab_Listing_Initial = Nothing

                    'Lecture ligne courante DGV et écriture dans tableau temp
                    Fonctions_DGV.Ecriture_Ligne_DGV_Dans_Tab(DGV_A_Lire, Nouveau_Tab_Listing, UBound(Nouveau_Tab_Listing))
                End If

                'Ecriture texte complet
                System.IO.File.WriteAllText(Lien_CSV, Fonctions_CSV.Mise_En_Forme_Tab_Pour_Ecriture_CSV _
                                            (Nouveau_Tab_Listing), System.Text.Encoding.Default)

                'Recherche numéro d'enregistrement le plus grand dans Tab
                Num_Enreg_Maxi_CSV = Fonctions_Tableau.Renvoi_Valeur_Maxi_Colonne_Tab(Nouveau_Tab_Listing, "Numéro enregistrement", False)

                'Vidage variable
                Nouveau_Tab_Listing = Nothing

                'Recherche numéro d'enregistrement le plus grand dans DGV et valuation variable
                Num_Enreg_Actuel = Fonctions_DGV.Renvoi_Valeur_Maxi_Colonne_DGV(DGV_A_Lire, "Numéro enregistrement", False)

                'Si numéro de colonne programme valué
                If Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, "Numéro programme") <> -1 Then

                    'Recherche numéro de programme le plus grand dans DGV et valuation variable
                    Num_Prog_Actuel = Fonctions_DGV.Renvoi_Valeur_Maxi_Colonne_DGV(DGV_A_Lire, "Numéro programme", True)
                End If

                'Message
                Fonctions_Messages.Appel_Msg(12, 1, , )

                'Si NOK
            Else

                'Message
                Fonctions_Messages.Appel_Msg(13, 1, , )

                'Retourne False
                Return False
            End If

            'Retourne True
            Return True

            'Si chemin dossier inexistant
        Catch Excep As DirectoryNotFoundException

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(0, 5, , Excep.Message)
            Return False

            'Si chemin de fichier inexistant = 0
        Catch Excep As ArgumentException

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(0, 5, , Excep.Message)
            Return False

            'Si chemin de fichier trop long
        Catch Excep As PathTooLongException

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(0, 5, , Excep.Message)
            Return False

            'Pour toutes les erreurs
        Catch

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(5, 5, , )
            Return False
        End Try
    End Function

    'Suppression ligne courante DGV et CSV

    Public Function Sup_Ligne_DGV_Et_CSV _
        (ByVal DGV_A_Lire As DataGridView, ByVal Lien_CSV As String,
         ByRef Num_Enreg_Actuel As Integer, ByRef Num_Enreg_Maxi_CSV As Integer,
         Optional ByRef Num_Prog_Actuel As String = Nothing) As Boolean

        'Gestion des exceptions
        Try

            'Variable
            Dim Tab_Listing_Initial(,) As String = Nothing

            'Lecture CSV et si erreur, retourne False
            If Fonctions_CSV.Lecture_CSV(Lien_CSV, Tab_Listing_Initial, 5) = False Then Return False

            'Variable
            Dim Result_MsgBox As MsgBoxResult = Fonctions_Messages.Appel_Msg(14, 2, , )

            'Si OK
            If Result_MsgBox = MsgBoxResult.Yes Then

                'Variable
                Dim Nouveau_Tab_Listing(,) As String = Nothing

                'Si numéro d'enregistrement trouvé dans tableau
                If Fonctions_Tableau.Presence_Chaine_Dans_Tab _
                    (Tab_Listing_Initial, DGV_A_Lire.CurrentRow.Cells.Item _
                     (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, "Numéro enregistrement")).Value,
                     Fonctions_Tableau.Renvoi_Num_Colonne_Tab(Tab_Listing_Initial, "Numéro enregistrement"), , ) = True Then

                    'Variable
                    Dim Compteur As Integer = 0
                    ReDim Nouveau_Tab_Listing(UBound(Tab_Listing_Initial, 1) - 1, UBound(Tab_Listing_Initial, 2))

                    'Boucle lecture lignes
                    For Boucle_Ligne = LBound(Tab_Listing_Initial, 1) To UBound(Tab_Listing_Initial, 1)

                        'Si différent du numéro d'enregistrement, écrire dans le tableau temp
                        If Tab_Listing_Initial(Boucle_Ligne, Fonctions_Tableau.Renvoi_Num_Colonne_Tab(Tab_Listing_Initial, "Numéro enregistrement")).ToString <>
                            DGV_A_Lire.CurrentRow.Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, "Numéro enregistrement")).Value.ToString Then

                            'Boucle lecture colonnes
                            For Boucle_Colonne = LBound(Tab_Listing_Initial, 2) To UBound(Tab_Listing_Initial, 2)

                                'Remplissage tableau
                                Nouveau_Tab_Listing(Compteur, Boucle_Colonne) = Tab_Listing_Initial(Boucle_Ligne, Boucle_Colonne)
                            Next

                            'Incrémentation compteur
                            Compteur = Compteur + 1
                        End If
                    Next

                    'Vidage variable
                    Tab_Listing_Initial = Nothing

                    'Ecriture texte complet
                    System.IO.File.WriteAllText(Lien_CSV, Fonctions_CSV.Mise_En_Forme_Tab_Pour_Ecriture_CSV(Nouveau_Tab_Listing), System.Text.Encoding.Default)

                    'Recherche numéro d'enregistrement le plus grand dans Tab
                    Num_Enreg_Maxi_CSV = Fonctions_Tableau.Renvoi_Valeur_Maxi_Colonne_Tab(Nouveau_Tab_Listing, "Numéro enregistrement", False)

                    'Vidage variable
                    Nouveau_Tab_Listing = Nothing

                    'Sinon
                Else

                    'Recherche du numéro d'enregistrement le plus grand dans CSV et si erreur, Return False
                    If Fonctions_CSV.Renvoi_Valeur_Maxi_CSV(Lien_CSV, "Numéro enregistrement", Num_Enreg_Maxi_CSV, 5) = False Then Return False
                End If

                'Suppression de la ligne DGV
                DGV_A_Lire.Rows.Remove(DGV_A_Lire.CurrentRow)

                'Recherche numéro d'enregistrement le plus grand dans DGV et valuation variable
                Num_Enreg_Actuel = Fonctions_DGV.Renvoi_Valeur_Maxi_Colonne_DGV(DGV_A_Lire, "Numéro enregistrement", False)

                'Si numéro de colonne programme trouvée dans DGV
                If Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, "Numéro programme") <> -1 Then

                    'Recherche numéro de programme le plus grand dans DGV et valuation variable
                    Num_Prog_Actuel = Fonctions_DGV.Renvoi_Valeur_Maxi_Colonne_DGV(DGV_A_Lire, "Numéro programme", True)
                End If

                'Message
                Fonctions_Messages.Appel_Msg(15, 1, , )

                'Si NOK
            Else

                'Message
                Fonctions_Messages.Appel_Msg(16, 1, , )

                'Retourne True
                Return False
            End If

            'Retourne True
            Return True

            'Si chemin dossier inexistant
        Catch Excep As DirectoryNotFoundException

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(0, 5, , Excep.Message)
            Return False

            'Si chemin de fichier inexistant = 0
        Catch Excep As ArgumentException

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(0, 5, , Excep.Message)
            Return False

            'Si chemin de fichier trop long
        Catch Excep As PathTooLongException

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(0, 5, , Excep.Message)
            Return False

            'Pour toutes les erreurs
        Catch

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(5, 5, , )
            Return False
        End Try
    End Function

    'Configuration DGV suivant fichier de config

    Public Sub Config_Colonnes_DGV _
        (ByVal Tab_Config(,) As String, ByVal Nom_Colonne_Debut As String, ByVal Nom_Colonne_Fin As String,
         ByVal Nom_Colonne_Config_Selectionnee As String, ByVal Config_Selectionnee As String,
         ByRef DGV_A_Configurer As DataGridView)

        'Boucle de masquage des colonnes suivant fichier de configuration
        For Boucle = Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Configurer, Nom_Colonne_Debut) To _
            Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Configurer, Nom_Colonne_Fin)

            'Masquage colonne
            DGV_A_Configurer.Columns(Boucle).Visible = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab _
                (Tab_Config, Config_Selectionnee, Nom_Colonne_Config_Selectionnee,
                 DGV_A_Configurer.Columns(Boucle).HeaderText, 1)
        Next
    End Sub

    'Ecriture dernière ligne dans DGV

    Public Sub Ecriture_Derniere_Ligne_DGV _
        (ByRef DGV_A_Remplir As DataGridView, ByRef Num_Enreg_Actuel As Integer,
         Optional ByRef Num_Prog_Actuel As String = Nothing, Optional ByVal Index_Machine As String = Nothing,
         Optional ByVal Tab_Chaines_Supplementaires(,) As String = Nothing,
         Optional ByVal Tab_Noms_Colonnes_Check() As String = Nothing)

        'Boucle lecture colonne
        For Boucle_Colonne = 1 To DGV_A_Remplir.Columns.Count

            'Si colonne est de type TextBox et unité présente dans ToolTip colonne
            If DGV_A_Remplir.Columns.Item(Boucle_Colonne - 1).GetType.ToString = "System.Windows.Forms.DataGridViewTextBoxColumn" And
                InStr(DGV_A_Remplir.Columns.Item(Boucle_Colonne - 1).ToolTipText, "(10)") <> 0 And
                DGV_A_Remplir.Columns.Item(Boucle_Colonne - 1).Visible = True Then

                'Remplissage valeur par défaut et unité
                DGV_A_Remplir.Rows.Item(DGV_A_Remplir.NewRowIndex - 1).Cells.Item(Boucle_Colonne - 1).Value = 0 &
                    Fonctions_Diverses.Renvoi_Chaine_Encadree(DGV_A_Remplir.Columns.Item(Boucle_Colonne - 1).ToolTipText,
                                                              "(10) Unités de la valeur : """, """")
            End If
        Next

        'Incrémentation compteur N° enregistrement
        Num_Enreg_Actuel = Num_Enreg_Actuel + 1

        'Ecriture numéro d'enregistrement
        DGV_A_Remplir.Rows.Item(DGV_A_Remplir.NewRowIndex - 1).Cells.Item _
                (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Remplir, "Numéro enregistrement")).Value = Num_Enreg_Actuel

        'Si numéro de programme actuel <> Nothing
        If Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Remplir, "Numéro programme") <> -1 Then

            'Incrémentation compteur N° programme
            Num_Prog_Actuel = Num_Prog_Actuel + 1

            'Boucle pour compter le nombre de caractère dans numéro de programme
            While Len(Num_Prog_Actuel) < 4

                'Valuation numéro
                Num_Prog_Actuel = 0 & Num_Prog_Actuel
            End While

            'Ecriture numéro de programme
            DGV_A_Remplir.Rows.Item(DGV_A_Remplir.NewRowIndex - 1).Cells.Item _
                (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Remplir, "Numéro programme")).Value =
                Index_Machine & Num_Prog_Actuel
        End If

        'Si tableau des colonnes à checker valué
        If Not Tab_Noms_Colonnes_Check Is Nothing Then

            'Boucle relecture ligne Tab
            For Boucle_Ligne = LBound(Tab_Noms_Colonnes_Check, 1) To UBound(Tab_Noms_Colonnes_Check, 1)

                'Valuation des CheckBox à la création (pour désactivation des cases BT si pièce dans PLM)
                DGV_A_Remplir.Rows.Item(DGV_A_Remplir.NewRowIndex - 1).Cells.Item _
                    (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Remplir, Tab_Noms_Colonnes_Check(Boucle_Ligne))).Value = False
            Next
        End If

        'Si Tab chaines supplémentaires <> Nothing
        If Not Tab_Chaines_Supplementaires Is Nothing Then

            'Boucle relecture ligne Tab
            For Boucle_Ligne = LBound(Tab_Chaines_Supplementaires, 1) To UBound(Tab_Chaines_Supplementaires, 1)

                'Ecriture chaîne supplémentaire
                DGV_A_Remplir.Rows.Item(DGV_A_Remplir.NewRowIndex - 1).Cells.Item _
                    (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Remplir, Tab_Chaines_Supplementaires(Boucle_Ligne, 0))).Value =
                    Tab_Chaines_Supplementaires(Boucle_Ligne, 1)
            Next
        End If

        'Ecriture date de création et auteur
        DGV_A_Remplir.Rows.Item(DGV_A_Remplir.NewRowIndex - 1).Cells.Item _
                (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Remplir, "Date création")).Value =
                Fonctions_Diverses.Renvoi_Maintenant_Et_Auteur(1)
        DGV_A_Remplir.Rows.Item(DGV_A_Remplir.NewRowIndex - 1).Cells.Item _
                (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Remplir, "Auteur création")).Value =
                Fonctions_Diverses.Renvoi_Maintenant_Et_Auteur(2)
    End Sub

    'Contrôle si ce pas la dernière ligne DGV et niveau d'ouverture

    Public Function Ctrl_Ligne_DGV_Niv_Ouverture(ByVal DGV_A_Lire As DataGridView) As Integer

        'Si ce n'est pas la dernière ligne
        If DGV_A_Lire.CurrentRow.Index <> DGV_A_Lire.NewRowIndex Then

            'Si pas en mode visualisation
            If Niveau_Ouverture >= 2 Then

                'Retourne True
                Return 1
            End If

            'Sinon
        Else

            'Si pas en mode visualisation
            If Niveau_Ouverture >= 2 Then

                'Retourne True
                Return 2
            End If
        End If

        'Retourne False
        Return Nothing
    End Function

    'Entrée en édition case DGV (extraction valeur numérique sans unité)

    Public Sub Entree_Edition_Case_DGV _
        (ByRef DGV_A_Lire As DataGridView,
         ByVal Evenement_Case As System.Windows.Forms.DataGridViewCellEventArgs)

        'Si ce n'est pas la dernière ligne, si pas en mode visualisation et si colonne de type TextBox
        If Fonctions_DGV.Ctrl_Ligne_DGV_Niv_Ouverture(DGV_A_Lire) = 1 And
            DGV_A_Lire.Columns.Item(Evenement_Case.ColumnIndex).GetType.ToString =
            "System.Windows.Forms.DataGridViewTextBoxColumn" Then

            'Extraction valeur numérique sans unité
            DGV_A_Lire.CurrentCell.Value = Fonctions_Diverses.
                Renvoi_Valeur_Sans_Unite(DGV_A_Lire.CurrentCell.Value,
                                         DGV_A_Lire.Columns.Item(Evenement_Case.ColumnIndex).ToolTipText)
        End If
    End Sub

    'Sortie d'édition case DGV (extraction valeur numérique et écriture unité)

    Public Sub Sortie_Edition_Case_DGV _
        (ByRef DGV_A_Lire As DataGridView,
         ByVal Evenement_Case As System.Windows.Forms.DataGridViewCellEventArgs)

        'Si ce n'est pas la dernière ligne et si pas en mode visualisation
        If Fonctions_DGV.Ctrl_Ligne_DGV_Niv_Ouverture(DGV_A_Lire) = 1 And
            DGV_A_Lire.Columns.Item(Evenement_Case.ColumnIndex).GetType.ToString =
            "System.Windows.Forms.DataGridViewTextBoxColumn" Then

            'Ecriture valeur avec unité
            DGV_A_Lire.CurrentCell.Value =
                Fonctions_Diverses.Renvoi_Valeur_Avec_Unite(DGV_A_Lire.CurrentCell.Value,
                                         DGV_A_Lire.Columns.Item(Evenement_Case.ColumnIndex).ToolTipText)
        End If
    End Sub

    Public Sub Gestion_Dynamique_CheckBox _
        (ByRef DGV_A_Lire As DataGridView, ByVal Nom_Colonne_Checkee As String,
         ByVal Tab_Noms_Colonnes_A_Traiter() As String, ByVal Vidage_Colonne As Boolean,
         ByVal Type_Traitement As Integer, ByVal Lecture_Colonnes_Intermediaires As Boolean,
         Optional ByRef Num_Ligne_DGV_A_Traiter As Integer = Nothing)

        'Si colonne vérouillée, ne rien faire (pas de modif de colonne vérouillée)
        If DGV_A_Lire.CurrentRow.Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                                            (DGV_A_Lire, Nom_Colonne_Checkee)).ReadOnly = False Then

            'Si colonne checkée
            If DGV_A_Lire.CurrentRow.Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                                            (DGV_A_Lire, Nom_Colonne_Checkee)).Value = True Then

                'Valuation colonne checkée
                DGV_A_Lire.CurrentRow.Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                                                 (DGV_A_Lire, Nom_Colonne_Checkee)).Value = False

                'Sinon
            Else

                'Valuation colonne checkée
                DGV_A_Lire.CurrentRow.Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                                                 (DGV_A_Lire, Nom_Colonne_Checkee)).Value = True
            End If

            'Vidage et blocage des colonnes si Check
            Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
                (DGV_A_Lire, Nom_Colonne_Checkee, Tab_Noms_Colonnes_A_Traiter,
                 Vidage_Colonne, Type_Traitement, Lecture_Colonnes_Intermediaires,
                 Num_Ligne_DGV_A_Traiter)
        End If
    End Sub

    'Blocage et coloriage colonnes dans DGV si autre colonne vide ou remplie

    Public Sub Controle_Et_Blocage_Coloriage_Colonnes_DGV _
        (ByRef DGV_A_Lire As DataGridView, ByVal Nom_Colonne_Maitresse As String,
         ByVal Tab_Noms_Colonnes_A_Traiter() As String, ByVal Vidage_Colonnes As Boolean,
         ByVal Type_Traitement As Integer, ByVal Lecture_Colonnes_Intermediaires As Boolean,
         Optional ByRef Num_Ligne_DGV_A_Traiter As Integer = -1)

        'Si lecture des colonnes intermédiaires
        If Lecture_Colonnes_Intermediaires = True And UBound(Tab_Noms_Colonnes_A_Traiter, 1) = 1 Then

            'Déclaration variables
            Dim Tab_Noms_Colonnes_Temp((Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, Tab_Noms_Colonnes_A_Traiter(1)) -
                                   Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, Tab_Noms_Colonnes_A_Traiter(0)))) As String
            Dim Compteur As Integer = 0

            'Boucle remplissage Tab
            For Boucle_Tab = Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, Tab_Noms_Colonnes_A_Traiter(0)) To _
                Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, Tab_Noms_Colonnes_A_Traiter(1))

                'Remplissage Tab avec entête DGV des colonnes à traiter
                Tab_Noms_Colonnes_Temp(Compteur) = DGV_A_Lire.Columns.Item(Boucle_Tab).HeaderText

                'Incrément compteur
                Compteur = Compteur + 1
            Next

            'Redimensionnement Tab
            ReDim Tab_Noms_Colonnes_A_Traiter(UBound(Tab_Noms_Colonnes_Temp))

            'Valuation variable
            Tab_Noms_Colonnes_A_Traiter = Tab_Noms_Colonnes_Temp
        End If

        'Variable
        Dim Result_Traitement As Boolean = False

        'Boucle Tab des colonnes
        For Boucle_Tab = LBound(Tab_Noms_Colonnes_A_Traiter, 1) To UBound(Tab_Noms_Colonnes_A_Traiter, 1)

            'Si numéro de ligne à traiter = -1
            If Num_Ligne_DGV_A_Traiter = -1 Then

                'Boucle lecture ligne
                For Boucle_Ligne = 0 To DGV_A_Lire.Rows.Count - 1

                    'Sortie de boucle si ajout dernière ligne possible (ligne vide)
                    If DGV_A_Lire.AllowUserToAddRows = True And
                        Boucle_Ligne = DGV_A_Lire.Rows.Count - 1 Then Exit For

                    'Test de la valeur de la cellule courante suivant le type de colonne DGV
                    Result_Traitement = Fonctions_DGV.Test_Valeur_Suivant_Type_Colonne_DGV _
                        (DGV_A_Lire.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, Nom_Colonne_Maitresse)),
                         DGV_A_Lire.Rows.Item(Boucle_Ligne).Cells.Item _
                         (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, Nom_Colonne_Maitresse)))

                    'Si type traitement = 1 et Test = True ou si type traitement = 2 et Test = False
                    If Type_Traitement = 1 And Result_Traitement = True Or
                        Type_Traitement = 2 And Result_Traitement = False Then

                        'Vidage/coloriage/blocage cellule DGV
                        Fonctions_DGV.Coloriage_Blocage_Celulle_DGV _
                            (DGV_A_Lire, Tab_Noms_Colonnes_A_Traiter(Boucle_Tab), 1, Boucle_Ligne, Vidage_Colonnes)

                        'Si type traitement = 1 et Test = False ou si type traitement = 2 et Test = True
                    ElseIf Type_Traitement = 1 And Result_Traitement = False Or
                        Type_Traitement = 2 And Result_Traitement = True Then

                        'Déblocage/suppression coloriage cellule DGV
                        Fonctions_DGV.Coloriage_Blocage_Celulle_DGV _
                            (DGV_A_Lire, Tab_Noms_Colonnes_A_Traiter(Boucle_Tab), 2, Boucle_Ligne, Vidage_Colonnes)
                    End If
                Next

                'Sinon
            Else

                'Test de la valeur de la cellule courante suivant le type de colonne DGV
                Result_Traitement = Fonctions_DGV.Test_Valeur_Suivant_Type_Colonne_DGV _
                    (DGV_A_Lire.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, Nom_Colonne_Maitresse)),
                     DGV_A_Lire.Rows.Item(Num_Ligne_DGV_A_Traiter).Cells.Item _
                     (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, Nom_Colonne_Maitresse)))

                'Si type traitement = 1 et Test = True ou si type traitement = 2 et Test = False
                If Type_Traitement = 1 And Result_Traitement = True Or
                    Type_Traitement = 2 And Result_Traitement = False Then

                    'Vidage/coloriage/blocage cellule DGV
                    Fonctions_DGV.Coloriage_Blocage_Celulle_DGV _
                        (DGV_A_Lire, Tab_Noms_Colonnes_A_Traiter(Boucle_Tab), 1, Num_Ligne_DGV_A_Traiter, Vidage_Colonnes)

                    'Si type traitement = 1 et Test = False ou si type traitement = 2 et Test = True
                ElseIf Type_Traitement = 1 And Result_Traitement = False Or
                    Type_Traitement = 2 And Result_Traitement = True Then

                    'Déblocage/suppression coloriage cellule DGV
                    Fonctions_DGV.Coloriage_Blocage_Celulle_DGV _
                        (DGV_A_Lire, Tab_Noms_Colonnes_A_Traiter(Boucle_Tab), 2, Num_Ligne_DGV_A_Traiter, Vidage_Colonnes)
                End If
            End If
        Next
    End Sub

    'Blocage/coloriage cellule DGV

    Public Sub Coloriage_Blocage_Celulle_DGV _
        (ByRef DGV_A_Traiter As DataGridView, ByVal Nom_Colonne_DGV As String,
         ByVal Type_Traitement As Integer, Optional ByVal Num_Ligne_DGV As Integer = -1,
         Optional ByVal Vidage_Colonnes As Boolean = Nothing)

        'Si numéro de ligne valué
        If Num_Ligne_DGV <> -1 Then

            'Si vidage colonnes autorisé
            If Vidage_Colonnes = True Then _
                DGV_A_Traiter.Rows.Item(Num_Ligne_DGV).Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                    (DGV_A_Traiter, Nom_Colonne_DGV)).Value = Nothing
        End If

        'Si traitement = 1, coloriage/blocage cellule
        If Type_Traitement = 1 Then

            'Si numéro de ligne valué
            If Num_Ligne_DGV <> -1 Then

                'Teinte du fond de la case
                DGV_A_Traiter.Rows.Item(Num_Ligne_DGV).Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                    (DGV_A_Traiter, Nom_Colonne_DGV)).Style.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvDesact)

                'Désactivation colonnes
                DGV_A_Traiter.Rows.Item(Num_Ligne_DGV).Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                    (DGV_A_Traiter, Nom_Colonne_DGV)).ReadOnly = True

                'Sinon
            Else

                'Teinte du fond de la case
                DGV_A_Traiter.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                    (DGV_A_Traiter, Nom_Colonne_DGV)).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(ParamsApp.DgvDesact)

                'Désactivation colonnes
                DGV_A_Traiter.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                    (DGV_A_Traiter, Nom_Colonne_DGV)).ReadOnly = True
            End If

            'Si traitement = 2, suppression coloriage/déblocage cellule
        ElseIf Type_Traitement = 2 Then

            'Si numéro de ligne valué
            If Num_Ligne_DGV <> -1 Then

                'Teinte du fond de la case
                DGV_A_Traiter.Rows.Item(Num_Ligne_DGV).Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                    (DGV_A_Traiter, Nom_Colonne_DGV)).Style.BackColor = Nothing

                'Désactivation colonnes
                DGV_A_Traiter.Rows.Item(Num_Ligne_DGV).Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                    (DGV_A_Traiter, Nom_Colonne_DGV)).ReadOnly = False

                'Sinon
            Else

                'Teinte du fond de la case
                DGV_A_Traiter.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                    (DGV_A_Traiter, Nom_Colonne_DGV)).DefaultCellStyle.BackColor = Nothing

                'Désactivation colonnes
                DGV_A_Traiter.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                    (DGV_A_Traiter, Nom_Colonne_DGV)).ReadOnly = False
            End If
        End If
    End Sub

    'Contrôle valeur cellule DGV (suivant type de colonne) et renvoi True ou False

    Public Function Test_Valeur_Suivant_Type_Colonne_DGV _
        (ByVal Colonne_DGV_A_Tester As DataGridViewColumn,
         ByVal Cellule_DGV_A_Tester As DataGridViewCell) As Boolean

        'Si colonne est de type TextBox
        If Colonne_DGV_A_Tester.GetType.ToString = "System.Windows.Forms.DataGridViewTextBoxColumn" Or
            Colonne_DGV_A_Tester.GetType.ToString = "System.Windows.Forms.DataGridViewComboBoxColumn" Then

            'Si valeur de cellule <> Nothing
            If Cellule_DGV_A_Tester.Value <> Nothing Then

                'Retourne True
                Return True
            End If

            'Si colonne est de type CheckBox
        ElseIf Colonne_DGV_A_Tester.GetType.ToString = "System.Windows.Forms.DataGridViewCheckBoxColumn" Then

            'Si valeur de cellule = True
            If Cellule_DGV_A_Tester.Value = True Then

                'Retourne True
                Return True
            End If
        End If

        'Retourne False
        Return False
    End Function

    'Remplissage colonne ComboBox DGV avec entêtes d'une autre DGV suivant cellules Checkées et ToolTips colonne

    Public Sub Remp_ComboBox_Avec_Entetes_DGV _
        (ByRef DGV_A_Lire As DataGridView, ByRef ComboBox_A_Remplir As Object,
         ByVal Nom_Colonne_Debut As String, ByVal Nom_Colonne_Fin As String,
         ByVal Vidage_ComboBox As Boolean, ByVal Ajout_Ligne_Vide As Boolean,
         ByVal Doublons_Autorises As Boolean)

        'Si vidage ComboBox autorisé
        If Vidage_ComboBox = True Then

            'Vidage ComboBox
            ComboBox_A_Remplir.Items.Clear()
        End If

        'Si ajout d'une ligne vide autorisé
        If Ajout_Ligne_Vide = True Then

            'Remplissage ComboBox
            ComboBox_A_Remplir.Items.Add("")
        End If

        'Boucle colonne
        For Boucle_Colonne = Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, Nom_Colonne_Debut) To _
            Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, Nom_Colonne_Fin) + 1

            'Si paramètre CATIA trouvé dans ToolTip
            If InStr(DGV_A_Lire.Columns.Item(Boucle_Colonne - 1).ToolTipText, "(11)") <> 0 Then

                'Si doublons autorisés dans la liste
                If Doublons_Autorises = True Then

                    'Remplissage ComboBox
                    ComboBox_A_Remplir.Items.Add(DGV_A_Lire.Columns.Item(Boucle_Colonne - 1).HeaderText)

                    'Sinon
                Else

                    'Si chaîne déjà présente dans liste
                    If ComboBox_A_Remplir.Items.Contains(DGV_A_Lire.Columns.Item(Boucle_Colonne - 1).HeaderText) = False Then

                        'Remplissage ComboBox
                        ComboBox_A_Remplir.Items.Add(DGV_A_Lire.Columns.Item(Boucle_Colonne - 1).HeaderText)
                    End If
                End If
            End If
        Next
    End Sub

    'Ecriture entête DGV dans Tab (a tester)

    Public Sub EcrtEnttDGVTab(ByVal DGV_A_Lire As DataGridView, ByRef tabTemp(,) As String)

        'Boucle lecture colonne
        For incColonne = 1 To DGV_A_Lire.Columns.Count

            'Si colonne différente d'un bouton ou si présence colonne programme dans DGV
            If DGV_A_Lire.Columns.Item(incColonne - 1).GetType.ToString <> "System.Windows.Forms.DataGridViewButtonColumn" Or
                incColonne - 1 = Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, "Numéro programme") Then

                'Incrément compteur
                tabTemp(0, incColonne - 1) = DGV_A_Lire.Columns.Item(incColonne - 1).HeaderText
            End If
        Next
    End Sub
End Module
