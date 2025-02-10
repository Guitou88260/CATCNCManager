Public Module Fonctions_Filtres_DGV

    'Démasquage de toutes les lignes DGV

    Public Sub Demasquage_Ligne_DGV(ByRef DGV_A_Filtrer As DataGridView)

        'Boucle lecture ligne
        For Boucle_Ligne = 1 To DGV_A_Filtrer.Rows.Count - 1

            'Démasquage ligne
            DGV_A_Filtrer.Rows.Item(Boucle_Ligne - 1).Visible = True
        Next
    End Sub

    'Masquage des lignes DGV suivant filtre ou suivant mode (visualisation ou modification)

    Public Sub Masquage_Ligne_DGV _
        (ByRef DGV_A_Filtrer As DataGridView, ByVal Nom_Colonne_A_Trier As String, _
         ByVal Chaine_Filtre As String, ByVal Type_Traitement As Integer)

        'Si type de traitement = 1, chaîne exacte recherchée
        If Type_Traitement = 1 Then

            'Boucle lecture ligne
            For Boucle_Ligne = 1 To DGV_A_Filtrer.Rows.Count - 1

                'Si chaine présente
                If DGV_A_Filtrer.Rows.Item(Boucle_Ligne - 1).Cells.Item _
                    (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Filtrer, Nom_Colonne_A_Trier)).Value <> Chaine_Filtre Then

                    'Masquage ligne
                    DGV_A_Filtrer.Rows.Item(Boucle_Ligne - 1).Visible = False
                End If
            Next

            'Si type de traitement = 2, chaîne présente recherchée
        ElseIf Type_Traitement = 2 Then

            'Boucle lecture ligne
            For Boucle_Ligne = 1 To DGV_A_Filtrer.Rows.Count - 1

                'Si chaine présente
                If InStr(LCase(DGV_A_Filtrer.Rows.Item(Boucle_Ligne - 1).Cells.Item _
                    (Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                     (DGV_A_Filtrer, Nom_Colonne_A_Trier)).Value), LCase(Chaine_Filtre)) = 0 Then

                    'Masquage ligne
                    DGV_A_Filtrer.Rows.Item(Boucle_Ligne - 1).Visible = False
                End If
            Next
        End If
    End Sub

    'Remplissage des nom de colonne, redimensionement et ouverture Form filtre

    Public Sub Ouverture_DGV_Filtre _
        (ByVal DGV_A_Filtrer As DataGridView, ByRef DGV_Filtre As DataGridView, _
         ByVal Tab_Filtre_DGV(,) As String, ByVal Tab_Noms_Colonnes_Filtre() As String)

        'Filtrage DGV
        Fonctions_Filtres_DGV.Relecture_Tous_Filtres_DGV _
            (DGV_A_Filtrer, Tab_Filtre_DGV, DGV_Filtre, 2)

        'Vidage DGV filtre
        DGV_Filtre.Rows.Clear()

        'Variable
        Dim Compteur As Integer = 0

        'Boucle lecture colonne
        For Boucle_Colonne = 1 To DGV_A_Filtrer.Columns.Count

            'Si colonne différente d'un bouton ou si présence colonne programme dans DGV
            If DGV_A_Filtrer.Columns.Item(Boucle_Colonne - 1).GetType.ToString <> "System.Windows.Forms.DataGridViewButtonColumn" Or _
                Boucle_Colonne - 1 = Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Filtrer, "Numéro programme") And _
                DGV_A_Filtrer.Columns.Item(Boucle_Colonne - 1).Visible = True Then

                'Ajout ligne DGV
                DGV_Filtre.Rows.Add()

                'Remplissage colonne CheckBox DGV
                DGV_Filtre.Rows.Item(Compteur).Cells.Item _
                    (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_Filtre, "Activer")).Value = False

                'Remplissage colonne DGV avec entête colonne DGV mère
                DGV_Filtre.Rows.Item(Compteur).Cells.Item _
                    (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_Filtre, "Colonne")).Value = _
                    DGV_A_Filtrer.Columns.Item(Boucle_Colonne - 1).HeaderText

                'Incrément compteur
                Compteur = Compteur + 1
            End If
        Next

        'Blocage des colonnes pour correspondre au Check par défaut
        Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
            (DGV_Filtre, "Activer", Tab_Noms_Colonnes_Filtre, False, 2, False, )

        'Ajustement des largeurs de colonne
        DGV_Filtre.AutoResizeColumns()
    End Sub

    'Remplissage colonne ComboBox DGV filtre avec données DGV à filtrer

    Public Sub Gestion_ComboBox_DGV_Filtre _
        (ByVal DGV_A_Filtrer As DataGridView, ByRef DGV_Filtre As DataGridView, _
         ByVal Remplissage_ComboBox_DGV As Boolean, ByVal Tab_Noms_Colonnes_Filtre() As String)

        'Gestion CheckBox et colonnes concernées
        Fonctions_DGV.Gestion_Dynamique_CheckBox _
            (DGV_Filtre, "Activer", Tab_Noms_Colonnes_Filtre, True, 2, False, DGV_Filtre.CurrentRow.Index)

        'Variable
        Dim ComboBox_DGV_A_Traiter As DataGridViewComboBoxCell

        'Valuation variable
        ComboBox_DGV_A_Traiter = DGV_Filtre.CurrentRow.Cells.Item _
            (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_Filtre, "Filtre suivant liste"))

        'Si remplissage ComboBox = True
        If Remplissage_ComboBox_DGV = True Then

            'Remplissage ComboBox DGV avec liste données colonne DGV mère
            Fonctions_DGV.Remp_ComboBox_Avec_Colonne_DGV _
                (DGV_A_Filtrer, DGV_Filtre.CurrentRow.Cells.Item _
                 (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_Filtre, "Colonne")).Value, ComboBox_DGV_A_Traiter)

            'Sinon
        Else

            'Vidage ComboBox
            ComboBox_DGV_A_Traiter.Items.Clear()
        End If
    End Sub

    'Ajout ou retrait d'un numéro de filtre dans liste des filtre

    Public Sub Gestion_Liste_Filtre _
        (ByRef DGV_Filtre As DataGridView, ByRef Tab_Filtre_DGV(,) As String, _
         ByVal Type_Traitement As Integer, Optional ByVal Type_Recherche As Integer = Nothing, _
         Optional ByVal Index_Colonne As Integer = Nothing)

        'Valeur de condition à trouver
        Select Case Type_Traitement

            'Si type de traitement = 1, ajout d'un filtre
            Case 1

                'Si valeur filtre différente de Nothing
                If DGV_Filtre.CurrentRow.Cells.Item(Index_Colonne).Value <> Nothing Then

                    'Si Tab des filtres vide
                    If Tab_Filtre_DGV Is Nothing Then

                        'Ajout d'une ligne vide dans variable
                        Fonctions_Tableau.Ajout_Ou_Retrait_Ligne_Dans_Tab(Tab_Filtre_DGV, 3, 1, , , , )

                        'Valuation variable
                        Tab_Filtre_DGV(UBound(Tab_Filtre_DGV, 1), 0) = DGV_Filtre.CurrentRow.Cells.Item _
                            (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_Filtre, "Colonne")).Value
                        Tab_Filtre_DGV(UBound(Tab_Filtre_DGV, 1), 1) = Type_Recherche
                        Tab_Filtre_DGV(UBound(Tab_Filtre_DGV, 1), 2) = _
                            DGV_Filtre.CurrentRow.Cells.Item(Index_Colonne).Value
                        Tab_Filtre_DGV(UBound(Tab_Filtre_DGV, 1), 3) = False

                        'Sinon
                    Else

                        'Si nom de colonne DGV à filtrer non-enregistré dans Tab
                        If Fonctions_Tableau.Presence_Chaine_Dans_Tab _
                            (Tab_Filtre_DGV, DGV_Filtre.CurrentRow.Cells.Item _
                             (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_Filtre, "Colonne")).Value, 0, , ) = False Or _
                         Fonctions_Tableau.Presence_Chaine_Dans_Tab _
                                (Tab_Filtre_DGV, DGV_Filtre.CurrentRow.Cells.Item _
                                 (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_Filtre, "Colonne")).Value, 0, True, 3) = True And _
                             Fonctions_Tableau.Presence_Chaine_Dans_Tab _
                             (Tab_Filtre_DGV, DGV_Filtre.CurrentRow.Cells.Item _
                              (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_Filtre, "Colonne")).Value, 0, False, 3) = False Then

                            'Ajout d'une ligne vide dans variable
                            Fonctions_Tableau.Ajout_Ou_Retrait_Ligne_Dans_Tab(Tab_Filtre_DGV, 3, 1, , , , )

                            'Valuation variable
                            Tab_Filtre_DGV(UBound(Tab_Filtre_DGV, 1), 0) = DGV_Filtre.CurrentRow.Cells.Item _
                                (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_Filtre, "Colonne")).Value
                            Tab_Filtre_DGV(UBound(Tab_Filtre_DGV, 1), 1) = Type_Recherche
                            Tab_Filtre_DGV(UBound(Tab_Filtre_DGV, 1), 2) = _
                                DGV_Filtre.CurrentRow.Cells.Item(Index_Colonne).Value
                            Tab_Filtre_DGV(UBound(Tab_Filtre_DGV, 1), 3) = False

                            'Sinon
                        Else

                            'Valuation variable
                            Tab_Filtre_DGV(UBound(Tab_Filtre_DGV, 1), 1) = Type_Recherche
                            Tab_Filtre_DGV(UBound(Tab_Filtre_DGV, 1), 2) = _
                                DGV_Filtre.CurrentRow.Cells.Item(Index_Colonne).Value
                        End If
                    End If
                    'End If
                End If

                'Si type de traitement = 2, retrait d'un filtre
            Case 2

                'Si nom de colonne DGV à filtrer enregistré dans Tab mais en filtre non-permanent
                If Fonctions_Tableau.Presence_Chaine_Dans_Tab _
                    (Tab_Filtre_DGV, DGV_Filtre.CurrentRow.Cells.Item _
                     (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_Filtre, "Colonne")).Value, 0, False, 3) = True Then

                    'Retrait d'une ligne dans variable
                    Fonctions_Tableau.Ajout_Ou_Retrait_Ligne_Dans_Tab _
                        (Tab_Filtre_DGV, 3, 2, 0, DGV_Filtre.CurrentRow.Cells.Item _
                         (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_Filtre, "Colonne")).Value, False, 3)
                End If
        End Select
    End Sub

    'Relecture de tous les filtres dans DGV et dévérouillage dernier filtre

    Public Sub Relecture_Tous_Filtres_DGV _
        (ByRef DGV_A_Filtrer As DataGridView, Optional ByRef Tab_Filtre_DGV(,) As String = Nothing, _
         Optional ByRef DGV_Filtre As DataGridView = Nothing, Optional ByVal Type_Traitement As Integer = Nothing)

        'Démasquage de toutes les lignes DGV
        Fonctions_Filtres_DGV.Demasquage_Ligne_DGV(DGV_A_Filtrer)

        'Si Tab différent de Nothing
        If Not Tab_Filtre_DGV Is Nothing Then

            'Boucle relecture tab
            For Boucle = LBound(Tab_Filtre_DGV, 1) To UBound(Tab_Filtre_DGV, 1)

                'Filtrage DGV
                Fonctions_Filtres_DGV.Masquage_Ligne_DGV _
                    (DGV_A_Filtrer, Tab_Filtre_DGV(Boucle, 0), Tab_Filtre_DGV(Boucle, 2), _
                     Tab_Filtre_DGV(Boucle, 1))

                'Si DGV filtre et type de traitement valués, déblocage dernière ligne de la DGV filtre
                If Not DGV_Filtre Is Nothing And Type_Traitement <> Nothing Then

                    'Déclaration variable
                    Dim Tab_Nom_Colonne(1) As String

                    'Valuation variable
                    Tab_Nom_Colonne(0) = "Filtre suivant liste"
                    Tab_Nom_Colonne(1) = "Filtre suivant chaîne libre"

                    'Si boucle = dernière ligne du tab, déverrouillage ligne
                    If Boucle = UBound(Tab_Filtre_DGV) Then Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
                        (DGV_Filtre, "Activer", Tab_Nom_Colonne, False, Type_Traitement, False, _
                            Fonctions_DGV.Renvoi_Num_Ligne_DGV _
                            (DGV_Filtre, "Colonne", Tab_Filtre_DGV(UBound(Tab_Filtre_DGV, 1), 0)))
                End If
            Next
        End If
    End Sub

    'Ajoute filtre par défaut dans Tab et relecture de tout le filtre

    Public Sub Ajout_Filtre_Et_Refiltrage_DGV _
        (ByRef Tab_Filtre_DGV(,) As String, ByVal Nom_Colonne_Filtree As String, ByVal Type_Filtrage As String, _
         ByVal Valeur_Filtre As String, ByVal Suppression_Interdite As Boolean, _
         Optional ByRef DGV_A_Filtrer As DataGridView = Nothing)

        'Ajout d'une ligne vide dans variable
        Fonctions_Tableau.Ajout_Ou_Retrait_Ligne_Dans_Tab(Tab_Filtre_DGV, 3, 1, , , , )

        'Valuation variable
        Tab_Filtre_DGV(UBound(Tab_Filtre_DGV, 1), 0) = Nom_Colonne_Filtree
        Tab_Filtre_DGV(UBound(Tab_Filtre_DGV, 1), 1) = Type_Filtrage
        Tab_Filtre_DGV(UBound(Tab_Filtre_DGV, 1), 2) = Valeur_Filtre
        Tab_Filtre_DGV(UBound(Tab_Filtre_DGV, 1), 3) = Suppression_Interdite

        'Si DGV à filtrer valuée
        If Not DGV_A_Filtrer Is Nothing Then

            'Filtrage DGV
            Fonctions_Filtres_DGV.Relecture_Tous_Filtres_DGV _
                (DGV_A_Filtrer, Tab_Filtre_DGV, , )
        End If
    End Sub
End Module
