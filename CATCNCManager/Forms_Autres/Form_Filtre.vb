Public Class Form_Filtre

    'Variables
    Private Tab_Filtre_DGV_Permanent(,), Tab_Filtre_DGV_Local(,), Tab_Noms_Colonnes_Filtre(1) As String
    Private Ouverture_Form_OK As Boolean = False
    Private DGV_A_Filtrer_Local As DataGridView

    '------------------------------------------------- Action sur événement ---------------------------------------------------

    'Action sur création nouvelle Form Filtre

    Public Sub New(ByRef DGV_A_Filtrer As DataGridView, Optional ByVal Tab_Filtre_DGV(,) As String = Nothing)

        'Initialisation composant
        InitializeComponent()

        'Si New Form créée sans les attributs (si créée juste pour récupérer des infos)
        If Not Tab_Filtre_DGV Is Nothing Then

            'Redimensionnement des Tab
            ReDim Tab_Filtre_DGV_Permanent(UBound(Tab_Filtre_DGV, 1), UBound(Tab_Filtre_DGV, 2))
            ReDim Tab_Filtre_DGV_Local(UBound(Tab_Filtre_DGV, 1), UBound(Tab_Filtre_DGV, 2))

            'Valuation variables
            Tab_Filtre_DGV_Permanent = Tab_Filtre_DGV
            Tab_Filtre_DGV_Local = Tab_Filtre_DGV
        End If

        'Valuation variable
        DGV_A_Filtrer_Local = DGV_A_Filtrer
        Tab_Noms_Colonnes_Filtre(0) = "Filtre suivant liste"
        Tab_Noms_Colonnes_Filtre(1) = "Filtre suivant chaîne libre"

        'Ouverture DGV filtre
        Fonctions_Filtres_DGV.Ouverture_DGV_Filtre _
            (DGV_A_Filtrer_Local, Me.DataGridView1, Tab_Filtre_DGV_Local, Tab_Noms_Colonnes_Filtre)

        'Valuation variable
        Ouverture_Form_OK = True
    End Sub

    'Gestion fenêtre si click DGV

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As DataGridViewCellEventArgs) _
        Handles DataGridView1.CellContentClick

        'Valeur de condition à trouver
        Select Case e.ColumnIndex

            'Si colonne activer filtre
            Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Activer")

                'De colonne non-checkée => checkée
                If Me.DataGridView1.CurrentRow.Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                                                          (Me.DataGridView1, "Activer")).Value = False Then

                    'Si Tab des filtres non-vide
                    If Not Tab_Filtre_DGV_Local Is Nothing Then

                        'Si dernière ligne Tab n'est pas un filtre permanent
                        If Tab_Filtre_DGV_Local(UBound(Tab_Filtre_DGV_Local, 1), 3) = "False" Then

                            'Blocage colonne précédente checkée
                            Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
                                (Me.DataGridView1, "Activer", Tab_Noms_Colonnes_Filtre, False, 1, False, _
                                 Fonctions_DGV.Renvoi_Num_Ligne_DGV _
                                 (Me.DataGridView1, "Colonne", Tab_Filtre_DGV_Local(UBound(Tab_Filtre_DGV_Local, 1), 0)))
                        End If
                    End If

                    'Remplissage ComboBox et déverrouillage filtre
                    Fonctions_Filtres_DGV.Gestion_ComboBox_DGV_Filtre _
                        (DGV_A_Filtrer_Local, Me.DataGridView1, True, Tab_Noms_Colonnes_Filtre)

                    'De colonne checkée => non-checkée
                Else

                    'Remplissage tab liste des filtres et vérrouillage filtre précédent (si existant)
                    Fonctions_Filtres_DGV.Gestion_Liste_Filtre _
                        (Me.DataGridView1, Tab_Filtre_DGV_Local, 2, , )

                    'Si Tab des filtres non-vide
                    If Not Tab_Filtre_DGV_Local Is Nothing Then

                        'Relecture de tous les filtres et dévérouillage dernier filtre dans DGV
                        Fonctions_Filtres_DGV.Relecture_Tous_Filtres_DGV _
                            (DGV_A_Filtrer_Local, Tab_Filtre_DGV_Local, Me.DataGridView1, 1)

                        'Sinon
                    Else

                        'Si Tab des filtres non-vide
                        If Not Tab_Filtre_DGV_Permanent Is Nothing Then

                            'Redimensionnement Tab
                            ReDim Tab_Filtre_DGV_Local(UBound(Tab_Filtre_DGV_Permanent, 1), UBound(Tab_Filtre_DGV_Permanent, 2))

                            'Valuation variable
                            Tab_Filtre_DGV_Local = Tab_Filtre_DGV_Permanent

                            'Relecture de tous les filtres et dévérouillage dernier filtre dans DGV
                            Fonctions_Filtres_DGV.Relecture_Tous_Filtres_DGV _
                                (DGV_A_Filtrer_Local, Tab_Filtre_DGV_Local, Me.DataGridView1, 1)

                            'Sinon
                        Else

                            'Démasquage de toutes les lignes DGV
                            Fonctions_Filtres_DGV.Demasquage_Ligne_DGV(DGV_A_Filtrer_Local)
                        End If
                    End If

                    'Vidage ComboBox et vérrouillage filtre
                    Fonctions_Filtres_DGV.Gestion_ComboBox_DGV_Filtre _
                        (DGV_A_Filtrer_Local, Me.DataGridView1, False, Tab_Noms_Colonnes_Filtre)

                    'Si Tab des filtres non-vide
                    If Not Tab_Filtre_DGV_Local Is Nothing Then

                        'Blocage colonne précédente checkée
                        Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
                            (Me.DataGridView1, "Activer", Tab_Noms_Colonnes_Filtre, False, 2, False, _
                             Fonctions_DGV.Renvoi_Num_Ligne_DGV _
                             (Me.DataGridView1, "Colonne", Tab_Filtre_DGV_Local(UBound(Tab_Filtre_DGV_Local, 1), 0)))
                    End If
                End If
        End Select
    End Sub

    'Gestion fenêtre en sortie d'édition

    Private Sub DataGridView1_CellEndEdit(ByVal sender As System.Object, ByVal e As DataGridViewCellEventArgs) _
        Handles DataGridView1.CellEndEdit

        'Si 1er ouverture Form OK
        If Ouverture_Form_OK = True Then

            'Si valeur filtre présente
            If Me.DataGridView1.CurrentRow.Cells.Item(e.ColumnIndex).Value <> Nothing Then

                'Valeur de condition à trouver
                Select Case e.ColumnIndex

                    'Si colonne Filtre suivant liste
                    Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, Tab_Noms_Colonnes_Filtre(0))



                        'Vidage cellule filtre suivant chaîne
                        Me.DataGridView1.CurrentRow.Cells.Item _
                            (Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                             (Me.DataGridView1, Tab_Noms_Colonnes_Filtre(1))).Value = Nothing


                        'Déclaration variable
                        'Dim Tab_Colonne_Temp(0) As String

                        'Valuation variable
                        'Tab_Colonne_Temp(0) = Tab_Noms_Colonnes_Filtre(1)


                        'Vidage et blocage des colonnes si autre colonne vide
                        'Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
                        '(Me.DataGridView1, Tab_Noms_Colonnes_Filtre(0), Tab_Colonne_Temp, True, 1, False, _
                        'Me.DataGridView1.CurrentRow.Index)




                        'Remplissage tab liste des filtres et vérrouillage filtre précédent (si existant)
                        Fonctions_Filtres_DGV.Gestion_Liste_Filtre _
                            (Me.DataGridView1, Tab_Filtre_DGV_Local, 1, 1, e.ColumnIndex)

                        'Relecture de tous les filtres et dévérouillage dernier filtre dans DGV
                        Fonctions_Filtres_DGV.Relecture_Tous_Filtres_DGV _
                            (DGV_A_Filtrer_Local, Tab_Filtre_DGV_Local, Me.DataGridView1, 2)

                        'Si colonne Filtre suivant chaîne libre
                    Case Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, Tab_Noms_Colonnes_Filtre(1))



                        'Vidage cellule filtre suivant chaîne
                        Me.DataGridView1.CurrentRow.Cells.Item _
                            (Fonctions_DGV.Renvoi_Num_Colonne_DGV _
                             (Me.DataGridView1, Tab_Noms_Colonnes_Filtre(0))).Value = Nothing

                        'Déclaration variable
                        'Dim Tab_Colonne_Temp(0) As String

                        'Valuation variable
                        'Tab_Colonne_Temp(0) = Tab_Noms_Colonnes_Filtre(0)


                        'Vidage et blocage des colonnes si autre colonne vide
                        'Fonctions_DGV.Controle_Et_Blocage_Coloriage_Colonnes_DGV _
                        '(Me.DataGridView1, Tab_Noms_Colonnes_Filtre(1), Tab_Colonne_Temp, True, 1, False, _
                        'Me.DataGridView1.CurrentRow.Index)




                        'Remplissage tab liste des filtres et vérrouillage filtre précédent (si existant)
                        Fonctions_Filtres_DGV.Gestion_Liste_Filtre _
                            (Me.DataGridView1, Tab_Filtre_DGV_Local, 1, 2, e.ColumnIndex)

                        'Relecture de tous les filtres et dévérouillage dernier filtre dans DGV
                        Fonctions_Filtres_DGV.Relecture_Tous_Filtres_DGV _
                            (DGV_A_Filtrer_Local, Tab_Filtre_DGV_Local, Me.DataGridView1, 2)
                End Select
            End If
        End If
    End Sub

    'Suppression de tous les filtres

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click

        'Si Tab des filtres non-vide
        If Not Tab_Filtre_DGV_Permanent Is Nothing Then

            'Redimensionnement Tab
            ReDim Tab_Filtre_DGV_Local(UBound(Tab_Filtre_DGV_Permanent, 1), UBound(Tab_Filtre_DGV_Permanent, 2))

            'Valuation variable
            Tab_Filtre_DGV_Local = Tab_Filtre_DGV_Permanent

            'Sinon
        Else

            'Vidage variable
            Tab_Filtre_DGV_Local = Nothing
        End If

        'Ouverture DGV filtre
        Fonctions_Filtres_DGV.Ouverture_DGV_Filtre _
            (DGV_A_Filtrer_Local, Me.DataGridView1, Tab_Filtre_DGV_Local, Tab_Noms_Colonnes_Filtre)
    End Sub
End Class