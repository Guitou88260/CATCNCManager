Public Module Fonctions_Tableau

    'Recherche chaine dans tableau et renvoi une valeur d'une autre colonne contenue sur la même ligne

    Public Function Renvoi_Chaine_Autre_Colonne_Tab _
        (ByVal Tab_A_Lire As String(,), ByVal Chaine_A_Trouver As String, _
         ByVal Colonne_Chaine_A_Trouver As String, ByVal Colonne_Chaine_A_Renvoyer As String, _
         ByVal Type_Traitement As Integer) As String

        'Déclaration variables
        Dim Num_Colonne_Chaine_A_Trouver As Integer = Nothing
        Dim Num_Colonne_Chaine_A_Renvoyer As Integer = Nothing

        'Valuation variable
        Renvoi_Chaine_Autre_Colonne_Tab = Nothing

        'Si type de traitement = 1, récupération numéro de colonne depuis le nom
        If Type_Traitement = 1 Then

            'Valuation variables
            Num_Colonne_Chaine_A_Trouver = Fonctions_Tableau.Renvoi_Num_Colonne_Tab _
                (Tab_A_Lire, Colonne_Chaine_A_Trouver)
            Num_Colonne_Chaine_A_Renvoyer = Fonctions_Tableau.Renvoi_Num_Colonne_Tab _
                (Tab_A_Lire, Colonne_Chaine_A_Renvoyer)

            'Si type de traitement = 2, valuation numéro de colonne
        ElseIf Type_Traitement = 2 Then

            'Valuation variables
            Num_Colonne_Chaine_A_Trouver = Colonne_Chaine_A_Trouver
            Num_Colonne_Chaine_A_Renvoyer = Colonne_Chaine_A_Renvoyer
        End If

        'Boucle recherche chaîne
        For Boucle = LBound(Tab_A_Lire, 1) To UBound(Tab_A_Lire, 1)

            'Si chaîne trouvée
            If Tab_A_Lire(Boucle, Num_Colonne_Chaine_A_Trouver) = Chaine_A_Trouver Then

                'Valuation variable et sortie de fonction
                Renvoi_Chaine_Autre_Colonne_Tab = Tab_A_Lire(Boucle, Num_Colonne_Chaine_A_Renvoyer)

                'Sortie de fonction
                Exit Function
            End If
        Next
    End Function

    'Recherche chaine dans tableau et renvoi Boolean

    Public Function Presence_Chaine_Dans_Tab _
        (ByVal Tab_A_Lire As Object(,), ByVal Chaine_A_Trouver As String, _
         ByVal Num_Colonne_Chaine_A_Trouver As Integer, _
         Optional ByVal Condition_A_Trouver As String = Nothing, _
         Optional ByVal Colonne_Condition_A_Trouver As Integer = Nothing) As Boolean

        'Boucle recherche dans Tab
        For Boucle = LBound(Tab_A_Lire, 1) To UBound(Tab_A_Lire, 1)

            'Si condition à trouver et colonne condition valués
            If Condition_A_Trouver <> Nothing And Colonne_Condition_A_Trouver <> Nothing Then

                'Si condition trouvée
                If Tab_A_Lire(Boucle, Colonne_Condition_A_Trouver) = Condition_A_Trouver And _
                    Tab_A_Lire(Boucle, Num_Colonne_Chaine_A_Trouver) = Chaine_A_Trouver Then

                    'Retourne résultat
                    Return True
                End If

                'Sinon
            Else

                'Si chaîne trouvée
                If Tab_A_Lire(Boucle, Num_Colonne_Chaine_A_Trouver) = Chaine_A_Trouver Then

                    'Retourne résultat
                    Return True
                End If
            End If
        Next

        'Retourne résultat
        Return False
    End Function

    'Recherche valeur la plus élevée dans colonne d'un tableau

    Public Function Renvoi_Valeur_Maxi_Colonne_Tab _
        (ByVal Tab_A_Lire As String(,), ByVal Nom_Colonne_A_Controler As String, _
         ByVal Traitement_Colonne_Prog As Boolean) As Integer

        'Valuation variable
        Renvoi_Valeur_Maxi_Colonne_Tab = 0

        'Boucle recherche (LBound(Tab_A_Lire, 1) + 1 pour saut ligne des entêtes)
        For Boucle = LBound(Tab_A_Lire, 1) + 1 To UBound(Tab_A_Lire, 1)

            'Si colonne numéro de programme
            If Traitement_Colonne_Prog = True Then

                'Extraction des 4 chiffres uniquement et si valeur dans la DGV supérieure à la valeur du numéro
                If Mid(Tab_A_Lire(Boucle, Fonctions_Tableau.Renvoi_Num_Colonne_Tab(Tab_A_Lire, Nom_Colonne_A_Controler)), 3) > Renvoi_Valeur_Maxi_Colonne_Tab Then _
                    Renvoi_Valeur_Maxi_Colonne_Tab = Mid(Tab_A_Lire(Boucle, Fonctions_Tableau.Renvoi_Num_Colonne_Tab(Tab_A_Lire, Nom_Colonne_A_Controler)), 3)

                'Sinon
            Else

                'Si valeur dans la DGV supérieure à la valeur du numéro
                If Tab_A_Lire(Boucle, Fonctions_Tableau.Renvoi_Num_Colonne_Tab(Tab_A_Lire, Nom_Colonne_A_Controler)) > Renvoi_Valeur_Maxi_Colonne_Tab Then _
                    Renvoi_Valeur_Maxi_Colonne_Tab = Tab_A_Lire(Boucle, Fonctions_Tableau.Renvoi_Num_Colonne_Tab(Tab_A_Lire, Nom_Colonne_A_Controler))
            End If
        Next
    End Function

    'Recherche chaine dans ligne 0 d'un tableau et renvoi le numéro de la colonne, et si non présente retourne -1

    Public Function Renvoi_Num_Colonne_Tab _
        (ByVal Tab_A_Lire As String(,), ByVal Chaine_A_Trouver As String) As Integer

        'Valuation variable
        Renvoi_Num_Colonne_Tab = Nothing

        'Boucle colonne
        For Boucle = LBound(Tab_A_Lire, 2) To UBound(Tab_A_Lire, 2)

            'Si chaîne trouvée
            If Tab_A_Lire(0, Boucle) = Chaine_A_Trouver Then

                'Valuation variable
                Renvoi_Num_Colonne_Tab = Boucle

                'Sortie de Fonction
                Exit Function
            End If
        Next

        'Valuation variable
        Renvoi_Num_Colonne_Tab = -1
    End Function

    'Ajout d'une ligne vierge à la fin d'un tableau à 2 dimensions

    Public Sub Ajout_Ou_Retrait_Ligne_Dans_Tab _
        (ByRef Tab_A_Traiter(,) As String, ByVal Taille_Deuxieme_Dimension_Tab As Integer, _
         ByVal Type_Traitement As Integer, Optional ByVal Num_Colonne_Chaine_A_Enlever As Integer = Nothing, _
         Optional ByVal Chaine_A_Enlever As String = Nothing, Optional ByVal Condition_A_Trouver As String = Nothing, _
         Optional ByVal Num_Colonne_Condition_A_Trouver As Integer = Nothing)

        'Si type de traitement = 1, ajout d'une ligne
        If Type_Traitement = 1 Then

            'Si aucune valeur dans tab
            If Tab_A_Traiter Is Nothing Then

                'Redimensionnement Tab
                ReDim Tab_A_Traiter(0, Taille_Deuxieme_Dimension_Tab)

                'Sinon
            Else

                'Déclaration variable
                Dim Tab_Temp(,) As String = Nothing

                'Redimensionnement Tab
                ReDim Tab_Temp(UBound(Tab_A_Traiter, 1) + 1, UBound(Tab_A_Traiter, 2))

                'Boucle lecture lignes
                For Boucle_Ligne = LBound(Tab_A_Traiter, 1) To UBound(Tab_A_Traiter, 1)

                    'Boucle lecture colonnes
                    For Boucle_Colonne = LBound(Tab_A_Traiter, 2) To UBound(Tab_A_Traiter, 2)

                        'Valuation variable
                        Tab_Temp(Boucle_Ligne, Boucle_Colonne) = Tab_A_Traiter(Boucle_Ligne, Boucle_Colonne)
                    Next
                Next

                'Redimensionnement Tab
                ReDim Tab_A_Traiter(UBound(Tab_Temp, 1), UBound(Tab_Temp, 2))

                'Valuation variable
                Tab_A_Traiter = Tab_Temp
            End If

            'Si type de traitement = 2, retrait d'une ligne
        ElseIf Type_Traitement = 2 Then

            'Si tab différent de Nothing
            If Not Tab_A_Traiter Is Nothing Then

                'Si une ligne dans tab
                If UBound(Tab_A_Traiter, 1) = 0 Then

                    'Vidage variable tab
                    Tab_A_Traiter = Nothing

                    'Sinon
                Else

                    'Déclaration variables
                    Dim Tab_Temp(,) As String = Nothing
                    Dim Compteur As Integer = 0

                    'Redimensionnement Tab
                    ReDim Tab_Temp(UBound(Tab_A_Traiter, 1) - 1, UBound(Tab_A_Traiter, 2))

                    'Boucle lecture lignes
                    For Boucle_Ligne = LBound(Tab_A_Traiter, 1) To UBound(Tab_A_Traiter, 1)

                        'Si condition à trouver et colonne condition valués
                        If Condition_A_Trouver <> Nothing And Num_Colonne_Condition_A_Trouver <> Nothing Then

                            'Si chaîne différente de celle recherchée
                            If Tab_A_Traiter(Boucle_Ligne, Num_Colonne_Chaine_A_Enlever) <> Chaine_A_Enlever And _
                                Tab_A_Traiter(Boucle_Ligne, Num_Colonne_Condition_A_Trouver) <> Condition_A_Trouver Or _
                                Tab_A_Traiter(Boucle_Ligne, Num_Colonne_Chaine_A_Enlever) <> Chaine_A_Enlever And _
                                Tab_A_Traiter(Boucle_Ligne, Num_Colonne_Condition_A_Trouver) = Condition_A_Trouver Or _
                                Tab_A_Traiter(Boucle_Ligne, Num_Colonne_Condition_A_Trouver) <> Condition_A_Trouver Then

                                'Boucle lecture colonnes
                                For Boucle_Colonne = LBound(Tab_A_Traiter, 2) To UBound(Tab_A_Traiter, 2)

                                    'Valuation variable
                                    Tab_Temp(Compteur, Boucle_Colonne) = Tab_A_Traiter(Boucle_Ligne, Boucle_Colonne)
                                Next

                                'Incrément compteur
                                Compteur = Compteur + 1
                            End If

                            'Sinon
                        Else

                            'Boucle lecture colonnes
                            For Boucle_Colonne = LBound(Tab_A_Traiter, 2) To UBound(Tab_A_Traiter, 2)

                                'Valuation variable
                                Tab_Temp(Compteur, Boucle_Colonne) = Tab_A_Traiter(Boucle_Ligne, Boucle_Colonne)
                            Next

                            'Incrément compteur
                            Compteur = Compteur + 1
                        End If
                    Next

                    'Valuation variable
                    Tab_A_Traiter = Tab_Temp
                End If
            End If
        End If
    End Sub

    'Récupère toutes les valeurs d'une colonne d'un Tab

    Public Sub Envoi_Colonne_Tab_Vers_Tab _
        (ByVal Tab_A_Lire As String(,), ByVal Num_Colonne_Tab_A_Lire As Integer, _
         ByRef Tab_A_Valuer As String(,), ByVal Num_Colonne_Tab_A_Valuer As Integer, _
         ByVal Lecture_Premiere_Ligne As Boolean)

        'Déclaration variable
        Dim Compteur As Integer = 0

        'Boucle lecture ligne
        For Boucle_Ligne = LBound(Tab_A_Lire, 1) To UBound(Tab_A_Lire, 1)

            'Si lecture première ligne = False
            If Lecture_Premiere_Ligne = False And Boucle_Ligne = LBound(Tab_A_Lire, 1) Then

                'Ne rien faire

                'Sinon
            Else

                'Si différent de Nothing
                If Tab_A_Lire(Boucle_Ligne, Num_Colonne_Tab_A_Lire) <> Nothing Then

                    'Valuation Tab
                    Tab_A_Valuer(Compteur, Num_Colonne_Tab_A_Valuer) = Tab_A_Lire(Boucle_Ligne, Num_Colonne_Tab_A_Lire)
                End If

                'Incrément compteur
                Compteur = Compteur + 1
            End If
        Next
    End Sub
End Module
