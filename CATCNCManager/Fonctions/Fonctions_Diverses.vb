Public Module Fonctions_Diverses

    'Ouverture de fichier PDF

    Public Function Ouverture_Fichier_PDF _
        (ByVal Lien_PDF As String) As Boolean

        'Gestion des exceptions
        Try

            'Si lien du ficher PDF existant
            If System.IO.File.Exists(Lien_PDF) = True Then

                'Ouverture fichier PDF avec lien dans Excel
                System.Diagnostics.Process.Start(Lien_PDF)

                'Sinon
            Else

                'Message
                Fonctions_Messages.Appel_Msg(5, 5, , )

                'Retourne False
                Return False
            End If

            'Retourne True
            Return True

            'Pour toutes les erreurs
        Catch

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(5, 5, , )
            Return False
        End Try
    End Function

    'Ouverture fichier Excel

    Public Function Ouverture_Fichier_Excel _
        (ByVal Fichier_A_Lire As String, ByRef EXCEL As Object,
         ByRef Fichier_Courant As Workbook, ByRef Feuille_Courante As Worksheet) As Boolean

        'Gestion des exceptions
        Try

            'Excel application
            EXCEL = CreateObject("Excel.application")

            'Masquage Excel
            EXCEL.Visible = False

            'Ouverture fichier
            Fichier_Courant = EXCEL.Workbooks.Open(Fichier_A_Lire)
            Feuille_Courante = Fichier_Courant.ActiveSheet

            'Retourne True
            Return True

            'Pour toutes les erreurs
        Catch

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(5, 5, , )
            Return False
        End Try
    End Function

    'Enregistrement et fermeture Excel

    Public Function Enregistrement_Fermeture_Excel _
        (ByVal EXCEL As Object, ByVal Fichier_Courant As Workbook) As Boolean

        'Gestion des exceptions
        Try

            'Enregistrement fichier Excel
            Fichier_Courant.Save()

            'Fermeture Excel
            EXCEL.DisplayAlerts = False
            Fichier_Courant.Close()
            EXCEL.Quit()

            'Vidage variable
            EXCEL = Nothing

            'Retourne True
            Return True

            'Pour toutes les erreurs
        Catch

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(5, 5, , )
            Return False
        End Try
    End Function

    'Renvoi d'une chaine encadrée par 2 autres

    Public Function Renvoi_Chaine_Encadree _
        (ByVal Chaine_A_Lire As String, ByVal Chaine_Debut As String,
         ByVal Chaine_Fin As String) As String

        'Variables
        Dim Position_Debut As Integer
        Dim Position_Fin As Integer
        Renvoi_Chaine_Encadree = Nothing

        'Valuation variable
        Position_Debut = InStr(Chaine_A_Lire, Chaine_Debut)
        Position_Fin = InStr(Position_Debut + Len(Chaine_Debut), Chaine_A_Lire, Chaine_Fin)

        'Si position de début et de fin trouvées
        If Position_Debut <> Nothing And Position_Fin <> Nothing Then

            'Extraction chaine
            Renvoi_Chaine_Encadree = Mid(Chaine_A_Lire, Position_Debut + Len(Chaine_Debut),
                                         Position_Fin - (Position_Debut + Len(Chaine_Debut)))
        End If
    End Function

    'Renvoi chaîne moins une autre

    Public Function Renvoi_Chaine_Moins_Une_Autre _
        (ByVal Chaine_A_Lire As String, ByVal Chaine_A_Enlever As String) As String

        'Variable
        Dim Position_Debut As Integer
        Renvoi_Chaine_Moins_Une_Autre = Nothing

        'Valuation variable
        Position_Debut = InStr(Chaine_A_Lire, Chaine_A_Enlever)

        'Si position de début trouvée
        If Position_Debut <> Nothing Then

            'Extraction chaines
            Renvoi_Chaine_Moins_Une_Autre = Mid(Chaine_A_Lire, 1, Position_Debut - 1) & Mid(Chaine_A_Lire, Position_Debut + Len(Chaine_A_Enlever))
        End If
    End Function

    'Génération nom de fichier Maintenant

    Public Function Renvoi_Maintenant_Et_Auteur _
        (ByVal Type_Traitement As Integer) As String

        'Valuation variable
        Renvoi_Maintenant_Et_Auteur = Nothing

        'Valeur de condition à trouver
        Select Case Type_Traitement

            'Si type de traitement = 1 ou 3, date et heure seulement ou date, heure et auteur
            Case 1, 3

                'Valuation variable
                Renvoi_Maintenant_Et_Auteur = "Le " & Now.Date & " à " & Renvoi_Entier_2_Chiffres _
                    (Now.Hour.ToString) & ":" & Renvoi_Entier_2_Chiffres(Now.Minute.ToString)

                'Si 3
                If Type_Traitement = 3 Then

                    'Valuation variable
                    Renvoi_Maintenant_Et_Auteur = Renvoi_Maintenant_Et_Auteur & ", par " &
                        UCase(Mid(Environment.UserName, 1, 1)) & Mid(Environment.UserName, 2)
                End If

                'Si type de traitement = 2, auteur seulement
            Case 2

                'Valuation variable
                Renvoi_Maintenant_Et_Auteur = "Par " &
                    UCase(Mid(Environment.UserName, 1, 1)) & Mid(Environment.UserName, 2)
        End Select
    End Function

    'Renvoi chaîne à 2 caractères (ajout d'un 0 si seulement 1 caratère)

    Public Function Renvoi_Entier_2_Chiffres(ByVal Entier_A_Traiter As String) As String

        'Valuation variable
        Renvoi_Entier_2_Chiffres = Nothing

        'Si valeur = 1 caractère
        If Len(Entier_A_Traiter) = 1 Then

            'Valuation variable
            Renvoi_Entier_2_Chiffres = "0" & Entier_A_Traiter

            'Sinon
        Else

            'Valuation variable
            Renvoi_Entier_2_Chiffres = Entier_A_Traiter
        End If
    End Function

    'Vérification arborescence et création dossiers si besoin et en option, création fichier

    Public Function Ctrl_Arbor_Ou_Creation_Fichier _
        (ByVal Dossier_A_Controler As String, ByVal Remplacement_Fichier_Possible As Boolean,
         Optional ByVal Fichier_A_Creer As String = Nothing, Optional ByVal Chemin_Fichier_Modele As String = Nothing,
         Optional ByVal DGV_A_Lire As DataGridView = Nothing) As Boolean

        'Variables
        Dim Chemin_Temp As String = Nothing
        Dim Result_MsgBox As MsgBoxResult

        'Gestion des exceptions
        Try

            'Variables
            Dim Fichier_Syst As Object
            Dim Split_Arborenscence() As String

            'Création de l'object
            Fichier_Syst = CreateObject("Scripting.FileSystemObject")

            'Split arborescence
            Split_Arborenscence = Split(Dossier_A_Controler, "\", )

            'Boucle pour contrôle présence de tous les dossiers
            For Boucle = LBound(Split_Arborenscence) To UBound(Split_Arborenscence)

                'Si premier passage dans boucle (donc C:)...
                If Boucle = 0 Then

                    'Valuation variable
                    Chemin_Temp = Split_Arborenscence(Boucle)

                    'Sinon
                Else

                    'Valuation variable
                    Chemin_Temp = Chemin_Temp & "\" & Split_Arborenscence(Boucle)

                    'Si dossier inexistant
                    If Fichier_Syst.FolderExists(Chemin_Temp) = False Then

                        'MsgBox comfirmation de création de dossier
                        Result_MsgBox = Fonctions_Messages.Appel_Msg(33, 2, , Chemin_Temp)

                        'Si OK, suppression du fichier
                        If Result_MsgBox = MsgBoxResult.Yes Then

                            'Création du dossier
                            Fichier_Syst.CreateFolder(Chemin_Temp)

                            'Si pas d'erreur, Msg de confirmation de création
                            Fonctions_Messages.Appel_Msg(34, 1, , Chemin_Temp)

                            'Sinon
                        Else

                            'Sinon, Msg et retourne False pour quitter Sub
                            Fonctions_Messages.Appel_Msg(36, 1, , )
                            Return False
                        End If
                    End If
                End If
            Next

            'Pour toutes les erreurs
        Catch

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(25, 5, , Chemin_Temp)
            Return False
        End Try

        'Gestion des exceptions
        Try

            'Si fichier à créer...
            If Fichier_A_Creer <> Nothing Then

                'Valuation variable
                Chemin_Temp = Chemin_Temp & "\" & Fichier_A_Creer

                'Si fichier présent
                If System.IO.File.Exists(Chemin_Temp) = True Then

                    'Si remplacement de fichier possible
                    If Remplacement_Fichier_Possible = True Then

                        'MsgBox comfirmation remplacement fichier
                        Result_MsgBox = Fonctions_Messages.Appel_Msg(21, 2, , Fichier_A_Creer)

                        'Si OK, suppression du fichier
                        If Result_MsgBox = MsgBoxResult.Yes Then

                            'Suppression ancien fichier
                            System.IO.File.Delete(Chemin_Temp)

                            'Sinon
                        Else

                            'Retourne False pour quitter Sub
                            Return False
                        End If

                        'Sinon
                    Else

                        'Message et retourne True
                        Fonctions_Messages.Appel_Msg(48, 1, , Fichier_A_Creer)
                        Return True
                    End If
                End If

                'MsgBox confirmation création fichier
                Result_MsgBox = Fonctions_Messages.Appel_Msg(18, 2, , Fichier_A_Creer)

                'Si OK, copie ou création du fichier
                If Result_MsgBox = MsgBoxResult.Yes Then

                    'Si fichier modèle
                    If Chemin_Fichier_Modele <> Nothing Then

                        'Copie du fichier
                        System.IO.File.Copy(Chemin_Fichier_Modele, Chemin_Temp)

                        'Sinon
                    Else

                        'Création du fichier
                        System.IO.File.Create(Chemin_Temp).Dispose()

                        'Si DGV valuée, écriture des entêtes dans CSV créé
                        If IsNothing(DGV_A_Lire) = False Then

                            'Ecriture des entêtes DGV dans CSV et si erreur, renvoi False
                            If Fonctions_DGV.Ecriture_Entete_DGV_Dans_CSV(DGV_A_Lire, Chemin_Temp) = False Then Return False
                        End If
                    End If

                    'Si pas d'erreur, Msg de confirmation de création
                    Fonctions_Messages.Appel_Msg(35, 1, , Fichier_A_Creer)

                    'Sinon
                Else

                    'Sinon, Msg et retourne False pour quitter Sub
                    Fonctions_Messages.Appel_Msg(36, 1, , )
                    Return False
                End If
            End If

            'Retourne True
            Return True

            'Si chemin de fichier inexistant
        Catch Excep As DirectoryNotFoundException

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(0, 5, , Excep.Message)
            Return False

            'Si chemin de fichier inexistant
        Catch Excep As FileNotFoundException

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

    'Contrôle de chaîne et renvoi résultat

    Public Function Ctrl_Chaine_Seul _
        (ByVal Chaine_A_Controler As String, ByVal Type_Traitement As Integer,
         ByVal Generation_Msg As Boolean, Optional ByVal LG_Chaine_Egale As Integer = Nothing,
         Optional ByVal LG_Chaine_Mini As Integer = Nothing, Optional ByVal LG_Chaine_Maxi As Integer = Nothing,
         Optional ByVal Chaine_Cherchee As String = Nothing) As Boolean

        'Valuation variable
        Ctrl_Chaine_Seul = True

        'Valeur de condition à trouver
        Select Case Type_Traitement

            'Traitement 1 = si nombre de caractère différent de
            Case 1

                'Si vrai
                If Len(Chaine_A_Controler) <> LG_Chaine_Egale Then

                    'Si message à renvoyer
                    If Generation_Msg = True Then

                        'Message
                        Fonctions_Messages.Appel_Msg(6, 3, , LG_Chaine_Egale)
                    End If

                    'Valuation variable
                    Ctrl_Chaine_Seul = False
                End If

                'Traitement 2 = si nombre de caractère égale à
            Case 2

                'Si vrai
                If Len(Chaine_A_Controler) = LG_Chaine_Egale Then

                    'Si message à renvoyer
                    If Generation_Msg = True Then

                        'Message
                        Fonctions_Messages.Appel_Msg(7, 3, , LG_Chaine_Egale)
                    End If

                    'Valuation variable
                    Ctrl_Chaine_Seul = False
                End If

                'Traitement 3 = si nombre de caractère non-compris entre...
            Case 3

                'Si vrai
                If Len(Chaine_A_Controler) < LG_Chaine_Mini Or Len(Chaine_A_Controler) > LG_Chaine_Maxi Then

                    'Si message à renvoyer
                    If Generation_Msg = True Then

                        'Variable message
                        Dim Msg(1) As String
                        Msg(0) = LG_Chaine_Mini
                        Msg(1) = LG_Chaine_Maxi

                        'Message
                        Fonctions_Messages.Appel_Msg(9, 3, Msg, )
                    End If

                    'Valuation variable
                    Ctrl_Chaine_Seul = False
                End If

                'Traitement 4 = si caractère ou chaîne différent de
            Case 4

                'Si vrai
                If Chaine_A_Controler <> Chaine_Cherchee Then

                    'Si message à renvoyer
                    If Generation_Msg = True Then

                        'Message
                        Fonctions_Messages.Appel_Msg(51, 3, , Chaine_Cherchee)
                    End If

                    'Valuation variable
                    Ctrl_Chaine_Seul = False
                End If

                'Traitement 5 = si caractère ou chaîne interdit trouver
            Case 5

                'Si vrai
                If InStr(Chaine_A_Controler, Chaine_Cherchee) <> 0 Then

                    'Si message à renvoyer
                    If Generation_Msg = True Then

                        'Message
                        Fonctions_Messages.Appel_Msg(8, 3, , Chaine_Cherchee)
                    End If

                    'Valuation variable
                    Ctrl_Chaine_Seul = False
                End If

                'Traitement 6 = si nombre de caractère non-compris entre chiffres et si caractère ou chaîne différent de
            Case 6

                'Si vrai
                If (Len(Chaine_A_Controler) < LG_Chaine_Mini Or Len(Chaine_A_Controler) > LG_Chaine_Maxi) And
                    Chaine_A_Controler <> Chaine_Cherchee Then

                    'Si message à renvoyer
                    If Generation_Msg = True Then

                        'Variable message
                        Dim Msg(2) As String
                        Msg(0) = LG_Chaine_Mini
                        Msg(1) = LG_Chaine_Maxi
                        Msg(2) = Chaine_Cherchee

                        'Message
                        Fonctions_Messages.Appel_Msg(32, 3, Msg, )
                    End If

                    'Valuation variable
                    Ctrl_Chaine_Seul = False
                End If

                'Traitement 7 = si caractère autre que ceux autorisés dans chaîne donnée
            Case 7

                'Boucle lecture chaîne
                For Boucle_Chaine = 1 To Len(Chaine_A_Controler)

                    'Si un caractère interdit est trouvé
                    If InStr(Chaine_Cherchee, Mid(Chaine_A_Controler, Boucle_Chaine, 1)) = 0 Then

                        'Si message à renvoyer
                        If Generation_Msg = True Then

                            'Variable message
                            Dim Msg(1) As String
                            Msg(0) = Mid(Chaine_A_Controler, Boucle_Chaine, 1)
                            Msg(1) = Chaine_Cherchee

                            'Message
                            Fonctions_Messages.Appel_Msg(49, 3, Msg, )
                        End If

                        'Valuation variable
                        Ctrl_Chaine_Seul = False
                    End If
                Next

                'Traitement 8 = si nombre de caractère non-égale à, et si caractère ou chaîne différent de
            Case 8

                'Si vrai
                If Len(Chaine_A_Controler) <> LG_Chaine_Egale And Chaine_A_Controler <> Chaine_Cherchee Then

                    'Si message à renvoyer
                    If Generation_Msg = True Then

                        'Variable message
                        Dim Msg(1) As String
                        Msg(0) = LG_Chaine_Egale
                        Msg(1) = Chaine_Cherchee

                        'Message
                        Fonctions_Messages.Appel_Msg(50, 3, Msg, )
                    End If

                    'Valuation variable
                    Ctrl_Chaine_Seul = False
                End If

                'Traitement 9 = si caractère ou chaîne non-trouvé
            Case 9

                'Si vrai
                If InStr(Chaine_A_Controler, Chaine_Cherchee) = 0 Then

                    'Si message à renvoyer
                    If Generation_Msg = True Then

                        'Message
                        Fonctions_Messages.Appel_Msg(10, 3, , Chaine_Cherchee)
                    End If

                    'Valuation variable
                    Ctrl_Chaine_Seul = False
                End If
        End Select
    End Function

    'Contrôle de chaîne suivant paramètres (contenu dans ToolTip) et renvoi résultat

    Public Function Ctrl_Chaine_Suivant_Param _
        (ByVal Chaine_A_Controler As String, ByVal Liste_Param As String,
         ByVal Generation_Msg As Boolean) As Boolean

        'Valuation variable
        Ctrl_Chaine_Suivant_Param = True

        'Si longueur de chaîne différente de, disponible dans ToolTip
        If InStr(Liste_Param, "(1)") <> 0 Then

            'Contrôle de la chaine et valuation variable
            Dim Result_Temp As Boolean = Fonctions_Diverses.Ctrl_Chaine_Seul _
                                         (Chaine_A_Controler, 1, Generation_Msg, Fonctions_Diverses.Renvoi_Chaine_Encadree _
                                          (Liste_Param, "(1) Longueur de chaîne différente de """, """ caractère(s) interdite"), , , )

            'Si résultat temp = True, valuation variable
            If Ctrl_Chaine_Suivant_Param = True Then Ctrl_Chaine_Suivant_Param = Result_Temp
        End If

        'Si longueur de chaîne égale à, disponible dans ToolTip
        If InStr(Liste_Param, "(2)") <> 0 Then

            'Contrôle de la chaine et valuation variable
            Dim Result_Temp As Boolean = Fonctions_Diverses.Ctrl_Chaine_Seul _
                                         (Chaine_A_Controler, 2, Generation_Msg, Fonctions_Diverses.Renvoi_Chaine_Encadree _
                                          (Liste_Param, "(2) Longueur de chaîne égale à """, """ caractère(s) interdite"), , , )

            'Si résultat temp = True, valuation variable
            If Ctrl_Chaine_Suivant_Param = True Then Ctrl_Chaine_Suivant_Param = Result_Temp
        End If

        'Si nombre de caractère non-compris entre, disponible dans ToolTip
        If InStr(Liste_Param, "(3)") <> 0 Then

            'Contrôle de la chaine et valuation variable
            Dim Result_Temp As Boolean = Fonctions_Diverses.Ctrl_Chaine_Seul _
                                         (Chaine_A_Controler, 3, Generation_Msg, , Fonctions_Diverses.Renvoi_Chaine_Encadree _
                                          (Liste_Param, "(3) Longueur de chaîne inférieure à """, """ et supérieure à """),
                                          Fonctions_Diverses.Renvoi_Chaine_Encadree _
                                          (Liste_Param, "(3) Longueur de chaîne inférieure à """ & Fonctions_Diverses.Renvoi_Chaine_Encadree _
                                           (Liste_Param, "(3) Longueur de chaîne inférieure à """, """ et supérieure à """) &
                                           """ et supérieure à """, """ caractère(s) interdite"), )

            'Si résultat temp = True, valuation variable
            If Ctrl_Chaine_Suivant_Param = True Then Ctrl_Chaine_Suivant_Param = Result_Temp
        End If

        'Si chaîne différente de, disponible dans ToolTip
        If InStr(Liste_Param, "(4)") <> 0 Then

            'Contrôle de la chaine et valuation variable
            Dim Result_Temp As Boolean = Fonctions_Diverses.Ctrl_Chaine_Seul _
                                         (Chaine_A_Controler, 4, Generation_Msg, , , , Fonctions_Diverses.Renvoi_Chaine_Encadree _
                                          (Liste_Param, "(4) Caractère ou chaîne différent de """, """ interdit"))

            'Si résultat temp = True, valuation variable
            If Ctrl_Chaine_Suivant_Param = True Then Ctrl_Chaine_Suivant_Param = Result_Temp
        End If

        'Si chaîne égale à, disponible dans ToolTip
        If InStr(Liste_Param, "(5)") <> 0 Then

            'Contrôle de la chaine et valuation variable
            Dim Result_Temp As Boolean = Fonctions_Diverses.Ctrl_Chaine_Seul _
                                         (Chaine_A_Controler, 5, Generation_Msg, , , , Fonctions_Diverses.Renvoi_Chaine_Encadree _
                                          (Liste_Param, "(5) Caractère ou chaîne """, """ interdit"))

            'Si résultat temp = True, valuation variable
            If Ctrl_Chaine_Suivant_Param = True Then Ctrl_Chaine_Suivant_Param = Result_Temp
        End If

        'Si longueur de chaîne non-comprise entre et si différente de, disponible dans ToolTip
        If InStr(Liste_Param, "(6)") <> 0 Then

            'Contrôle de la chaine et valuation variable
            Dim Result_Temp As Boolean = Fonctions_Diverses.Ctrl_Chaine_Seul _
                                         (Chaine_A_Controler, 6, Generation_Msg, , Fonctions_Diverses.Renvoi_Chaine_Encadree _
                                          (Liste_Param, "(6) Longueur de chaîne inférieure à """, """ et supérieure à """),
                                          Fonctions_Diverses.Renvoi_Chaine_Encadree _
                                          (Liste_Param, "(6) Longueur de chaîne inférieure à """ & Fonctions_Diverses.Renvoi_Chaine_Encadree _
                                           (Liste_Param, "(6) Longueur de chaîne inférieure à """, """ et supérieure à """) & """ et supérieure à """,
                                           """ caractère(s) ou caractère ou chaîne différent de """), Fonctions_Diverses.Renvoi_Chaine_Encadree _
                                       (Liste_Param, "(6) Longueur de chaîne inférieure à """ & Fonctions_Diverses.Renvoi_Chaine_Encadree _
                                        (Liste_Param, "(6) Longueur de chaîne inférieure à """, """ et supérieure à """) & """ et supérieure à """ &
                                        Fonctions_Diverses.Renvoi_Chaine_Encadree(Liste_Param, "(6) Longueur de chaîne inférieure à """ & Fonctions_Diverses.Renvoi_Chaine_Encadree _
                                                                                  (Liste_Param, "(6) Longueur de chaîne inférieure à """, """ et supérieure à """) &
                                                                                  """ et supérieure à """, """ caractère(s) ou caractère ou chaîne différent de """) &
                                                                              """ caractère(s) ou caractère ou chaîne différent de """, """ interdits"))

            'Si résultat temp = True, valuation variable
            If Ctrl_Chaine_Suivant_Param = True Then Ctrl_Chaine_Suivant_Param = Result_Temp
        End If

        'Si caractère autre que ceux autorisés dans chaîne donnée et si caractère autre que ceux autorisés dans chaîne donnée et unité présente
        If InStr(Liste_Param, "(7)") <> 0 Or (InStr(Liste_Param, "(7)") <> 0 And InStr(Liste_Param, "(10)") <> 0) Then

            'Si unité trouvée dans chaîne à contrôler
            If InStr(Liste_Param, "(10)") <> 0 And
                InStr(Chaine_A_Controler, Fonctions_Diverses.Renvoi_Chaine_Encadree _
                      (Liste_Param, "(10) Unités de la valeur : """, """")) <> 0 Then

                'Contrôle de la chaine et valuation variable
                Dim Result_Temp As Boolean = Fonctions_Diverses.Ctrl_Chaine_Seul _
                                             (Fonctions_Diverses.Renvoi_Chaine_Moins_Une_Autre _
                                              (Chaine_A_Controler, Fonctions_Diverses.Renvoi_Chaine_Encadree _
                                               (Liste_Param, "(10) Unités de la valeur : """, """")), 7, Generation_Msg, , , ,
                                              Fonctions_Diverses.Renvoi_Chaine_Encadree(Liste_Param, "(7) Caractère(s) autre que """, """ interdit(s)"))

                'Si résultat temp = True, valuation variable
                If Ctrl_Chaine_Suivant_Param = True Then Ctrl_Chaine_Suivant_Param = Result_Temp

                'Sinon
            Else

                'Contrôle de la chaine et valuation variable
                Dim Result_Temp As Boolean = Fonctions_Diverses.Ctrl_Chaine_Seul _
                                             (Chaine_A_Controler, 7, Generation_Msg, , , , Fonctions_Diverses.Renvoi_Chaine_Encadree _
                                              (Liste_Param, "(7) Caractère(s) autre que """, """ interdit(s)"))

                'Si résultat temp = True, valuation variable
                If Ctrl_Chaine_Suivant_Param = True Then Ctrl_Chaine_Suivant_Param = Result_Temp
            End If
        End If

        'Si nombre de caractère différent de et si chaîne différente de, disponible dans ToolTip
        If InStr(Liste_Param, "(8)") <> 0 Then

            'Contrôle de la chaine et valuation variable
            Dim Result_Temp As Boolean = Fonctions_Diverses.Ctrl_Chaine_Seul _
                                         (Chaine_A_Controler, 8, Generation_Msg, Fonctions_Diverses.Renvoi_Chaine_Encadree _
                                          (Liste_Param, "(8) Longueur de chaîne différente de """, """ ou caractère(s) différent(s) de """), , ,
                                          Fonctions_Diverses.Renvoi_Chaine_Encadree _
                                          (Liste_Param, "(8) Longueur de chaîne différente de """ & Fonctions_Diverses.Renvoi_Chaine_Encadree _
                                           (Liste_Param, "(8) Longueur de chaîne différente de """, """ ou caractère(s) différent(s) de """) &
                                           """ ou caractère(s) différent(s) de """, """ interdites"))

            'Si résultat temp = True, valuation variable
            If Ctrl_Chaine_Suivant_Param = True Then Ctrl_Chaine_Suivant_Param = Result_Temp
        End If

        'Si chaîne non-présente, disponible dans ToolTip
        If InStr(Liste_Param, "(9)") <> 0 Then

            'Contrôle de la chaine et valuation variable
            Dim Result_Temp As Boolean = Fonctions_Diverses.Ctrl_Chaine_Seul _
                                         (Chaine_A_Controler, 9, Generation_Msg, , , , Fonctions_Diverses.Renvoi_Chaine_Encadree _
                                          (Liste_Param, "(9) Caractère ou chaîne """, """ obligatoire"))

            'Si résultat temp = True, valuation variable
            If Ctrl_Chaine_Suivant_Param = True Then Ctrl_Chaine_Suivant_Param = Result_Temp
        End If
    End Function

    'Renvoi valeur sans unité

    Public Function Renvoi_Valeur_Sans_Unite _
        (ByVal Chaine_Courante As String, ByVal ToolTip_Associe As String) As String

        'Si unité trouvée dans ToolTip colonne
        If InStr(ToolTip_Associe, "(10)") <> 0 Then

            'Valuation variable
            Renvoi_Valeur_Sans_Unite = Val(Chaine_Courante)

            'Sinon
        Else

            'Valuation variable
            Renvoi_Valeur_Sans_Unite = Chaine_Courante
        End If
    End Function

    'Renvoi valeur avec unité

    Public Function Renvoi_Valeur_Avec_Unite _
        (ByVal Chaine_Courante As String, ByVal ToolTip_Associe As String) As String

        'Si unité trouvée dans ToolTip colonne
        If InStr(ToolTip_Associe, "(10)") <> 0 Then

            'Extraction valeur numérique arrondie avec unité
            Renvoi_Valeur_Avec_Unite = System.Math.Round(Val(Chaine_Courante), 3) &
                Fonctions_Diverses.Renvoi_Chaine_Encadree(ToolTip_Associe, "(10) Unités de la valeur : """, """")

            'Sinon
        Else

            'Valuation variable
            Renvoi_Valeur_Avec_Unite = Chaine_Courante
        End If
    End Function

    'Calcul une chaîne complète avec parenthèses et écrit le résultat avec unité dans Tab et si erreur, retourne False

    Public Function Calcule_Tab_Chaines_Completes_Vers_Tab _
        (ByRef Tab_A_Valuer(,) As String, ByVal Num_Colonne_Tab_A_Lire As Integer,
         ByVal Num_Colonne_Tab_A_Valuer As Integer,
         Optional ByVal Num_Colonne_Tab_Unite As Integer = Nothing) As Boolean

        'Gestion des exceptions
        Try

            'Boucle ligne Tab
            For Boucle_Ligne_Tab = LBound(Tab_A_Valuer, 1) To UBound(Tab_A_Valuer, 1)

                'Déclaration variables
                Dim Chaine_Depart As String
                Dim Chaine_Temp As String = Nothing

                'Valuation variable
                Chaine_Depart = Tab_A_Valuer(Boucle_Ligne_Tab, Num_Colonne_Tab_A_Lire)

                'Boucle recherche présence parenthèse ouverte
                While InStr(Chaine_Depart, "(") <> 0

                    'Si parenthèse fermée existante
                    If InStr(Chaine_Depart, ")") <> 0 Then

                        'Valuation variable
                        Chaine_Temp = Mid(Chaine_Depart, InStr(Chaine_Depart, "(") + 1, InStr(Chaine_Depart, ")") -
                                          (InStr(Chaine_Depart, "(") + 1))

                        'Déclaration variable
                        Dim Chaine_Calculee_Temp As String = Nothing

                        'Calcul de la chaîne et si erreur, retourne False
                        If Fonctions_Diverses.Calcul_Chaine_Simple _
                            (Chaine_Temp, Chaine_Calculee_Temp, ) = False Then Return False

                        'Valuation variable
                        Chaine_Depart = Replace(Chaine_Depart, "(" & Chaine_Temp & ")", Chaine_Calculee_Temp)

                        'Sinon
                    Else

                        'Message et retourne False
                        Fonctions_Messages.Appel_Msg(58, 5, , )
                        Return False
                    End If
                End While

                'Calcul de la chaîne et si erreur, retourne False
                If Fonctions_Diverses.Calcul_Chaine_Simple _
                    (Chaine_Depart, Chaine_Temp, Tab_A_Valuer(Boucle_Ligne_Tab, Num_Colonne_Tab_Unite)) =
                    False Then Return False

                'Valuation variable
                Tab_A_Valuer(Boucle_Ligne_Tab, Num_Colonne_Tab_A_Valuer) = Chaine_Temp
            Next

            'Retourne True
            Return True

            'Pour toutes les erreurs
        Catch

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(59, 5, , )
            Return False
        End Try
    End Function

    'Calcul simple chaîne avec ajout d'unité vers Tab et si erreur, retourne False

    Public Function Calcul_Chaine_Simple _
        (ByVal Chaine_A_Traiter As String, ByRef Chaine_Resultat As String,
         Optional ByVal Unite_A_Inserer As String = Nothing) As Boolean

        'Gestion des exceptions
        Try

            'Déclaration variables
            Dim Chaine_Depart As String
            Dim Chaine_Temp As String = Nothing
            Dim Chiffre_Calcule As Double
            Dim Premier_Passage As Boolean = True

            'Valuation variables
            Chaine_Depart = Chaine_A_Traiter

            'Boucle recherche chiffre
            While Len(Chaine_Depart) > 0

                'Si premier passage
                If Premier_Passage = True Then

                    'Boucle premier chiffre
                    While Mid(Chaine_Depart, 1, 1) <> "+" And Mid(Chaine_Depart, 1, 1) <> "-" _
                        And Mid(Chaine_Depart, 1, 1) <> "*" And Mid(Chaine_Depart, 1, 1) <> "/" And
                        Len(Chaine_Depart) > 0

                        'Valuation variables
                        Chaine_Temp = Chaine_Temp & Mid(Chaine_Depart, 1, 1)
                        Chaine_Depart = Mid(Chaine_Depart, 2)
                    End While

                    'Valuation variables
                    Chiffre_Calcule = Val(Chaine_Temp)
                    Chaine_Temp = Nothing
                    Premier_Passage = False

                    'Si deuxième passage
                ElseIf Premier_Passage = False Then

                    'Valeur de condition à trouver
                    Select Case Mid(Chaine_Depart, 1, 1)

                        'Si opérateur calcul = +
                        Case "+"

                            'Valuation variable
                            Chaine_Depart = Mid(Chaine_Depart, 2)

                            'Boucle recherche chiffre
                            While Mid(Chaine_Depart, 1, 1) <> "+" And Mid(Chaine_Depart, 1, 1) <> "-" _
                            And Mid(Chaine_Depart, 1, 1) <> "*" And Mid(Chaine_Depart, 1, 1) <> "/" And
                            Len(Chaine_Depart) > 0

                                'Valuation variables
                                Chaine_Temp = Chaine_Temp & Mid(Chaine_Depart, 1, 1)
                                Chaine_Depart = Mid(Chaine_Depart, 2)
                            End While

                            'Valuation variable
                            Chiffre_Calcule = Chiffre_Calcule + Val(Chaine_Temp)

                            'Si opérateur calcul = -
                        Case "-"

                            'Valuation variable
                            Chaine_Depart = Mid(Chaine_Depart, 2)

                            'Boucle recherche chiffre
                            While Mid(Chaine_Depart, 1, 1) <> "+" And Mid(Chaine_Depart, 1, 1) <> "-" _
                            And Mid(Chaine_Depart, 1, 1) <> "*" And Mid(Chaine_Depart, 1, 1) <> "/" And
                            Len(Chaine_Depart) > 0

                                'Valuation variable
                                Chaine_Temp = Chaine_Temp & Mid(Chaine_Depart, 1, 1)
                                Chaine_Depart = Mid(Chaine_Depart, 2)
                            End While

                            'Valuation variable
                            Chiffre_Calcule = Chiffre_Calcule - Val(Chaine_Temp)

                            'Si opérateur calcul = *
                        Case "*"

                            'Valuation variable
                            Chaine_Depart = Mid(Chaine_Depart, 2)

                            'Boucle recherche chiffre
                            While Mid(Chaine_Depart, 1, 1) <> "+" And Mid(Chaine_Depart, 1, 1) <> "-" _
                            And Mid(Chaine_Depart, 1, 1) <> "*" And Mid(Chaine_Depart, 1, 1) <> "/" And
                            Len(Chaine_Depart) > 0

                                'Valuation variable
                                Chaine_Temp = Chaine_Temp & Mid(Chaine_Depart, 1, 1)
                                Chaine_Depart = Mid(Chaine_Depart, 2)
                            End While

                            'Valuation variable
                            Chiffre_Calcule = Chiffre_Calcule * Val(Chaine_Temp)

                            'Si opérateur calcul = /
                        Case "/"

                            'Valuation variable
                            Chaine_Depart = Mid(Chaine_Depart, 2)

                            'Boucle recherche chiffre
                            While Mid(Chaine_Depart, 1, 1) <> "+" And Mid(Chaine_Depart, 1, 1) <> "-" _
                            And Mid(Chaine_Depart, 1, 1) <> "*" And Mid(Chaine_Depart, 1, 1) <> "/" And
                            Len(Chaine_Depart) > 0

                                'Valuation variable
                                Chaine_Temp = Chaine_Temp & Mid(Chaine_Depart, 1, 1)
                                Chaine_Depart = Mid(Chaine_Depart, 2)
                            End While

                            'Valuation variable
                            Chiffre_Calcule = Chiffre_Calcule / Val(Chaine_Temp)
                    End Select

                    'Vidage variable
                    Chaine_Temp = Nothing
                End If
            End While

            'Si unité valuée
            If Unite_A_Inserer <> Nothing Then

                'Valuation variable
                Chaine_Resultat = System.Math.Round(Chiffre_Calcule, 3) &
                    Unite_A_Inserer

                'Sinon
            Else

                'Valuation variable
                Chaine_Resultat = System.Math.Round(Chiffre_Calcule, 3)
            End If

            'Retourne True
            Return True

            'Pour toutes les erreurs
        Catch

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(59, 5, , )
            Return False
        End Try
    End Function

    'Renvoi chaine concaténée x fois

    Public Function Renvoi_Concatenation_Chaine _
        (ByVal Chaine_A_Concatener As String, ByVal Nb_Concatenation As Integer) As String

        'Valuation variable
        Renvoi_Concatenation_Chaine = Nothing

        'Boucle pour concaténation
        For Boucle = 1 To Nb_Concatenation

            'Si 1er passage dans boucle (pour mise en forme)
            If Boucle = 1 Then

                'Concaténation
                Renvoi_Concatenation_Chaine = Chaine_A_Concatener

                'Sinon
            Else

                'Concaténation
                Renvoi_Concatenation_Chaine = Renvoi_Concatenation_Chaine & "," & Chaine_A_Concatener
            End If
        Next
    End Function

    'Conversion code ASCII
    Function ConvASCII(ByVal valASCII() As Integer) As String

        'Var
        Dim nameTool As String = Nothing

        'Boucle lecture Tab code ASCII
        For Each incChar As Integer In valASCII

            'Si <> 0
            If incChar <> 0 Then

                'Ecriture nom
                nameTool = nameTool & Chr(incChar)
            End If
        Next

        'Retourne nom
        Return nameTool
    End Function


    'Conversion couleur en string -> entier
    Function ConvColorStrVersInt(ByVal colorStr As String) As Integer

        'Split string
        Dim tmp() As String
        tmp = Split(colorStr, ",", 3)
        'Conversion
        ConvColorStrVersInt = RGB(CInt(tmp(0)), CInt(tmp(1)), CInt(tmp(2)))
    End Function

    'Reformulation plage de colonne vers plage de cellule
    Function ReformColonnCell(ByVal plageColonn As String, ByVal numLignCrt As Integer) As String

        'Si LG string = 1
        If Len(plageColonn) = 1 Then

            'Assemblage
            ReformColonnCell = plageColonn & numLignCrt

            'Si LG string = 3 (2 pas pris en compte)
        Else

            'Split & assenblage
            Dim tmp() As String
            tmp = Split(plageColonn, ":", 2)
            ReformColonnCell = tmp(0) & CStr(numLignCrt) & ":" & tmp(1) & CStr(numLignCrt)
        End If
    End Function


    'Convertir nombre vers lettre
    Function ConvNumbToLetter(ByVal numb As Integer) As String
        'Renvoi letter
        If numb > 26 Then
            ConvNumbToLetter = Chr(Int((numb - 1) / 26) + 64) & Chr(((numb - 1) Mod 26) + 65)
        Else
            ConvNumbToLetter = Chr(numb + 64)
        End If
    End Function
End Module
