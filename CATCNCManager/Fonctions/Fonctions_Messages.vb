Public Module Fonctions_Messages

    'Boites de dialogue communes

    Public Function Appel_Msg(ByVal Num_Msg As Integer, ByVal Num_Box As Integer, _
                              Optional ByVal Tab_Valeurs_Retours() As String = Nothing, _
                              Optional ByVal Valeur_Retour_Seule As String = Nothing) As Object

        'Variable
        Dim Msg_A_Afficher As String = Nothing

        'Choix du message à afficher
        Select Case Num_Msg

            'Ecriture valeur de retour seule uniquement
            Case 0
                Msg_A_Afficher = Valeur_Retour_Seule

                'Si ComboBox de choix de machine non valuée
            Case 1
                Msg_A_Afficher = "Veuillez sélectionner machine avant de générer la fiche outils."

                'Confirmation de génération
            Case 2
                Msg_A_Afficher = "Génération fiche outils effectuée pour le programme """ & Tab_Valeurs_Retours(0) & _
                    """, de la phase """ & Tab_Valeurs_Retours(1) & """, dans le dossier """ & Tab_Valeurs_Retours(2) & """."

                'Si pas d'opération de programme
            Case 3
                Msg_A_Afficher = "Aucune opération dans le programme """ & Valeur_Retour_Seule & """."

                'Si plusieurs correcteur entrés
            Case 4
                Msg_A_Afficher = "Ne mettre qu'un seul correcteur à l'outil """ & Valeur_Retour_Seule & """."

                'Si erreur de lecture de fichier extérieurs à l'application (CSV, PDF, ...)
            Case 5
                Msg_A_Afficher = "Erreur lors de la création, la lecture ou l'écriture du fichier." & vbCrLf & vbCrLf & _
                    "Fichier en lecture seul, déplacé, supprimé ou inaccessible."

                'Si nombre caractères obligatoire dans chaîne
            Case 6
                Msg_A_Afficher = Valeur_Retour_Seule & " caractère(s) obligatoire(s)."

                'Si nombre caractères impossible dans chaîne
            Case 7
                Msg_A_Afficher = Valeur_Retour_Seule & " caractère(s) impossible(s)."

                'Si caractère ou chaine interdit trouvés dans chaîne
            Case 8
                Msg_A_Afficher = "Caractère ou chaîne """ & Valeur_Retour_Seule & """ interdit(e)."

                'Si nombre caractères dans chaîne non-compris entre...
            Case 9
                Msg_A_Afficher = "Le nombre de caractères doit-être compris entre " & _
                    Tab_Valeurs_Retours(0) & " et " & Tab_Valeurs_Retours(1) & """."

                'Si caractère ou chaîne introuvable dans autre chaîne
            Case 10
                Msg_A_Afficher = "Caractère ou chaîne """ & Valeur_Retour_Seule & """ non-présente dans la chaîne."

                'Demande de confirmation d'enregistrement
            Case 11
                Msg_A_Afficher = "Etes-vous sur de vouloir enregistrer la ligne ?"

                'Confirmation d'enregistrement
            Case 12
                Msg_A_Afficher = "Enregistrement effectué avec succès."

                'Confirmation d'annulation d'enregistrement
            Case 13
                Msg_A_Afficher = "Enregistrement abandonné."

                'Demande de confirmation de suppression
            Case 14
                Msg_A_Afficher = "Etes-vous sur de vouloir supprimer la ligne ?"

                'Confirmation de suppression
            Case 15
                Msg_A_Afficher = "Suppression effectuée avec succès."

                'Confirmation d'annulation de suppression
            Case 16
                Msg_A_Afficher = "Suppression abandonnée."

                'Si pas de CATProcess ouvert dans CATIA
            Case 17
                Msg_A_Afficher = "Pas de CATProcess actif dans CATIA."

                'Demande de confirmation de création de fichier
            Case 18
                Msg_A_Afficher = "Voulez-vous créer le fichier """ & Valeur_Retour_Seule & """ ?"

                'Si numéro d'origine programme différent de 1 ou 2
            Case 19
                Msg_A_Afficher = "Numéro d'origine différent de 1 et de 2 dans la phase """ & Valeur_Retour_Seule & """." & _
                    vbCrLf & vbCrLf & "Pas d'interruption de génération donc vérifier le contenu."

                'Aucun élément sélectionné dans liste déroulante
            Case 20
                Msg_A_Afficher = "Veuillez sélectionner un élément dans la liste déroulante avant de valider."

                'Si fichier à créer déja existant et autorisation de le supprimer
            Case 21
                Msg_A_Afficher = "Le fichier """ & Valeur_Retour_Seule & """ existe déjà, voulez-vous l'écraser ?"

                'Si diverses erreurs possibles lors de la génération de fiche outils
            Case 22
                Msg_A_Afficher = "Erreur lors de l'écriture de la fiche outils." & _
                    vbCrLf & vbCrLf & "Génération partielle de la fiche."

                'Si erreur lors de la génération de l'aperçu de la fiche outils
            Case 23
                Msg_A_Afficher = "Erreur lors de la génération du JPG de l'aperçu 3D." & _
                    vbCrLf & vbCrLf & "Génération de la fiche sans l'aperçu."

                'Si erreur lors de la convertion des vitesse de coupe
            Case 24
                Msg_A_Afficher = "Erreur lors de la convertion des vitesses de coupe." & _
                    vbCrLf & vbCrLf & "Génération partielle des vitesses de coupe."

                'Si erreur lors de la création d'un dossier ou fichier
            Case 25
                Msg_A_Afficher = "Erreur lors de la création du dossier """ & Valeur_Retour_Seule & """." & _
                    vbCrLf & vbCrLf & "Dossier en lecture seul, déplacé, supprimé ou inaccessible."

                'Confirmation de calcul du temps de cycle programme
            Case 26
                Msg_A_Afficher = "Calcul du temps de cycle effectué pour le programme """ & Valeur_Retour_Seule & """."

                'Si pas de machine déclarée dans Phase du Process
            Case 27
                Msg_A_Afficher = "Pas de nom de machine définie." & vbCrLf & vbCrLf & "Poursuite de la génération."

                'Si mot de passe non-OK
            Case 28
                Msg_A_Afficher = "Mot de passe érroné."

                'Si erreur lors de l'écriture de l'entête de la fiche outils
            Case 29
                Msg_A_Afficher = "Erreur lors de l'écriture de l'entête de la fiche outils." & _
                    vbCrLf & vbCrLf & "Génération partielle de l'entête."

                'Si pas de programme sélectionné
            Case 30
                Msg_A_Afficher = "Veuillez sélectionner un programme dans la liste déroulante avant de valider."

                'Si pas la dernière version de l'application
            Case 31
                Msg_A_Afficher = "L'application lancée n'est pas la dernière disponible."

                'Si nombre caractères non-compris entre, ou NA non-trouvé
            Case 32
                Msg_A_Afficher = "Nombre de caractère non-compris entre " & Tab_Valeurs_Retours(0) & " et " & _
                    Tab_Valeurs_Retours(1) & ", ou " & Tab_Valeurs_Retours(2) & " non-trouvé."

                'Demande de confirmation de création de dossier
            Case 33
                Msg_A_Afficher = "Voulez-vous créer le dossier """ & Valeur_Retour_Seule & """ ?"

                'Confirmation de création dossier
            Case 34
                Msg_A_Afficher = "Création du dossier """ & Valeur_Retour_Seule & """ effectuée avec succès."

                'Confirmation de création fichier
            Case 35
                Msg_A_Afficher = "Création du fichier """ & Valeur_Retour_Seule & """ effectuée avec succès."

                'Confirmation d'abandon d'opération
            Case 36
                Msg_A_Afficher = "Opération abandonnée."

                'Si la ligne à enregistrée est précédée d'une ligne vierge dans DGV
            Case 37
                Msg_A_Afficher = "Plusieurs nouvelles lignes ont été créées." & vbCrLf & vbCrLf & _
                    "Enregistrer d'abord la ligne correspondant au numéro d'enregistrement " & Valeur_Retour_Seule & "."

                'Si ligne dans Listing à été créée par quelqu'un d'autre le temps de l'ouverture
            Case 38
                Msg_A_Afficher = "Une ligne à été créée dans le listing par un autre utilisateur." & _
                    vbCrLf & vbCrLf & "Veuillez redémmarer l'application."

                'Validation de l'envoi des données d'outil vers CATIA
            Case 39
                Msg_A_Afficher = "Envoi dans CATIA effectué avec succès."

                'Validation de l'envoi des données d'outil CATIA vers DGV
            Case 40
                Msg_A_Afficher = "Envoi dans la grille effectué avec succès."

                'Si rien de checké ou pas de valeur dans DGV Listing programme
            Case 41
                Msg_A_Afficher = "Veuillez sélectionner les éléments à modifier."

                'Si valeur dans case DGV NOK
            Case 42
                Msg_A_Afficher = "La valeur entrée n'est pas de type numérique ou le séparateur décimal n'est pas un point." & _
                    vbCrLf & vbCrLf & "La valeur précédente à été rétablie."

                'Outil primaire et secondaire obligatoire
            Case 43
                Msg_A_Afficher = "Type d'outil principal et secondaire (fraisage ou tournage + plaquette tournage) obligatoires pour pouvoir configurer les conversions."

                'Si diverses erreurs lors de la renumérotation des outils dans CATIA
            Case 44
                Msg_A_Afficher = "Erreur lors de la renumérotation des outils dans CATIA." & _
                    vbCrLf & vbCrLf & "Pas d'interruption de génération donc vérifier le contenu."

                'Si erreur de la lecture ou l'écriture du Process CATIA
            Case 45
                Msg_A_Afficher = "Erreur lors de la lecture ou l'écriture du fichier CATIA." & _
                    vbCrLf & vbCrLf & "Application CATIA non-lancée, dossier en lecture seul, déplacé, supprimé ou inaccessible."

                'Si erreur sur l'extraction du temps de cycle sur une opération
            Case 46
                Msg_A_Afficher = "Erreur lors de la lecture d'une durée d'opération." & _
                    vbCrLf & vbCrLf & "Relancer un calcul complet du programme """ & Valeur_Retour_Seule & """ ."

                'Si ComboBox de choix du temps de cycle d'un chgt d'outil
            Case 47
                Msg_A_Afficher = "Veuillez sélectionner un temps de changement d'outil avant de calculer le temps de cycle."

                'Si fichier à créer déja existant et pas de droit de la supprimer
            Case 48
                Msg_A_Afficher = "Le fichier """ & Valeur_Retour_Seule & """ existe déjà."

                'Si caractère ou chaine interdit trouvés dans chaîne données
            Case 49
                Msg_A_Afficher = "Caractère ou chaîne """ & Tab_Valeurs_Retours(0) & _
                    """ interdit. Caractère ou chaîne autorisés: """ & Tab_Valeurs_Retours(1) & """."

                'Si nombre caractères non-égale à, ou NA non-trouvé
            Case 50
                Msg_A_Afficher = "Nombre de caractère non-égale à " & Tab_Valeurs_Retours(0) & _
                    ", ou """ & Tab_Valeurs_Retours(1) & """ non-trouvé."

                'Si caractère ou chaîne différent d'une autre
            Case 51
                Msg_A_Afficher = "Caractère ou chaîne """ & Valeur_Retour_Seule & """ impossible."

                'Type outil primaire obligatoire dans les configurations outils
            Case 52
                Msg_A_Afficher = "Type d'outil principal (fraisage ou tournage + plaquette tournage) obligatoire dans les configurations d'outils."

                'Si plusieurs éléments sélectionnés dans CATIA
            Case 53
                Msg_A_Afficher = "Plusieurs éléments sélectionnés dans CATIA. Seul le premier dans l'ordre de sélection sera prit en compte."

                'Si l'activité sélectionnée dans Process n'est pas un changement d'outil
            Case 54
                Msg_A_Afficher = "L'activité sélectionnée n'est pas un changement d'outil."

                'Si le type de l'outil sélectionné dans Process n'est pas le même 
                'que celui de l'outil de fraisage principal ou secondaire
            Case 55
                Msg_A_Afficher = "Le changement d'outil sélectionné n'est pas de type """ & _
                    Valeur_Retour_Seule & """."

                'Si le type de l'outil sélectionné dans Process n'est pas le même 
                'que celui de l'outil de tournage principal ou secondaire
            Case 56
                Msg_A_Afficher = "Le changement d'outil sélectionné n'est pas de type """ & _
                    Tab_Valeurs_Retours(0) & """ avec """ & Tab_Valeurs_Retours(1) & """."

                'Si le type de l'outil sélectionné dans Process n'est pas le même 
                'que celui de l'outil de fraisage et de tournage, principal ou secondaire
            Case 57
                Msg_A_Afficher = "Le changement d'outil sélectionné n'est pas de type """ & _
                    Tab_Valeurs_Retours(0) & """, ni de type """ & Tab_Valeurs_Retours(1) & """ avec """ & Tab_Valeurs_Retours(2) & """."

                'Si manque de parenthèse fermante dans la formule de conversion des outils dans CATIA
            Case 58
                Msg_A_Afficher = "Manque parenthèse fermante dans la formule de conversion de l'outil."

                'Si erreur dans la formule de conversion des outils dans CATIA
            Case 59
                Msg_A_Afficher = "Erreur dans la formule de conversion de l'outil. Formule impossible à calculer"

                'Si erreur lors de la connexion FOCAS
            Case 60
                Msg_A_Afficher = "Erreur lors de la connexion FOCAS (retour erreur N°" & Valeur_Retour_Seule & ")."

                'Si erreur lors de la déconnexion FOCAS
            Case 61
                Msg_A_Afficher = "Erreur lors de la déconnexion FOCAS (retour erreur N°" & Valeur_Retour_Seule & ")."

                'Si erreur lors de la lecture FOCAS
            Case 62
                Msg_A_Afficher = "Erreur lors de la lecture FOCAS (retour erreur N°" & Valeur_Retour_Seule & ")."
        End Select

        '------------------------------------------------------- Box standard ------------------------------------------------------------

        'Choix de la Box à afficher

        Select Case Num_Box

            'Box "Information" avec bouton "Ok"
            Case 1
                Return MessageBox.Show(Msg_A_Afficher, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                'Box "Question" avec bouton Oui/Non
            Case 2
                Return MessageBox.Show(Msg_A_Afficher, "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                'Box "Exclamation" avec bouton "Ok"
            Case 3
                Return MessageBox.Show(Msg_A_Afficher, "Avertissement", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

                'Box "Erreur critique" avec bouton "OK"
            Case 4
                Return MessageBox.Show(Msg_A_Afficher, "Erreur critique", MessageBoxButtons.OK, MessageBoxIcon.Stop)

                'Box "Erreur critique avec sortie de commande" avec bouton "OK" (ne pas oublier d'ajouter un Exit Sub après cette commande)
            Case 5
                Return MessageBox.Show(Msg_A_Afficher & vbCrLf & vbCrLf & "Sortie de commande.", _
                                       "Erreur critique avec sortie de commande", MessageBoxButtons.OK, MessageBoxIcon.Stop)

                'Box "Erreur critique avec fermeture d'application" avec bouton "OK" (ne pas oublier d'ajouter un End après cette commande)
            Case 6
                Return MessageBox.Show(Msg_A_Afficher & vbCrLf & vbCrLf & "L'application va quitter.", _
                                       "Erreur critique avec fermeture d'application", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        End Select
    End Function
End Module
