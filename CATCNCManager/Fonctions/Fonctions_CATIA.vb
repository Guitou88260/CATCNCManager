Public Module Fonctions_CATIA


    'Génération de la fiche outils

    Public Function Generation_Fiche_Outils _
        (ByVal CATIA As Object, ByVal docCatiaCrt As Object, ByVal phaseCrt As Object, ByVal progCrt As Object,
         ByVal comboBoxMachine As String, ByVal dossierGeneration As String) As Boolean

        'Variables
        Dim EXCEL As Object = Nothing
        Dim fileCrt As Workbook = Nothing
        Dim feuilleCrt As Worksheet = Nothing
        Dim listeOP As Object
        Dim compteurOP As Integer

        'Valuation variables
        listeOP = progCrt.Activities
        compteurOP = listeOP.Count

        'Si opération trouvée...
        If compteurOP > 0 Then

            'Création des dossiers et fichiers si besion et si erreur, renvoi False
            If Fonctions_Diverses.Ctrl_Arbor_Ou_Creation_Fichier _
                (dossierGeneration, True, progCrt.Name & ".xlsx", Dossier_Reseau &
                 "\FO_Modeles\" & comboBoxMachine & ".xlsx", ) = False Then Return False

            'Ouverture Excel et si erreur, renvoi False pour sortie de Sub
            If Fonctions_Diverses.Ouverture_Fichier_Excel _
                (dossierGeneration & "\" & progCrt.Name & ".xlsx", EXCEL,
                 fileCrt, feuilleCrt) = False Then Return False

            'Lecture CSV config FO et si erreur, retourne faux
            Dim tabConfigTmp(,) As String = Nothing
            If Fonctions_CSV.Lecture_CSV(Dossier_Reseau & "\Config_FO\Fichier_Config_FO.csv",
                                             tabConfigTmp, 5) = False Then Return False

            'Si traitement entête True
            If CBool(Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMachine, "Machine", "Traiter entête", 1)) = True Then _
                Fonctions_CATIA.RemplEnteteFO(docCatiaCrt, phaseCrt, progCrt, feuilleCrt, comboBoxMachine, tabConfigTmp)

            'Si insertion image différent de NA
            Dim cellImage As String
            cellImage = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMachine, "Machine", "Cellule image pièce", 1)
            If cellImage <> "NA" Then Fonctions_CATIA.ImagePC(CATIA, progCrt, dossierGeneration, feuilleCrt, cellImage)

            'Appel Sub remplissage corps de FO
            Fonctions_CATIA.RemplCorpsProgFO(phaseCrt, listeOP, compteurOP, feuilleCrt, tabConfigTmp, comboBoxMachine)

            'Enregistrement et fermeture Excel et si erreur, renvoi False
            If Fonctions_Diverses.Enregistrement_Fermeture_Excel(EXCEL, fileCrt) = False Then Return False

            'Variable message
            Dim Msg(2) As String
            Msg(0) = progCrt.Name
            Msg(1) = phaseCrt.Name
            Msg(2) = dossierGeneration

            'Message
            Fonctions_Messages.Appel_Msg(2, 1, Msg, )

            'Sinon
        Else

            'Message et renvoi False
            Fonctions_Messages.Appel_Msg(3, 5, , progCrt.Name)
            Return False
        End If

        'Retourne True
        Return True
    End Function


    'Insertion aperçu 3D de CATIA puis insertion dans Excel

    Public Sub ImagePC(ByVal CATIA As Object, ByVal progCrt As Object, ByVal dossierGeneration As String, ByVal feuilleCrt As Worksheet, ByVal cellImage As String)

        'Gestion des exceptions
        Try

            'Variables
            Dim Vue_WINDOWS As SpecsAndGeomWindow
            Dim Vue_Initiale As CatSpecsAndGeomWindowLayout
            Dim Larg_Fenetre, Haut_Fenetre, Larg_Viewer, Haut_Viewer As Integer
            Dim Ecran As Viewer3D
            Dim Couleur_Initiale_Fond_Ecran(2), Transparence_Fond_Ecran(2) As Object

            'Valuation variable transparence
            Transparence_Fond_Ecran(0) = 1
            Transparence_Fond_Ecran(1) = 1
            Transparence_Fond_Ecran(2) = 1

            'Vue Windows
            Vue_WINDOWS = CATIA.ActiveWindow

            'Chargement des anciens paramètres d'affichage
            Vue_Initiale = Vue_WINDOWS.Layout
            Larg_Fenetre = Vue_WINDOWS.Width
            Haut_Fenetre = Vue_WINDOWS.Height

            'Récup coté le + petit
            Dim DimCell As DimCellXLS
            DimCell = RecupDimPosCell(feuilleCrt.Range(cellImage))
            Dim coteMinus As Integer

            If DimCell.Larg > DimCell.Haut Then
                coteMinus = DimCell.Haut
                DimCell.posX = ((DimCell.Larg - DimCell.Haut) / 2) + DimCell.posX
            Else
                coteMinus = DimCell.Larg
                DimCell.posY = ((DimCell.Haut - DimCell.Larg) / 2) + DimCell.posY
            End If

            'Définition taille photo à obtenir
            Vue_WINDOWS.Width = coteMinus * 3
            Vue_WINDOWS.Height = coteMinus * 3
            Vue_WINDOWS.Layout = 1

            'Activation de la prise de vue
            Ecran = Vue_WINDOWS.ActiveViewer

            'Activation de certains paramètres de visu (mode 3D, couleurs, ...)
            Ecran.GetBackgroundColor(Couleur_Initiale_Fond_Ecran)
            Ecran.PutBackgroundColor(Transparence_Fond_Ecran)
            Ecran.RenderingMode = CatRenderingMode.catRenderShadingWithEdges
            Ecran.Reframe()

            'TODO Voir pour l'aperçu des FO: dimensionement instable. A l'air de marcher pour l'instant. A suivre
            'Dim timer1 As New Timer
            'timer1.Interval = 30000
            'timer1.Start()

            'Si aperçu présent, le supprimer
            If System.IO.File.Exists(dossierGeneration & "\" & progCrt.Name & ".jpg") = True Then _
                System.IO.File.Delete(dossierGeneration & "\" & progCrt.Name & ".jpg")

            'Génération de l'aperçu en .jpg
            Ecran.CaptureToFile(5, (dossierGeneration & "\" & progCrt.Name & ".jpg"))

            'Insertion aperçu dans fiche (+1 & -2 pour marge image)
            feuilleCrt.Shapes.AddPicture(dossierGeneration & "\" & progCrt.Name & ".jpg", True, True, DimCell.posX + 1, DimCell.posY + 1, coteMinus - 2, coteMinus - 2)

            'Retour des paramètres initiaux
            Ecran.PutBackgroundColor(Couleur_Initiale_Fond_Ecran)
            Larg_Viewer = Ecran.Width
            Haut_Viewer = Ecran.Height
            Vue_WINDOWS.Width = Larg_Fenetre
            Vue_WINDOWS.Height = Haut_Fenetre
            Vue_WINDOWS.Layout = Vue_Initiale

            'Vidage variables
            Ecran = Nothing
            Vue_WINDOWS = Nothing

            'Pour toutes les erreurs
        Catch

            'Message et sortie de Sub
            Fonctions_Messages.Appel_Msg(23, 5, , )
            Exit Sub
        End Try
    End Sub


    'Remplissage entête fiche outils

    Public Sub RemplEnteteFO(ByVal Document_CATIA_Courant As Object, ByVal Phase_Courante As Object,
                             ByVal Prog_Courant As Object, ByVal Feuille_Courante As Worksheet,
                             ByVal comboBoxMach As String, ByVal tabConfigTmp(,) As String)

        'Gestion des exceptions
        Try

            'Val config FO
            Dim tabValTmp(10, 1) As Object
            tabValTmp(0, 0) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Traiter entête", 1)
            tabValTmp(1, 0) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Cellule nom programme", 1)
            tabValTmp(2, 0) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Cellule commentaire programme", 1)
            tabValTmp(3, 0) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Cellule nom phase", 1)
            tabValTmp(4, 0) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Cellule commentaire phase", 1)
            tabValTmp(5, 0) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Cellule nom process", 1)
            tabValTmp(6, 0) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Cellule nom product", 1)
            tabValTmp(7, 0) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Cellule nom machine", 1)
            tabValTmp(8, 0) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Cellule date", 1)
            tabValTmp(9, 0) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Cellule auteur", 1)
            tabValTmp(10, 0) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Couleur RGB fond entête", 1)
            'Param CATIA
            tabValTmp(1, 1) = Prog_Courant.Name
            tabValTmp(2, 1) = Prog_Courant.Comment
            tabValTmp(3, 1) = Phase_Courante.Name
            tabValTmp(4, 1) = Phase_Courante.Comment
            tabValTmp(5, 1) = Replace(Document_CATIA_Courant.Name, ".CATProcess", "")

            'TODO GetPartName à l'air de marcher mais à suivre à l'utilisation
            tabValTmp(6, 1) = Phase_Courante.GetPartName
            tabValTmp(7, 1) = Phase_Courante.Machine.Name
            tabValTmp(8, 1) = Renvoi_Maintenant_Et_Auteur(1)
            tabValTmp(9, 1) = Renvoi_Maintenant_Et_Auteur(2)

            'Si remplissage entête = True
            If tabValTmp(0, 0) = True Then

                'Boucle remplissage entête
                For i = 1 To 9

                    'Si param <> NA
                    If CStr(tabValTmp(i, 0)) <> "NA" Then

                        'Val & color
                        Feuille_Courante.Range(tabValTmp(i, 0)).Value = tabValTmp(i, 1)
                        If CStr(tabValTmp(10, 0)) <> "NA" Then Feuille_Courante.Range(tabValTmp(i, 0)).Interior.Color = Fonctions_Diverses.ConvColorStrVersInt(tabValTmp(10, 0))
                    End If
                Next
            End If

            'Pour toutes les erreurs
        Catch

            'Message et sortie de Sub
            Fonctions_Messages.Appel_Msg(29, 3, , )
            Exit Sub
        End Try
    End Sub


    'Remplissage corps de fiche outils

    Public Sub RemplCorpsProgFO(ByVal phaseCrt As Object, ByVal listeOP As Object, ByVal cptOP As Integer,
                                ByVal feuilleCrt As Worksheet, ByVal tabConfigTmp(,) As String, ByVal comboBoxMach As String)

        'Gestion des exceptions
        Try

            'Val config FO
            Dim tabValTmp(29) As Object
            tabValTmp(0) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Première ligne programme", 1)
            tabValTmp(1) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Plage colonnes programme", 1)
            tabValTmp(2) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Traiter opérations", 1)
            tabValTmp(3) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Colonne nom opération", 1)
            tabValTmp(4) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Colonne commentaire opération", 1)
            tabValTmp(5) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Couleur RGB fond opération", 1)
            tabValTmp(6) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Blocs par opération", 1)
            tabValTmp(7) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Colonne bloc", 1)
            tabValTmp(8) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Bloc départ", 1)
            tabValTmp(9) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Incrément bloc", 1)
            tabValTmp(10) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Texte précédent bloc", 1)
            tabValTmp(11) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Traiter changements outil", 1)
            tabValTmp(12) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Colonne nom outil", 1)
            tabValTmp(13) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Colonne commentaire outil", 1)
            tabValTmp(14) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Colonne N° outil", 1)
            tabValTmp(15) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Texte précédent N° outil", 1)
            tabValTmp(16) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Couleur RGB fond outil", 1)
            tabValTmp(17) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Correcteurs type BUMOTEC", 1)
            tabValTmp(18) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Colonne N° correcteur", 1)
            tabValTmp(19) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Texte précédent N° correcteur", 1)
            tabValTmp(20) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Sous-programmes par outil", 1)
            tabValTmp(21) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Colonne sous-programme", 1)
            tabValTmp(22) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Sous-programme départ ST1", 1)
            tabValTmp(23) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Sous-programme départ ST2", 1)
            tabValTmp(24) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Incrément sous-programme", 1)
            tabValTmp(25) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Texte précédent sous-programme", 1)
            tabValTmp(26) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Traiter repères usinage", 1)
            tabValTmp(27) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Colonne repère", 1)
            tabValTmp(28) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Texte précédent N° repère", 1)
            tabValTmp(29) = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(tabConfigTmp, comboBoxMach, "Machine", "Couleur RGB fond repère", 1)

            Dim lignCrt As Integer
            lignCrt = tabValTmp(0)

            'Si traitement repères
            If CBool(tabValTmp(26)) = True Then

                'Ecriture texte avec N° origine, couleur et incrément ligne
                feuilleCrt.Range(Fonctions_Diverses.ReformColonnCell(tabValTmp(27), lignCrt)).Value =
                    tabValTmp(28) & phaseCrt.MachiningAxisSystem.OriginNumber
                If tabValTmp(29) <> "NA" Then feuilleCrt.Range(Fonctions_Diverses.ReformColonnCell(tabValTmp(1), lignCrt)).Interior.Color =
                    Fonctions_Diverses.ConvColorStrVersInt(tabValTmp(29))
                lignCrt += 1
            End If

            Dim numProgCrt As Integer

            'Si traitement chgt outil & sous-prog par outil
            If CBool(tabValTmp(11)) = True And CBool(tabValTmp(20)) = True Then

                'Val num prog courant
                If phaseCrt.MachiningAxisSystem.OriginNumber = 1 Then numProgCrt = CInt(tabValTmp(22))
                If phaseCrt.MachiningAxisSystem.OriginNumber = 2 And tabValTmp(23) <> "NA" Then _
                    numProgCrt = CInt(tabValTmp(23))
            End If

            Dim OPCrt, paramOP As Object
            Dim numBlocCrt As Integer
            Dim premChgtRep As Boolean
            premChgtRep = False

            'Boucle OP
            For i = 1 To cptOP

                'Récup OP
                OPCrt = listeOP.GetElement(i)

                'Si traitement chgt outil
                If CBool(tabValTmp(11)) = True And OPCrt.Type = "ToolChange" Or OPCrt.Type = "ToolChangeLathe" Then

                    'Ecriture nom OP chgt outil
                    feuilleCrt.Range(Fonctions_Diverses.ReformColonnCell(tabValTmp(3), lignCrt)).Value = OPCrt.Name

                    'Ecriture nom outil
                    If tabValTmp(12) <> "NA" Then _
                        feuilleCrt.Range(Fonctions_Diverses.ReformColonnCell(tabValTmp(12), lignCrt)).Value = OPCrt.Tool.Name

                    'Ecriture commentaire outil
                    If tabValTmp(13) <> "NA" Then _
                        feuilleCrt.Range(Fonctions_Diverses.ReformColonnCell(tabValTmp(13), lignCrt)).Value = OPCrt.Tool.Comment

                    'Ecriture N° outil
                    If tabValTmp(14) <> "NA" Then

                        'Outil fraisage
                        If OPCrt.Type = "ToolChange" Then feuilleCrt.Range(Fonctions_Diverses.ReformColonnCell(tabValTmp(14), lignCrt)).Value =
                        tabValTmp(15) & OPCrt.Tool.ToolNumber

                        'Outil tournage
                        If OPCrt.Type = "ToolChangeLathe" Then feuilleCrt.Range(Fonctions_Diverses.ReformColonnCell(tabValTmp(14), lignCrt)).Value =
                            tabValTmp(15) & OPCrt.ToolAssembly.ToolNumber
                    End If

                    'Si sous-prog par outil
                    If CBool(tabValTmp(20)) = True Then

                        'Si origine N°1 ou N°2
                        If phaseCrt.MachiningAxisSystem.OriginNumber = 1 Or
                            (phaseCrt.MachiningAxisSystem.OriginNumber = 2 And tabValTmp(23) <> "NA") Then

                            'Ecriture & incrément N° sous-prog
                            feuilleCrt.Range(Fonctions_Diverses.ReformColonnCell(tabValTmp(21), lignCrt)).Value =
                                tabValTmp(25) & numProgCrt

                            'Inc num sous-prog
                            numProgCrt += tabValTmp(24)
                        End If
                    End If

                    'Coloration fond outil
                    If tabValTmp(16) <> "NA" Then feuilleCrt.Range(Fonctions_Diverses.ReformColonnCell(tabValTmp(1), lignCrt)).Interior.Color =
                        Fonctions_Diverses.ConvColorStrVersInt(tabValTmp(16))

                    'Si bloc par OP, reval départ de bloc
                    If CBool(tabValTmp(6)) = True Then numBlocCrt = CInt(tabValTmp(8))

                    'Incrément ligne
                    lignCrt += 1

                    'Si traitement changement de repère
                ElseIf OPCrt.Type = "CoordinateSystem" And CBool(tabValTmp(26)) = True Then

                    'Feature OP
                    paramOP = OPCrt.GetFeature

                    'Si origine 1 cochée et = 2
                    If paramOP.Origin = 1 And paramOP.OriginNumber = 2 And
                        paramOP.OriginNumber <> phaseCrt.MachiningAxisSystem.OriginNumber And
                        premChgtRep = False Then

                        'Val ctrl passage
                        premChgtRep = True
                        'Ecriture texte avec N° origine, couleur et incrément ligne
                        feuilleCrt.Range(Fonctions_Diverses.ReformColonnCell(tabValTmp(27), lignCrt)).Value =
                            tabValTmp(28) & paramOP.OriginNumber

                        'Coloration ligne
                        If tabValTmp(29) <> "NA" Then feuilleCrt.Range _
                            (Fonctions_Diverses.ReformColonnCell(tabValTmp(1), lignCrt)).Interior.Color =
                            Fonctions_Diverses.ConvColorStrVersInt(tabValTmp(29))

                        'Inc ligne
                        lignCrt += 1

                        'Val num prog ST2
                        If tabValTmp(23) <> "NA" Then numProgCrt = CInt(tabValTmp(23))

                        'Si bloc par OP, reval départ de bloc
                        If CBool(tabValTmp(6)) = True Then numBlocCrt = CInt(tabValTmp(8))
                    End If

                    'Si Instruction PP et balise Tracut et Copy (pour les transformations), ne rien faire...
                ElseIf OPCrt.Type = "PPInstruction" Or OPCrt.Type = "MfgTracutEnd" Or
                OPCrt.Type = "MfgCopyEnd" Or OPCrt.Type = "MfgCopyBegin" Or OPCrt.Type = "MfgCopyOrder" Then

                    'Sinon, traitement par défaut...
                Else

                    'Si traitement OP
                    If CBool(tabValTmp(2)) = True Then

                        'Ecriture nom OP
                        If tabValTmp(3) <> "NA" Then feuilleCrt.Range(Fonctions_Diverses.ReformColonnCell(tabValTmp(3), lignCrt)).Value = OPCrt.Name

                        'Ecriture commentaire OP
                        If tabValTmp(4) <> "NA" Then feuilleCrt.Range(Fonctions_Diverses.ReformColonnCell(tabValTmp(4), lignCrt)).Value = OPCrt.Description

                        'Ecriture N° bloc
                        If tabValTmp(6) = True Then

                            'Val num bloc & inc
                            feuilleCrt.Range(Fonctions_Diverses.ReformColonnCell(tabValTmp(7), lignCrt)).Value = tabValTmp(10) & numBlocCrt
                            numBlocCrt += CInt(tabValTmp(9))
                        End If

                        'Coloration fond OP
                        If tabValTmp(5) <> "NA" Then feuilleCrt.Range _
                            (Fonctions_Diverses.ReformColonnCell(tabValTmp(1), lignCrt)).Interior.Color =
                            Fonctions_Diverses.ConvColorStrVersInt(tabValTmp(5))

                        'Ecriture infos relative au numéro d'origine
                        'Si traiter les N° de repère
                        If CBool(tabValTmp(26)) = True Then

                            'Si origine N°1
                            If phaseCrt.MachiningAxisSystem.OriginNumber = 1 And paramOP Is Nothing Then

                                'Ecriture correcteur
                                feuilleCrt.Range(Fonctions_Diverses.ReformColonnCell(tabValTmp(18), lignCrt)).Value = tabValTmp(19) & "0"

                                'Si origine N°2
                            ElseIf phaseCrt.MachiningAxisSystem.OriginNumber = 2 Or Not paramOP Is Nothing Then

                                'Ecriture correcteur
                                feuilleCrt.Range(Fonctions_Diverses.ReformColonnCell(tabValTmp(18), lignCrt)).Value = tabValTmp(19) & "4"
                            End If
                        End If

                        'Inc ligne
                        lignCrt += 1
                    End If
                End If
            Next

            'Pour toutes les erreurs
        Catch
            'MsgBox(Err.Description)
            'Message et sortie de Sub
            Fonctions_Messages.Appel_Msg(22, 3, , )
            Exit Sub
        End Try
    End Sub


    'Renvoi VC en m/s dans fiche outils Excel

    Public Function Renvoi_VC_Calculee(ByVal Outil_Courant As Object) As Double

        'Gestion des exceptions
        Try

            'Variables
            Dim Compteur_Param As Integer
            Dim Liste_Param() As Object
            Dim VC_Calculee As Object = Nothing

            'Compte le nombre de paramètres
            Compteur_Param = Outil_Courant.NumberOfAttributes

            'Redimenssionnement tab
            ReDim Liste_Param(Compteur_Param - 1)

            'Valuation tab
            Outil_Courant.GetListOfAttributes(Liste_Param)

            'Boucle lecture tab
            For Compteur = LBound(Liste_Param) To UBound(Liste_Param)

                'Si VC pour outil axial trouvée
                If Liste_Param(Compteur) = "MFG_VC" Then

                    'Calcul
                    VC_Calculee = Outil_Courant.GetAttribute("MFG_VC").Value / 1000
                    VC_Calculee = VC_Calculee / 60

                    'Si VC pour autre outil trouvée
                ElseIf Liste_Param(Compteur) = "MFG_VC_ROUGH" Then

                    'Calcul
                    VC_Calculee = Outil_Courant.GetAttribute("MFG_VC_ROUGH").Value / 1000
                    VC_Calculee = VC_Calculee / 60
                End If
            Next

            'Retourne résultat
            Return VC_Calculee

            'Pour toutes les erreurs
        Catch

            'Message, sortie de Sub et retourne Nothing
            Fonctions_Messages.Appel_Msg(24, 3, , )
            Return Nothing
        End Try
    End Function


    'Création machine dans CATIA

    Public Function Creation_Machine_Dans_CATIA _
        (ByVal Document_CATIA_Courant As Object, ByVal Phase_Courante As Object,
         ByVal Type_Machine As String, ByVal Nom_Machine As String) As Boolean

        'Gestion des exceptions
        Try

            'Variables
            Dim Compteur_Ressource As Integer = 1
            Dim Nom_Complet_Nouvelle_Machine As String

            'Recherche si ressource déjà existante et si erreur, retourne False
            If Fonctions_CATIA.Renvoi_Num_Ressource _
                (Document_CATIA_Courant, Nom_Machine & " *",
                 "*", Compteur_Ressource) = False Then Return False

            'Valuation variable
            Nom_Complet_Nouvelle_Machine = Nom_Machine & " *" & Compteur_Ressource & "*"

            'Si pas de machine dans la phase
            If IsNothing(Phase_Courante.Machine) Then

                'Création d'une machine du bon type
                Phase_Courante.CreateMachine(Type_Machine)

                'Sinon
            Else

                'Si machine n'est pas du type spécifié, en créer une
                If Phase_Courante.Machine.MachineType <> Type_Machine Then

                    'Création de la machine
                    Phase_Courante.CreateMachine(Type_Machine)
                End If
            End If

            'Variable
            Dim Machine_Courante As Object

            'Valuation machine courante
            Machine_Courante = Phase_Courante.Machine

            'Renommage ressource et si erreur, retourne False
            If Fonctions_CATIA.Renommage_Ressource_CATIA _
                (Document_CATIA_Courant, Machine_Courante.Name, Nom_Complet_Nouvelle_Machine) = False Then Return False

            'Renommage machine courante et si erreur, retourne False
            If Fonctions_CATIA.Modif_Valeur_Param _
                (Machine_Courante, "MFG_NAME", Nom_Complet_Nouvelle_Machine) = False Then Return False

            'Retourne True
            Return True

            'Pour toutes les erreurs
        Catch

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(45, 5, , )
            Return False
        End Try
    End Function


    'Renumérotation de tous les outils ou recherche de l'opération suivant le changement d'outil courant dans programme CATIA

    Public Function Renum_Outil_Ou_Recherche_OP_CATIA _
        (ByVal Prog_Courant As Object, ByVal Type_Traitement As Integer,
         Optional ByVal Nom_Chgt_Outil_Courant As String = Nothing,
         Optional ByRef OP_Suivante As Object = Nothing) As Boolean

        'Gestion des exceptions
        Try

            'Variables
            Dim Liste_OP As Object
            Dim Compteur_OP As Integer
            Dim Compteur_Outil As Integer

            'Valuation variables
            Liste_OP = Prog_Courant.Activities
            Compteur_OP = Liste_OP.Count

            'Si opération trouvée...
            If Compteur_OP > 0 Then

                'Variable
                Compteur_Outil = 1

                'Boucle lecture des opérations
                For Boucle_Operation = 1 To Compteur_OP

                    'Variable
                    Dim OP_Courante As Object

                    'Valuation variable
                    OP_Courante = Liste_OP.GetElement(Boucle_Operation)

                    'Si "ToolChange" ou "ToolChangeLathe"
                    If OP_Courante.Type = "ToolChange" Or OP_Courante.Type = "ToolChangeLathe" Then

                        'Si type de traitement 1, renommage outil seulement
                        If Type_Traitement = 1 Then

                            'Si "ToolChange"
                            If OP_Courante.Type = "ToolChange" Then

                                'Modification valeur paramètre
                                If Fonctions_CATIA.Modif_Valeur_Param _
                                    (OP_Courante.Tool, "MFG_TOOL_NUMBER", Compteur_Outil.ToString) = False Then Return False

                                'Si "ToolChangeLathe"
                            ElseIf OP_Courante.Type = "ToolChangeLathe" Then

                                'Modification valeur paramètre
                                If Fonctions_CATIA.Modif_Valeur_Param _
                                    (OP_Courante.ToolAssembly, "MFG_TOOL_NUMBER", Compteur_Outil.ToString) = False Then Return False
                            End If

                            'Incrémentation compteur
                            Compteur_Outil = Compteur_Outil + 1

                            'Si type de traitement 2, recherche opération suivant changement d'outil
                        ElseIf Type_Traitement = 2 Then

                            'Si nom OP courante = nom du chgt d'outil courant
                            If OP_Courante.Tool.Name = Nom_Chgt_Outil_Courant Then

                                'Si OP suivante présente
                                If Boucle_Operation + 1 < Liste_OP.Count Then

                                    'Valuation variable
                                    OP_Suivante = Liste_OP.GetElement(Boucle_Operation + 1)

                                    'Sortie de boucle
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next
            End If

            'Retourne True
            Return True

            'Pour toutes les erreurs
        Catch

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(45, 5, , )
            Return False
        End Try
    End Function


    'Envoi des données communes des outils dans CATIA (renommage ressource et changement d'outil et valuation des paramètres suivant TollTips)

    Public Function Envoi_Donnees_Communes_Outil _
        (ByVal Document_CATIA_Courant As Object, ByVal DGV_A_Utiliser As DataGridView,
         ByVal Objet_A_Valuer As Object, ByVal Nouveau_Nom_Outil As String,
         ByVal Indice_Nom_Ressource As String) As Boolean

        'Renommage ressource outil (avec passage en majuscule des textes) et si erreur, retourne False
        If Fonctions_CATIA.Renommage_Ressource_CATIA _
            (Document_CATIA_Courant, Objet_A_Valuer.GetAttribute("MFG_NAME").Value,
             Indice_Nom_Ressource & "-" & Nouveau_Nom_Outil) =
         False Then Return False

        'Renommage changement outil et commentaire (avec passage en majuscule des textes) et si erreur, retourne False
        If Fonctions_CATIA.Modif_Valeur_Param(Objet_A_Valuer, "MFG_NAME", Nouveau_Nom_Outil) =
            False Then Return False
        If Fonctions_CATIA.Modif_Valeur_Param _
            (Objet_A_Valuer, "MFG_COMMENT", UCase _
             (DGV_A_Utiliser.Columns.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Utiliser, "Numéro d'article fournisseur")).HeaderText) &
             ": " & UCase(DGV_A_Utiliser.CurrentRow.Cells.Item _
                          (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Utiliser, "Numéro d'article fournisseur")).Value)) =
                  False Then Return False

        'Retourne True
        Return True
    End Function


    'Recherche nom de paramètre CATIA dans ToolTips de toutes colonnes DGV et valuation variables correspondantes dans CATIA

    Public Function Modif_Variable_Selon_ToolTips _
        (ByVal DGV_A_Lire As DataGridView, ByVal Object_A_Valuer As Object,
         ByVal Nom_Colonne_Debut As String, ByVal Nom_Colonne_Fin As String,
         Optional ByVal Tab_Des_Valeurs(,) As String = Nothing,
         Optional ByVal Chaine_A_Trouver As String = Nothing,
         Optional ByVal Nom_Colonne_Chaine_A_Trouver As String = Nothing) As Boolean

        'Gestion des exceptions
        Try

            'Déclaration et dimensionnement variable
            Dim Liste_Attributs_CATIA(Object_A_Valuer.NumberOfAttributes - 1, 0) As Object

            'Valuation variable
            Object_A_Valuer.GetListOfAttributes(Liste_Attributs_CATIA)

            'Boucle colonne
            For Boucle_Colonne = Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, Nom_Colonne_Debut) To _
                Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Lire, Nom_Colonne_Fin)

                'Si num paramètre trouvé dans ToolTip colonne, si colonne visible et différente d'un type bouton
                If InStr(DGV_A_Lire.Columns.Item(Boucle_Colonne).ToolTipText, "(11)") <> 0 And
                DGV_A_Lire.Columns.Item(Boucle_Colonne).Visible = True And
                DGV_A_Lire.Columns.Item(Boucle_Colonne).GetType.ToString <>
                "System.Windows.Forms.DataGridViewButtonColumn" Then

                    'Déclaration variable
                    Dim Parametre_CATIA_A_Valuer As String

                    'Valuation variable
                    Parametre_CATIA_A_Valuer = Fonctions_Diverses.Renvoi_Chaine_Encadree _
                        (DGV_A_Lire.Columns.Item(Boucle_Colonne).ToolTipText, "(11) Paramètre CATIA : """, """")

                    'Si paramètre présent dans l'objet à modifier
                    If Fonctions_Tableau.Presence_Chaine_Dans_Tab _
                        (Liste_Attributs_CATIA, Parametre_CATIA_A_Valuer, 0, , ) = True Then

                        'Si Tab des valeurs valué, pioche dans la DGV et le Tab
                        If Not IsNothing(Tab_Des_Valeurs) And Chaine_A_Trouver <>
                            Nothing And Nom_Colonne_Chaine_A_Trouver <> Nothing Then

                            'Si colonne de type CheckBox
                            If DGV_A_Lire.Columns.Item(Boucle_Colonne).GetType.ToString =
                                "System.Windows.Forms.DataGridViewCheckBoxColumn" Then

                                'Passage en minuscule de la valeur, valuation paramètre et si erreur, retourne False
                                If Fonctions_CATIA.Modif_Valeur_Param(Object_A_Valuer, Parametre_CATIA_A_Valuer,
                                LCase(Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(Tab_Des_Valeurs, Chaine_A_Trouver, Nom_Colonne_Chaine_A_Trouver,
                                DGV_A_Lire.Columns.Item(Boucle_Colonne).HeaderText, 1))) = False Then Return False

                                'Sinon
                            Else

                                'Valuation paramètre et si erreur, retourne False
                                If Fonctions_CATIA.Modif_Valeur_Param(Object_A_Valuer, Parametre_CATIA_A_Valuer,
                                Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab(Tab_Des_Valeurs, Chaine_A_Trouver, Nom_Colonne_Chaine_A_Trouver,
                                DGV_A_Lire.Columns.Item(Boucle_Colonne).HeaderText, 1)) = False Then Return False
                            End If

                            'Sinon, se contente de la DGV
                        Else

                            'Si colonne de type CheckBox
                            If DGV_A_Lire.Columns.Item(Boucle_Colonne).GetType.ToString =
                                "System.Windows.Forms.DataGridViewCheckBoxColumn" Then

                                'Passage en minuscule de la valeur, valuation paramètre et si erreur, retourne False
                                If Fonctions_CATIA.Modif_Valeur_Param(Object_A_Valuer, Parametre_CATIA_A_Valuer,
                                LCase(DGV_A_Lire.CurrentRow.Cells.Item(Boucle_Colonne).Value)) = False Then Return False

                                'Sinon
                            Else

                                'Valuation paramètre et si erreur, retourne False
                                If Fonctions_CATIA.Modif_Valeur_Param(Object_A_Valuer, Parametre_CATIA_A_Valuer,
                                DGV_A_Lire.CurrentRow.Cells.Item(Boucle_Colonne).Value) = False Then Return False
                            End If
                        End If
                    End If
                End If
            Next

            'Retourne True
            Return True

            'Pour toutes les erreurs
        Catch

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(45, 5, , )
            Return False
        End Try
    End Function


    'Recherche nom de paramètre CATIA dans Tab et valuation variables correspondantes dans CATIA

    Public Function Modif_Variable_Selon_Tab _
        (ByVal Tab_Des_Valeurs(,) As String, ByVal Object_A_Valuer As Object,
         ByVal Num_Colonne_Chaine_A_Trouver As Integer,
         ByVal Num_Colonne_Nouvelle_Valeur As Integer) As Boolean

        'Gestion des exceptions
        Try

            'Déclaration et dimensionnement variable
            Dim Liste_Attributs_CATIA(Object_A_Valuer.NumberOfAttributes - 1, 0) As Object

            'Valuation variable
            Object_A_Valuer.GetListOfAttributes(Liste_Attributs_CATIA)

            'Boucle ligne Tab
            For Boucle_Ligne_Tab = LBound(Tab_Des_Valeurs, 1) To UBound(Tab_Des_Valeurs, 1)

                'Si paramètre présent dans l'objet à modifier
                If Fonctions_Tableau.Presence_Chaine_Dans_Tab _
                    (Liste_Attributs_CATIA, Tab_Des_Valeurs(Boucle_Ligne_Tab, Num_Colonne_Chaine_A_Trouver),
                     0, , ) = True Then

                    'Valuation paramètre et si erreur, retourne False
                    If Fonctions_CATIA.Modif_Valeur_Param _
                        (Object_A_Valuer, Tab_Des_Valeurs(Boucle_Ligne_Tab, Num_Colonne_Chaine_A_Trouver),
                         Tab_Des_Valeurs(Boucle_Ligne_Tab, Num_Colonne_Nouvelle_Valeur)) = False Then Return False
                End If
            Next

            'Retourne True
            Return True

            'Pour toutes les erreurs
        Catch

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(45, 5, , )
            Return False
        End Try
    End Function


    'Recherche nom de paramètre CATIA dans ToolTips de toutes colonnes DGV et valuation variables correspondantes dans CATIA

    Public Function Envoi_Donnees_Outil_Vers_DGV _
        (ByVal DGV_A_Remplir As DataGridView, ByVal Object_A_Lire As Object,
         ByVal Nom_Colonne_Debut As String, ByVal Nom_Colonne_Fin As String) As Boolean

        'Gestion des exceptions
        Try

            'Déclaration et dimensionnement variable
            Dim Liste_Attributs_CATIA(Object_A_Lire.NumberOfAttributes - 1, 0) As Object

            'Valuation variable
            Object_A_Lire.GetListOfAttributes(Liste_Attributs_CATIA)

            'Boucle colonne
            For Boucle_Colonne = Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Remplir, Nom_Colonne_Debut) To _
                Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_A_Remplir, Nom_Colonne_Fin)

                'Si num paramètre trouvé dans ToolTip colonne, si colonne visible et différente d'un type bouton
                If InStr(DGV_A_Remplir.Columns.Item(Boucle_Colonne).ToolTipText, "(11)") <> 0 And
                DGV_A_Remplir.Columns.Item(Boucle_Colonne).Visible = True And
                DGV_A_Remplir.Columns.Item(Boucle_Colonne).GetType.ToString <>
                "System.Windows.Forms.DataGridViewButtonColumn" Then

                    'Déclaration variable
                    Dim Parametre_CATIA_A_Lire As String

                    'Valuation variable
                    Parametre_CATIA_A_Lire = Fonctions_Diverses.Renvoi_Chaine_Encadree _
                        (DGV_A_Remplir.Columns.Item(Boucle_Colonne).ToolTipText, "(11) Paramètre CATIA : """, """")

                    'Si paramètre présent dans l'objet à modifier
                    If Fonctions_Tableau.Presence_Chaine_Dans_Tab _
                        (Liste_Attributs_CATIA, Parametre_CATIA_A_Lire, 0, , ) = True Then

                        'Si colonne de type CheckBox
                        If DGV_A_Remplir.Columns.Item(Boucle_Colonne).GetType.ToString =
                            "System.Windows.Forms.DataGridViewCheckBoxColumn" Then

                            'Passage en majuscule de la première lettre et valuation DGV
                            DGV_A_Remplir.CurrentRow.Cells.Item(Boucle_Colonne).Value =
                                UCase(Mid(Object_A_Lire.GetAttribute(Parametre_CATIA_A_Lire).Value, 1, 1)) &
                                Mid(Object_A_Lire.GetAttribute(Parametre_CATIA_A_Lire).ValueAsString, 2)

                            'Sinon
                        Else

                            'Valuation DGV
                            DGV_A_Remplir.CurrentRow.Cells.Item(Boucle_Colonne).Value =
                                Object_A_Lire.GetAttribute(Parametre_CATIA_A_Lire).ValueAsString
                        End If
                    End If
                End If
            Next

            'Retourne True
            Return True

            'Pour toutes les erreurs
        Catch

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(45, 5, , )
            Return False
        End Try
    End Function


    'Valuation des paramètres d'outil dans CATIA

    Public Function Valuation_Param_Outil _
        (ByRef Document_CATIA_Courant As Object, ByRef Prog_CATIA_Courant As Object,
         ByVal DGV_Outil_Principal As DataGridView, ByRef Outil_Selectionne As Object,
         ByRef Assemblage_Outil_Selectionne As Object, ByRef Plaq_Selectionnee As Object,
         ByVal Valeur_Bris_Outil As String, ByVal Inversion_Outil_Tournage As Boolean,
         ByVal Renumerotation_Outil As Boolean,
         ByVal Tab_Chaines() As String, Optional ByVal Tab_Config_Conversion_Outil(,) As String = Nothing,
         Optional ByVal DGV_Outil_Secondaire As DataGridView = Nothing) As Boolean

        'Gestion des exceptions
        Try

            'Variables
            Dim Compteur_Ressource As Integer = 1
            Dim Nom_Complet_Nouvel_Outil As String = Nothing

            'Recherche et renvoi compteur nom de ressource et si erreur, retourne False (avec passage en majuscule)
            If Fonctions_CATIA.Renvoi_Num_Ressource _
                (Document_CATIA_Courant, "T-" & UCase _
                 (DGV_Outil_Principal.CurrentRow.Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_Outil_Principal, "Outil")).Value) &
                 " *" & UCase(DGV_Outil_Principal.CurrentRow.Cells.Item _
                              (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_Outil_Principal, "Numéro enregistrement")).Value) &
                          "-", "*", Compteur_Ressource) = False Then Return False

            'Valuation variable (avec passage en majuscule des textes)
            Nom_Complet_Nouvel_Outil = UCase _
                (DGV_Outil_Principal.CurrentRow.Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_Outil_Principal, "Outil")).Value) &
                " *" & UCase(DGV_Outil_Principal.CurrentRow.Cells.Item _
                             (Fonctions_DGV.Renvoi_Num_Colonne_DGV(DGV_Outil_Principal, "Numéro enregistrement")).Value) &
                         "-" & Compteur_Ressource & "*"

            'Envoi des données outil dans CATIA et si erreur, retourne False
            If Fonctions_CATIA.Envoi_Donnees_Communes_Outil _
                (Document_CATIA_Courant, DGV_Outil_Principal, Outil_Selectionne,
                 Nom_Complet_Nouvel_Outil, "T") = False Then Return False

            'Si renumérotation d'outils checkée
            If Renumerotation_Outil = True Then

                'Renumérotation des outils et si erreur, sortie de sub
                If Fonctions_CATIA.Renum_Outil_Ou_Recherche_OP_CATIA(Prog_CATIA_Courant, 1, , ) =
                    False Then Return False
            End If

            'Si plaquette valuée, type d'outil tournage
            If Not Plaq_Selectionnee Is Nothing Then

                'Envoi des données assemblage dans CATIA et si erreur, retourne False
                If Fonctions_CATIA.Envoi_Donnees_Communes_Outil _
                    (Document_CATIA_Courant, DGV_Outil_Principal, Assemblage_Outil_Selectionne,
                     Nom_Complet_Nouvel_Outil, "A") = False Then Return False

                'Envoi des données plaquette dans CATIA et si erreur, retourne False
                If Fonctions_CATIA.Envoi_Donnees_Communes_Outil _
                    (Document_CATIA_Courant, DGV_Outil_Principal, Plaq_Selectionnee,
                     Nom_Complet_Nouvel_Outil, "I") = False Then Return False

                'Modification de paramètres d'inversion outil et si erreur, retourne False
                If Fonctions_CATIA.Modif_Valeur_Param _
                    (Assemblage_Outil_Selectionne, "MFG_TOOL_INVERT", LCase(Inversion_Outil_Tournage.ToString)) =
                    False Then Return False

                'Modification de paramètres d'inversion outil et si erreur, retourne False
                If Fonctions_CATIA.Modif_Valeur_Param _
                    (Outil_Selectionne, "MFG_HAND_STYLE", "RIGHT_HAND") =
                    False Then Return False

                'Sinon
            Else

                'Si valeur bris d'outil valué
                If Valeur_Bris_Outil <> Nothing Then

                    'Modification de paramètres d'outil et si erreur, retourne False
                    If Fonctions_CATIA.Modif_Valeur_Param _
                        (Outil_Selectionne, "MFG_WEIGHT_SNTX", "BRIS," & Valeur_Bris_Outil) =
                        False Then Return False
                End If
            End If

            'Si Tab de conversion et outil secondaire valué
            If Not Tab_Config_Conversion_Outil Is Nothing And Not DGV_Outil_Secondaire Is Nothing Then

                'Si plaquette de tournage valué
                If Not Plaq_Selectionnee Is Nothing Then

                    'Déclaration variable
                    Dim Tab_Valeurs_Temp(UBound(Tab_Config_Conversion_Outil, 1) - 1, 6) As String

                    'Envoi colonne Tab conversions vers Tab temporaire
                    Fonctions_Tableau.Envoi_Colonne_Tab_Vers_Tab _
                        (Tab_Config_Conversion_Outil, 1, Tab_Valeurs_Temp, 0, False)
                    Fonctions_Tableau.Envoi_Colonne_Tab_Vers_Tab _
                        (Tab_Config_Conversion_Outil, 2, Tab_Valeurs_Temp, 1, False)
                    Fonctions_Tableau.Envoi_Colonne_Tab_Vers_Tab _
                        (Tab_Config_Conversion_Outil, 3, Tab_Valeurs_Temp, 4, False)

                    'Recherche des attributs outil et plaquette de l'outil secondaire
                    Fonctions_DGV.Envoi_Element_ToolTip_DGV_Dans_Tab _
                        (DGV_Outil_Secondaire, "Style d'orientation",
                         "Largeur de plaquette (la)", Tab_Valeurs_Temp, 0, 3, "(11) Paramètre CATIA : """, """")
                    Fonctions_DGV.Envoi_Element_ToolTip_DGV_Dans_Tab _
                        (DGV_Outil_Secondaire, "Type",
                         "Profil de filetage", Tab_Valeurs_Temp, 1, 3, "(11) Paramètre CATIA : """, """")

                    'Extraction unité des valeurs
                    Fonctions_DGV.Envoi_Element_ToolTip_DGV_Dans_Tab _
                        (DGV_Outil_Secondaire, "Style d'orientation",
                         "Largeur de plaquette (la)", Tab_Valeurs_Temp, 0, 2, "(10) Unités de la valeur : """, """")
                    Fonctions_DGV.Envoi_Element_ToolTip_DGV_Dans_Tab _
                        (DGV_Outil_Secondaire, "Type",
                         "Profil de filetage", Tab_Valeurs_Temp, 1, 2, "(10) Unités de la valeur : """, """")

                    'Ecriture formule avec remplacement des paramètres par valeur DGV
                    Fonctions_DGV.Remplacement_Param_Par_Valeur_Dans_Chaine _
                        (DGV_Outil_Principal, "Outil à bout sphérique",
                         "Nombre de lèvres", Tab_Valeurs_Temp, 4, 5, "'")

                    'Calcul de la formule et si erreur, retourne False
                    If Fonctions_Diverses.Calcule_Tab_Chaines_Completes_Vers_Tab _
                        (Tab_Valeurs_Temp, 5, 6, 2) = False Then Return False

                    'Envoi des données propres à l'outil sélectionné et si erreur, retourne False
                    If Fonctions_CATIA.Modif_Variable_Selon_Tab _
                        (Tab_Valeurs_Temp, Outil_Selectionne, 3, 6) = False Then Return False

                    'Envoi des données propres à l'outil sélectionné et si erreur, retourne False
                    If Fonctions_CATIA.Modif_Variable_Selon_Tab _
                        (Tab_Valeurs_Temp, Plaq_Selectionnee, 3, 6) = False Then Return False

                    'Sinon
                Else

                    'Déclaration variable
                    Dim Tab_Valeurs_Temp(UBound(Tab_Config_Conversion_Outil, 1) - 1, 5) As String

                    'Envoi colonne Tab conversions vers Tab temporaire et si erreur, retourne False
                    Fonctions_Tableau.Envoi_Colonne_Tab_Vers_Tab _
                        (Tab_Config_Conversion_Outil, 1, Tab_Valeurs_Temp, 0, False)
                    Fonctions_Tableau.Envoi_Colonne_Tab_Vers_Tab _
                        (Tab_Config_Conversion_Outil, 3, Tab_Valeurs_Temp, 3, False)

                    'Recherche des attributs outil et plaquette de l'outil secondaire
                    Fonctions_DGV.Envoi_Element_ToolTip_DGV_Dans_Tab _
                        (DGV_Outil_Secondaire, "Outil à bout sphérique",
                         "Nombre de lèvres", Tab_Valeurs_Temp, 0, 2, "(11) Paramètre CATIA : """, """")

                    'Extraction unité des valeurs
                    Fonctions_DGV.Envoi_Element_ToolTip_DGV_Dans_Tab _
                        (DGV_Outil_Secondaire, "Outil à bout sphérique",
                         "Nombre de lèvres", Tab_Valeurs_Temp, 0, 1, "(10) Unités de la valeur : """, """")

                    'Ecriture formule avec remplacement des paramètres par valeur DGV
                    Fonctions_DGV.Remplacement_Param_Par_Valeur_Dans_Chaine _
                        (DGV_Outil_Principal, "Style d'orientation",
                         "Profil de filetage", Tab_Valeurs_Temp, 3, 4, "'")

                    'Calcul de la formule et si erreur, retourne False et si erreur, retourne False
                    If Fonctions_Diverses.Calcule_Tab_Chaines_Completes_Vers_Tab _
                        (Tab_Valeurs_Temp, 4, 5, 1) = False Then Return False

                    'Envoi des données propres à l'outil sélectionné et si erreur, retourne False
                    If Fonctions_CATIA.Modif_Variable_Selon_Tab _
                        (Tab_Valeurs_Temp, Outil_Selectionne, 2, 5) = False Then Return False
                End If

                'Sinon
            Else

                'Envoi des données propres à l'outil sélectionné et si erreur, retourne False
                If Fonctions_CATIA.Modif_Variable_Selon_ToolTips _
                    (DGV_Outil_Principal, Outil_Selectionne, Tab_Chaines(5), Tab_Chaines(6), , , ) = False Then Return False

                'Si plaquette de tournage valué
                If Not Plaq_Selectionnee Is Nothing Then

                    'Envoi des données propres à la plaquette sélectionnée et si erreur, retourne False
                    If Fonctions_CATIA.Modif_Variable_Selon_ToolTips _
                        (DGV_Outil_Principal, Plaq_Selectionnee, Tab_Chaines(10), Tab_Chaines(11), , , ) = False Then Return False
                End If
            End If

            'Retourne True
            Return True

            'Pour toutes les erreurs
        Catch

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(45, 5, , )
            Return False
        End Try
    End Function


    'Renvoi un numéro de ressource non-utilisé

    Public Function Renvoi_Num_Ressource _
        (ByVal Document_CATIA_Courant As Object, ByRef Debut_Nom_Ressource As String,
         ByRef Fin_Nom_Ressource As String, ByRef Compteur_Ressource As Integer) As Boolean

        'Gestion des exceptions
        Try

            'Variables
            Dim Document_PPR, Ressources_PPR As Object
            Dim Nb_Ressources As Integer

            'Valuation variables
            Document_PPR = Document_CATIA_Courant.PPRDocument
            Ressources_PPR = Document_PPR.Resources
            Nb_Ressources = Ressources_PPR.Count

            'Boucle lecture ressources pour comptage nom d'outil identique
            For Boucle = 1 To Nb_Ressources

                'Si nom outil = nom ressource, +1 num ressource
                If Debut_Nom_Ressource & Compteur_Ressource & Fin_Nom_Ressource =
                    Ressources_PPR.Item(Boucle).PartNumber Then

                    'Incrémentation compteur ressource et relecture liste des ressources pour vérif
                    Compteur_Ressource = Compteur_Ressource + 1
                    Boucle = 1
                End If
            Next

            'Retourne True
            Return True

            'Pour toutes les erreurs
        Catch

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(45, 5, , )
            Return False
        End Try
    End Function


    'Renommage d'une ressource dans CATIA

    Public Function Renommage_Ressource_CATIA _
        (ByRef Document_CATIA_Courant As Object, ByVal Ancien_Nom As String,
         ByVal Nouveau_Nom As String) As Boolean

        'Gestion des exceptions
        Try

            'Variables
            Dim Document_PPR As Object
            Dim Ressources_PPR As Object
            Dim Nb_Ressources As Integer

            'Valuation variables
            Document_PPR = Document_CATIA_Courant.PPRDocument
            Ressources_PPR = Document_PPR.Resources
            Nb_Ressources = Ressources_PPR.Count

            'Boucle lecture
            For Boucle = 1 To Nb_Ressources

                'Si nom outil = nom ressource
                If InStr(Ressources_PPR.Item(Boucle).PartNumber, Ancien_Nom) <> 0 Then

                    'Valuation variable
                    Ressources_PPR.Item(Boucle).PartNumber = Nouveau_Nom.ToString

                    'Sortie de boucle
                    Exit For
                End If
            Next

            'Retourne True
            Return True

            'Pour toutes les erreurs
        Catch

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(45, 5, , )
            Return False
        End Try
    End Function


    'Recherche phase d'après le programme sélectionné dans CATIA

    Public Function Recherche_Phase_Courante_Process _
        (ByVal Process_Courant As Object, ByVal Prog_Courant As Object,
         ByRef Phase_Courante As Object) As Boolean

        'Gestion des exceptions
        Try

            'Si phases présentes dans arbre Process
            If Process_Courant.IsSubTypeOf("PhysicalActivity") Then

                'Variables
                Dim Activites_Process As Object
                Dim Compteur_Activites As Integer
                Dim Activite_Temp As Object
                Dim Compteur_Phase As Integer = 1

                'Rechercher si phases présentes et les compter
                Activites_Process = Process_Courant.ChildrenActivities
                Compteur_Activites = Activites_Process.Count

                'Boucle lecture des enfants
                For Boucle_Activites = 1 To Compteur_Activites

                    'Valuation enfant actif
                    Activite_Temp = Activites_Process.Item(Boucle_Activites)

                    'Si phase
                    If Activite_Temp.IsSubTypeOf("ManufacturingSetup") Then

                        'Variables
                        Dim Compteur_Prog As Integer
                        Dim Liste_Prog As Object

                        'Valuation des variables, extraction des programmes et les compter
                        Liste_Prog = Activite_Temp.Programs
                        Compteur_Prog = Liste_Prog.Count

                        'Si programmes présents
                        If Compteur_Prog > 0 Then

                            'Boucle lecture des programmes
                            For Boucle_Prog = 1 To Compteur_Prog

                                'Si nom de programe = nom du programme sélectionné
                                If Prog_Courant.Name = Liste_Prog.GetElement(Boucle_Prog).Name Then

                                    'Valuation de la phase courante
                                    Phase_Courante = Activite_Temp

                                    'Retourne True
                                    Return True
                                End If
                            Next
                        End If

                        'Compteur phase
                        Compteur_Phase = Compteur_Phase + 1
                    End If
                Next

            End If

            'Retourne True
            Return True

            'Pour toutes les erreurs
        Catch

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(45, 5, , )
            Return False
        End Try
    End Function


    'Modification paramètre

    Public Function Modif_Valeur_Param _
        (ByVal Objet_A_Modifier As Object, ByVal Param_A_Modifier As String,
         ByVal Nouvelle_Valeur_Param As String) As Boolean

        'Gestion des exceptions
        Try

            'Modification du paramètre
            Objet_A_Modifier.GetAttribute(Param_A_Modifier).ValuateFromString(Nouvelle_Valeur_Param)

            'Retourne True
            Return True

            'Pour toutes les erreurs
        Catch

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(45, 5, , )
            Return False
        End Try
    End Function


    'Contrôle CATIA ouvert et CATProcess actif et valuation des variables

    Public Function Ctrl_Recherche_Donnees_CATIA _
        (ByRef CATIA_App As Object, ByRef Doc_CATIA_Courant As Object,
         ByRef Process_Courant As Object) As Boolean

        'Gestion des exceptions
        Try

            'Extraction objet
            CATIA_App = GetObject(, "CATIA.Application")

            'Si pas de CATProcess trouvé, retourne False
            If TypeName(CATIA_App.ActiveDocument) <> "ProcessDocument" Then

                'Message et renvoi False
                Appel_Msg(17, 5, , )
                Return False
            End If

            'Valuation variables
            Doc_CATIA_Courant = CATIA_App.ActiveDocument
            Process_Courant = Doc_CATIA_Courant.GetItem("Process")

            'Retourne True
            Return True

            'Pour toutes les erreurs
        Catch

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(45, 5, , )
            Return False
        End Try
    End Function


    'Demande de sélection d'un élément dans CATIA

    Public Function Selection_Element_CATIA _
        (ByVal Doc_CATIA_Courant As Object, ByVal Element_A_Selectionner As String,
         ByVal Msg_Demande_Select As String, ByRef Selection_CATIA As Object) As Boolean

        'Variables
        Dim Element_Select As INFITF.Selection
        Dim Entree_Object_Type(0) As Object
        Dim Status As String

        'Valuation sélection
        Element_Select = Doc_CATIA_Courant.Selection
        Entree_Object_Type(0) = Element_A_Selectionner

        'Demande de sélection
        Status = Element_Select.SelectElement2 _
            (Entree_Object_Type, Msg_Demande_Select, True)

        'Contrôle de la sélection
        If Status = "Cancel" Then

            'Retourne False
            Return False
        End If

        'Si nombre d'éléments sélectionné > 1
        If Element_Select.Count > 1 Then

            'Message d'information
            Fonctions_Messages.Appel_Msg(53, 1, , )
        End If

        'Valuation sélection
        Selection_CATIA = Element_Select.Item(1).Value

        'Vidage sélection
        Element_Select.Clear()

        'Retourne True
        Return True
    End Function

    'Contrôle la sélection de changement d'outil dans CATIA et contrôle du type d'outil si OK

    Public Function Controle_Selection_ToolChange _
        (ByVal Activite_Selectionnee As Object, ByRef Resultat_Test_Type As Boolean, ByVal Type_Outil_Principal As String,
         ByVal Type_Plaq_Principal As String, Optional ByVal Type_Outil_Secondaire As String = Nothing,
         Optional ByVal Type_Plaq_Secondaire As String = Nothing,
         Optional ByRef Conversion_Outil As Boolean = False) As Boolean

        'Valeur de condition à trouver
        Select Case Activite_Selectionnee.Type

            'Si le type de l'activité sélectionnée est un changement d'outil de fraisage
            Case "ToolChange"

                'Variable
                Dim Tab_Temp_Config(,) As String = Nothing
                Dim Tool_Type_CATIA As String = Nothing

                'Lecture CSV et si erreur, retourne faux
                If Fonctions_CSV.Lecture_CSV(Dossier_Reseau & "\Config_Outils\Fichier_Config_Outils_Fraisage_Stds.csv",
                                             Tab_Temp_Config, 5) = False Then Return False

                'Valuation variable
                Tool_Type_CATIA = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab _
                  (Tab_Temp_Config, Type_Outil_Principal, "Type outil standard", "Type outil CATIA", 1)

                'Si le type de l'outil sélectionné correspond au type principal ou secondaire
                If Activite_Selectionnee.Tool.ToolType = Tool_Type_CATIA Then

                    'Valuation variable et retourne True
                    Resultat_Test_Type = True
                    Return True

                    'Sinon
                Else

                    'Si outil secondaire valué
                    If Type_Outil_Secondaire <> Nothing Then

                        'Si outil secondaire du bon type
                        If Activite_Selectionnee.Tool.GetAttributeNLSName _
                            (Activite_Selectionnee.Tool.ToolType) = Type_Outil_Secondaire Then

                            'Valuation variables et retourne True
                            Conversion_Outil = True
                            Resultat_Test_Type = True
                            Return True

                            'Sinon
                        Else

                            'Message d'erreur de sélection
                            Fonctions_CATIA.Message_Erreur_Selection_ToolChange _
                                (Type_Outil_Principal, Type_Plaq_Principal,
                                 Type_Outil_Secondaire, Type_Plaq_Secondaire)

                            'Valuation variable et retourne False
                            Resultat_Test_Type = False
                            Return True
                        End If
                    End If

                    'Message d'erreur de sélection
                    Fonctions_CATIA.Message_Erreur_Selection_ToolChange _
                        (Type_Outil_Principal, Type_Plaq_Principal,
                         Type_Outil_Secondaire, Type_Plaq_Secondaire)

                    'Valuation variable et retourne False
                    Resultat_Test_Type = False
                    Return True
                End If

                'Si le type de l'activité sélectionnée est un changement d'outil de tournage
            Case "ToolChangeLathe"

                'Si le type de l'outil et de la plaquette sélectionnés correspond au type principal
                If Activite_Selectionnee.Tool.GetAttributeNLSName _
                    (Activite_Selectionnee.Tool.ToolType) = Type_Outil_Principal And
                    Activite_Selectionnee.ToolAssembly.Insert.GetAttributeNLSName _
                    (Activite_Selectionnee.ToolAssembly.Insert.InsertType) = Type_Plaq_Principal Then

                    'Valuation variable et retourne true
                    Resultat_Test_Type = True
                    Return True

                    'Sinon
                Else

                    'Si outil secondaire valué
                    If Type_Outil_Secondaire <> Nothing Then

                        'Si outil et plaquette secondaire du bon type
                        If Activite_Selectionnee.Tool.GetAttributeNLSName _
                            (Activite_Selectionnee.Tool.ToolType) = Type_Outil_Secondaire And
                            Activite_Selectionnee.ToolAssembly.Insert.GetAttributeNLSName _
                            (Activite_Selectionnee.ToolAssembly.Insert.InsertType) = Type_Plaq_Secondaire Then

                            'Valuation variables et retourne true
                            Conversion_Outil = True
                            Resultat_Test_Type = True
                            Return True

                            'Sinon
                        Else

                            'Message d'erreur de sélection
                            Fonctions_CATIA.Message_Erreur_Selection_ToolChange _
                                (Type_Outil_Principal, Type_Plaq_Principal,
                                 Type_Outil_Secondaire, Type_Plaq_Secondaire)

                            'Valuation variable et retourne False
                            Resultat_Test_Type = False
                            Return True
                        End If
                    End If

                    'Message d'erreur de sélection
                    Fonctions_CATIA.Message_Erreur_Selection_ToolChange _
                        (Type_Outil_Principal, Type_Plaq_Principal,
                         Type_Outil_Secondaire, Type_Plaq_Secondaire)

                    'Valuation variable et retourne False
                    Resultat_Test_Type = False
                    Return True
                End If

                'Sinon
            Case Else

                'Message d'information
                Fonctions_Messages.Appel_Msg(54, 3, , )

                'Valuation variable et retourne False
                Resultat_Test_Type = False
                Return True
        End Select
    End Function

    'Construction du message d'erreur de sélection des changements d'outil dans CATIA

    Public Sub Message_Erreur_Selection_ToolChange _
        (ByVal Type_Outil_Principal As String, ByVal Type_Plaq_Principal As String,
         ByVal Type_Outil_Secondaire As String, ByVal Type_Plaq_Secondaire As String)

        'Variable message
        Dim Msg() As String

        'Si outil secondaire est valuée
        If Type_Outil_Secondaire <> Nothing Then

            'Redimensionnement Tab message
            ReDim Msg(2)

            'Si le type de plaquette principal est valué
            If Type_Plaq_Principal = Nothing Then

                'Valuation variable message
                Msg(0) = Type_Outil_Principal
                Msg(1) = Type_Outil_Secondaire
                Msg(2) = Type_Plaq_Secondaire

                'Message
                Fonctions_Messages.Appel_Msg(57, 3, Msg, )

                'Sinon
            Else

                'Valuation variable message
                Msg(0) = Type_Outil_Secondaire
                Msg(1) = Type_Outil_Principal
                Msg(2) = Type_Plaq_Principal

                'Message
                Fonctions_Messages.Appel_Msg(57, 3, Msg, )
            End If


            'Sinon
        Else

            'Si le type de plaquette principal est valué
            If Type_Plaq_Principal = Nothing Then

                'Message
                Fonctions_Messages.Appel_Msg(55, 3, , Type_Outil_Principal)

                'Sinon
            Else

                'Redimensionnement Tab message
                ReDim Msg(1)

                'Valuation variable message
                Msg(0) = Type_Outil_Principal
                Msg(1) = Type_Plaq_Principal

                'Message
                Fonctions_Messages.Appel_Msg(56, 3, Msg, )
            End If
        End If
    End Sub

    'Extrait un tableau des temps de cycle d'un programme

    Public Function Extract_Tab_Tps_Cycle_Prog _
        (ByVal Prog_Courant As Object, ByVal Tps_ToolChange As Integer,
         ByRef Tab_Tps_OP(,) As String) As Boolean

        'Gestion des exceptions
        Try

            'Déclaration + valuation variables
            Dim Liste_Operations As Object
            Dim Nb_Operations As Integer

            Liste_Operations = Prog_Courant.Activities
            Nb_Operations = Liste_Operations.Count

            'Si opérations présentes...
            If Nb_Operations > 0 Then

                'Déclaration + valuation variables
                Dim Tab_Num_OP() As String = Nothing
                Dim Inc_Tab_Tps_OP As Integer
                Inc_Tab_Tps_OP = -1

                'Boucle OP
                For i = 1 To Nb_Operations

                    'Si OP active
                    If Liste_Operations.GetElement(i).Active = True Then

                        'Si OP avec temps de cycle...
                        If Liste_Operations.GetElement(i).Type <> "PPInstruction" And
                            Liste_Operations.GetElement(i).Type <> "CoordinateSystem" Then

                            'Incrément pour dimenssionnement Tab
                            Inc_Tab_Tps_OP += 1

                            'Redimenssionnement Tab
                            ReDim Preserve Tab_Num_OP(Inc_Tab_Tps_OP)

                            'Valuation numéro OP
                            Tab_Num_OP(Inc_Tab_Tps_OP) = i
                        End If
                    End If
                Next

                'Déclaration variables
                Dim OP_Courante As Object

                'Redimenssionnement Tab
                ReDim Tab_Tps_OP(UBound(Tab_Num_OP), 2)

                'Boucle OP
                For i = 0 To UBound(Tab_Num_OP)

                    'Extraction OP de liste OP
                    OP_Courante = Liste_Operations.GetElement(Tab_Num_OP(i))

                    'Valuation variables
                    Tab_Tps_OP(i, 0) = OP_Courante.Type
                    Tab_Tps_OP(i, 1) = OP_Courante.Name

                    'Si ToolChange, valuation suivant temps donné dans Form
                    If OP_Courante.Type = "ToolChange" Or OP_Courante.Type = "ToolChangeLathe" Then

                        'Valuation variable
                        Tab_Tps_OP(i, 2) = Tps_ToolChange

                        'Sinon, extraction temps OP
                    Else

                        'Valuation variable
                        Tab_Tps_OP(i, 2) = OP_Courante.TotalTime * 60
                    End If
                Next

                'Message
                Fonctions_Messages.Appel_Msg(26, 1, , Prog_Courant.Name)

                'Sinon
            Else

                'Message et renvoi False
                Fonctions_Messages.Appel_Msg(3, 5, , Prog_Courant.Name)
                Return False
            End If

            'Retourne True
            Return True

            'Pour l'erreur sur une OP pas à jour
        Catch

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(46, 5, , Prog_Courant.Name)
            Return False
        End Try
    End Function

    'Fonction convertion temps de cycle en secondes vers heures/minutes/seceondes

    Public Function Conv_Tps_Cycle(ByVal Tps_Cycle_Sec As Double) As String

        'Déclaration variables temporaires
        Dim Temps_Temp, H_Temp, Min_Temp, Sec_Temp As Double

        'Extraction des heures
        Temps_Temp = Tps_Cycle_Sec / 3600
        H_Temp = Int(Temps_Temp)

        'Extraction des minutes
        Temps_Temp = (Temps_Temp - H_Temp) * 60
        Min_Temp = Int(Temps_Temp)

        'Extraction des secondes
        Temps_Temp = (Temps_Temp - Min_Temp) * 60
        Sec_Temp = Int(Temps_Temp)

        'Si heure, minute, et seconde dispo
        If H_Temp > 0 And Min_Temp > 0 And Sec_Temp > 0 Then

            'Renvoi valeur
            Return H_Temp & " heure(s), " & Min_Temp & " minute(s) et " & Sec_Temp & " seconde(s)"

            'Si minute, et seconde dispo
        ElseIf Min_Temp > 0 And Sec_Temp > 0 Then

            'Renvoi valeur
            Return Min_Temp & " minute(s) et " & Sec_Temp & " seconde(s)"

            'Si seconde dispo
        ElseIf Sec_Temp > 0 Then

            'Renvoi valeur
            Return Sec_Temp & " seconde(s)"

            'Sinon...
        Else

            'Renvoi valeur
            Return "Pas de durée disponible."
        End If
    End Function

    'Remplissage barre des temps de cycle

    Public Function Remp_Barre_Temps(ByVal Tab_Temps(,) As String, ByVal PictureBox_Form As PictureBox,
                                     ByVal Couleur_OP() As Color, ByRef Coeff_Image_Tps As Double) As Bitmap

        'Variable
        Dim Total_Tps As Double = Nothing

        'Boucle cumul des temps
        For i = 0 To UBound(Tab_Temps)

            'Cumul des temps
            Total_Tps = Total_Tps + Tab_Temps(i, 2)
        Next

        'Variables
        Coeff_Image_Tps = PictureBox_Form.Size.Width / Total_Tps
        Dim Image_Temp As New Bitmap(PictureBox_Form.Size.Width, PictureBox_Form.Size.Height, Imaging.PixelFormat.Format32bppArgb)
        Dim Inc_Pixel_Hor As Integer = 0

        'Boucle lecture Tab
        For i = 0 To UBound(Tab_Temps, 1)

            'Boucle coloration horizontale
            For j = 1 To Int(Tab_Temps(i, 2) * Coeff_Image_Tps)

                'Boucle coloration verticale
                For k = 0 To PictureBox_Form.Size.Height - 1

                    'Si changement d'outil
                    If Tab_Temps(i, 0) = "ToolChange" Or Tab_Temps(i, 0) = "ToolChangeLathe" Then

                        'Pixel ToolChange
                        Image_Temp.SetPixel(Inc_Pixel_Hor, k, Color.FromArgb(255, Couleur_OP(0)))

                        'Sinon...
                    Else

                        'Pixel OP
                        Image_Temp.SetPixel(Inc_Pixel_Hor, k, Color.FromArgb(255, Couleur_OP(1)))
                    End If
                Next

                'Incrément pixel
                Inc_Pixel_Hor += 1
            Next
        Next

        'Retourne l'image finale
        Return Image_Temp
    End Function




    'Structure dimensions cellule
    Structure DimCellXLS
        Dim Larg, Haut, posX, posY As Integer
    End Structure

    'Fonction dimensions cellule
    Function RecupDimPosCell(plagCell As Range) As DimCellXLS

        'Récup dim + pos plage cellule
        RecupDimPosCell.Larg = plagCell.Width
        RecupDimPosCell.Haut = plagCell.Height
        RecupDimPosCell.posX = plagCell.Left
        RecupDimPosCell.posY = plagCell.Top
    End Function
End Module
