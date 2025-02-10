Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Tab

Public Module Fonctions_CSV

    Public Function Mise_En_Forme_Tab_Pour_Ecriture_CSV _
        (ByVal Tab_A_Lire(,) As String) As String

        'Valuation variable
        Mise_En_Forme_Tab_Pour_Ecriture_CSV = Nothing

        'Boucle lecture lignes
        For Boucle_Ligne = LBound(Tab_A_Lire, 1) To UBound(Tab_A_Lire, 1)

            'Boucle lecture colonnes
            For Boucle_Colonne = LBound(Tab_A_Lire, 2) To UBound(Tab_A_Lire, 2)

                'Si différent du dernier passage dans la boucle
                If Boucle_Colonne <> UBound(Tab_A_Lire, 2) Then

                    'Mise en forme
                    Mise_En_Forme_Tab_Pour_Ecriture_CSV = Mise_En_Forme_Tab_Pour_Ecriture_CSV &
                        Tab_A_Lire(Boucle_Ligne, Boucle_Colonne) & ";"

                    'sinon
                Else

                    'Mise en forme dernière chaîne
                    Mise_En_Forme_Tab_Pour_Ecriture_CSV = Mise_En_Forme_Tab_Pour_Ecriture_CSV &
                        Tab_A_Lire(Boucle_Ligne, Boucle_Colonne) & vbCrLf
                End If
            Next
        Next
    End Function

    'Lecture CSV et création tableau de stockage

    Public Function Lecture_CSV _
        (ByVal Fichier_A_Lire As String, ByRef Tab_A_Remplir As String(,),
         ByVal Type_MsgBox As Integer) As Boolean

        'Gestion des exceptions
        Try

            'Variables
            Dim Split_Ligne_Lue() As String
            Dim Dimension_Tab(1) As Integer
            Dim Compteur As Integer = 0

            'Si erreur lors de la valuation, renvoi False pour Exit Sub
            If Fonctions_CSV.Dim_Tab_Depuis_CSV(Fichier_A_Lire, Dimension_Tab) = False Then Return False

            'Redimenssionnment Tab
            ReDim Split_Ligne_Lue(Dimension_Tab(1))
            ReDim Tab_A_Remplir(Dimension_Tab(0), Dimension_Tab(1))

            'Boucle lecture de toutes les lignes du Listing
            For Each Boucle_Ligne_Lue As String In System.IO.File.ReadLines(Fichier_A_Lire, System.Text.Encoding.Default)

                'Split de la ligne lue
                Split_Ligne_Lue = Split(Boucle_Ligne_Lue, ";", UBound(Split_Ligne_Lue, 1) + 1)

                'Boucle remplissage tab
                For Boucle_Remp = LBound(Split_Ligne_Lue, 1) To UBound(Split_Ligne_Lue, 1)

                    'Remplissage tab
                    Tab_A_Remplir(Compteur, Boucle_Remp) = Split_Ligne_Lue(Boucle_Remp)
                Next

                'Incrémentation compteur
                Compteur = Compteur + 1
            Next

            'Retourne True
            Return True

            'Si chemin de dossier inexistant
        Catch Excep As DirectoryNotFoundException

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(0, Type_MsgBox, , Excep.Message)
            Return False

            'Si chemin de fichier inexistant
        Catch Excep As FileNotFoundException

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(0, Type_MsgBox, , Excep.Message)
            Return False

            'Si chemin de fichier inexistant = 0
        Catch Excep As ArgumentException

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(0, Type_MsgBox, , Excep.Message)
            Return False

            'Si chemin de fichier trop long
        Catch Excep As PathTooLongException

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(0, Type_MsgBox, , Excep.Message)
            Return False

            'Pour toutes les erreurs
        Catch

            'Message et retourne False
            Fonctions_Messages.Appel_Msg(5, Type_MsgBox, , )
            Return False
        End Try
    End Function

    'Dimensionnement Tab depuis CSV

    Public Function Dim_Tab_Depuis_CSV _
        (ByVal Lien_Fichier As String, ByRef Dimension_Tab() As Integer) As Boolean

        'Gestion des exceptions
        Try

            'Variables
            Dim Split_Ligne_Temp() As String
            Dimension_Tab(0) = 0

            'Boucle pour dimenssionnement tab
            For Each Boucle As String In System.IO.File.ReadLines(Lien_Fichier, System.Text.Encoding.Default)

                'Si 1er passage dans boucle
                If Dimension_Tab(0) = 0 Then

                    'Split de la 1er ligne
                    Split_Ligne_Temp = Split(Boucle, ";", , )

                    'Valuation variable
                    Dimension_Tab(1) = UBound(Split_Ligne_Temp, 1)
                End If

                'Valuation variable
                Dimension_Tab(0) = Dimension_Tab(0) + 1
            Next

            'Enlever 1 ligne à la fin dû à la boucle de lecture
            Dimension_Tab(0) = Dimension_Tab(0) - 1

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

    'Lis toutes les lignes d'un CSV et renvoi le nombre maximum trouvé d'une colonne

    Public Function Renvoi_Valeur_Maxi_CSV _
        (ByVal Fichier_A_Lire As String, ByVal Nom_Colonne_A_Lire As String,
         ByRef Valeur_Maxi_Trouvee As Integer, ByVal Type_MsgBox As Integer) As Boolean

        'Variable
        Dim Tab_Temp As String(,) = Nothing
        Valeur_Maxi_Trouvee = 0

        'Lecture CSV et si erreur, retourne False
        If Fonctions_CSV.Lecture_CSV(Fichier_A_Lire, Tab_Temp, Type_MsgBox) = False Then Return False

        'Boucle lecture lignes
        For Boucle_Ligne = LBound(Tab_Temp, 1) + 1 To UBound(Tab_Temp, 1)

            'Si trouve valeur supérieur à la dernière lue
            If Tab_Temp(Boucle_Ligne, Fonctions_Tableau.Renvoi_Num_Colonne_Tab(Tab_Temp, Nom_Colonne_A_Lire)) > Valeur_Maxi_Trouvee Then _
                Valeur_Maxi_Trouvee = Tab_Temp(Boucle_Ligne, Fonctions_Tableau.Renvoi_Num_Colonne_Tab(Tab_Temp, Nom_Colonne_A_Lire))
        Next

        'Retourne True
        Return True
    End Function

    'Ecriture résultat CheckBox
    Public Function Ecriture_Result_CheckBox _
        (ByVal DGV_A_Lire As DataGridView, ByVal Num_Colonne As Integer) As String

        'Si case checkée
        If DGV_A_Lire.CurrentRow.Cells.Item(Num_Colonne).Value = True Then

            'Valuation variable
            Ecriture_Result_CheckBox = "True"

            'Sinon
        Else

            'Valuation variable
            Ecriture_Result_CheckBox = "False"
        End If
    End Function


    'Lecture CSV complet et écriture entête & autres dans autre CSV
    Public Function recupEnteteCSV() As Boolean

        'Lecture CSV
        Dim tabCSV(,), strTmp As String
        If Fonctions_CSV.Lecture_CSV(Dossier_Reseau & "\Config_FO\Fichier_Config_FO.csv", tabCSV, 5) = False Then Return False

        For i = LBound(tabCSV, 2) To UBound(tabCSV, 2)

            'Ecriture string
            strTmp = strTmp & tabCSV(0, i) & vbCrLf
        Next

        'Ecriture string dans CSV
        System.IO.File.WriteAllText("C:\Users\Guitou\Documents\CATCNCManager\Ressources_Dev\Config_FO\ListParamFO.csv", strTmp, System.Text.Encoding.Default)
        Return True
    End Function
End Module
