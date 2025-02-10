Public Class Form_CATIA_Vers_Listing_Outils

    'Variables
    Private Doc_CATIA_Courant, Process_Courant, Assembly_Outil_Courant, Outil_Courant, Plaquette_Courante As Object
    Private DGV_Maitresse_Local As DataGridView
    Private Tab_Chaine_Local(), Nom_Listing_Outil, Type_Outil_Principal, _
        Type_Plaq_Principal As String

    'Action sur création nouvelle Form

    Public Sub New(ByVal DGV_Maitresse As DataGridView, ByVal Tab_Chaines() As String)

        'Initialisation composant
        InitializeComponent()

        'Valuation variables
        DGV_Maitresse_Local = DGV_Maitresse
        Tab_Chaine_Local = Tab_Chaines
        Type_Outil_Principal = Tab_Chaine_Local(4)

        'Si plaquette valuée
        If UBound(Tab_Chaine_Local, 1) > 6 Then

            Type_Plaq_Principal = Tab_Chaine_Local(9)
        End If

        'Variable
        Dim CATIA_App As Object = Nothing

        'Si pas de CATIA ouvert et CATProcess inactif, sortie de Sub
        If Fonctions_CATIA.Ctrl_Recherche_Donnees_CATIA(CATIA_App, Doc_CATIA_Courant, Process_Courant) = _
            False Then Exit Sub
    End Sub

    'Action sur fermeture Form

    Private Sub Form_Listing_Outils_Vers_CATIA_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) _
        Handles MyBase.FormClosed

        'Fermeture Form
        Me.Dispose()
    End Sub

    'Click bouton sélection dans CATIA

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'Variables
        Dim Activite_Select As Object = Nothing
        Dim Result_Test_Type As Boolean

        'Début boucle
        Do

            'Demande de sélection dans CATIA
            If Fonctions_CATIA.Selection_Element_CATIA _
                (Doc_CATIA_Courant, "ManufacturingActivity", _
                 "Sélectionner un changement d'outil parmi les activités de l'arbre", _
                 Activite_Select) = False Then Exit Sub

            'Contrôle type de la sélection et retourne faux si problème de lecture fichier CSV
            If Fonctions_CATIA.Controle_Selection_ToolChange _
                (Activite_Select, Result_Test_Type, Type_Outil_Principal, Type_Plaq_Principal, , , ) = False Then Exit Sub

            'Fin de boucle et test résultat
        Loop While Result_Test_Type = False

        'Si outil fraisage
        If Activite_Select.Type = "ToolChange" Then

            'Valuation variables
            Outil_Courant = Activite_Select.Tool

            'Si outil tournage
        ElseIf Activite_Select.Type = "ToolChangeLathe" Then

            'Valuation variables
            Assembly_Outil_Courant = Activite_Select.ToolAssembly
            Outil_Courant = Activite_Select.Tool
            Plaquette_Courante = Assembly_Outil_Courant.Insert
        End If

        'Remplissage TextBox et activation bouton d'envoi
        Me.TextBox1.Text = Outil_Courant.Name
        Me.Button2.Enabled = True
    End Sub

    'Click bouton envoyer dans Listing outils

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        'Si outil sélectionné est de type tournage
        If Not Plaquette_Courante Is Nothing Then

            'Valuation des paramètres outils et si erreur, sortie de sub
            If Fonctions_CATIA.Envoi_Donnees_Outil_Vers_DGV _
                    (DGV_Maitresse_Local, Outil_Courant, Tab_Chaine_Local(5), Tab_Chaine_Local(6)) = _
                 False Then Exit Sub

            'Valuation des paramètres outils et si erreur, sortie de sub
            If Fonctions_CATIA.Envoi_Donnees_Outil_Vers_DGV _
                    (DGV_Maitresse_Local, Plaquette_Courante, Tab_Chaine_Local(10), Tab_Chaine_Local(11)) = _
                 False Then Exit Sub

            'Sinon
        Else

            'Valuation des paramètres outils et si erreur, sortie de sub
            If Fonctions_CATIA.Envoi_Donnees_Outil_Vers_DGV _
                    (DGV_Maitresse_Local, Outil_Courant, Tab_Chaine_Local(5), Tab_Chaine_Local(6)) = _
                 False Then Exit Sub
        End If

        'Message de confirmation
        Fonctions_Messages.Appel_Msg(40, 1, , )

        'Fermeture Form
        Me.Dispose()
    End Sub
End Class