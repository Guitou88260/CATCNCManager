

'Modifs à prévoir:
'- Passer par tab pour remplissage DGV (commencé mais pas terminé)
'- Permettre l'écriture dans tab outils machine également
'- Ajouter génération fichier de mesure (avec G300)


Public Class Form_Tab_Outils

    'FOCAS
    'Var
    Private retVal As Short
    Private libHndl As Integer
    Private retPMacro As ODBPM

    'Librairie handle 3
    Declare Function cnc_allclibhndl4 Lib "FWLIB32.DLL" (ByVal ip As String, ByVal port As Short, ByVal timeout As Integer, ByVal id As Integer, ByRef FlibHndl As Integer) As Short
    Declare Function cnc_freelibhndl Lib "FWLIB32.DLL" (ByRef FlibHndl As Integer) As Short
    Declare Function cnc_rdpmacro Lib "FWLIB32.DLL" (ByVal FlibHndl As Integer, ByVal a As Integer, ByRef b As ODBPM) As Short
    Declare Function cnc_rdpmacror Lib "FWLIB32.DLL" (ByVal FlibHndl As Integer, ByVal a As Integer, ByVal b As Integer, ByVal c As Integer, ByRef d As IODBPR) As Short

    'Structure lecture var 1 à 1
    Public Structure ODBPM
        Public datano As Integer    ' variable number 
        Public dummy As Short       ' dummy 
        Public mcr_val As Integer   ' macro variable 
        Public dec_val As Short     ' decimal point 
    End Structure

    'Structure lecture var en paquet
    Public Structure IODBPR_data
        Public mcr_val As Integer   ' macro variable 
        Public dec_val As Short     ' decimal point 
    End Structure

    'Structure nom outil
    Public Structure IODBPR
        Public datano_s As Integer  ' start macro number 
        Public dummy As Short       ' dummy 
        Public datano_e As Integer  ' end macro number 
        Public data As IODBPR1
    End Structure

    Public Structure IODBPR1
        Public LtrNomTool1 As IODBPR_data
        Public LtrNomTool2 As IODBPR_data
        Public LtrNomTool3 As IODBPR_data
        Public LtrNomTool4 As IODBPR_data
        Public LtrNomTool5 As IODBPR_data
        Public LtrNomTool6 As IODBPR_data
        Public LtrNomTool7 As IODBPR_data
        Public LtrNomTool8 As IODBPR_data
        Public LtrNomTool9 As IODBPR_data
        Public LtrNomTool10 As IODBPR_data
        Public LtrNomTool11 As IODBPR_data
        Public LtrNomTool12 As IODBPR_data
        Public LtrNomTool13 As IODBPR_data
        Public LtrNomTool14 As IODBPR_data
        Public LtrNomTool15 As IODBPR_data
        Public LtrNomTool16 As IODBPR_data
        Public LtrNomTool17 As IODBPR_data
        Public LtrNomTool18 As IODBPR_data
        Public LtrNomTool19 As IODBPR_data
        Public LtrNomTool20 As IODBPR_data
        Public LtrNomTool21 As IODBPR_data
        Public LtrNomTool22 As IODBPR_data
        Public LtrNomTool23 As IODBPR_data
        Public LtrNomTool24 As IODBPR_data
        Public LtrNomTool25 As IODBPR_data
        Public LtrNomTool26 As IODBPR_data
        Public LtrNomTool27 As IODBPR_data
        Public LtrNomTool28 As IODBPR_data
        Public LtrNomTool29 As IODBPR_data
        Public LtrNomTool30 As IODBPR_data
        Public LtrNomTool31 As IODBPR_data
        Public LtrNomTool32 As IODBPR_data
    End Structure

    'Variables
    Private Bouton_Ouverture_Local As System.Windows.Forms.Button
    Private ComboBox_Selection_Local As System.Windows.Forms.ComboBox
    Private Form_Filtre_Locale As Form_Filtre
    Private tabTemp(,) As String
    Private paramMachTemp As paramMach
    'Private tabToolsTemp(,) As String

    'Var param machine
    Public Structure paramMach
        Public nomMach As String
        Public IPMach As String
        Public nbTools As Integer
    End Structure

    '------------------------------------------------- Action sur événement ---------------------------------------------------

    'Action sur création nouvelle Form

    Public Sub New(Optional ByRef Bouton_Ouverture As System.Windows.Forms.Button = Nothing,
                   Optional ByRef ComboBox_Selection As System.Windows.Forms.ComboBox = Nothing)

        'Initialisation composant
        InitializeComponent()

        'Valuation variables
        Bouton_Ouverture_Local = Bouton_Ouverture
        ComboBox_Selection_Local = ComboBox_Selection
        paramMachTemp.nomMach = ComboBox_Selection_Local.Text

        'Lecture CSV et si erreur, sortie de Sub
        If Fonctions_CSV.Lecture_CSV _
            (Dossier_Reseau & "\Config_Machines\Fichier_Config_Machines.csv",
             tabTemp, 5) = False Then Exit Sub

        'Valuation var param machine
        paramMachTemp.IPMach = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab _
            (tabTemp, ComboBox_Selection.Text, "Machine", "Adresse IP", 1)
        paramMachTemp.nbTools = Fonctions_Tableau.Renvoi_Chaine_Autre_Colonne_Tab _
            (tabTemp, ComboBox_Selection.Text, "Machine", "Nb outils", 1)

        'Valuation var
        'ReDim tabToolsTemp(paramMachTemp.nbTools, Me.DataGridView1.Columns.Count - 1)

        'Gestion Button de sélection
        Fonctions_Form.Gestion_Bouton_Et_ComboBox_Form(Bouton_Ouverture_Local, 1, ComboBox_Selection, paramMachTemp.nomMach)
    End Sub

    'Action sur fermeture Form

    Private Sub Form_Listing_Prog_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) _
        Handles MyBase.FormClosed

        'Traitement fermeture DGV
        Fonctions_Form.Fermeture_Form_Enfant(Me, False, )

        'Gestion Button de sélection
        Fonctions_Form.Gestion_Bouton_Et_ComboBox_Form(Bouton_Ouverture_Local, 2, ComboBox_Selection_Local, paramMachTemp.nomMach)
    End Sub

    'Click bouton connexion FOCAS
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        'Connexion
        retVal = cnc_allclibhndl4(paramMachTemp.IPMach, 8193, 10, 0, libHndl)

        'Si erreur msg et Exit Sub
        If retVal <> 0 Then

            'Msg & Exit Sub
            Fonctions_Messages.Appel_Msg(60, 5, , retVal)
            Exit Sub
        End If

        'Gestion des boutons
        Me.Button5.Enabled = False
        Me.Button1.Enabled = True
        Me.Button3.Enabled = True
    End Sub

    'Click bouton déconnexion FOCAS
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'Déconnexion
        retVal = cnc_freelibhndl(libHndl)

        'Si erreur msg et Exit Sub
        If retVal <> 0 Then

            'Msg & Exit Sub
            Fonctions_Messages.Appel_Msg(61, 5, , retVal)
            Exit Sub
        End If

        'Gestion des boutons
        Me.Button5.Enabled = True
        Me.Button1.Enabled = False
        Me.Button3.Enabled = False
    End Sub

    'Click lecture tableau outils
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        'Effacement DGV
        Me.DataGridView1.Rows.Clear()

        'Autorisation ajout ligne
        'Me.DataGridView1.AllowUserToAddRows = True

        'Boucle outils
        For incTools = 1 To paramMachTemp.nbTools Step 1

            'Ajout ligne
            Me.DataGridView1.Rows.Add()

            'Ecriture num outil
            Me.DataGridView1.Rows.Item(incTools - 1).Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Numéro outil")).Value = Str(incTools)

            'Lecture var paramètres outils
            retVal = cnc_rdpmacro(libHndl, 28000 + (16 * (incTools - 1)), retPMacro)
            If retVal <> 0 Then
                'Si erreur msg et Exit Sub
                Fonctions_Messages.Appel_Msg(62, 5, , retVal)
                Exit Sub
            End If

            'Var
            Dim toolType As Integer = retPMacro.mcr_val / 10 ^ retPMacro.dec_val

            'Select suivant type d'outil pour géo outil
            Select Case toolType

                'Si outil de fraisage
                Case 2

                    'Lecture et écriture valeur dans DGV et si erreur, saut vers msg
                    'Z
                    retVal = cnc_rdpmacro(libHndl, 14000 + (120 * (incTools - 1)), retPMacro)
                    If retVal <> 0 Then
                        'Si erreur msg et Exit Sub
                        Fonctions_Messages.Appel_Msg(62, 5, , retVal)
                        Exit Sub
                    End If
                    'Ecriture DGV
                    DataGridView1.Rows.Item(incTools - 1).Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Z")).Value =
                        CDbl(retPMacro.mcr_val / 10 ^ retPMacro.dec_val) & "mm"

                    'R
                    retVal = cnc_rdpmacro(libHndl, 14000 + (120 * (incTools - 1)) + 1, retPMacro)
                    If retVal <> 0 Then
                        'Si erreur msg et Exit Sub
                        Fonctions_Messages.Appel_Msg(62, 5, , retVal)
                        Exit Sub
                    End If
                    'Ecriture DGV
                    DataGridView1.Rows.Item(incTools - 1).Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "R")).Value =
                        CDbl(retPMacro.mcr_val / 10 ^ retPMacro.dec_val) & "mm"

                    'Si outil de tournage
                Case -1, 1, 3

                    'Lecture et écriture valeur dans DGV et si erreur, saut vers msg
                    'X
                    retVal = cnc_rdpmacro(libHndl, 14000 + (120 * (incTools - 1)), retPMacro)
                    If retVal <> 0 Then
                        'Si erreur msg et Exit Sub
                        Fonctions_Messages.Appel_Msg(62, 5, , retVal)
                        Exit Sub
                    End If
                    'Ecriture DGV
                    DataGridView1.Rows.Item(incTools - 1).Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "X")).Value =
                        CDbl(retPMacro.mcr_val / 10 ^ retPMacro.dec_val) & "mm"

                    'Y
                    retVal = cnc_rdpmacro(libHndl, 14000 + (120 * (incTools - 1)) + 1, retPMacro)
                    If retVal <> 0 Then
                        'Si erreur msg et Exit Sub
                        Fonctions_Messages.Appel_Msg(62, 5, , retVal)
                        Exit Sub
                    End If
                    'Ecriture DGV
                    DataGridView1.Rows.Item(incTools - 1).Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Y")).Value =
                        CDbl(retPMacro.mcr_val / 10 ^ retPMacro.dec_val) & "mm"

                    'Z
                    retVal = cnc_rdpmacro(libHndl, 14000 + (120 * (incTools - 1)) + 2, retPMacro)
                    If retVal <> 0 Then
                        'Si erreur msg et Exit Sub
                        Fonctions_Messages.Appel_Msg(62, 5, , retVal)
                        Exit Sub
                    End If
                    'Ecriture DGV
                    DataGridView1.Rows.Item(incTools - 1).Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Z")).Value =
                        CDbl(retPMacro.mcr_val / 10 ^ retPMacro.dec_val) & "mm"

                    'R
                    retVal = cnc_rdpmacro(libHndl, 14000 + (120 * (incTools - 1)) + 3, retPMacro)
                    If retVal <> 0 Then
                        'Si erreur msg et Exit Sub
                        Fonctions_Messages.Appel_Msg(62, 5, , retVal)
                        Exit Sub
                    End If
                    'Ecriture DGV
                    DataGridView1.Rows.Item(incTools - 1).Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "R")).Value =
                        CDbl(retPMacro.mcr_val / 10 ^ retPMacro.dec_val) & "mm"
            End Select

            'Select suivant type d'outil pour nom outil
            Select Case toolType

                    'Tous les outils
                Case -1, 1, 2, 3

                    'Var
                    Dim tabASCII(31) As Integer
                    Dim varNameTool As IODBPR

                    'Lecture var noms d'outil
                    retVal = cnc_rdpmacror(libHndl, 25000 + (32 * (incTools - 1)), 25031 + (32 * (incTools - 1)), 12 + 8 * 32, varNameTool)
                    If retVal <> 0 Then
                        'Si erreur msg et Exit Sub
                        Fonctions_Messages.Appel_Msg(62, 5, , retVal)
                        Exit Sub
                    End If

                    'Remp Tab
                    tabASCII(0) = varNameTool.data.LtrNomTool1.mcr_val / 10 ^ varNameTool.data.LtrNomTool1.dec_val
                    tabASCII(1) = varNameTool.data.LtrNomTool2.mcr_val / 10 ^ varNameTool.data.LtrNomTool2.dec_val
                    tabASCII(2) = varNameTool.data.LtrNomTool3.mcr_val / 10 ^ varNameTool.data.LtrNomTool3.dec_val
                    tabASCII(3) = varNameTool.data.LtrNomTool4.mcr_val / 10 ^ varNameTool.data.LtrNomTool4.dec_val
                    tabASCII(4) = varNameTool.data.LtrNomTool5.mcr_val / 10 ^ varNameTool.data.LtrNomTool5.dec_val
                    tabASCII(5) = varNameTool.data.LtrNomTool6.mcr_val / 10 ^ varNameTool.data.LtrNomTool6.dec_val
                    tabASCII(6) = varNameTool.data.LtrNomTool7.mcr_val / 10 ^ varNameTool.data.LtrNomTool7.dec_val
                    tabASCII(7) = varNameTool.data.LtrNomTool8.mcr_val / 10 ^ varNameTool.data.LtrNomTool8.dec_val
                    tabASCII(8) = varNameTool.data.LtrNomTool9.mcr_val / 10 ^ varNameTool.data.LtrNomTool9.dec_val
                    tabASCII(9) = varNameTool.data.LtrNomTool10.mcr_val / 10 ^ varNameTool.data.LtrNomTool10.dec_val
                    tabASCII(10) = varNameTool.data.LtrNomTool11.mcr_val / 10 ^ varNameTool.data.LtrNomTool11.dec_val
                    tabASCII(11) = varNameTool.data.LtrNomTool12.mcr_val / 10 ^ varNameTool.data.LtrNomTool12.dec_val
                    tabASCII(12) = varNameTool.data.LtrNomTool13.mcr_val / 10 ^ varNameTool.data.LtrNomTool13.dec_val
                    tabASCII(13) = varNameTool.data.LtrNomTool14.mcr_val / 10 ^ varNameTool.data.LtrNomTool14.dec_val
                    tabASCII(14) = varNameTool.data.LtrNomTool15.mcr_val / 10 ^ varNameTool.data.LtrNomTool15.dec_val
                    tabASCII(15) = varNameTool.data.LtrNomTool16.mcr_val / 10 ^ varNameTool.data.LtrNomTool16.dec_val
                    tabASCII(16) = varNameTool.data.LtrNomTool17.mcr_val / 10 ^ varNameTool.data.LtrNomTool17.dec_val
                    tabASCII(17) = varNameTool.data.LtrNomTool18.mcr_val / 10 ^ varNameTool.data.LtrNomTool18.dec_val
                    tabASCII(18) = varNameTool.data.LtrNomTool19.mcr_val / 10 ^ varNameTool.data.LtrNomTool19.dec_val
                    tabASCII(19) = varNameTool.data.LtrNomTool20.mcr_val / 10 ^ varNameTool.data.LtrNomTool20.dec_val
                    tabASCII(20) = varNameTool.data.LtrNomTool21.mcr_val / 10 ^ varNameTool.data.LtrNomTool21.dec_val
                    tabASCII(21) = varNameTool.data.LtrNomTool22.mcr_val / 10 ^ varNameTool.data.LtrNomTool22.dec_val
                    tabASCII(22) = varNameTool.data.LtrNomTool23.mcr_val / 10 ^ varNameTool.data.LtrNomTool23.dec_val
                    tabASCII(23) = varNameTool.data.LtrNomTool24.mcr_val / 10 ^ varNameTool.data.LtrNomTool24.dec_val
                    tabASCII(24) = varNameTool.data.LtrNomTool25.mcr_val / 10 ^ varNameTool.data.LtrNomTool25.dec_val
                    tabASCII(25) = varNameTool.data.LtrNomTool26.mcr_val / 10 ^ varNameTool.data.LtrNomTool26.dec_val
                    tabASCII(26) = varNameTool.data.LtrNomTool27.mcr_val / 10 ^ varNameTool.data.LtrNomTool27.dec_val
                    tabASCII(27) = varNameTool.data.LtrNomTool28.mcr_val / 10 ^ varNameTool.data.LtrNomTool28.dec_val
                    tabASCII(28) = varNameTool.data.LtrNomTool29.mcr_val / 10 ^ varNameTool.data.LtrNomTool29.dec_val
                    tabASCII(29) = varNameTool.data.LtrNomTool30.mcr_val / 10 ^ varNameTool.data.LtrNomTool30.dec_val
                    tabASCII(30) = varNameTool.data.LtrNomTool31.mcr_val / 10 ^ varNameTool.data.LtrNomTool31.dec_val
                    tabASCII(31) = varNameTool.data.LtrNomTool32.mcr_val / 10 ^ varNameTool.data.LtrNomTool32.dec_val

                    'Ecriture dans DGV
                    DataGridView1.Rows.Item(incTools - 1).Cells.Item(Fonctions_DGV.Renvoi_Num_Colonne_DGV(Me.DataGridView1, "Nom")).Value = ConvASCII(tabASCII)
            End Select
        Next

        'Ajustement des largeurs de colonne
        DataGridView1.AutoResizeColumns()
    End Sub

    '---------------------------------------------------- Filtres DGV -----------------------------------------------------

    'Si click bouton filtre

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click

        'Si Form Filtre = Nothing
        If Form_Filtre_Locale Is Nothing Then

            'Création nouvelle Form Filtre
            Form_Filtre_Locale = New Form_Filtre(Me.DataGridView1, )

            'Ouverture Form filtre
            Form_Filtre_Locale.ShowDialog()

            'Sinon
        Else

            'Ouverture Form filtre existante
            Form_Filtre_Locale.ShowDialog()
        End If
    End Sub

    'Si click bouton suppression filtre

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click

        'Vidage variable
        Form_Filtre_Locale = Nothing

        'Relecture de tous les filtres
        Fonctions_Filtres_DGV.Relecture_Tous_Filtres_DGV(Me.DataGridView1, , , )
    End Sub
End Class