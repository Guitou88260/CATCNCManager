Public Class Form_Choix_Num_Prog_Listing

    'Variable
    Private DGV_Maitresse_Local As DataGridView

    '------------------------------------------------- Action sur événement ---------------------------------------------------

    'Action sur création nouvelle Form

    Public Sub New(ByRef DGV_Maitresse As DataGridView)

        'Initialisation composant
        InitializeComponent()

        'Valuation variable
        DGV_Maitresse_Local = DGV_Maitresse

        'Remplissage des items de la colonne des numéro de programme
        Fonctions_DGV.Remp_ComboBox_Avec_Colonne_DGV(DGV_Maitresse_Local, "Numéro programme", Me.ComboBox1)

        'Remplissage ComboBox avec programme sélectionné
        Me.ComboBox1.Text = DGV_Maitresse_Local.CurrentCell.Value
    End Sub

    'Action sur fermeture Form

    Private Sub Form_Choix_Num_Prog_Listing_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) _
        Handles MyBase.FormClosed

        'Fermeture Form
        Me.Dispose()
    End Sub

    'Click bouton "Insérer numéro"

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'Si ComboBox non-vide
        If Me.ComboBox1.Text <> Nothing Then

            'Remplissage bouton programme DGV avec numéro sélectionné
            DGV_Maitresse_Local.CurrentCell.Value = Me.ComboBox1.Text

            'Message de confirmation
            Fonctions_Messages.Appel_Msg(12, 1, , )

            'Fermeture la Form
            Me.Dispose()

            'Sinon
        Else

            'Message et sortie de Sub
            Fonctions_Messages.Appel_Msg(30, 3, , )
            Exit Sub
        End If
    End Sub
End Class
