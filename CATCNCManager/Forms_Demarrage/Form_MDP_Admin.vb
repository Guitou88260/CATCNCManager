Public Class Form_MDP_Admin

    '------------------------------------------------- Action sur événement ---------------------------------------------------

    'Action sur fermeture Form et click bouton annuler

    Private Sub Form_MDP_Admin_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) _
        Handles MyBase.FormClosed, Button3.Click

        'Fermeture Form
        Me.Dispose()
    End Sub

    'Click bouton "Ouvrir en mode administrateur"

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        'Variable
        Dim Resultat_MDP As Boolean = False

        'Contrôle MDP
        Fonctions_Form.Ctrl_MDP(Me, Me.TextBox1, ParamsApp.MdpAdmin, Resultat_MDP)

        'Si MDP OK
        If Resultat_MDP = True Then

            'Valuation variable
            Niveau_Ouverture = 3

            'Ecriture mode ouverture dans entête Form principale
            Fonctions_Form.Ecriture_Mode_Ouverture(Form_Principale)
        End If
    End Sub

    'Si touche entrer enfoncée au clavier lorsque TextBox1 à le focus

    Private Sub TextBox1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown

        'Si détecte enfoncement touche entrer
        If e.KeyCode = Keys.Enter Then

            'Provoque le click Button2
            Button2.PerformClick()
        End If
    End Sub
End Class
