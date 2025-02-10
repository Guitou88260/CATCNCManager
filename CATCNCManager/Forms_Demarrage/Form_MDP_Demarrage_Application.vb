Public Class Form_MDP_Demarrage_Application

    '------------------------------------------------- Action sur événement ---------------------------------------------------

    'Click bouton "Ouvrir en mode visualisation"

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'Valuation variable
        Niveau_Ouverture = 1

        'Fermeture Form MDP
        Me.Dispose()
    End Sub

    'Click bouton "Ouvrir en mode modification"

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        'Variable
        Dim Resultat_MDP As Boolean

        'Contrôle MDP
        Fonctions_Form.Ctrl_MDP(Me, Me.TextBox1, ParamsApp.Mdp, Resultat_MDP)

        'Si MDP OK
        If Resultat_MDP = True Then

            'Valuation variable
            Niveau_Ouverture = 2

            'Fermeture Form MDP
            Me.Dispose()
        End If
    End Sub

    'Click bouton "Annuler"

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        'Fermeture application
        End
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
