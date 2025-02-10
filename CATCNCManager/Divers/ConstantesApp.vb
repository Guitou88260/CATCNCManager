Public Module ConstantesApp
    'Constantes valuées dans l'app
    Public Dossier_Reseau As String
    Public Dossier_Temp As String
    Public Niveau_Ouverture As Integer
    Public Bouton_Ouverture_Temp As System.Windows.Forms.Button
    'Constantes valuées au démarrage de l'app
    Public ParamsApp As New ListParamsApp
End Module


'Class paramètres de l'app
Public Class ListParamsApp
    'Mdp
    Public Mdp As String
    Public MdpAdmin As String
    'Data & FO test
    Public DataApp As String
    Public FoTmpApp As String
    'Data & FO
    Public DataAppTest As String
    Public FoTmpAppTest As String
    'Name app
    Public NameApp As String
    'Couleurs spéciales DGV
    Public DgvDesact As String
    Public DgvBloq As String
    Public DgvStock As String
    'Paramètrage colonnes DGV config machine & listing programme
    Public ColRefA As String
    Public VerifColRefA As String
    Public ColRefB As String
    Public VerifColRefB As String
    'Paramètrage colonnes DGV programme
    Public ColParamA As String
    Public VerifColParamA As String
    Public ColParamB As String
    Public VerifColParamB As String
    Public ColParamC As String
    Public VerifColParamC As String
    Public ColParamD As String
    Public VerifColParamD As String
    Public ColParamE As String
    Public VerifColParamE As String
    Public ColParamF As String
    Public VerifColParamF As String
    Public ColParamG As String
    Public VerifColParamG As String
    Public ColParamH As String
    Public VerifColParamH As String
End Class