

'Paramètrage application
Module FctXMLParamApp

    'Sérialisation
    Public Sub SerialXML()
        'Serializer & new class
        Dim serializer As New XmlSerializer(GetType(ListParamsApp))
        'Mdp
        ParamsApp.Mdp = "NiFa2400"
        ParamsApp.MdpAdmin = "NiFa2400"
        'Data & FO test
        'TODO Voir pour mettre adresse swatchgroup au lieu de P:/Proto
        ParamsApp.DataApp = "P:\Proto\CATCNCManager"
        ParamsApp.FoTmpApp = "C:\Temp"
        'Data & FO
        ParamsApp.DataAppTest = "C:\Users\Guitou\Documents\CATCNCManager\Ressources_App"
        ParamsApp.FoTmpAppTest = "C:\Temp"
        'Name app
        ParamsApp.NameApp = "CATCNCManager"
        'Couleurs spéciales DGV
        ParamsApp.DgvDesact = ColorTranslator.ToHtml(Color.LavenderBlush)
        ParamsApp.DgvBloq = ColorTranslator.ToHtml(SystemColors.Control)
        ParamsApp.DgvStock = ColorTranslator.ToHtml(Color.MintCream)
        'Paramètrage colonnes DGV config machine & listing programme
        ParamsApp.ColRefA = "Numéro article, BT ou projet"
        ParamsApp.VerifColRefA = "(6) Longueur de chaîne inférieure à ""5"" et supérieure à ""11"" caractère(s) ou caractère ou chaîne différent de ""NA"" interdits, (7) Caractère(s) autre que ""0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-."" interdit(s)"
        ParamsApp.ColRefB = "Numéro NF"
        ParamsApp.VerifColRefB = "(6) Longueur de chaîne inférieure à ""10"" et supérieure à ""10"" caractère(s) ou caractère ou chaîne différent de ""NA"" interdits, (7) Caractère(s) autre que ""0123456789NF"" interdit(s)"
        'Paramètrage colonnes DGV programme
        ParamsApp.ColParamA = "Posage 1"
        ParamsApp.VerifColParamA = "(7) Caractère(s) autre que ""0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-."" interdit(s)"
        ParamsApp.ColParamB = "Posage 2"
        ParamsApp.VerifColParamB = "(7) Caractère(s) autre que ""0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-."" interdit(s)"
        ParamsApp.ColParamC = "Posage 3"
        ParamsApp.VerifColParamC = "(7) Caractère(s) autre que ""0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-."" interdit(s)"
        ParamsApp.ColParamD = "ParamD"
        ParamsApp.VerifColParamD = "(7) Caractère(s) autre que ""0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-."" interdit(s)"
        ParamsApp.ColParamE = "ParamE"
        ParamsApp.VerifColParamE = "(7) Caractère(s) autre que ""0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-."" interdit(s)"
        ParamsApp.ColParamF = "ParamF"
        ParamsApp.VerifColParamF = "(7) Caractère(s) autre que ""0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-."" interdit(s)"
        ParamsApp.ColParamG = "ParamG"
        ParamsApp.VerifColParamG = "(7) Caractère(s) autre que ""0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-."" interdit(s)"
        ParamsApp.ColParamH = "ParamH"
        ParamsApp.VerifColParamH = "(7) Caractère(s) autre que ""0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-."" interdit(s)"

        'Ecriture xml
        Dim writer As New StreamWriter("configApp.xml")
        serializer.Serialize(writer, ParamsApp)
        writer.Close()
    End Sub

    'Désérialisation
    Public Sub DeserialXML()
        'Serializer
        Dim serializer As New XmlSerializer(GetType(ListParamsApp))
        Using reader As New FileStream("configApp.xml", FileMode.Open)
            ParamsApp = CType(serializer.Deserialize(reader), ListParamsApp)
        End Using
    End Sub
End Module
