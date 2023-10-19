Imports OfficeOpenXml
Imports OfficeOpenXml.VBA
Imports System.Security.Cryptography.X509Certificates

Namespace EPPlusSamples.EncryptionProtectionAndVba
    Friend Class SigningYourVBAProject
        ''' <summary>
        ''' Opens the Battleships sample and sign it with the certificate from the pfx file.
        ''' </summary>
        Public Shared Sub Run()
            'Load our test certificate from the pfx file. 
            'In a real production environment, make to store your certificate in a secure way.
            Dim certFile = FileUtil.GetFileInfo("08-Encryption Protection and VBA\02-VBA", "SampleCertificate.pfx")
            Dim cert = New X509Certificate2(certFile.FullName, "EPPlus")

            'Open the workbook created in the previous sample.
            Using p = New ExcelPackage(FileUtil.GetFileInfo("8.2-03-CreateABattleShipsGameVba.xlsm"))
                Dim signature = p.Workbook.VbaProject.Signature

                'The only thing you need to do to sign your project is to set the signatures 'Certificate' property with your code-signing certificate.
                'The certificate must have access to the private key to sign the project.
                signature.Certificate = cert

                'If the file is unsigned, EPPlus will by default create all three signatures - Legacy, Agile and V3.
                'You can use the property 'CreateSignatureOnSave' to decide which signature version you want to create on saving the workbook.
                'For example 'signature.LegacySignature.CreateSignatureOnSave = false' to remove the legacy signature.

                'You can also set the hash algorithm for each signature version.
                'The Excel and EPPlus default is MD5 for the legacy signature and SHA1 for the Agile and V3 signature.
                'We want to change it to SHA256 to get better and more modern hash algorithm.
                signature.LegacySignature.HashAlgorithm = VbaSignatureHashAlgorithm.SHA256
                signature.AgileSignature.HashAlgorithm = VbaSignatureHashAlgorithm.SHA256
                signature.V3Signature.HashAlgorithm = VbaSignatureHashAlgorithm.SHA256

                p.SaveAs(FileUtil.GetFileInfo("8.2-04-Signed-CreateABattleShipsGameVba.xlsm"))
            End Using
        End Sub

    End Class
End Namespace
