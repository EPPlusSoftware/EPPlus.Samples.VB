Imports EPPlusSamples.EncryptionProtectionAndVba
Imports System

Namespace EPPlusSamples
    Public Module EncryptionProtectionAndVBASample
        Public Sub Run()
            'Sample 8.1 Swedish Quiz : Shows Encryption, workbook- and worksheet protection.
            Call EncryptionAndProtection.Run()

            'Sample 8.2 - Shows how to work with macro-enabled workbooks(VBA) and how to sign the code with a certificate.
            Console.WriteLine("Running sample 8.2-VBA")
            Call WorkingWithVbaSample.Run()
            Call SigningYourVBAProject.Run()

            'Sample 8.3 shows how to sign workbooks
            Call DigitalSignatureSample.Run()

            Console.WriteLine("Sample 8.2 created {0}", FileUtil.OutputDir.Name)
            Console.WriteLine()
        End Sub
    End Module
End Namespace
