Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing
Imports OfficeOpenXml.DigitalSignatures
Imports System.Security.Cryptography.X509Certificates

Namespace EPPlusSamples.EncryptionProtectionAndVba
    Public Module DigitalSignatureSample
        Public Sub Run()
            'Sign wb with minimal details.
            Call SignWorkbookSimple()
            'Sign wb with details such as commitment type, title and address
            Call SignWorkbookWithDetails()
            'How to create and add setup to a signature line
            Call CreateSignatureline()
            'How to sign one or multiple signature lines using EPPlus
            Call SignSignatureLines()
        End Sub

        Private Sub SignWorkbookSimple()
            Using pck = New ExcelPackage()
                Dim wb = pck.Workbook
                Dim ws = wb.Worksheets.Add("SomeWorksheet")

                'A digital signature requires a certificate with a private key.
                'In this case we'll sign with a certificate stored in a .pfx file.
                'In a real production environment, make to store your certificate in a secure way.
                Dim certFile = FileUtil.GetFileInfo("08-Encryption Protection and VBA\02-VBA", "SampleCertificate.pfx")
                Dim cert = New X509Certificate2(certFile.FullName, "EPPlus")

                'Add a digital signature and sign it with the certificate
                Dim digitalSignature = wb.DigitalSignatures.Add(cert)

                'It is recommended to set a more secure digest method. As default is SHA-1.
                digitalSignature.SetDigestMethod(DigitalSignatureHashAlgorithm.SHA512)

                Dim fi = FileUtil.GetCleanFileInfo("8.3-01-SignWorkbook.xlsx")
                pck.SaveAs(fi)

                ' Note: Because this is a test certificate, it will count as a 'recoverable' signature
                ' unless you choose to trust the certificate.
                ' If you add your own certificate it should count as a 'valid' signature.
            End Using
        End Sub

        Private Function GetCert() As X509Certificate2
            Dim certFile = FileUtil.GetFileInfo("08-Encryption Protection and VBA\02-VBA", "SampleCertificate.pfx")

            'Excel lists different certificates for you.
            'You should be able to access the same list like this:
            'X509Store store = new X509Store(StoreLocation.CurrentUser);
            'store.Open(OpenFlags.ReadOnly);
            'var certAlt = store.Certificates[0];

            'If it doesn't work please double-check that you have certificates stored correctly
            'and double check the .net X509Store implementation for your OS

            Return New X509Certificate2(certFile.FullName, "EPPlus")
        End Function

        Private Sub SignWorkbookWithDetails()
            Using pck = New ExcelPackage(FileUtil.GetCleanFileInfo("8.3-02-SigningDetails.xlsx"))
                Dim wb = pck.Workbook
                Dim ws = wb.Worksheets.Add("DetailsWs")

                'Same as above sample
                Dim cert = GetCert()

                'The method to add a signature also includes optional parameters for the comments commitment type and reason for signing
                'That represent the 'commitment type' and 'purpose for signing this document' fields from Excel.
                Dim digitalSignature = wb.DigitalSignatures.Add(cert)

                digitalSignature.CommitmentTyping = CommitmentType.Approved
                digitalSignature.PurposeForSigning = "My reason for signing"

                'You can also add signer details via the Details property.
                'This represents the 'details' button in excel for example:
                Dim details = digitalSignature.Details

                details.SignerRoleTitle = "Developer"
                details.Address1 = "ExampleStreet 2"
                details.Address2 = "Floor 2"
                details.ZipOrPostalCode = "114 51"
                details.City = "Stockholm"
                details.CountryOrRegion = "Sweden"
                details.StateOrProvince = "Stockholm"

                'The signature xml is not truly created until after the file has been saved
                Dim isTheSignatureValid = digitalSignature.IsValid
                pck.Save()

                'And so won't be valid until after save:
                Dim signatureIsValid = digitalSignature.IsValid
            End Using
        End Sub


        Private Sub CreateSignatureline()
            Using pck = New ExcelPackage(FileUtil.GetCleanFileInfo("8.3-03-CreateSignatureLine.xlsx"))
                Dim wb = pck.Workbook
                Dim ws = wb.Worksheets.Add("SignatureLinesEmpty")

                'From a worksheet you can create a signature line
                'A visual representation via a vmldrawing object for signing.
                Dim signatureLine = ws.SignatureLines.Add()

                'As in excel, you can set a few options for a suggested signer.
                signatureLine.Signer = "FirstName LastName"
                signatureLine.Title = "Engineer"
                signatureLine.Email = "FirstName@yourdomain.com"
                signatureLine.SigningInstructions = "Please mr.LastName check and approve this document."
                signatureLine.AllowComments = True
                signatureLine.ShowSignDate = False

                'You can set the size and position of a signature line via from and to for columns and rows.
                signatureLine.From.Column = 5
                signatureLine.To.Column = 9
                signatureLine.From.Row = 0
                signatureLine.To.Row = 6

                'If opened in Excel someone can now double-click and sign this signature line.
                pck.Save()
            End Using
        End Sub

        Private Sub SignSignatureLines()
            'Open package from previous sample
            Using pck = New ExcelPackage(FileUtil.GetFileInfo("8.3-03-CreateSignatureLine.xlsx"))
                Dim wb = pck.Workbook
                Dim ws = wb.Worksheets(0)

                ws.Name = "SignedWorksheet"

                Dim cert = GetCert()

                Dim signatureLine = ws.SignatureLines(0)

                'Sign the signature line from the previous sample
                signatureLine.AsSignatureLine.SignWithText(cert, "FirstName LastName")

                'The reason for '.AsSignatureline' is because Signature Line is actually a child-class.
                'The parent class is SignatureLineStamp A signatureLineStamp can only be signed with an image and has a different look
                'Let's add one and sign that too
                Dim stamp = ws.SignatureLines.AddStamp()

                stamp.Signer = "FirstName LastName"
                stamp.Title = "Engineer"

                'Note that only .bmp file format are supported for digital signatures
                Dim sampleImage = FileUtil.GetFileInfo("08-Encryption Protection and VBA\03-DigitalSignatures", "SignatureImgExample.bmp")
                Dim image = New ExcelImage(sampleImage)
                stamp.SignWithImage(cert, image)


                'Stamps can also be resized and moved
                stamp.From.Column = 2
                stamp.To.Column = 4
                stamp.From.Row = 8
                stamp.To.Row = 17

                'Naturally a non-stamp can also be signed with an image.
                'Let's add one so we can see all variations.
                Dim SignatureLineTwo = ws.SignatureLines.Add()
                SignatureLineTwo.SignWithImage(cert, image)

                pck.SaveAs(FileUtil.GetCleanFileInfo("8.3-04-SignSignatureLines.xlsx"))
            End Using
        End Sub
    End Module
End Namespace
