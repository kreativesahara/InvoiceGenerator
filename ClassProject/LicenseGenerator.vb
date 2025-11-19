Imports System
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

''' <summary>
''' Simple standalone LicenseGenerator utility for vendors.
''' Usage (console):
'''   - To generate a new RSA keypair:
'''       LicenseGenerator.exe genkeys <privateKeyOut.xml> <publicKeyOut.xml>
'''
'''   - To create a license file (payload + signature):
'''       LicenseGenerator.exe makelic <privateKey.xml> <clientId> <expiry-iso8601> [out.lic]
'''
''' Notes / guidance for vendor:
'''   - The application (client) embeds the public key XML (ToXmlString(FALSE) output) into
'''     LicenseManager.PublicKeyXml. Replace the placeholder there with the generated public key XML.
'''   - The license format expected by the client is two lines:
'''       1) payload: "expiry=<ISO-8601>;clientid=<GUID>"
'''       2) base64(signature) where signature = RSA-SHA256(payload) using the vendor private key
'''   - Example payload: expiry=2026-12-31T23:59:59Z;clientid=3f3a1b2c-... (use UTC in ISO 'o' or 's')
'''   - Use a secure process to protect the private key. Do NOT distribute the private key.
'''   - Consider including additional fields (license version, features) in the payload. If you
'''     change payload structure, update LicenseManager.TryValidateLicense parsing accordingly.
''' </summary>
Module LicenseGenerator

    Private Sub ShowUsage()
        Console.WriteLine("LicenseGenerator utility")
        Console.WriteLine()
        Console.WriteLine("Commands:")
        Console.WriteLine("  genkeys <privateOut.xml> <publicOut.xml>   - generate RSA4096 keypair and save XML")
        Console.WriteLine("  makelic <privateKey.xml> <clientId> <expiry-iso8601> [out.lic]    - make license file")
        Console.WriteLine()
        Console.WriteLine("Examples:")
        Console.WriteLine("  LicenseGenerator genkeys vendor_private.xml vendor_public.xml")
        Console.WriteLine("  LicenseGenerator makelic vendor_private.xml 3f3a1b2c-... 2026-12-31T23:59:59Z customer.lic")
    End Sub

    Private Function GenerateKeyPair() As String()
        ' Use RSA with 4096 bits for reasonable security
        Using rsa As New RSACryptoServiceProvider(4096)
            Try
                rsa.PersistKeyInCsp = False
                Dim priv = rsa.ToXmlString(True) ' include private parameters
                Dim pub = rsa.ToXmlString(False) ' public only
                Return New String() {priv, pub}
            Finally
                rsa.Clear()
            End Try
        End Using
    End Function

    Private Function SignPayload(payload As String, privateKeyXml As String) As Byte()
        Dim payloadBytes = Encoding.UTF8.GetBytes(payload)
        Using rsa As New RSACryptoServiceProvider()
            rsa.PersistKeyInCsp = False
            rsa.FromXmlString(privateKeyXml)
            Dim shaOid = CryptoConfig.MapNameToOID("SHA256")
            Dim signature = rsa.SignData(payloadBytes, shaOid)
            Return signature
        End Using
    End Function

    Private Function MakePayload(clientId As String, expiryIso As String) As String
        ' canonical payload format - use the provided ISO expiry (prefer 'o' UTC)
        Return $"expiry={expiryIso};clientid={clientId}"
    End Function

    Sub Main(args As String())
        If args.Length = 0 Then
            ShowUsage()
            Return
        End If

        Dim cmd = args(0).ToLowerInvariant()
        Try
            If cmd = "genkeys" Then
                If args.Length < 3 Then
                    Console.WriteLine("genkeys requires two file paths: <privateOut.xml> <publicOut.xml>")
                    Return
                End If
                Dim privOut = args(1)
                Dim pubOut = args(2)
                Dim pair = GenerateKeyPair()
                Dim privXml = pair(0)
                Dim pubXml = pair(1)
                File.WriteAllText(privOut, privXml)
                File.WriteAllText(pubOut, pubXml)
                Console.WriteLine("Generated keys:")
                Console.WriteLine("  Private: " & Path.GetFullPath(privOut))
                Console.WriteLine("  Public:  " & Path.GetFullPath(pubOut))
                Console.WriteLine()
                Console.WriteLine("Important: Keep the private key secure. Embed the public key XML into the client application.")
                Return
            End If

            If cmd = "makelic" Then
                If args.Length < 4 Then
                    Console.WriteLine("makelic requires: <privateKey.xml> <clientId> <expiry-iso8601> [out.lic]")
                    Return
                End If
                Dim privateKeyPath = args(1)
                Dim clientId = args(2)
                Dim expiryIso = args(3)
                Dim outPath As String = If(args.Length >= 5, args(4), "license.lic")

                If Not File.Exists(privateKeyPath) Then
                    Console.WriteLine("Private key file not found: " & privateKeyPath)
                    Return
                End If

                Dim privateXml = File.ReadAllText(privateKeyPath)
                ' Build canonical payload
                Dim payload = MakePayload(clientId, expiryIso)
                Dim signature = SignPayload(payload, privateXml)
                Dim signatureBase64 = Convert.ToBase64String(signature)
                Dim licenseText = payload & vbCrLf & signatureBase64
                File.WriteAllText(outPath, licenseText, Encoding.UTF8)
                Console.WriteLine("Written license to: " & Path.GetFullPath(outPath))
                Console.WriteLine("Payload:")
                Console.WriteLine(payload)
                Console.WriteLine("Signature (base64):")
                Console.WriteLine(signatureBase64)
                Console.WriteLine()
                Console.WriteLine("Deliver the license file to the customer. They must place it as 'license.lic' under %APPDATA%\InvoiceGenerator or use the app's Load License feature.")
                Return
            End If

            ShowUsage()
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.Message)
        End Try
    End Sub
End Module
