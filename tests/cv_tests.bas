Attribute VB_Name = "cv_tests"
Option Explicit
Private Sub AssertEqual(ByVal actual As String, ByVal expected As String, ByVal label As String)
    If actual <> expected Then Err.Raise 5, , label & ": " & actual & " <> " & expected
End Sub
Private Function HexBytes(ByVal value As String) As Byte()
    Dim b() As Byte, i As Long
    ReDim b(0 To Len(value) \ 2 - 1)
    For i = 0 To UBound(b): b(i) = CByte("&H" & Mid$(value, i * 2 + 1, 2)): Next
    HexBytes = b
End Function
Private Function ByteHex(ByRef b() As Byte) As String
    Dim i As Long, s As String
    For i = 0 To UBound(b): s = s & Right$("0" & Hex$(b(i)), 2): Next
    ByteHex = LCase$(s)
End Function
Public Function CVCryptoSelfTest() As String
    On Error GoTo Failed
    Dim a() As Byte, b() As Byte, c() As Byte, d() As Byte, e() As Byte, f() As Byte
    Dim mode As Long, publicText As String, privateText As String, i As Long
    a = HexBytes("616263")
    b = CVHash(a)
    AssertEqual ByteHex(b), "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", "SHA256"
    a = HexBytes("6161616161616161616161616161616161616161616161616161616161616161616161616161616161616161616161616161616161616161")
    b = CVHash(a)
    AssertEqual ByteHex(b), "b35439a4ac6f0948b6d6f9e3c6af0f5f590ce20f1bde7090ef7970686ec6738a", "SHA256"
    a = HexBytes("61616161616161616161616161616161616161616161616161616161616161616161616161616161616161616161616161616161616161616161616161616161")
    b = CVHash(a)
    AssertEqual ByteHex(b), "ffe054fe7ae0cb6dc65c3af9b61d5209f439851db43d0ba5997337df154668eb", "SHA256"
    a = CVAscii(String$(1000, "a"))
    b = CVHash(a)
    AssertEqual ByteHex(b), "41edece42d63e8d9bf515a9ba6932e1c20cbc9f5a5d134645adb5db1b9737ea3", "SHA256"
    a = HexBytes("000102030405060708090a0b0c0d0e0f101112131415161718191a1b1c1d1e1f202122232425262728292a2b2c2d2e2f303132333435363738393a3b3c3d3e3f404142434445464748494a4b4c4d4e4f505152535455565758595a5b5c5d5e5f606162636465666768696a6b6c6d6e6f707172737475767778797a7b7c7d7e7f808182838485868788898a8b8c8d8e8f909192939495969798999a9b9c9d9e9fa0a1a2a3a4a5a6a7a8a9aaabacadaeafb0b1b2b3b4b5b6b7b8b9babbbcbdbebfc0c1c2c3c4c5c6c7c8c9cacbcccdcecfd0d1d2d3d4d5d6d7d8d9dadbdcdddedfe0e1e2e3e4e5e6e7e8e9eaebecedeeeff0f1f2f3f4f5f6f7f8f9fafbfcfdfeff")
    b = CVHash(a)
    AssertEqual ByteHex(b), "40aff2e9d2d8922e47afd4648e6967497158785fbd1da870e7110266bf944880", "SHA256"
    a = HexBytes("76cb70408a954f93fbe2001debc7b8c81052b1831ce62ad9355ba1584464f635"): b = HexBytes("dc82cc1bd4094442a463b51584eb281c56ba9bbce15e68d832c9802832760452")
    c = CVX25519(a, b)
    AssertEqual ByteHex(c), "37f9196ad831b413b2aab1740360aa44e89f92f831daa158c0cd934fe0d6722a", "X25519 Python interoperability"
    a = HexBytes("68985ff9c84498437f202e25e6e1f77553b937111fbe64c84e3f0a8084573e86"): b = HexBytes("c888521b3a96fb7fe81b99b3ccb7d4fab356a13a51baf08dc600adef787f8976")
    c = CVX25519(a, b)
    AssertEqual ByteHex(c), "9f18ef72a7a1fe5c26f48d26517d5fa0e1c8d6ba9a47284377cbd01c26b2b20d", "X25519 Python interoperability"
    a = HexBytes("2bc701516ce0454e7216d6f18c9f2ad0cf1702b52cdcde92f740fc6e6caafe0d"): b = HexBytes("3b3606af32708cf88604284ed2be81ff9650b346521b5a234eae9b02a6a40d11")
    c = CVX25519(a, b)
    AssertEqual ByteHex(c), "ad7296bc4ce3318a79634a08a31e24805afb6e40a328a399ba6c85f4b27d3647", "X25519 Python interoperability"
    a = HexBytes("414b26ea7cc886f719679b70b98aba2e487313421c71bba635a83c3937e99761"): b = HexBytes("7917c0d52f64dfe8ddaf5cc314e8b159882487039a73a7e0aacb40efa92df012")
    c = CVX25519(a, b)
    AssertEqual ByteHex(c), "7392b7e49cd86ea2933e53819f1bbbe15103097c1ffecbab5b9ec55bd2176d51", "X25519 Python interoperability"
    a = HexBytes("df9a24684ac434715db4f8559aa3572caf51d893cb4fac217986720e52cce72c"): b = HexBytes("a19efbf5af564e1cc3d3d898ad54f84583a85702fb024c930b0b0062b0184f07")
    c = CVX25519(a, b)
    AssertEqual ByteHex(c), "913db898df5b7ef193f986a2fb94e148c3073c0c3bd9ea2ac0bffeda69943031", "X25519 Python interoperability"
    a = HexBytes("2c"): b = HexBytes("cccc068bc3a5a2c2ebe6892d4359b2daf235bee1b4252348c8a9b7a6405da89f"): c = HexBytes("0000000000000000000000000000ffff")
    d = CVCTR(a, b, c)
    AssertEqual ByteHex(d), "ba", "AES256 CTR Python interoperability"
    e = CVCTR(d, b, c)
    AssertEqual ByteHex(e), ByteHex(a), "AES CTR reverse"
    a = HexBytes("195a967b686e5a36830350a25fd4f3c4"): b = HexBytes("ab067263d4059dac36c07f67750b88ee70fea99d5fd858923c0afd29598e92c6"): c = HexBytes("0000000000000000000000000000ffff")
    d = CVCTR(a, b, c)
    AssertEqual ByteHex(d), "ea0931f8b9b8e0d5025fe00bd6873f20", "AES256 CTR Python interoperability"
    e = CVCTR(d, b, c)
    AssertEqual ByteHex(e), ByteHex(a), "AES CTR reverse"
    a = HexBytes("5166291a421d7de1a59f382e452e431df0"): b = HexBytes("7df53ee5d89c2eccefd30910226b174dc532e76b36a2a0eeee92846d1f54b8fc"): c = HexBytes("0000000000000000000000000000ffff")
    d = CVCTR(a, b, c)
    AssertEqual ByteHex(d), "48b1bb2883892850a2b6a27bc82154c006", "AES256 CTR Python interoperability"
    e = CVCTR(d, b, c)
    AssertEqual ByteHex(e), ByteHex(a), "AES CTR reverse"
    a = HexBytes("8e8de2121579fe3117978912ca8e0851da4bbe93e3d81a3cd21d92a8c1e32029b9"): b = HexBytes("d00162bc5128f5358bb0484dab5679bdb1d7ab0848b842341452469c37c7cc57"): c = HexBytes("0000000000000000000000000000ffff")
    d = CVCTR(a, b, c)
    AssertEqual ByteHex(d), "e0f7456c828518e58e9d72e7a8d85659705ee940f2265d53c4557afb4360bdc295", "AES256 CTR Python interoperability"
    e = CVCTR(d, b, c)
    AssertEqual ByteHex(e), ByteHex(a), "AES CTR reverse"
    a = HexBytes("668e4010ad6d8339d4225d7582d2ac9a244abe8aaa6f90de854b6d69f9cb1a50c7e456eb0c2be63c3ac05d9ec62e8ff8be3def0ddf70edae1fac3c2ade3fefbecbc50f41e1017eb30d3ce48d5f89b25d4d84d9effc814e7f90236c3925d6fcea0fb7d0fcd08f85c17720e5b180c76e298c3db63c60de53550ba8e530f6beac1c"): b = HexBytes("970ef40f01e0cf1ec968352c5344e8c4140ff6aeeb4ae5b6c9bd2d00279f8993"): c = HexBytes("0000000000000000000000000000ffff")
    d = CVCTR(a, b, c)
    AssertEqual ByteHex(d), "b43ca14a92b6126107b33ee25dd10a6048c1530889af1e2833cc29b1ffa525b0f137194dd0097b32b812b40740cf8eaa75ef62d365dc7b2a5d92a37ae06a97121d506e93b28405e8c120aff853ca55f992632b0eb34df74c8df28adf2c9ca5a46e63cb87785e4da86128d50a739a8d72cec785aab7a2b4882f96a21326473593", "AES256 CTR Python interoperability"
    e = CVCTR(d, b, c)
    AssertEqual ByteHex(e), ByteHex(a), "AES CTR reverse"
    privateText = "8cK4JMHDpf_kjCk": publicText = "BTB6LXorTePBHTI"
    a = CVDecode(privateText, 11)
    b = CVPublic(a, 1)
    AssertEqual CVEncode(b), publicText, "Python demo public key"
    a = CVAscii("Unicode payload is tested in the carrier test")
    b = CVSeal(a, publicText, 1)
    c = CVOpen(b, privateText, 1)
    AssertEqual ByteHex(c), ByteHex(a), "Asymmetric round trip"
    d = CVSeal(a, publicText, 1)
    If ByteHex(b) = ByteHex(d) Then Err.Raise 5, , "Encryption repeated"
    a = HexBytes("57c169f255738228f513c0"): b = HexBytes("9eeec4fe2fa5fdb9c922d0"): c = HexBytes("a4d8d73baf546e53294ec7")
    d = CVKDF(a,b,c,1)
    AssertEqual ByteHex(d), "ac987ae726ed29735699e92a25066df704fafc1ddb6cfad7697dad91649dc94f", "HKDF Python interoperability"
    privateText = "CPL_5-vYCrOh-7oorslz1OIsf0V-FQ2mlRIlFk4oLEM": publicText = "-sA52O9X0r_Gh5DgEqtMbIf7TKFle_uDuf6vadvBTEM"
    a = CVDecode(privateText, 32)
    b = CVPublic(a, 2)
    AssertEqual CVEncode(b), publicText, "Python production public key"
    a = CVAscii("Unicode payload is tested in the carrier test")
    b = CVSeal(a, publicText, 2)
    c = CVOpen(b, privateText, 2)
    AssertEqual ByteHex(c), ByteHex(a), "Asymmetric round trip"
    d = CVSeal(a, publicText, 2)
    If ByteHex(b) = ByteHex(d) Then Err.Raise 5, , "Encryption repeated"
    a = HexBytes("c426d8e194ecb5a7fda5c125d6d781b9f3107c2c9b2ebaaef065618c4afa7d44"): b = HexBytes("e71cb90d15e3a5a4e175952fc82fd8aa3329489f47798d4cf7c36e48110b0745"): c = HexBytes("eb0efb2a087e42384f9f45ae08133bf5a5936b9375e15e96963f9640f1e2b178")
    d = CVKDF(a,b,c,2)
    AssertEqual ByteHex(d), "04fa296ca230f69aae18519c641384e1ae9d203c622000052cb2af57ec778f7f", "HKDF Python interoperability"
    For i = 1 To 100
        a = CVRandom(i): b = CVDecode(CVEncode(a), i)
        AssertEqual ByteHex(a), ByteHex(b), "Base64URL"
    Next
    CVCryptoSelfTest = "PASS: SHA256, X25519, AES256-CTR, HKDF, both key modes, Base64URL"
    Exit Function
Failed:
    CVCryptoSelfTest = "FAIL " & Err.Number & ": " & Err.Description
End Function
Public Sub CVSaveActiveTestDocument(ByVal filename As String)
    ActiveWorkbook.SaveAs Filename:=filename, FileFormat:=xlOpenXMLWorkbook
End Sub
Public Sub CVExportActiveTestDocument(ByVal filename As String)
    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=filename
End Sub
