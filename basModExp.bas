Attribute VB_Name = "basModExp"
Option Explicit
Option Base 0

' A VB6/VBA procedure to carry out modular exponentiation
' with examples of RSA encryption and Diffie-Hellman key exchange

' USAGE:
' Example: strResult = mpModExp("3c", "03", "face")
' computes (0x3c)^3 mod 0xface = 0x5b56
' or, in decimal, 60^3 mod 64206 = 23382
' Parameters may be hex strings of any length subject to limitations
' of VB and your computer. May take a long time!

' First published 23 September 2005.
' mpFromHex modified 13 October 2007.
' mpModExp fixed "0" issue 4 February 2009
'************************* COPYRIGHT NOTICE*************************
' This code was originally written in Visual Basic by David Ireland
' and is copyright (c) 2005-9 D.I. Management Services Pty Limited,
' all rights reserved.

' You are free to use this code as part of your own applications
' provided you keep this copyright notice intact and acknowledge
' its authorship with the words:

'   "Contains cryptography software by David Ireland of
'   DI Management Services Pty Ltd <www.di-mgt.com.au>."

' If you use it as part of a web site, please include a link
' to our site in the form
' <a href="http://www.di-mgt.com.au/crypto.html">Cryptography
' Software Code</a>

' This code may only be used as part of an application. It may
' not be reproduced or distributed separately by any means without
' the express written permission of the author.

' David Ireland and DI Management Services Pty Limited make no
' representations concerning either the merchantability of this
' software or the suitability of this software for any particular
' purpose. It is provided "as is" without express or implied
' warranty of any kind.

' The latest version of this source code can be downloaded from
' www.di-mgt.com.au/crypto.html.
' Comments and bug reports to http://www.di-mgt.com.au/contact.html
'****************** END OF COPYRIGHT NOTICE*************************
' *********
' * TESTS *
' *********
Public Function Test_mpModExp()
    Dim strResult As String
    strResult = mpModExp("3c", "03", "face")
    Debug.Print strResult & " (expected 5b56)"
    strResult = mpModExp("beef", "03", "1000000000000") ' beef^3 = beef cubed = OXO?
    Debug.Print strResult & " (expected 6A35DDD3C9CF)"
    strResult = mpModExp("beef", "03", "10000")
    Debug.Print strResult & " (expected C9CF)"
    ' Do a mini-RSA encryption with 32-bit key:
    ' Public key (n, e) = (0x5518f65d, 0x11)
    ' Private key d = 0x2309cd31
    ' Message m = 0x35b9a3cb
    ' Encrypt c = m^e mod n = 35b9a3cb^11 mod 5518f65d = 528C41E5
    ' Decrypt m' = c^e mod n = 528C41E5^2309cd31 mod 5518f65d = 35B9A3CB
    strResult = mpModExp("35b9a3cb", "11", "5518f65d")
    Debug.Print strResult & " (expected 528C41E5)"
    strResult = mpModExp("528C41E5", "2309cd31", "5518f65d")
    Debug.Print strResult & " (expected 35B9A3CB)"
    
End Function

Public Function Test_RSA508()
' An example of an RSA calculation using mpModExp from
' "Some Examples of the PKCS Standards",
' An RSA Laboratories Technical Note,
' Burton S. Kaliski Jr., November 1, 1993.
' RSA key is 508 bits long.
' WARNING: this may take some time!
    Dim strMod As String
    Dim strExp As String
    Dim strPri As String
    Dim strMsg As String
    Dim strSig As String
    Dim strOK As String
    Dim strVer As String
    
    strMod = "0A66791DC6988168" & _
        "DE7AB77419BB7FB0" & _
        "C001C62710270075" & _
        "142942E19A8D8C51" & _
        "D053B3E3782A1DE5" & _
        "DC5AF4EBE9946817" & _
        "0114A1DFE67CDC9A" & _
        "9AF55D655620BBAB"
    strExp = "010001"
    strPri = "0123C5B61BA36EDB" & _
        "1D3679904199A89E" & _
        "A80C09B9122E1400" & _
        "C09ADCF7784676D0" & _
        "1D23356A7D44D6BD" & _
        "8BD50E94BFC723FA" & _
        "87D8862B75177691" & _
        "C11D757692DF8881"
    strMsg = "1FFFFFFFFFFFF" & _
        "FFFFFFFFFFFFFFFF" & _
        "FFFFFFFFFFFFFFFF" & _
        "FFFFFFFFFF003020" & _
        "300C06082A864886" & _
        "F70D020205000410" & _
        "DCA9ECF1C15C1BD2" & _
        "66AFF9C8799365CD"
    strOK = "6DB36CB18D3475B" & _
        "9C01DB3C78952808" & _
        "0279BBAEFF2B7D55" & _
        "8ED6615987C85186" & _
        "3F8A6C2CFFBC89C3" & _
        "F75A18D96B127C71" & _
        "7D54D0D8048DA8A0" & _
        "544626D17A2A8FBE"
        
    ' Sign, i.e. Encrypt with private key, s = m^d mod n
    Debug.Print "Calculating signature (be patient)..."
    strSig = mpModExp(strMsg, strPri, strMod)
    Debug.Print strSig
    If strSig = strOK Then
        Debug.Print "Hooray! Signature matches."
    Else
        Debug.Print "BOO! Signature was wrong."
    End If
    
    ' Verify, i.e. Decrypt with public key m' = s^e mod n
    Debug.Print "Calculating verification (be patient)..."
    strVer = mpModExp(strSig, strExp, strMod)
    Debug.Print strVer
    If strVer = strMsg Then
        Debug.Print "Hooray! Verification was OK."
    Else
        Debug.Print "BOO! Verification failed."
    End If

End Function

Public Function Test_Diffie_Hellman()
    ' A very simple example of Diffie-Hellman key exchange.
    ' CAUTION: Practical use requires numbers of 1000-2000+ bits in length
    ' and other checks on suitability of p and g.
    ' EXPLANATION OF SIMPLE DIFFIE-HELLMAN
    ' 1. Both parties agree to select and share a public generator, say, g = 3
    '    and public prime modulus  p = 0xc773218c737ec8ee993b4f2ded30f48edace915f
    ' 2. Alice selects private key x = 0x849dbd59069bff80cf30d052b74beeefc285b46f
    ' 3. Alice's public key is Ya = g^x mod p. Alice sends this to Bob.
    ' 4. To send a concealed, shared secret key to Alice, Bob picks a secret random number
    '    say, y = 0x40a2cf7390f76c1f2eef39c33eb61fb11811d528
    ' 5. Bob computes Yb = g^y mod p and sends this to Alice.
    ' 6. Bob can computes the shared key k = Ya^y mod p,
    '    to use for further communications with Alice
    ' 7. Alice can compute the same shared key k = Yb^x mod p,
    '    to use for further communications with Bob.
    ' Note: k = Ya^y = (g^x)^y = g^(xy) = Yb^x = (g^y)^x = g^(xy) mod p
    ' An eavesdropper only sees g, p, Ya and Yb.
    ' It is easy to compute Y=g^x mod p but it is
    ' difficult to compute x given g^x mod p.
    ' This is the discrete logarithm problem.
    
    Dim Ya As String
    Dim Yb As String
    Dim Ka As String
    Dim Kb As String
    
    ' Alice computes Ya = g^x mod p
    Ya = mpModExp("3", "849dbd59069bff80cf30d052b74beeefc285b46f", "c773218c737ec8ee993b4f2ded30f48edace915f")
    Debug.Print "Ya = " & Ya
    ' Bob computes Yb = g^y mod p
    Yb = mpModExp("3", "40a2cf7390f76c1f2eef39c33eb61fb11811d528", "c773218c737ec8ee993b4f2ded30f48edace915f")
    Debug.Print "Yb = " & Yb
    ' Alice computes the secret key k = Yb^x mod p
    Ka = mpModExp(Yb, "849dbd59069bff80cf30d052b74beeefc285b46f", "c773218c737ec8ee993b4f2ded30f48edace915f")
    Debug.Print "Ka = " & Ka
    ' Bob computes the secret key k = Ya^y mod p
    Kb = mpModExp(Ya, "40a2cf7390f76c1f2eef39c33eb61fb11811d528", "c773218c737ec8ee993b4f2ded30f48edace915f")
    Debug.Print "Kb = " & Kb
    If Ka <> Kb Then
        Debug.Print "ERROR: keys do not match!"
    Else
        Debug.Print "Keys match OK."
    End If
    
End Function


' *********************
' * EXPORTED FUNCTION *
' *********************

Public Function mpModExp(strBaseHex As String, strExponentHex As String, strModulusHex As String) As String
' Computes b^e mod m given input (b, e, m) in hex format.
' Returns result as a hex string with all leading zeroes removed.

' Store numbers as byte arrays with
' least-significant byte in x[len-1]
' and most-significant byte in x[1]
' x[0] is initially zero and is used for overflow
    
    Dim abBase() As Byte
    Dim abExponent() As Byte
    Dim abModulus() As Byte
    Dim abResult() As Byte
    Dim nLen As Integer
    Dim n As Integer
    
    ' Convert hex strings to arrays of bytes
    abBase = mpFromHex(strBaseHex)
    abExponent = mpFromHex(strExponentHex)
    abModulus = mpFromHex(strModulusHex)
    
    ' We require all byte arrays to be the same length
    ' with the first byte left as zero
    nLen = UBound(abModulus) + 1
    n = UBound(abExponent) + 1
    If n > nLen Then nLen = n
    n = UBound(abBase) + 1
    If n > nLen Then nLen = n
    Call FixArrayDim(abModulus, nLen)
    Call FixArrayDim(abExponent, nLen)
    Call FixArrayDim(abBase, nLen)
    '''Debug.Print "b=" & mpToHex(abBase)
    '''Debug.Print "e=" & mpToHex(abExponent)
    '''Debug.Print "m=" & mpToHex(abModulus)
    
    ' Do the business
    abResult = aModExp(abBase, abExponent, abModulus, nLen)
    
    ' Convert result to hex
    mpModExp = mpToHex(abResult)
    '''Debug.Print "r=" & mpModExp
    ' Strip leading zeroes
    For n = 1 To Len(mpModExp)
        If Mid$(mpModExp, n, 1) <> "0" Then
            Exit For
        End If
    Next
    ' FIX: [2009-02-04] Changed from >= to >
    If n > Len(mpModExp) Then
        ' Answer is zero
        mpModExp = "0"
    ElseIf n > 1 Then
        ' Zeroes to strip
        mpModExp = Mid$(mpModExp, n)
    End If
    
End Function

' **********************
' * INTERNAL FUNCTIONS *
' **********************
Public Function aModExp(abBase() As Byte, abExponent() As Byte, abModulus() As Byte, nLen As Integer) As Variant
' Computes a = b^e mod m and returns the result in a byte array as a VARIANT
    Dim a() As Byte
    Dim e() As Byte
    Dim s() As Byte
    Dim nBits As Long
    
    ' Perform right-to-left binary exponentiation
    ' 1. Set A = 1, S = b
    ReDim a(nLen - 1)
    a(nLen - 1) = 1
    ' NB s and e are trashed so use copies
    s = abBase
    e = abExponent
    ' 2. While e != 0 do:
    For nBits = nLen * 8 To 1 Step -1
        ' 2.1 if e is odd then A = A*S mod m
        If (e(nLen - 1) And &H1) <> 0 Then
            a = aModMult(a, s, abModulus, nLen)
        End If
        ' 2.2 e = e / 2
        Call DivideByTwo(e)
        ' 2.3 if e != 0 then S = S*S mod m
        If aIsZero(e, nLen) Then Exit For
        s = aModMult(s, s, abModulus, nLen)
        DoEvents
    Next
    
    ' 3. Return(A)
    aModExp = a
    
End Function

Private Function aModMult(abX() As Byte, abY() As Byte, abMod() As Byte, nLen As Integer) As Variant
' Returns w = (x * y) mod m
    Dim w() As Byte
    Dim x() As Byte
    Dim y() As Byte
    Dim nBits As Integer
    
    ' 1. Set w = 0, and temps x = abX, y = abY
    ReDim w(nLen - 1)
    x = abX
    y = abY
    ' 2. From LS bit to MS bit of X do:
    For nBits = nLen * 8 To 1 Step -1
        ' 2.1 if x is odd then w = (w + y) mod m
        If (x(nLen - 1) And &H1) <> 0 Then
            Call aModAdd(w, y, abMod, nLen)
        End If
        ' 2.2 x = x / 2
        Call DivideByTwo(x)
        ' 2.3 if x != 0 then y = (y + y) mod m
        If aIsZero(x, nLen) Then Exit For
        Call aModAdd(y, y, abMod, nLen)
    Next
    aModMult = w
    
End Function

Private Function aIsZero(a() As Byte, ByVal nLen As Integer) As Boolean
' Returns true if a is zero
    aIsZero = True
    Do While nLen > 0
        If a(nLen - 1) <> 0 Then
            aIsZero = False
            Exit Do
        End If
        nLen = nLen - 1
    Loop
End Function

Private Sub aModAdd(a() As Byte, b() As Byte, m() As Byte, nLen As Integer)
' Computes a = (a + b) mod m
    Dim i As Integer
    Dim d As Long
    ' 1. Add a = a + b
    d = 0
    For i = nLen - 1 To 0 Step -1
        d = CLng(a(i)) + CLng(b(i)) + d
        a(i) = CByte(d And &HFF)
        d = d \ &H100
    Next
    ' 2. If a > m then a = a - m
    For i = 0 To nLen - 2
        If a(i) <> m(i) Then
            Exit For
        End If
    Next
    If a(i) >= m(i) Then
        Call aSubtract(a, m, nLen)
    End If
    ' 3. Return a in-situ
            
End Sub

Private Sub aSubtract(a() As Byte, b() As Byte, nLen As Integer)
' Computes a = a - b
    Dim i As Integer
    Dim borrow As Long
    Dim d As Long   ' NB d is signed
    
    borrow = 0
    For i = nLen - 1 To 0 Step -1
        d = CLng(a(i)) - CLng(b(i)) - borrow
        If d < 0 Then
            d = d + &H100
            borrow = 1
        Else
            borrow = 0
        End If
        a(i) = CByte(d And &HFF)
    Next
    
End Sub

Private Sub DivideByTwo(ByRef x() As Byte)
' Divides multiple-precision integer x by 2 by shifting to right by one bit
    Dim d As Long
    Dim i As Integer
    d = 0
    For i = 0 To UBound(x)
        d = d Or x(i)
        x(i) = CByte((d \ 2) And &HFF)
        If (d And &H1) Then
            d = &H100
        Else
            d = 0
        End If
    Next
End Sub

Public Function mpToHex(abNum() As Byte) As String
' Returns a string containg the mp number abNum encoded in hex
' with leading zeroes trimmed.
    Dim i As Integer
    Dim sHex As String
    sHex = ""
    For i = 0 To UBound(abNum)
        If abNum(i) < &H10 Then
            sHex = sHex & "0" & Hex(abNum(i))
        Else
            sHex = sHex & Hex(abNum(i))
        End If
    Next
    mpToHex = sHex
End Function

Public Function mpFromHex(ByVal strHex As String) As Variant
' Converts number encoded in hex in big-endian order to a multi-precision integer
' Returns an array of bytes as a VARIANT
' containing number in big-endian order
' but with the first byte always zero
' strHex must only contain valid hex digits [0-9A-Fa-f]
' [2007-10-13] Changed direct >= <= comparisons with strings.
    Dim abData() As Byte
    Dim ib As Long
    Dim ic As Long
    Dim ch As String
    Dim nch As Long
    Dim nLen As Long
    Dim t As Long
    Dim v As Long
    Dim j As Integer
    
    ' Cope with odd # of digits, e.g. "fed" => "0fed"
    If Len(strHex) Mod 2 > 0 Then
        strHex = "0" & strHex
    End If
    nLen = Len(strHex) \ 2 + 1
    ReDim abData(nLen - 1)
    ib = 1
    j = 0
    For ic = 1 To Len(strHex)
        ch = Mid$(strHex, ic, 1)
        nch = Asc(ch)
        ''If ch >= "0" And ch <= "9" Then
        If nch >= &H30 And nch <= &H39 Then
            ''t = Asc(ch) - Asc("0")
            t = nch - &H30
        ''ElseIf ch >= "a" And ch <= "f" Then
        ElseIf nch >= &H61 And nch <= &H66 Then
            ''t = Asc(ch) - Asc("a") + 10
            t = nch - &H61 + 10
        ''ElseIf ch >= "A" And ch <= "F" Then
        ElseIf nch >= &H41 And nch <= &H46 Then
            ''t = Asc(ch) - Asc("A") + 10
            t = nch - &H41 + 10
        Else
            ' Invalid digit
            ' Flag error?
            Debug.Print "ERROR: Invalid Hex character found!"
            Exit Function
        End If
        ' Store byte value on every alternate digit
        If j = 0 Then
            ' v = t << 8
            v = t * &H10
            j = 1
        Else
            ' b[i] = (v | t) & 0xff
            abData(ib) = CByte((v Or t) And &HFF)
            ib = ib + 1
            j = 0
        End If
    Next
        
    mpFromHex = abData
End Function

Private Sub FixArrayDim(ByRef abData() As Byte, ByVal nLen As Long)
' Redim abData to be nLen bytes long with existing contents
' aligned at the RHS of the extended array
    Dim oLen As Long
    Dim i As Long
    
    oLen = UBound(abData) + 1
    If oLen > nLen Then
        ' Truncate
        ReDim Preserve abData(nLen - 1)
    ElseIf oLen < nLen Then
        ' Shift right
        ReDim Preserve abData(nLen - 1)
        For i = oLen - 1 To 0 Step -1
            abData(i + nLen - oLen) = abData(i)
        Next
        For i = 0 To nLen - oLen - 1
            abData(i) = 0
        Next
    End If
        
End Sub

Public Function TestConvFromHex()
    Dim abData() As Byte
    
    abData = mpFromHex("deadbeef")
    Debug.Print mpToHex(abData)
    abData = mpFromHex("FfeE01")
    Debug.Print mpToHex(abData)
    abData = mpFromHex("1")
    Debug.Print mpToHex(abData)
    
End Function

