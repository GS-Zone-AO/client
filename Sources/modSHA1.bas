Attribute VB_Name = "modSHA1"
Option Explicit

' Source: https://gist.github.com/anonymous/573a875dac68a4af560d


Private Type FourBytes
    a As Byte
    b As Byte
    c As Byte
    d As Byte
End Type
Private Type OneLong
    l As Long
End Type

Private Function HexDefaultSHA1(Message() As Byte) As String
 Dim H1 As Long, H2 As Long, H3 As Long, H4 As Long, H5 As Long
 DefaultSHA1 Message, H1, H2, H3, H4, H5
 HexDefaultSHA1 = DecToHex5(H1, H2, H3, H4, H5)
End Function

Private Function HexSHA1(Message() As Byte, ByVal Key1 As Long, ByVal Key2 As Long, ByVal Key3 As Long, ByVal Key4 As Long) As String
 Dim H1 As Long, H2 As Long, H3 As Long, H4 As Long, H5 As Long
 xSHA1 Message, Key1, Key2, Key3, Key4, H1, H2, H3, H4, H5
 HexSHA1 = DecToHex5(H1, H2, H3, H4, H5)
End Function

Private Sub DefaultSHA1(Message() As Byte, H1 As Long, H2 As Long, H3 As Long, H4 As Long, H5 As Long)
 xSHA1 Message, &H5A827999, &H6ED9EBA1, &H8F1BBCDC, &HCA62C1D6, H1, H2, H3, H4, H5
End Sub

Private Sub xSHA1(Message() As Byte, ByVal Key1 As Long, ByVal Key2 As Long, ByVal Key3 As Long, ByVal Key4 As Long, H1 As Long, H2 As Long, H3 As Long, H4 As Long, H5 As Long)

 Dim U As Long, P As Long
 Dim FB As FourBytes, OL As OneLong
 Dim i As Integer
 Dim w(80) As Long
 Dim a As Long, b As Long, c As Long, d As Long, E As Long
 Dim t As Long

 H1 = &H67452301: H2 = &HEFCDAB89: H3 = &H98BADCFE: H4 = &H10325476: H5 = &HC3D2E1F0

 U = UBound(Message) + 1: OL.l = U32ShiftLeft3(U): a = U \ &H20000000: LSet FB = OL 'U32ShiftRight29(U)

 ReDim Preserve Message(0 To (U + 8 And -64) + 63)
 Message(U) = 128

 U = UBound(Message)
 Message(U - 4) = a
 Message(U - 3) = FB.d
 Message(U - 2) = FB.c
 Message(U - 1) = FB.b
 Message(U) = FB.a

 While P < U
     For i = 0 To 15
         FB.d = Message(P)
         FB.c = Message(P + 1)
         FB.b = Message(P + 2)
         FB.a = Message(P + 3)
         LSet OL = FB
         w(i) = OL.l
         P = P + 4
     Next i

     For i = 16 To 79
         w(i) = U32RotateLeft1(w(i - 3) Xor w(i - 8) Xor w(i - 14) Xor w(i - 16))
     Next i

     a = H1: b = H2: c = H3: d = H4: E = H5

     For i = 0 To 19
         t = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), E), w(i)), Key1), ((b And c) Or ((Not b) And d)))
         E = d: d = c: c = U32RotateLeft30(b): b = a: a = t
     Next i
     For i = 20 To 39
         t = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), E), w(i)), Key2), (b Xor c Xor d))
         E = d: d = c: c = U32RotateLeft30(b): b = a: a = t
     Next i
     For i = 40 To 59
         t = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), E), w(i)), Key3), ((b And c) Or (b And d) Or (c And d)))
         E = d: d = c: c = U32RotateLeft30(b): b = a: a = t
     Next i
     For i = 60 To 79
         t = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), E), w(i)), Key4), (b Xor c Xor d))
         E = d: d = c: c = U32RotateLeft30(b): b = a: a = t
     Next i

     H1 = U32Add(H1, a): H2 = U32Add(H2, b): H3 = U32Add(H3, c): H4 = U32Add(H4, d): H5 = U32Add(H5, E)
 Wend
End Sub

Private Function U32Add(ByVal a As Long, ByVal b As Long) As Long
 If (a Xor b) < 0 Then
     U32Add = a + b
 Else
     U32Add = (a Xor &H80000000) + b Xor &H80000000
 End If
End Function

Private Function U32ShiftLeft3(ByVal a As Long) As Long
 U32ShiftLeft3 = (a And &HFFFFFFF) * 8
 If a And &H10000000 Then U32ShiftLeft3 = U32ShiftLeft3 Or &H80000000
End Function

Private Function U32ShiftRight29(ByVal a As Long) As Long
 U32ShiftRight29 = (a And &HE0000000) \ &H20000000 And 7
End Function

Private Function U32RotateLeft1(ByVal a As Long) As Long
 U32RotateLeft1 = (a And &H3FFFFFFF) * 2
 If a And &H40000000 Then U32RotateLeft1 = U32RotateLeft1 Or &H80000000
 If a And &H80000000 Then U32RotateLeft1 = U32RotateLeft1 Or 1
End Function

Private Function U32RotateLeft5(ByVal a As Long) As Long
 U32RotateLeft5 = (a And &H3FFFFFF) * 32 Or (a And &HF8000000) \ &H8000000 And 31
 If a And &H4000000 Then U32RotateLeft5 = U32RotateLeft5 Or &H80000000
End Function

Private Function U32RotateLeft30(ByVal a As Long) As Long
 U32RotateLeft30 = (a And 1) * &H40000000 Or (a And &HFFFC) \ 4 And &H3FFFFFFF
 If a And 2 Then U32RotateLeft30 = U32RotateLeft30 Or &H80000000
End Function

Private Function DecToHex5(ByVal H1 As Long, ByVal H2 As Long, ByVal H3 As Long, ByVal H4 As Long, ByVal H5 As Long) As String
 Dim H As String, l As Long
 DecToHex5 = "00000000 00000000 00000000 00000000 00000000"
 H = Hex(H1): l = Len(H): Mid(DecToHex5, 9 - l, l) = H
 H = Hex(H2): l = Len(H): Mid(DecToHex5, 18 - l, l) = H
 H = Hex(H3): l = Len(H): Mid(DecToHex5, 27 - l, l) = H
 H = Hex(H4): l = Len(H): Mid(DecToHex5, 36 - l, l) = H
 H = Hex(H5): l = Len(H): Mid(DecToHex5, 45 - l, l) = H
End Function

Public Function SHA1TRUNC(str)
  Dim i As Integer
  Dim arr() As Byte
  ReDim arr(0 To Len(str) - 1) As Byte
  Const cutoff As Integer = 6
  
  For i = 0 To Len(str) - 1
   arr(i) = Asc(mid(str, i + 1, 1))
  Next i
  
  SHA1TRUNC = Replace(LCase(HexDefaultSHA1(arr)), " ", "")
  SHA1TRUNC = Left(SHA1TRUNC, cutoff)
  
End Function
