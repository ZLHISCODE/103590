VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMD5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes As Long)

'/* Constants for MD5Transform routine. */
Private Const S11 As Long = 7
Private Const S12 As Long = 12
Private Const S13 As Long = 17
Private Const S14 As Long = 22
Private Const S21 As Long = 5
Private Const S22 As Long = 9
Private Const S23 As Long = 14
Private Const S24 As Long = 20
Private Const S31 As Long = 4
Private Const S32 As Long = 11
Private Const S33 As Long = 16
Private Const S34 As Long = 23
Private Const S41 As Long = 6
Private Const S42 As Long = 10
Private Const S43 As Long = 15
Private Const S44 As Long = 21

Private LongBits(0 To 31) As Long

Private x(0 To 15) As Long
Private State(0 To 3) As Long
Private Count As Long

'/* ROTATE_LEFT rotates x left n bits.*/
Private Function ROTATE_LEFT(ByVal x As Long, ByVal n As Long) As Long
    Dim After_n As Long, Before_n As Long
    After_n = x And (LongBits(&H1F Xor n) - 1)
    Before_n = (x And &H7FFFFFFF) \ LongBits(32 - n)
    ROTATE_LEFT = After_n * LongBits(n) Or Before_n
    If x And LongBits(&H1F Xor n) Then ROTATE_LEFT = ROTATE_LEFT Or &H80000000
    If x And &H80000000 Then ROTATE_LEFT = ROTATE_LEFT Or LongBits(n - 1)
End Function

'/* FF, GG, HH, and II transformations for rounds 1, 2, 3, and 4.
'Rotation is separate from addition to prevent recomputation.*/
Private Sub FF(ByRef a As Long, _
               ByVal b As Long, _
               ByVal c As Long, _
               ByVal d As Long, _
               ByVal x As Long, _
               ByVal s As Long, _
               ByVal ac As Long)
    a = LongOverflowAdd4(a, b And c Or Not b And d, x, ac)
    a = ROTATE_LEFT(a, s)
    a = LongOverflowAdd(a, b)
End Sub

Private Sub GG(ByRef a As Long, _
               ByVal b As Long, _
               ByVal c As Long, _
               ByVal d As Long, _
               ByVal x As Long, _
               ByVal s As Long, _
               ByVal ac As Long)
    a = LongOverflowAdd4(a, b And d Or c And Not d, x, ac)
    a = ROTATE_LEFT(a, s)
    a = LongOverflowAdd(a, b)
End Sub

Private Sub HH(ByRef a As Long, _
               ByVal b As Long, _
               ByVal c As Long, _
               ByVal d As Long, _
               ByVal x As Long, _
               ByVal s As Long, _
               ByVal ac As Long)
    a = LongOverflowAdd4(a, b Xor c Xor d, x, ac)
    a = ROTATE_LEFT(a, s)
    a = LongOverflowAdd(a, b)
End Sub

Private Sub II(ByRef a As Long, _
               ByVal b As Long, _
               ByVal c As Long, _
               ByVal d As Long, _
               ByVal x As Long, _
               ByVal s As Long, _
               ByVal ac As Long)
    a = LongOverflowAdd4(a, c Xor b Or Not d, x, ac)
    a = ROTATE_LEFT(a, s)
    a = LongOverflowAdd(a, b)
End Sub



'/* Construct a MD5 object with a input buffer. */
Public Function DigestByteToHexStr(SourceByte() As Byte) As String
    Call ReSet
    Call Update(SourceByte, UBound(SourceByte) + 1)
    Call Final
    DigestByteToHexStr = LongToString(State(0)) & LongToString(State(1)) & LongToString(State(2)) & LongToString(State(3))
End Function

'/* Construct a MD5 object with a string. */
Public Function DigestStrToHexStr(ByRef SourceString As String) As String
    DigestStrToHexStr = DigestByteToHexStr(StrConv(SourceString, vbFromUnicode))
End Function

'/* Construct a MD5 object with a file. */
'Public Function DigestFileToHexStr(ByRef FilePath As String) As String
'    If Len(Dir$(FilePath)) Then
'        Dim F1 As Long, FileBuffer(0 To 65535) As Byte, i As Long, j As Long
'        F1 = FreeFile
'        Open FilePath For Binary Access Read As #F1
'        Call ReSet
'
'        For i = 1 To LOF(F1) \ 65536
'            Form1.Cls
'            Form1.Print i
'            DoEvents
'            Get #F1, , FileBuffer
'            Call Update(FileBuffer, UBound(FileBuffer) + 1)
'        Next
'        i = LOF(F1) And &HFFFF&
'        If i Then
'            Get #F1, , FileBuffer
'            Call Update(FileBuffer, i)
'        End If
'
'        Count = LOF(F1)
'        Close #F1
'        Call Final
'        DigestFileToHexStr = LongToString(State(0)) & LongToString(State(1)) & LongToString(State(2)) & LongToString(State(3))
'    End If
'End Function

'/* Reset the calculate state */
Private Sub ReSet()
    '/* reset number of bits. */
    Count = 0
    '/* Load magic initialization constants. */
    State(0) = &H67452301
    State(1) = &HEFCDAB89
    State(2) = &H98BADCFE
    State(3) = &H10325476
End Sub

'/* MD5 block update operation. Continues an MD5 message-digest
'operation, processing another message block, and updating thecontext.*/
Private Sub Update(InputArr() As Byte, ByVal Length As Long)
    Dim i As Long, index As Long, PartLen As Long
    '/* Compute number of bytes mod 64 */
    index = Count And &H3F
    '/* update number of bits */
    Count = Count + Length
    PartLen = 64 - index
    Dim Address As Long
    Address = VarPtr(x(0))
    '/* transform as many times as possible. */
    If Length >= PartLen Then
        Call CopyMemory(ByVal Address + index, InputArr(0), PartLen)
        Call Transform '(Buffer)
        For i = PartLen To Length - 64 Step 64
            Call CopyMemory(ByVal Address + index, InputArr(i), 64)
            Call Transform '(Buffer)
        Next
        index = 0
    'Else
        'i = 0
    End If
    '/* Buffer remaining input */
    If Length - i > 0 Then Call CopyMemory(ByVal Address + index, InputArr(i), Length - i)
End Sub


'/* MD5 finalization. Ends an MD5 message-_digest operation, writing the
'the message _digest and zeroizing the context.*/
Private Sub Final()
    'byte bits[8];
    'uint32 oldState[4];
    'uint32 oldCount[2];
    Dim Padding(0 To 63) As Byte
    Dim index As Long, PadLen As Long
    Dim TempCount As Long
    TempCount = Count
    '/* Save current state and count. */
    'memcpy(oldState, _state, 16);
    'memcpy(oldCount, _count, 8);

    '/* Save number of bits */
    'encode(_count, bits, 8);
    'Call ZeroMemory(Padding(1), 7)
    Padding(0) = &H80
    '/* Pad out to 56 mod 64. */
    index = Count And &H3F
    PadLen = IIf(index < 56, 56 - index, 120 - index)
    Call Update(Padding, PadLen)

    '/* Append length (before padding) */
    Call CopyMemory(Padding(0), TempCount * 0.0008@, 8)
    Call Update(Padding, 8)

    '/* Store state in digest */
    'encode(_state, _digest, 16);

    '/* Restore current state and count. */
    'memcpy(_state, oldState, 16);
    'memcpy(_count, oldCount, 8);
End Sub

'/* MD5 basic transformation. Transforms _state based on block. */
Private Sub Transform() 'Block() As Byte)

    Dim a As Long, b As Long, c As Long, d As Long
    a = State(0)
    b = State(1)
    c = State(2)
    d = State(3)
    'Dim x(0 To 15) As Long
    'Call Decode(Block, x, 64)
    'Call CopyMemory(x(0), Block(0), 64)
    '/* Round 1 */
    Call FF(a, b, c, d, x(0), S11, &HD76AA478)  ' /* 1 */
    Call FF(d, a, b, c, x(1), S12, &HE8C7B756)  ' /* 2 */
    Call FF(c, d, a, b, x(2), S13, &H242070DB)  ' /* 3 */
    Call FF(b, c, d, a, x(3), S14, &HC1BDCEEE)  ' /* 4 */
    Call FF(a, b, c, d, x(4), S11, &HF57C0FAF)  ' /* 5 */
    Call FF(d, a, b, c, x(5), S12, &H4787C62A)  ' /* 6 */
    Call FF(c, d, a, b, x(6), S13, &HA8304613)  ' /* 7 */
    Call FF(b, c, d, a, x(7), S14, &HFD469501)  ' /* 8 */
    Call FF(a, b, c, d, x(8), S11, &H698098D8)  ' /* 9 */
    Call FF(d, a, b, c, x(9), S12, &H8B44F7AF)  ' /* 10 */
    Call FF(c, d, a, b, x(10), S13, &HFFFF5BB1) ' /* 11 */
    Call FF(b, c, d, a, x(11), S14, &H895CD7BE) ' /* 12 */
    Call FF(a, b, c, d, x(12), S11, &H6B901122) ' /* 13 */
    Call FF(d, a, b, c, x(13), S12, &HFD987193) ' /* 14 */
    Call FF(c, d, a, b, x(14), S13, &HA679438E) ' /* 15 */
    Call FF(b, c, d, a, x(15), S14, &H49B40821) ' /* 16 */

    '/* Round 2 */
    Call GG(a, b, c, d, x(1), S21, &HF61E2562)  ' /* 17 */
    Call GG(d, a, b, c, x(6), S22, &HC040B340)  ' /* 18 */
    Call GG(c, d, a, b, x(11), S23, &H265E5A51) ' /* 19 */
    Call GG(b, c, d, a, x(0), S24, &HE9B6C7AA)  ' /* 20 */
    Call GG(a, b, c, d, x(5), S21, &HD62F105D)  ' /* 21 */
    Call GG(d, a, b, c, x(10), S22, &H2441453)  ' /* 22 */
    Call GG(c, d, a, b, x(15), S23, &HD8A1E681) ' /* 23 */
    Call GG(b, c, d, a, x(4), S24, &HE7D3FBC8)  ' /* 24 */
    Call GG(a, b, c, d, x(9), S21, &H21E1CDE6)  ' /* 25 */
    Call GG(d, a, b, c, x(14), S22, &HC33707D6) ' /* 26 */
    Call GG(c, d, a, b, x(3), S23, &HF4D50D87)  ' /* 27 */
    Call GG(b, c, d, a, x(8), S24, &H455A14ED)  ' /* 28 */
    Call GG(a, b, c, d, x(13), S21, &HA9E3E905) ' /* 29 */
    Call GG(d, a, b, c, x(2), S22, &HFCEFA3F8)  ' /* 30 */
    Call GG(c, d, a, b, x(7), S23, &H676F02D9)  ' /* 31 */
    Call GG(b, c, d, a, x(12), S24, &H8D2A4C8A) ' /* 32 */

    '/* Round 3 */
    Call HH(a, b, c, d, x(5), S31, &HFFFA3942)  ' /* 33 */
    Call HH(d, a, b, c, x(8), S32, &H8771F681)  ' /* 34 */
    Call HH(c, d, a, b, x(11), S33, &H6D9D6122) ' /* 35 */
    Call HH(b, c, d, a, x(14), S34, &HFDE5380C) ' /* 36 */
    Call HH(a, b, c, d, x(1), S31, &HA4BEEA44)  ' /* 37 */
    Call HH(d, a, b, c, x(4), S32, &H4BDECFA9)  ' /* 38 */
    Call HH(c, d, a, b, x(7), S33, &HF6BB4B60)  ' /* 39 */
    Call HH(b, c, d, a, x(10), S34, &HBEBFBC70) ' /* 40 */
    Call HH(a, b, c, d, x(13), S31, &H289B7EC6) ' /* 41 */
    Call HH(d, a, b, c, x(0), S32, &HEAA127FA)  ' /* 42 */
    Call HH(c, d, a, b, x(3), S33, &HD4EF3085)  ' /* 43 */
    Call HH(b, c, d, a, x(6), S34, &H4881D05)   ' /* 44 */
    Call HH(a, b, c, d, x(9), S31, &HD9D4D039)  ' /* 45 */
    Call HH(d, a, b, c, x(12), S32, &HE6DB99E5) ' /* 46 */
    Call HH(c, d, a, b, x(15), S33, &H1FA27CF8) ' /* 47 */
    Call HH(b, c, d, a, x(2), S34, &HC4AC5665)  ' /* 48 */

    '/* Round 4 */
    Call II(a, b, c, d, x(0), S41, &HF4292244)  ' /* 49 */
    Call II(d, a, b, c, x(7), S42, &H432AFF97)  ' /* 50 */
    Call II(c, d, a, b, x(14), S43, &HAB9423A7) ' /* 51 */
    Call II(b, c, d, a, x(5), S44, &HFC93A039)  ' /* 52 */
    Call II(a, b, c, d, x(12), S41, &H655B59C3) ' /* 53 */
    Call II(d, a, b, c, x(3), S42, &H8F0CCC92)  ' /* 54 */
    Call II(c, d, a, b, x(10), S43, &HFFEFF47D) ' /* 55 */
    Call II(b, c, d, a, x(1), S44, &H85845DD1)  ' /* 56 */
    Call II(a, b, c, d, x(8), S41, &H6FA87E4F)  ' /* 57 */
    Call II(d, a, b, c, x(15), S42, &HFE2CE6E0) ' /* 58 */
    Call II(c, d, a, b, x(6), S43, &HA3014314)  ' /* 59 */
    Call II(b, c, d, a, x(13), S44, &H4E0811A1) ' /* 60 */
    Call II(a, b, c, d, x(4), S41, &HF7537E82)  ' /* 61 */
    Call II(d, a, b, c, x(11), S42, &HBD3AF235) ' /* 62 */
    Call II(c, d, a, b, x(2), S43, &H2AD7D2BB)  ' /* 63 */
    Call II(b, c, d, a, x(9), S44, &HEB86D391)  ' /* 64 */

    State(0) = LongOverflowAdd(State(0), a)
    State(1) = LongOverflowAdd(State(1), b)
    State(2) = LongOverflowAdd(State(2), c)
    State(3) = LongOverflowAdd(State(3), d)
End Sub

'/* Encodes input (ulong) into output (byte). Assumes length is a multiple of 4.*/
'Private Sub Encode(InputArr() As Long, OutputArr() As Byte, Length As Long)
    'Call CopyMemory(OutputArr(0), InputArr(0), 64)
'End Sub

'/* Decodes input (byte) into output (ulong). Assumes length is a multiple of 4.*/
'Private Sub Decode(InputArr() As Byte, OutputArr() As Long, Length As Long)
    'Call CopyMemory(OutputArr(0), InputArr(0), 64)
'End Sub

'Private Function LongOverflowAdd(ByVal a As Double, ByVal b As Double) As Long
'    Dim t As Double
'    t = a + b
'    If t < &H80000000 Then t = t + 4294967296#
'    If t > &H7FFFFFFF Then t = t - 4294967296#
'    LongOverflowAdd = t
'End Function

Private Function LongOverflowAdd(ByVal a As Long, ByVal b As Long) As Long
    If (a Xor b) And &H80000000 Then
        LongOverflowAdd = a + b
    Else
        LongOverflowAdd = (a Xor &H80000000) + b Xor &H80000000
    End If
End Function



'Private Function LongOverflowAdd4(ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double) As Long
'    Dim t As Double
'    t = a + b + c + d
'    If t < &H80000000 Then t = t + 4294967296#
'    If t < &H80000000 Then t = t + 4294967296#
'    If t > &H7FFFFFFF Then t = t - 4294967296#
'    If t > &H7FFFFFFF Then t = t - 4294967296#
'    LongOverflowAdd4 = t
'End Function
Private Function LongOverflowAdd4(ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long) As Long
    'LongOverflowAdd4 = LongOverflowAdd(LongOverflowAdd(a, b), LongOverflowAdd(c, d))

    If (a Xor b) And &H80000000 Then
        a = a + b
    Else
        a = (a Xor &H80000000) + b Xor &H80000000
    End If
    If (c Xor d) And &H80000000 Then
        c = c + d
    Else
        c = (c Xor &H80000000) + d Xor &H80000000
    End If
    If (a Xor c) And &H80000000 Then
        LongOverflowAdd4 = a + c
    Else
        LongOverflowAdd4 = (a Xor &H80000000) + c Xor &H80000000
    End If

End Function

Private Function LongToString(ByVal Value As Long) As String
    Dim Arr(0 To 3) As Byte, t As Byte
    Call CopyMemory(Arr(0), Value, 4)
    t = Arr(0)
    Arr(0) = Arr(3)
    Arr(3) = t
    t = Arr(1)
    Arr(1) = Arr(2)
    Arr(2) = t
    Call CopyMemory(Value, Arr(0), 4)
    LongToString = Hex$(Value)
    If Len(LongToString) < 8 Then LongToString = String$(8 - Len(LongToString), "0") & LongToString
End Function


Private Sub Class_Initialize()
    Dim i As Long
    LongBits(0) = 1
    For i = 1 To 30
        LongBits(i) = LongBits(i - 1) * 2
    Next
    LongBits(31) = &H80000000
End Sub
Public Function Unicode_To_UTF8(ByRef Unicode() As Byte) As Byte()
    Dim i As Long, j As Long, k As Long
    Dim Buffer() As Byte
    Dim Length As Long
    Dim Temp As Long
    Length = (UBound(Unicode) - LBound(Unicode) + 1) \ 2
    If Length > 0 Then ReDim Buffer(0 To Length * 3 - 1) Else Buffer = vbNullString
    For i = LBound(Unicode) To UBound(Unicode) Step 2
        Temp = Unicode(i) Or Unicode(i + 1) * &H100&
        k = 1
        If Temp >= &H800& Then
            Buffer(j + 2) = Temp And &H3F& Or &H80&
            Buffer(j + 0) = Buffer(j + 0) Or &HE0&
            Temp = Temp \ &H40&
            k = k + 1
        End If
        If Temp >= &H80& Then
            Buffer(j + 1) = Temp And &H3F& Or &H80&
            Buffer(j + 0) = Buffer(j + 0) Or &HC0&
            Temp = Temp \ &H40&
            k = k + 1
        End If
        
        Buffer(j + 0) = Buffer(j + 0) Or Temp
        j = j + k
    Next
    If j Then ReDim Preserve Buffer(0 To j - 1)
    Unicode_To_UTF8 = Buffer
End Function



