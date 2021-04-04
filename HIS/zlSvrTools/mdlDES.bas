Attribute VB_Name = "mdlDES"
Option Explicit
'DES加密需要变量-------------------------------------------
'置换表
Private IP(63) As Byte 'IP置换码
Private IP_1(63) As Byte 'IP-1置换码


Private E(47) As Byte 'E膨胀码
Private p(31) As Byte 'P变换码


Private s1(3, 15) As Byte 'S盒1
Private s2(3, 15) As Byte 'S盒2
Private s3(3, 15) As Byte 'S盒3
Private S4(3, 15) As Byte 'S盒4
Private S5(3, 15) As Byte 'S盒5
Private S6(3, 15) As Byte 'S盒6
Private S7(3, 15) As Byte 'S盒7
Private S8(3, 15) As Byte 'S盒8

Private PC_1(55) As Byte
Private PC_2(47) As Byte

Private Lsi(16) As Byte '循环左移位
'------------------------------------------------------------


'sCode 待加密字串
'sKey 密钥字串(前8位有效)
Public Sub DES_Encode(sCode() As Byte, ByRef bReturn() As Byte, ByVal sKey As String)

    Dim tempKey() As Byte '存放密钥
    Dim BinKey(63) As Byte '64位二进制原始密钥
    Dim KeyPC_1(55) As Byte '存放56位密钥
    Dim tempCode(7) As Byte '存放8位原始明文
    Dim tempReturn(7) As Byte '存放8位密文
    
    Dim BinCode(63) As Byte '存放64位明文
    Dim CodeIP(63) As Byte '存放IP置换结果
    Dim CodeE(47) As Byte 'E膨胀结果
    Dim CodeP(31) As Byte 'P变换结果
    Dim RetS(47) As Byte 'S盒运算32位结果
    Dim S(7) As Byte 'S盒运算8个结果
    Dim CodeS1(5) As Byte: Dim CodeS2(5) As Byte: Dim CodeS3(5) As Byte: Dim CodeS4(5) As Byte
    Dim CodeS5(5) As Byte: Dim CodeS6(5) As Byte: Dim CodeS7(5) As Byte: Dim CodeS8(5) As Byte
    
    Dim L0(31) As Byte: Dim R0(31) As Byte
    Dim l1(31) As Byte: Dim R1(31) As Byte
    Dim L2(31) As Byte: Dim R2(31) As Byte
    Dim L3(31) As Byte: Dim R3(31) As Byte
    Dim L4(31) As Byte: Dim R4(31) As Byte
    Dim L5(31) As Byte: Dim R5(31) As Byte
    Dim L6(31) As Byte: Dim R6(31) As Byte
    Dim L7(31) As Byte: Dim R7(31) As Byte
    Dim L8(31) As Byte: Dim R8(31) As Byte
    Dim L9(31) As Byte: Dim R9(31) As Byte
    Dim L10(31) As Byte: Dim R10(31) As Byte
    Dim L11(31) As Byte: Dim R11(31) As Byte
    Dim L12(31) As Byte: Dim R12(31) As Byte
    Dim L13(31) As Byte: Dim R13(31) As Byte
    Dim L14(31) As Byte: Dim R14(31) As Byte
    Dim L15(31) As Byte: Dim R15(31) As Byte
    Dim L16(31) As Byte: Dim R16(31) As Byte
    
    Dim C0(27) As Byte: Dim D0(27) As Byte '16个密钥
    Dim C1(27) As Byte: Dim D1(27) As Byte
    Dim C2(27) As Byte: Dim D2(27) As Byte:
    Dim C3(27) As Byte: Dim D3(27) As Byte:
    Dim C4(27) As Byte: Dim D4(27) As Byte:
    Dim C5(27) As Byte: Dim D5(27) As Byte:
    Dim C6(27) As Byte: Dim D6(27) As Byte:
    Dim C7(27) As Byte: Dim D7(27) As Byte:
    Dim C8(27) As Byte: Dim D8(27) As Byte:
    Dim C9(27) As Byte: Dim D9(27) As Byte:
    Dim C10(27) As Byte: Dim D10(27) As Byte:
    Dim C11(27) As Byte: Dim D11(27) As Byte:
    Dim C12(27) As Byte: Dim D12(27) As Byte:
    Dim C13(27) As Byte: Dim D13(27) As Byte:
    Dim C14(27) As Byte: Dim D14(27) As Byte:
    Dim C15(27) As Byte: Dim D15(27) As Byte:
    Dim C16(27) As Byte: Dim D16(27) As Byte:
    
    Dim C_D(55) As Byte 'Cn,Dn合并后的存放处
    
    Dim K1(47) As Byte: Dim K2(47) As Byte: Dim K3(47) As Byte: Dim K4(47) As Byte:
    Dim K5(47) As Byte: Dim K6(47) As Byte: Dim K7(47) As Byte: Dim K8(47) As Byte:
    Dim K9(47) As Byte: Dim K10(47) As Byte: Dim K11(47) As Byte: Dim K12(47) As Byte:
    Dim K13(47) As Byte: Dim K14(47) As Byte: Dim K15(47) As Byte: Dim K16(47) As Byte:
    
    Dim i As Integer
    Dim j As Integer
    
    Call InitDES
    
    '取密钥的前8字节
    tempKey = StrConv(sKey, vbFromUnicode)
    ReDim Preserve tempKey(7)
    
    For i = 0 To 7
    BinKey(i * 8 + 0) = (tempKey(i) And &H80) \ &H80
    BinKey(i * 8 + 1) = (tempKey(i) And &H40) \ &H40
    BinKey(i * 8 + 2) = (tempKey(i) And &H20) \ &H20
    BinKey(i * 8 + 3) = (tempKey(i) And &H10) \ &H10
    BinKey(i * 8 + 4) = (tempKey(i) And &H8) \ &H8
    BinKey(i * 8 + 5) = (tempKey(i) And &H4) \ &H4
    BinKey(i * 8 + 6) = (tempKey(i) And &H2) \ &H2
    BinKey(i * 8 + 7) = (tempKey(i) And &H1) \ &H1
    Next
    
    'PC_1转换
    For i = 0 To 55
    KeyPC_1(i) = BinKey(PC_1(i))
    Next
    
    '生成C0,D0
    For i = 0 To 27
    C0(i) = KeyPC_1(i)
    D0(i) = KeyPC_1(i + 28)
    Next
    
    '***************************************************K1
    '生成C1,D1
    For i = 0 To 26
    C1(i) = C0(i + Lsi(1))
    D1(i) = D0(i + Lsi(1))
    Next
    C1(27) = C0(0)
    D1(27) = D0(0)
    
    '组合C1,D1成C_D
    For i = 0 To 27
    C_D(i) = C1(i)
    C_D(i + 28) = D1(i)
    Next
    
    'PC_2转换,生成K1
    For i = 0 To 47
    K1(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K2
    '生成C2,D2
    For i = 0 To 26
    C2(i) = C1(i + Lsi(2))
    D2(i) = D1(i + Lsi(2))
    Next
    C2(27) = C1(0)
    D2(27) = D1(0)
    
    '组合C2,D2成C_D
    For i = 0 To 27
    C_D(i) = C2(i)
    C_D(i + 28) = D2(i)
    Next
    
    'PC_2转换,生成K2
    For i = 0 To 47
    K2(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K3
    '生成C3,D3
    For i = 0 To 25
    C3(i) = C2(i + Lsi(3))
    D3(i) = D2(i + Lsi(3))
    Next
    C3(26) = C2(0)
    D3(26) = D2(0)
    C3(27) = C2(1)
    D3(27) = D2(1)
    
    '组合C3,D3成C_D
    For i = 0 To 27
    C_D(i) = C3(i)
    C_D(i + 28) = D3(i)
    Next
    
    'PC_2转换,生成K3
    For i = 0 To 47
    K3(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K4
    '生成C4,D4
    For i = 0 To 25
    C4(i) = C3(i + Lsi(4))
    D4(i) = D3(i + Lsi(4))
    Next
    C4(26) = C3(0)
    D4(26) = D3(0)
    C4(27) = C3(1)
    D4(27) = D3(1)
    
    '组合C4,D4成C_D
    For i = 0 To 27
    C_D(i) = C4(i)
    C_D(i + 28) = D4(i)
    Next
    
    'PC_2转换,生成K4
    For i = 0 To 47
    K4(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K5
    '生成C5,D5
    For i = 0 To 25
    C5(i) = C4(i + Lsi(5))
    D5(i) = D4(i + Lsi(5))
    Next
    C5(26) = C4(0)
    D5(26) = D4(0)
    C5(27) = C4(1)
    D5(27) = D4(1)
    
    '组合C5,D5成C_D
    For i = 0 To 27
    C_D(i) = C5(i)
    C_D(i + 28) = D5(i)
    Next
    
    'PC_2转换,生成K5
    For i = 0 To 47
    K5(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K6
    '生成C6,D6
    For i = 0 To 25
    C6(i) = C5(i + Lsi(6))
    D6(i) = D5(i + Lsi(6))
    Next
    C6(26) = C5(0)
    D6(26) = D5(0)
    C6(27) = C5(1)
    D6(27) = D5(1)
    
    '组合C6,D6成C_D
    For i = 0 To 27
    C_D(i) = C6(i)
    C_D(i + 28) = D6(i)
    Next
    
    'PC_2转换,生成K6
    For i = 0 To 47
    K6(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K7
    '生成C7,D7
    For i = 0 To 25
    C7(i) = C6(i + Lsi(7))
    D7(i) = D6(i + Lsi(7))
    Next
    C7(26) = C6(0)
    D7(26) = D6(0)
    C7(27) = C6(1)
    D7(27) = D6(1)
    
    '组合C7,D7成C_D
    For i = 0 To 27
    C_D(i) = C7(i)
    C_D(i + 28) = D7(i)
    Next
    
    'PC_2转换,生成K7
    For i = 0 To 47
    K7(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K8
    '生成C8,D8
    For i = 0 To 25
    C8(i) = C7(i + Lsi(8))
    D8(i) = D7(i + Lsi(8))
    Next
    C8(26) = C7(0)
    D8(26) = D7(0)
    C8(27) = C7(1)
    D8(27) = D7(1)
    
    '组合C8,D8成C_D
    For i = 0 To 27
    C_D(i) = C8(i)
    C_D(i + 28) = D8(i)
    Next
    
    'PC_2转换,生成K8
    For i = 0 To 47
    K8(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K9
    '生成C9,D9
    For i = 0 To 26
    C9(i) = C8(i + Lsi(9))
    D9(i) = D8(i + Lsi(9))
    Next
    C9(27) = C8(0)
    D9(27) = D8(0)
    
    '组合C9,D9成C_D
    For i = 0 To 27
    C_D(i) = C9(i)
    C_D(i + 28) = D9(i)
    Next
    
    'PC_2转换,生成K9
    For i = 0 To 47
    K9(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K10
    '生成C10,D10
    For i = 0 To 25
    C10(i) = C9(i + Lsi(10))
    D10(i) = D9(i + Lsi(10))
    Next
    C10(26) = C9(0)
    D10(26) = D9(0)
    C10(27) = C9(1)
    D10(27) = D9(1)
    
    '组合C10,D10成C_D
    For i = 0 To 27
    C_D(i) = C10(i)
    C_D(i + 28) = D10(i)
    Next
    
    'PC_2转换,生成K10
    For i = 0 To 47
    K10(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K11
    '生成C11,D11
    For i = 0 To 25
    C11(i) = C10(i + Lsi(11))
    D11(i) = D10(i + Lsi(11))
    Next
    C11(26) = C10(0)
    D11(26) = D10(0)
    C11(27) = C10(1)
    D11(27) = D10(1)
    
    '组合C11,D11成C_D
    For i = 0 To 27
    C_D(i) = C11(i)
    C_D(i + 28) = D11(i)
    Next
    
    'PC_2转换,生成K11
    For i = 0 To 47
    K11(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K12
    '生成C12,D12
    For i = 0 To 25
    C12(i) = C11(i + Lsi(12))
    D12(i) = D11(i + Lsi(12))
    Next
    C12(26) = C11(0)
    D12(26) = D11(0)
    C12(27) = C11(1)
    D12(27) = D11(1)
    
    '组合C12,D12成C_D
    For i = 0 To 27
    C_D(i) = C12(i)
    C_D(i + 28) = D12(i)
    Next
    
    'PC_2转换,生成K12
    For i = 0 To 47
    K12(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K13
    '生成C13,D13
    For i = 0 To 25
    C13(i) = C12(i + Lsi(13))
    D13(i) = D12(i + Lsi(13))
    Next
    C13(26) = C12(0)
    D13(26) = D12(0)
    C13(27) = C12(1)
    D13(27) = D12(1)
    
    '组合C13,D13成C_D
    For i = 0 To 27
    C_D(i) = C13(i)
    C_D(i + 28) = D13(i)
    Next
    
    'PC_2转换,生成K13
    For i = 0 To 47
    K13(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K14
    '生成C14,D14
    For i = 0 To 25
    C14(i) = C13(i + Lsi(14))
    D14(i) = D13(i + Lsi(14))
    Next
    C14(26) = C13(0)
    D14(26) = D13(0)
    C14(27) = C13(1)
    D14(27) = D13(1)
    
    '组合C14,D14成C_D
    For i = 0 To 27
    C_D(i) = C14(i)
    C_D(i + 28) = D14(i)
    Next
    
    'PC_2转换,生成K14
    For i = 0 To 47
    K14(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K15
    '生成C15,D15
    For i = 0 To 25
    C15(i) = C14(i + Lsi(15))
    D15(i) = D14(i + Lsi(15))
    Next
    C15(26) = C14(0)
    D15(26) = D14(0)
    C15(27) = C14(1)
    D15(27) = D14(1)
    
    '组合C15,D15成C_D
    For i = 0 To 27
    C_D(i) = C15(i)
    C_D(i + 28) = D15(i)
    Next
    
    'PC_2转换,生成K15
    For i = 0 To 47
    K15(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K16
    '生成C16,D16
    For i = 0 To 26
    C16(i) = C15(i + Lsi(16))
    D16(i) = D15(i + Lsi(16))
    Next
    C16(27) = C15(0)
    D16(27) = D15(0)
    
    '组合C16,D16成C_D
    For i = 0 To 27
    C_D(i) = C16(i)
    C_D(i + 28) = D16(i)
    Next
    
    'PC_2转换,生成K16
    For i = 0 To 47
    K16(i) = C_D(PC_2(i))
    Next
    
    '**************************************************************************************
    
    '将明文位数扩展为8的倍数
    If (UBound(sCode) + 1) Mod 8 > 0 Then ReDim Preserve sCode(((UBound(sCode) + 1) \ 8 + 1) * 8 - 1)
    '定义返回密文长度
    ReDim bReturn(UBound(sCode))
    
    For j = 0 To UBound(sCode) Step 8
    '加密过程
    '依次取8位加密
    CopyMemory tempCode(0), sCode(j), 8
    For i = 0 To 7
    BinCode(i * 8 + 0) = (tempCode(i) And &H80) \ &H80
    BinCode(i * 8 + 1) = (tempCode(i) And &H40) \ &H40
    BinCode(i * 8 + 2) = (tempCode(i) And &H20) \ &H20
    BinCode(i * 8 + 3) = (tempCode(i) And &H10) \ &H10
    BinCode(i * 8 + 4) = (tempCode(i) And &H8) \ &H8
    BinCode(i * 8 + 5) = (tempCode(i) And &H4) \ &H4
    BinCode(i * 8 + 6) = (tempCode(i) And &H2) \ &H2
    BinCode(i * 8 + 7) = (tempCode(i) And &H1) \ &H1
    Next
    
    'IP置换
    For i = 0 To 63
    CodeIP(i) = BinCode(IP(i))
    Next
    
    '分段
    For i = 0 To 31
    L0(i) = CodeIP(i)
    R0(i) = CodeIP(i + 32)
    Next
    
    '进行第一次迭代
    For i = 0 To 47
    CodeE(i) = R0(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K1(i) '与K1按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L1,R1
    R1(i) = L0(i) Xor CodeP(i)
    l1(i) = R0(i)
    Next
    
    
    '进行第二次迭代
    For i = 0 To 47
    CodeE(i) = R1(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K2(i) '与K2按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L2,R2
    R2(i) = l1(i) Xor CodeP(i)
    L2(i) = R1(i)
    Next
    
    '进行第三次迭代
    For i = 0 To 47
    CodeE(i) = R2(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K3(i) '与K3按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L3,R3
    R3(i) = L2(i) Xor CodeP(i)
    L3(i) = R2(i)
    Next
    
    
    '进行第四次迭代
    For i = 0 To 47
    CodeE(i) = R3(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K4(i) '与K4按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L4,R4
    R4(i) = L3(i) Xor CodeP(i)
    L4(i) = R3(i)
    Next
    
    
    '进行第五次迭代
    For i = 0 To 47
    CodeE(i) = R4(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K5(i) '与K5按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L5,R5
    R5(i) = L4(i) Xor CodeP(i)
    L5(i) = R4(i)
    Next
    
    
    
    '进行第六次迭代
    For i = 0 To 47
    CodeE(i) = R5(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K6(i) '与K6按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L6,R6
    R6(i) = L5(i) Xor CodeP(i)
    L6(i) = R5(i)
    Next
    
    
    
    '进行第7次迭代
    For i = 0 To 47
    CodeE(i) = R6(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K7(i) '与K7按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L7,R7
    R7(i) = L6(i) Xor CodeP(i)
    L7(i) = R6(i)
    Next
    
    
    
    '进行第8次迭代
    For i = 0 To 47
    CodeE(i) = R7(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K8(i) '与K8按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L8,R8
    R8(i) = L7(i) Xor CodeP(i)
    L8(i) = R7(i)
    Next
    
    
    
    '进行第9次迭代
    For i = 0 To 47
    CodeE(i) = R8(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K9(i) '与K9按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L9,R9
    R9(i) = L8(i) Xor CodeP(i)
    L9(i) = R8(i)
    Next
    
    
    
    '进行第10次迭代
    For i = 0 To 47
    CodeE(i) = R9(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K10(i) '与K10按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L10,R10
    R10(i) = L9(i) Xor CodeP(i)
    L10(i) = R9(i)
    Next
    
    
    
    '进行第11次迭代
    For i = 0 To 47
    CodeE(i) = R10(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K11(i) '与K11按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L11,R11
    R11(i) = L10(i) Xor CodeP(i)
    L11(i) = R10(i)
    Next
    
    
    '进行第12次迭代
    For i = 0 To 47
    CodeE(i) = R11(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K12(i) '与K12按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L12,R12
    R12(i) = L11(i) Xor CodeP(i)
    L12(i) = R11(i)
    Next
    
    '进行第13次迭代
    For i = 0 To 47
    CodeE(i) = R12(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K13(i) '与K13按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L13,R13
    R13(i) = L12(i) Xor CodeP(i)
    L13(i) = R12(i)
    Next
    
    
    '进行第14次迭代
    For i = 0 To 47
    CodeE(i) = R13(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K14(i) '与K14按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L14,R14
    R14(i) = L13(i) Xor CodeP(i)
    L14(i) = R13(i)
    Next
    
    
    '进行第15次迭代
    For i = 0 To 47
    CodeE(i) = R14(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K15(i) '与K15按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L15,R15
    R15(i) = L14(i) Xor CodeP(i)
    L15(i) = R14(i)
    Next
    
    '进行第16次迭代
    For i = 0 To 47
    CodeE(i) = R15(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K16(i) '与K16按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L11,R11
    R16(i) = L15(i) Xor CodeP(i)
    L16(i) = R15(i)
    Next
    
    For i = 0 To 31
    BinCode(i) = L16(i)
    BinCode(i + 32) = R16(i)
    Next
    
    For i = 0 To 63
    CodeIP(i) = BinCode(IP_1(i))
    Next
    
    For i = 0 To 7
    tempReturn(i) = CodeIP(i * 8 + 0) * &H80
    tempReturn(i) = tempReturn(i) + CodeIP(i * 8 + 1) * &H40
    tempReturn(i) = tempReturn(i) + CodeIP(i * 8 + 2) * &H20
    tempReturn(i) = tempReturn(i) + CodeIP(i * 8 + 3) * &H10
    tempReturn(i) = tempReturn(i) + CodeIP(i * 8 + 4) * &H8
    tempReturn(i) = tempReturn(i) + CodeIP(i * 8 + 5) * &H4
    tempReturn(i) = tempReturn(i) + CodeIP(i * 8 + 6) * &H2
    tempReturn(i) = tempReturn(i) + CodeIP(i * 8 + 7) * &H1
    Next
    
    CopyMemory bReturn(j), tempReturn(0), 8
    Next
End Sub

'sCode 待解密字串
'sKey 密钥字串
Public Sub DES_Decode(sCode() As Byte, ByRef bReturn() As Byte, ByVal sKey As String)

    Dim LenTimes As Integer '明文
    Dim tempKey() As Byte '存放密钥
    Dim BinKey(63) As Byte '64位二进制原始密钥
    Dim KeyPC_1(55) As Byte '存放56位密钥
    
    Dim tempCode(7) As Byte '存放8位原始密文
    Dim tempReturn(7) As Byte '存放8位明文
    Dim BinCode(63) As Byte '存放64位明文
    Dim CodeIP(63) As Byte '存放IP置换结果
    Dim CodeE(47) As Byte 'E膨胀结果
    Dim CodeP(31) As Byte 'P变换结果
    Dim RetS(47) As Byte 'S盒运算32位结果
    Dim S(7) As Byte 'S盒运算8个结果
    Dim CodeS1(5) As Byte: Dim CodeS2(5) As Byte: Dim CodeS3(5) As Byte: Dim CodeS4(5) As Byte
    Dim CodeS5(5) As Byte: Dim CodeS6(5) As Byte: Dim CodeS7(5) As Byte: Dim CodeS8(5) As Byte
    
    
    Dim L0(31) As Byte: Dim R0(31) As Byte
    Dim l1(31) As Byte: Dim R1(31) As Byte
    Dim L2(31) As Byte: Dim R2(31) As Byte
    Dim L3(31) As Byte: Dim R3(31) As Byte
    Dim L4(31) As Byte: Dim R4(31) As Byte
    Dim L5(31) As Byte: Dim R5(31) As Byte
    Dim L6(31) As Byte: Dim R6(31) As Byte
    Dim L7(31) As Byte: Dim R7(31) As Byte
    Dim L8(31) As Byte: Dim R8(31) As Byte
    Dim L9(31) As Byte: Dim R9(31) As Byte
    Dim L10(31) As Byte: Dim R10(31) As Byte
    Dim L11(31) As Byte: Dim R11(31) As Byte
    Dim L12(31) As Byte: Dim R12(31) As Byte
    Dim L13(31) As Byte: Dim R13(31) As Byte
    Dim L14(31) As Byte: Dim R14(31) As Byte
    Dim L15(31) As Byte: Dim R15(31) As Byte
    Dim L16(31) As Byte: Dim R16(31) As Byte
    
    Dim C0(27) As Byte: Dim D0(27) As Byte '16个密钥
    Dim C1(27) As Byte: Dim D1(27) As Byte
    Dim C2(27) As Byte: Dim D2(27) As Byte:
    Dim C3(27) As Byte: Dim D3(27) As Byte:
    Dim C4(27) As Byte: Dim D4(27) As Byte:
    Dim C5(27) As Byte: Dim D5(27) As Byte:
    Dim C6(27) As Byte: Dim D6(27) As Byte:
    Dim C7(27) As Byte: Dim D7(27) As Byte:
    Dim C8(27) As Byte: Dim D8(27) As Byte:
    Dim C9(27) As Byte: Dim D9(27) As Byte:
    Dim C10(27) As Byte: Dim D10(27) As Byte:
    Dim C11(27) As Byte: Dim D11(27) As Byte:
    Dim C12(27) As Byte: Dim D12(27) As Byte:
    Dim C13(27) As Byte: Dim D13(27) As Byte:
    Dim C14(27) As Byte: Dim D14(27) As Byte:
    Dim C15(27) As Byte: Dim D15(27) As Byte:
    Dim C16(27) As Byte: Dim D16(27) As Byte:
    
    Dim C_D(55) As Byte 'Cn,Dn合并后的存放处
    
    Dim K1(47) As Byte: Dim K2(47) As Byte: Dim K3(47) As Byte: Dim K4(47) As Byte:
    Dim K5(47) As Byte: Dim K6(47) As Byte: Dim K7(47) As Byte: Dim K8(47) As Byte:
    Dim K9(47) As Byte: Dim K10(47) As Byte: Dim K11(47) As Byte: Dim K12(47) As Byte:
    Dim K13(47) As Byte: Dim K14(47) As Byte: Dim K15(47) As Byte: Dim K16(47) As Byte:
    
    Dim i As Integer
    Dim j As Integer
    
    Call InitDES
    
    '取密钥的前8字节
    tempKey = StrConv(sKey, vbFromUnicode)
    ReDim Preserve tempKey(7)
    
    For i = 0 To 7
    BinKey(i * 8 + 0) = (tempKey(i) And &H80) \ &H80
    BinKey(i * 8 + 1) = (tempKey(i) And &H40) \ &H40
    BinKey(i * 8 + 2) = (tempKey(i) And &H20) \ &H20
    BinKey(i * 8 + 3) = (tempKey(i) And &H10) \ &H10
    BinKey(i * 8 + 4) = (tempKey(i) And &H8) \ &H8
    BinKey(i * 8 + 5) = (tempKey(i) And &H4) \ &H4
    BinKey(i * 8 + 6) = (tempKey(i) And &H2) \ &H2
    BinKey(i * 8 + 7) = (tempKey(i) And &H1) \ &H1
    Next
    
    'PC_1转换
    For i = 0 To 55
    KeyPC_1(i) = BinKey(PC_1(i))
    Next
    
    '生成C0,D0
    For i = 0 To 27
    C0(i) = KeyPC_1(i)
    D0(i) = KeyPC_1(i + 28)
    Next
    
    '***************************************************K1
    '生成C1,D1
    For i = 0 To 26
    C1(i) = C0(i + Lsi(1))
    D1(i) = D0(i + Lsi(1))
    Next
    C1(27) = C0(0)
    D1(27) = D0(0)
    
    '组合C1,D1成C_D
    For i = 0 To 27
    C_D(i) = C1(i)
    C_D(i + 28) = D1(i)
    Next
    
    'PC_2转换,生成K1
    For i = 0 To 47
    K1(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K2
    '生成C2,D2
    For i = 0 To 26
    C2(i) = C1(i + Lsi(2))
    D2(i) = D1(i + Lsi(2))
    Next
    C2(27) = C1(0)
    D2(27) = D1(0)
    
    '组合C2,D2成C_D
    For i = 0 To 27
    C_D(i) = C2(i)
    C_D(i + 28) = D2(i)
    Next
    
    'PC_2转换,生成K2
    For i = 0 To 47
    K2(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K3
    '生成C3,D3
    For i = 0 To 25
    C3(i) = C2(i + Lsi(3))
    D3(i) = D2(i + Lsi(3))
    Next
    C3(26) = C2(0)
    D3(26) = D2(0)
    C3(27) = C2(1)
    D3(27) = D2(1)
    
    '组合C3,D3成C_D
    For i = 0 To 27
    C_D(i) = C3(i)
    C_D(i + 28) = D3(i)
    Next
    
    'PC_2转换,生成K3
    For i = 0 To 47
    K3(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K4
    '生成C4,D4
    For i = 0 To 25
    C4(i) = C3(i + Lsi(4))
    D4(i) = D3(i + Lsi(4))
    Next
    C4(26) = C3(0)
    D4(26) = D3(0)
    C4(27) = C3(1)
    D4(27) = D3(1)
    
    '组合C4,D4成C_D
    For i = 0 To 27
    C_D(i) = C4(i)
    C_D(i + 28) = D4(i)
    Next
    
    'PC_2转换,生成K4
    For i = 0 To 47
    K4(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K5
    '生成C5,D5
    For i = 0 To 25
    C5(i) = C4(i + Lsi(5))
    D5(i) = D4(i + Lsi(5))
    Next
    C5(26) = C4(0)
    D5(26) = D4(0)
    C5(27) = C4(1)
    D5(27) = D4(1)
    
    '组合C5,D5成C_D
    For i = 0 To 27
    C_D(i) = C5(i)
    C_D(i + 28) = D5(i)
    Next
    
    'PC_2转换,生成K5
    For i = 0 To 47
    K5(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K6
    '生成C6,D6
    For i = 0 To 25
    C6(i) = C5(i + Lsi(6))
    D6(i) = D5(i + Lsi(6))
    Next
    C6(26) = C5(0)
    D6(26) = D5(0)
    C6(27) = C5(1)
    D6(27) = D5(1)
    
    '组合C6,D6成C_D
    For i = 0 To 27
    C_D(i) = C6(i)
    C_D(i + 28) = D6(i)
    Next
    
    'PC_2转换,生成K6
    For i = 0 To 47
    K6(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K7
    '生成C7,D7
    For i = 0 To 25
    C7(i) = C6(i + Lsi(7))
    D7(i) = D6(i + Lsi(7))
    Next
    C7(26) = C6(0)
    D7(26) = D6(0)
    C7(27) = C6(1)
    D7(27) = D6(1)
    
    '组合C7,D7成C_D
    For i = 0 To 27
    C_D(i) = C7(i)
    C_D(i + 28) = D7(i)
    Next
    
    'PC_2转换,生成K7
    For i = 0 To 47
    K7(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K8
    '生成C8,D8
    For i = 0 To 25
    C8(i) = C7(i + Lsi(8))
    D8(i) = D7(i + Lsi(8))
    Next
    C8(26) = C7(0)
    D8(26) = D7(0)
    C8(27) = C7(1)
    D8(27) = D7(1)
    
    '组合C8,D8成C_D
    For i = 0 To 27
    C_D(i) = C8(i)
    C_D(i + 28) = D8(i)
    Next
    
    'PC_2转换,生成K8
    For i = 0 To 47
    K8(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K9
    '生成C9,D9
    For i = 0 To 26
    C9(i) = C8(i + Lsi(9))
    D9(i) = D8(i + Lsi(9))
    Next
    C9(27) = C8(0)
    D9(27) = D8(0)
    
    '组合C9,D9成C_D
    For i = 0 To 27
    C_D(i) = C9(i)
    C_D(i + 28) = D9(i)
    Next
    
    'PC_2转换,生成K9
    For i = 0 To 47
    K9(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K10
    '生成C10,D10
    For i = 0 To 25
    C10(i) = C9(i + Lsi(10))
    D10(i) = D9(i + Lsi(10))
    Next
    C10(26) = C9(0)
    D10(26) = D9(0)
    C10(27) = C9(1)
    D10(27) = D9(1)
    
    '组合C10,D10成C_D
    For i = 0 To 27
    C_D(i) = C10(i)
    C_D(i + 28) = D10(i)
    Next
    
    'PC_2转换,生成K10
    For i = 0 To 47
    K10(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K11
    '生成C11,D11
    For i = 0 To 25
    C11(i) = C10(i + Lsi(11))
    D11(i) = D10(i + Lsi(11))
    Next
    C11(26) = C10(0)
    D11(26) = D10(0)
    C11(27) = C10(1)
    D11(27) = D10(1)
    
    '组合C11,D11成C_D
    For i = 0 To 27
    C_D(i) = C11(i)
    C_D(i + 28) = D11(i)
    Next
    
    'PC_2转换,生成K11
    For i = 0 To 47
    K11(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K12
    '生成C12,D12
    For i = 0 To 25
    C12(i) = C11(i + Lsi(12))
    D12(i) = D11(i + Lsi(12))
    Next
    C12(26) = C11(0)
    D12(26) = D11(0)
    C12(27) = C11(1)
    D12(27) = D11(1)
    
    '组合C12,D12成C_D
    For i = 0 To 27
    C_D(i) = C12(i)
    C_D(i + 28) = D12(i)
    Next
    
    'PC_2转换,生成K12
    For i = 0 To 47
    K12(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K13
    '生成C13,D13
    For i = 0 To 25
    C13(i) = C12(i + Lsi(13))
    D13(i) = D12(i + Lsi(13))
    Next
    C13(26) = C12(0)
    D13(26) = D12(0)
    C13(27) = C12(1)
    D13(27) = D12(1)
    
    '组合C13,D13成C_D
    For i = 0 To 27
    C_D(i) = C13(i)
    C_D(i + 28) = D13(i)
    Next
    
    'PC_2转换,生成K13
    For i = 0 To 47
    K13(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K14
    '生成C14,D14
    For i = 0 To 25
    C14(i) = C13(i + Lsi(14))
    D14(i) = D13(i + Lsi(14))
    Next
    C14(26) = C13(0)
    D14(26) = D13(0)
    C14(27) = C13(1)
    D14(27) = D13(1)
    
    '组合C14,D14成C_D
    For i = 0 To 27
    C_D(i) = C14(i)
    C_D(i + 28) = D14(i)
    Next
    
    'PC_2转换,生成K14
    For i = 0 To 47
    K14(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K15
    '生成C15,D15
    For i = 0 To 25
    C15(i) = C14(i + Lsi(15))
    D15(i) = D14(i + Lsi(15))
    Next
    C15(26) = C14(0)
    D15(26) = D14(0)
    C15(27) = C14(1)
    D15(27) = D14(1)
    
    '组合C15,D15成C_D
    For i = 0 To 27
    C_D(i) = C15(i)
    C_D(i + 28) = D15(i)
    Next
    
    'PC_2转换,生成K15
    For i = 0 To 47
    K15(i) = C_D(PC_2(i))
    Next
    
    '***************************************************K16
    '生成C16,D16
    For i = 0 To 26
    C16(i) = C15(i + Lsi(16))
    D16(i) = D15(i + Lsi(16))
    Next
    C16(27) = C15(0)
    D16(27) = D15(0)
    
    '组合C16,D16成C_D
    For i = 0 To 27
    C_D(i) = C16(i)
    C_D(i + 28) = D16(i)
    Next
    
    'PC_2转换,生成K16
    For i = 0 To 47
    K16(i) = C_D(PC_2(i))
    Next
    
    '**************************************************************************************
    
    '加密过程
    '将明文位数扩展为8的倍数
    If (UBound(sCode) + 1) Mod 8 > 0 Then ReDim Preserve sCode(((UBound(sCode) + 1) \ 8 + 1) * 8 - 1)
    '定义返回密文长度
    ReDim bReturn(UBound(sCode))
    
    For j = 0 To UBound(sCode) Step 8
    CopyMemory tempCode(0), sCode(j), 8
    For i = 0 To 7
    BinCode(i * 8 + 0) = (tempCode(i) And &H80) \ &H80
    BinCode(i * 8 + 1) = (tempCode(i) And &H40) \ &H40
    BinCode(i * 8 + 2) = (tempCode(i) And &H20) \ &H20
    BinCode(i * 8 + 3) = (tempCode(i) And &H10) \ &H10
    BinCode(i * 8 + 4) = (tempCode(i) And &H8) \ &H8
    BinCode(i * 8 + 5) = (tempCode(i) And &H4) \ &H4
    BinCode(i * 8 + 6) = (tempCode(i) And &H2) \ &H2
    BinCode(i * 8 + 7) = (tempCode(i) And &H1) \ &H1
    Next
    
    'IP置换
    For i = 0 To 63
    CodeIP(i) = BinCode(IP(i))
    Next
    
    '分段
    For i = 0 To 31
    R0(i) = CodeIP(i)
    L0(i) = CodeIP(i + 32)
    Next
    
    '进行第一次迭代
    For i = 0 To 47
    CodeE(i) = R0(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K16(i) '与K16按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L1,R1
    R1(i) = L0(i) Xor CodeP(i)
    l1(i) = R0(i)
    Next
    
    
    '进行第二次迭代
    For i = 0 To 47
    CodeE(i) = R1(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K15(i) '与K15按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L2,R2
    R2(i) = l1(i) Xor CodeP(i)
    L2(i) = R1(i)
    Next
    
    '进行第三次迭代
    For i = 0 To 47
    CodeE(i) = R2(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K14(i) '与K14按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L3,R3
    R3(i) = L2(i) Xor CodeP(i)
    L3(i) = R2(i)
    Next
    
    
    '进行第四次迭代
    For i = 0 To 47
    CodeE(i) = R3(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K13(i) '与K13按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L4,R4
    R4(i) = L3(i) Xor CodeP(i)
    L4(i) = R3(i)
    Next
    
    
    '进行第五次迭代
    For i = 0 To 47
    CodeE(i) = R4(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K12(i) '与K12按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L5,R5
    R5(i) = L4(i) Xor CodeP(i)
    L5(i) = R4(i)
    Next
    
    
    
    '进行第六次迭代
    For i = 0 To 47
    CodeE(i) = R5(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K11(i) '与K11按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L6,R6
    R6(i) = L5(i) Xor CodeP(i)
    L6(i) = R5(i)
    Next
    
    
    
    '进行第7次迭代
    For i = 0 To 47
    CodeE(i) = R6(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K10(i) '与K10按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L7,R7
    R7(i) = L6(i) Xor CodeP(i)
    L7(i) = R6(i)
    Next
    
    
    
    '进行第8次迭代
    For i = 0 To 47
    CodeE(i) = R7(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K9(i) '与K9按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L8,R8
    R8(i) = L7(i) Xor CodeP(i)
    L8(i) = R7(i)
    Next
    
    
    
    '进行第9次迭代
    For i = 0 To 47
    CodeE(i) = R8(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K8(i) '与K8按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L9,R9
    R9(i) = L8(i) Xor CodeP(i)
    L9(i) = R8(i)
    Next
    
    
    
    '进行第10次迭代
    For i = 0 To 47
    CodeE(i) = R9(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K7(i) '与K7按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L10,R10
    R10(i) = L9(i) Xor CodeP(i)
    L10(i) = R9(i)
    Next
    
    
    
    '进行第11次迭代
    For i = 0 To 47
    CodeE(i) = R10(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K6(i) '与K6按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L11,R11
    R11(i) = L10(i) Xor CodeP(i)
    L11(i) = R10(i)
    Next
    
    
    '进行第12次迭代
    For i = 0 To 47
    CodeE(i) = R11(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K5(i) '与K5按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L12,R12
    R12(i) = L11(i) Xor CodeP(i)
    L12(i) = R11(i)
    Next
    
    '进行第13次迭代
    For i = 0 To 47
    CodeE(i) = R12(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K4(i) '与K4按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L13,R13
    R13(i) = L12(i) Xor CodeP(i)
    L13(i) = R12(i)
    Next
    
    
    '进行第14次迭代
    For i = 0 To 47
    CodeE(i) = R13(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K3(i) '与K3按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L14,R14
    R14(i) = L13(i) Xor CodeP(i)
    L14(i) = R13(i)
    Next
    
    
    '进行第15次迭代
    For i = 0 To 47
    CodeE(i) = R14(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K2(i) '与K2按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L15,R15
    R15(i) = L14(i) Xor CodeP(i)
    L15(i) = R14(i)
    Next
    
    '进行第16次迭代
    For i = 0 To 47
    CodeE(i) = R15(E(i)) '经过E变换扩充，由32位变为48位
    CodeE(i) = CodeE(i) Xor K1(i) '与K1按位作不进位加法运算
    Next
    
    '分8组
    For i = 0 To 5
    CodeS1(i) = CodeE(i)
    CodeS2(i) = CodeE(i + 6)
    CodeS3(i) = CodeE(i + 12)
    CodeS4(i) = CodeE(i + 18)
    CodeS5(i) = CodeE(i + 24)
    CodeS6(i) = CodeE(i + 30)
    CodeS7(i) = CodeE(i + 36)
    CodeS8(i) = CodeE(i + 42)
    Next
    
    'S盒运算，得到8个数
    
    S(0) = s1(CodeS1(5) + CodeS1(0) * 2, CodeS1(4) + CodeS1(3) * 2 + CodeS1(2) * 4 + CodeS1(1) * 8)
    S(1) = s2(CodeS2(5) + CodeS2(0) * 2, CodeS2(4) + CodeS2(3) * 2 + CodeS2(2) * 4 + CodeS2(1) * 8)
    S(2) = s3(CodeS3(5) + CodeS3(0) * 2, CodeS3(4) + CodeS3(3) * 2 + CodeS3(2) * 4 + CodeS3(1) * 8)
    S(3) = S4(CodeS4(5) + CodeS4(0) * 2, CodeS4(4) + CodeS4(3) * 2 + CodeS4(2) * 4 + CodeS4(1) * 8)
    S(4) = S5(CodeS5(5) + CodeS5(0) * 2, CodeS5(4) + CodeS5(3) * 2 + CodeS5(2) * 4 + CodeS5(1) * 8)
    S(5) = S6(CodeS6(5) + CodeS6(0) * 2, CodeS6(4) + CodeS6(3) * 2 + CodeS6(2) * 4 + CodeS6(1) * 8)
    S(6) = S7(CodeS7(5) + CodeS7(0) * 2, CodeS7(4) + CodeS7(3) * 2 + CodeS7(2) * 4 + CodeS7(1) * 8)
    S(7) = S8(CodeS8(5) + CodeS8(0) * 2, CodeS8(4) + CodeS8(3) * 2 + CodeS8(2) * 4 + CodeS8(1) * 8)
    
    'S盒运算32位结果
    For i = 0 To 7
    RetS(i * 4 + 0) = (S(i) And &H8) \ &H8
    RetS(i * 4 + 1) = (S(i) And &H4) \ &H4
    RetS(i * 4 + 2) = (S(i) And &H2) \ &H2
    RetS(i * 4 + 3) = (S(i) And &H1) \ &H1
    Next
    
    
    For i = 0 To 31
    'P变换
    CodeP(i) = RetS(p(i))
    
    '产生L11,R11
    R16(i) = L15(i) Xor CodeP(i)
    L16(i) = R15(i)
    Next
    
    For i = 0 To 31
    BinCode(i) = R16(i)
    BinCode(i + 32) = L16(i)
    Next
    
    For i = 0 To 63
    CodeIP(i) = BinCode(IP_1(i))
    Next
    
    For i = 0 To 7
    tempReturn(i) = CodeIP(i * 8 + 0) * &H80
    tempReturn(i) = tempReturn(i) + CodeIP(i * 8 + 1) * &H40
    tempReturn(i) = tempReturn(i) + CodeIP(i * 8 + 2) * &H20
    tempReturn(i) = tempReturn(i) + CodeIP(i * 8 + 3) * &H10
    tempReturn(i) = tempReturn(i) + CodeIP(i * 8 + 4) * &H8
    tempReturn(i) = tempReturn(i) + CodeIP(i * 8 + 5) * &H4
    tempReturn(i) = tempReturn(i) + CodeIP(i * 8 + 6) * &H2
    tempReturn(i) = tempReturn(i) + CodeIP(i * 8 + 7) * &H1
    Next
    
    CopyMemory bReturn(j), tempReturn(0), 8
    Next

End Sub

Private Sub InitDES()
    Dim i As Integer

    IP(0) = 57 ' 58
    IP(1) = 49 ' 50
    IP(2) = 41 ' 42
    IP(3) = 33 ' 34
    IP(4) = 25 ' 26
    IP(5) = 17 ' 18
    IP(6) = 9 ' 10
    IP(7) = 1 ' 2
    IP(8) = 59 ' 60
    IP(9) = 51 ' 52
    IP(10) = 43 ' 44
    IP(11) = 35 ' 36
    IP(12) = 27 ' 28
    IP(13) = 19 ' 20
    IP(14) = 11 ' 12
    IP(15) = 3 ' 4
    IP(16) = 61 ' 62
    IP(17) = 53 ' 54
    IP(18) = 45 ' 46
    IP(19) = 37 ' 38
    IP(20) = 29 ' 30
    IP(21) = 21 ' 22
    IP(22) = 13 ' 14
    IP(23) = 5 ' 6
    IP(24) = 63 ' 64
    IP(25) = 55 ' 56
    IP(26) = 47 ' 48
    IP(27) = 39 ' 40
    IP(28) = 31 ' 32
    IP(29) = 23 ' 24
    IP(30) = 15 ' 16
    IP(31) = 7 ' 8
    IP(32) = 56 ' 57
    IP(33) = 48 ' 49
    IP(34) = 40 ' 41
    IP(35) = 32 ' 33
    IP(36) = 24 ' 25
    IP(37) = 16 ' 17
    IP(38) = 8 ' 9
    IP(39) = 0 ' 1
    IP(40) = 58 ' 59
    IP(41) = 50 ' 51
    IP(42) = 42 ' 43
    IP(43) = 34 ' 35
    IP(44) = 26 ' 27
    IP(45) = 18 ' 19
    IP(46) = 10 ' 11
    IP(47) = 2 ' 3
    IP(48) = 60 ' 61
    IP(49) = 52 ' 53
    IP(50) = 44 ' 45
    IP(51) = 36 ' 37
    IP(52) = 28 ' 29
    IP(53) = 20 ' 21
    IP(54) = 12 ' 13
    IP(55) = 4 ' 5
    IP(56) = 62 ' 63
    IP(57) = 54 ' 55
    IP(58) = 46 ' 47
    IP(59) = 38 ' 39
    IP(60) = 30 ' 31
    IP(61) = 22 ' 23
    IP(62) = 14 ' 15
    IP(63) = 6 ' 7
    
    
    IP_1(0) = 39 ' 40
    IP_1(1) = 7 ' 8
    IP_1(2) = 47 ' 48
    IP_1(3) = 15 ' 16
    IP_1(4) = 55 ' 56
    IP_1(5) = 23 ' 24
    IP_1(6) = 63 ' 64
    IP_1(7) = 31 ' 32
    IP_1(8) = 38 ' 39
    IP_1(9) = 6 ' 7
    IP_1(10) = 46 ' 47
    IP_1(11) = 14 ' 15
    IP_1(12) = 54 ' 55
    IP_1(13) = 22 ' 23
    IP_1(14) = 62 ' 63
    IP_1(15) = 30 ' 31
    IP_1(16) = 37 ' 38
    IP_1(17) = 5 ' 6
    IP_1(18) = 45 ' 46
    IP_1(19) = 13 ' 14
    IP_1(20) = 53 ' 54
    IP_1(21) = 21 ' 22
    IP_1(22) = 61 ' 62
    IP_1(23) = 29 ' 30
    IP_1(24) = 36 ' 37
    IP_1(25) = 4 ' 5
    IP_1(26) = 44 ' 45
    IP_1(27) = 12 ' 13
    IP_1(28) = 52 ' 53
    IP_1(29) = 20 ' 21
    IP_1(30) = 60 ' 61
    IP_1(31) = 28 ' 29
    IP_1(32) = 35 ' 36
    IP_1(33) = 3 ' 4
    IP_1(34) = 43 ' 44
    IP_1(35) = 11 ' 12
    IP_1(36) = 51 ' 52
    IP_1(37) = 19 ' 20
    IP_1(38) = 59 ' 60
    IP_1(39) = 27 ' 28
    IP_1(40) = 34 ' 35
    IP_1(41) = 2 ' 3
    IP_1(42) = 42 ' 43
    IP_1(43) = 10 ' 11
    IP_1(44) = 50 ' 51
    IP_1(45) = 18 ' 19
    IP_1(46) = 58 ' 59
    IP_1(47) = 26 ' 27
    IP_1(48) = 33 ' 34
    IP_1(49) = 1 ' 2
    IP_1(50) = 41 ' 42
    IP_1(51) = 9 ' 10
    IP_1(52) = 49 ' 50
    IP_1(53) = 17 ' 18
    IP_1(54) = 57 ' 58
    IP_1(55) = 25 ' 26
    IP_1(56) = 32 ' 33
    IP_1(57) = 0 ' 1
    IP_1(58) = 40 ' 41
    IP_1(59) = 8 ' 9
    IP_1(60) = 48 ' 49
    IP_1(61) = 16 ' 17
    IP_1(62) = 56 ' 57
    IP_1(63) = 24 ' 25
    
    
    E(0) = 31
    For i = 1 To 5
    E(i) = i - 1
    Next
    
    For i = 6 To 11
    E(i) = i - 3
    Next
    
    For i = 12 To 17
    E(i) = i - 5
    Next
    
    For i = 18 To 23
    E(i) = i - 7
    Next
    
    For i = 24 To 29
    E(i) = i - 9
    Next
    
    For i = 30 To 35
    E(i) = i - 11
    Next
    
    For i = 36 To 41
    E(i) = i - 13
    Next
    For i = 42 To 46
    E(i) = i - 15
    Next
    E(47) = 30
    
    p(0) = 15 ' 16
    p(1) = 6 ' 7
    p(2) = 19 ' 20
    p(3) = 20 ' 21
    p(4) = 28 ' 29
    p(5) = 11 ' 12
    p(6) = 27 ' 28
    p(7) = 16 ' 17
    p(8) = 0 ' 1
    p(9) = 14 ' 15
    p(10) = 22 ' 23
    p(11) = 25 ' 26
    p(12) = 4 ' 5
    p(13) = 17 ' 18
    p(14) = 30 ' 31
    p(15) = 9 ' 10
    p(16) = 1 ' 2
    p(17) = 7 ' 8
    p(18) = 23 ' 24
    p(19) = 13 ' 14
    p(20) = 31 ' 32
    p(21) = 26 ' 27
    p(22) = 2 ' 3
    p(23) = 8 ' 9
    p(24) = 18 ' 19
    p(25) = 12 ' 13
    p(26) = 29 ' 30
    p(27) = 5 ' 6
    p(28) = 21 ' 22
    p(29) = 10 ' 11
    p(30) = 3 ' 4
    p(31) = 24 ' 25
    
    s1(0, 0) = 14
    s1(0, 1) = 4
    s1(0, 2) = 13
    s1(0, 3) = 1
    s1(0, 4) = 2
    s1(0, 5) = 15
    s1(0, 6) = 11
    s1(0, 7) = 8
    s1(0, 8) = 3
    s1(0, 9) = 10
    s1(0, 10) = 6
    s1(0, 11) = 12
    s1(0, 12) = 5
    s1(0, 13) = 9
    s1(0, 14) = 0
    s1(0, 15) = 7
    s1(1, 0) = 0
    s1(1, 1) = 15
    s1(1, 2) = 7
    s1(1, 3) = 4
    s1(1, 4) = 14
    s1(1, 5) = 2
    s1(1, 6) = 13
    s1(1, 7) = 1
    s1(1, 8) = 10
    s1(1, 9) = 6
    s1(1, 10) = 12
    s1(1, 11) = 11
    s1(1, 12) = 9
    s1(1, 13) = 5
    s1(1, 14) = 3
    s1(1, 15) = 8
    s1(2, 0) = 4
    s1(2, 1) = 1
    s1(2, 2) = 14
    s1(2, 3) = 8
    s1(2, 4) = 13
    s1(2, 5) = 6
    s1(2, 6) = 2
    s1(2, 7) = 11
    s1(2, 8) = 15
    s1(2, 9) = 12
    s1(2, 10) = 9
    s1(2, 11) = 7
    s1(2, 12) = 3
    s1(2, 13) = 10
    s1(2, 14) = 5
    s1(2, 15) = 0
    s1(3, 0) = 15
    s1(3, 1) = 12
    s1(3, 2) = 8
    s1(3, 3) = 2
    s1(3, 4) = 4
    s1(3, 5) = 9
    s1(3, 6) = 1
    s1(3, 7) = 7
    s1(3, 8) = 5
    s1(3, 9) = 11
    s1(3, 10) = 3
    s1(3, 11) = 14
    s1(3, 12) = 10
    s1(3, 13) = 0
    s1(3, 14) = 6
    s1(3, 15) = 13
    
    s2(0, 0) = 15
    s2(0, 1) = 1
    s2(0, 2) = 8
    s2(0, 3) = 14
    s2(0, 4) = 6
    s2(0, 5) = 11
    s2(0, 6) = 3
    s2(0, 7) = 4
    s2(0, 8) = 9
    s2(0, 9) = 7
    s2(0, 10) = 2
    s2(0, 11) = 13
    s2(0, 12) = 12
    s2(0, 13) = 0
    s2(0, 14) = 5
    s2(0, 15) = 10
    s2(1, 0) = 3
    s2(1, 1) = 13
    s2(1, 2) = 4
    s2(1, 3) = 7
    s2(1, 4) = 15
    s2(1, 5) = 2
    s2(1, 6) = 8
    s2(1, 7) = 14
    s2(1, 8) = 12
    s2(1, 9) = 0
    s2(1, 10) = 1
    s2(1, 11) = 10
    s2(1, 12) = 6
    s2(1, 13) = 9
    s2(1, 14) = 11
    s2(1, 15) = 5
    s2(2, 0) = 0
    s2(2, 1) = 14
    s2(2, 2) = 7
    s2(2, 3) = 11
    s2(2, 4) = 10
    s2(2, 5) = 4
    s2(2, 6) = 13
    s2(2, 7) = 1
    s2(2, 8) = 5
    s2(2, 9) = 8
    s2(2, 10) = 12
    s2(2, 11) = 6
    s2(2, 12) = 9
    s2(2, 13) = 3
    s2(2, 14) = 2
    s2(2, 15) = 15
    s2(3, 0) = 13
    s2(3, 1) = 8
    s2(3, 2) = 10
    s2(3, 3) = 1
    s2(3, 4) = 3
    s2(3, 5) = 15
    s2(3, 6) = 4
    s2(3, 7) = 2
    s2(3, 8) = 11
    s2(3, 9) = 6
    s2(3, 10) = 7
    s2(3, 11) = 12
    s2(3, 12) = 0
    s2(3, 13) = 5
    s2(3, 14) = 14
    s2(3, 15) = 9
    
    s3(0, 0) = 10
    s3(0, 1) = 0
    s3(0, 2) = 9
    s3(0, 3) = 14
    s3(0, 4) = 6
    s3(0, 5) = 3
    s3(0, 6) = 15
    s3(0, 7) = 5
    s3(0, 8) = 1
    s3(0, 9) = 13
    s3(0, 10) = 12
    s3(0, 11) = 7
    s3(0, 12) = 11
    s3(0, 13) = 4
    s3(0, 14) = 2
    s3(0, 15) = 8
    s3(1, 0) = 13
    s3(1, 1) = 7
    s3(1, 2) = 0
    s3(1, 3) = 9
    s3(1, 4) = 3
    s3(1, 5) = 4
    s3(1, 6) = 6
    s3(1, 7) = 10
    s3(1, 8) = 2
    s3(1, 9) = 8
    s3(1, 10) = 5
    s3(1, 11) = 14
    s3(1, 12) = 12
    s3(1, 13) = 11
    s3(1, 14) = 15
    s3(1, 15) = 1
    s3(2, 0) = 13
    s3(2, 1) = 6
    s3(2, 2) = 4
    s3(2, 3) = 9
    s3(2, 4) = 8
    s3(2, 5) = 15
    s3(2, 6) = 3
    s3(2, 7) = 0
    s3(2, 8) = 11
    s3(2, 9) = 1
    s3(2, 10) = 2
    s3(2, 11) = 12
    s3(2, 12) = 5
    s3(2, 13) = 10
    s3(2, 14) = 14
    s3(2, 15) = 7
    s3(3, 0) = 1
    s3(3, 1) = 10
    s3(3, 2) = 13
    s3(3, 3) = 0
    s3(3, 4) = 6
    s3(3, 5) = 9
    s3(3, 6) = 8
    s3(3, 7) = 7
    s3(3, 8) = 4
    s3(3, 9) = 15
    s3(3, 10) = 14
    s3(3, 11) = 3
    s3(3, 12) = 11
    s3(3, 13) = 5
    s3(3, 14) = 2
    s3(3, 15) = 12
    
    S4(0, 0) = 7
    S4(0, 1) = 13
    S4(0, 2) = 14
    S4(0, 3) = 3
    S4(0, 4) = 0
    S4(0, 5) = 6
    S4(0, 6) = 9
    S4(0, 7) = 10
    S4(0, 8) = 1
    S4(0, 9) = 2
    S4(0, 10) = 8
    S4(0, 11) = 5
    S4(0, 12) = 11
    S4(0, 13) = 12
    S4(0, 14) = 4
    S4(0, 15) = 15
    S4(1, 0) = 13
    S4(1, 1) = 8
    S4(1, 2) = 11
    S4(1, 3) = 5
    S4(1, 4) = 6
    S4(1, 5) = 15
    S4(1, 6) = 0
    S4(1, 7) = 3
    S4(1, 8) = 4
    S4(1, 9) = 7
    S4(1, 10) = 2
    S4(1, 11) = 12
    S4(1, 12) = 1
    S4(1, 13) = 10
    S4(1, 14) = 14
    S4(1, 15) = 9
    S4(2, 0) = 10
    S4(2, 1) = 6
    S4(2, 2) = 9
    S4(2, 3) = 0
    S4(2, 4) = 12
    S4(2, 5) = 11
    S4(2, 6) = 7
    S4(2, 7) = 13
    S4(2, 8) = 15
    S4(2, 9) = 1
    S4(2, 10) = 3
    S4(2, 11) = 14
    S4(2, 12) = 5
    S4(2, 13) = 2
    S4(2, 14) = 8
    S4(2, 15) = 4
    S4(3, 0) = 3
    S4(3, 1) = 15
    S4(3, 2) = 0
    S4(3, 3) = 6
    S4(3, 4) = 10
    S4(3, 5) = 1
    S4(3, 6) = 13
    S4(3, 7) = 8
    S4(3, 8) = 9
    S4(3, 9) = 4
    S4(3, 10) = 5
    S4(3, 11) = 11
    S4(3, 12) = 12
    S4(3, 13) = 7
    S4(3, 14) = 2
    S4(3, 15) = 14
    
    S5(0, 0) = 2
    S5(0, 1) = 12
    S5(0, 2) = 4
    S5(0, 3) = 1
    S5(0, 4) = 7
    S5(0, 5) = 10
    S5(0, 6) = 11
    S5(0, 7) = 6
    S5(0, 8) = 8
    S5(0, 9) = 5
    S5(0, 10) = 3
    S5(0, 11) = 15
    S5(0, 12) = 13
    S5(0, 13) = 0
    S5(0, 14) = 14
    S5(0, 15) = 9
    S5(1, 0) = 14
    S5(1, 1) = 11
    S5(1, 2) = 2
    S5(1, 3) = 12
    S5(1, 4) = 4
    S5(1, 5) = 7
    S5(1, 6) = 13
    S5(1, 7) = 1
    S5(1, 8) = 5
    S5(1, 9) = 0
    S5(1, 10) = 15
    S5(1, 11) = 10
    S5(1, 12) = 3
    S5(1, 13) = 9
    S5(1, 14) = 8
    S5(1, 15) = 6
    S5(2, 0) = 4
    S5(2, 1) = 2
    S5(2, 2) = 1
    S5(2, 3) = 11
    S5(2, 4) = 10
    S5(2, 5) = 13
    S5(2, 6) = 7
    S5(2, 7) = 8
    S5(2, 8) = 15
    S5(2, 9) = 9
    S5(2, 10) = 12
    S5(2, 11) = 5
    S5(2, 12) = 6
    S5(2, 13) = 3
    S5(2, 14) = 0
    S5(2, 15) = 14
    S5(3, 0) = 11
    S5(3, 1) = 8
    S5(3, 2) = 12
    S5(3, 3) = 7
    S5(3, 4) = 1
    S5(3, 5) = 14
    S5(3, 6) = 2
    S5(3, 7) = 13
    S5(3, 8) = 6
    S5(3, 9) = 15
    S5(3, 10) = 0
    S5(3, 11) = 9
    S5(3, 12) = 10
    S5(3, 13) = 4
    S5(3, 14) = 5
    S5(3, 15) = 3
    
    S6(0, 0) = 12
    S6(0, 1) = 1
    S6(0, 2) = 10
    S6(0, 3) = 15
    S6(0, 4) = 9
    S6(0, 5) = 2
    S6(0, 6) = 6
    S6(0, 7) = 8
    S6(0, 8) = 0
    S6(0, 9) = 13
    S6(0, 10) = 3
    S6(0, 11) = 4
    S6(0, 12) = 14
    S6(0, 13) = 7
    S6(0, 14) = 5
    S6(0, 15) = 11
    S6(1, 0) = 10
    S6(1, 1) = 15
    S6(1, 2) = 4
    S6(1, 3) = 2
    S6(1, 4) = 7
    S6(1, 5) = 12
    S6(1, 6) = 9
    S6(1, 7) = 5
    S6(1, 8) = 6
    S6(1, 9) = 1
    S6(1, 10) = 13
    S6(1, 11) = 14
    S6(1, 12) = 0
    S6(1, 13) = 11
    S6(1, 14) = 3
    S6(1, 15) = 8
    S6(2, 0) = 9
    S6(2, 1) = 14
    S6(2, 2) = 15
    S6(2, 3) = 5
    S6(2, 4) = 2
    S6(2, 5) = 8
    S6(2, 6) = 12
    S6(2, 7) = 3
    S6(2, 8) = 7
    S6(2, 9) = 0
    S6(2, 10) = 4
    S6(2, 11) = 10
    S6(2, 12) = 1
    S6(2, 13) = 13
    S6(2, 14) = 11
    S6(2, 15) = 6
    S6(3, 0) = 4
    S6(3, 1) = 3
    S6(3, 2) = 2
    S6(3, 3) = 12
    S6(3, 4) = 9
    S6(3, 5) = 5
    S6(3, 6) = 15
    S6(3, 7) = 10
    S6(3, 8) = 11
    S6(3, 9) = 14
    S6(3, 10) = 1
    S6(3, 11) = 7
    S6(3, 12) = 6
    S6(3, 13) = 0
    S6(3, 14) = 8
    S6(3, 15) = 13
    
    S7(0, 0) = 4
    S7(0, 1) = 11
    S7(0, 2) = 2
    S7(0, 3) = 14
    S7(0, 4) = 15
    S7(0, 5) = 0
    S7(0, 6) = 8
    S7(0, 7) = 13
    S7(0, 8) = 3
    S7(0, 9) = 12
    S7(0, 10) = 9
    S7(0, 11) = 7
    S7(0, 12) = 5
    S7(0, 13) = 10
    S7(0, 14) = 6
    S7(0, 15) = 1
    S7(1, 0) = 13
    S7(1, 1) = 0
    S7(1, 2) = 11
    S7(1, 3) = 7
    S7(1, 4) = 4
    S7(1, 5) = 9
    S7(1, 6) = 1
    S7(1, 7) = 10
    S7(1, 8) = 14
    S7(1, 9) = 3
    S7(1, 10) = 5
    S7(1, 11) = 12
    S7(1, 12) = 2
    S7(1, 13) = 15
    S7(1, 14) = 8
    S7(1, 15) = 6
    S7(2, 0) = 1
    S7(2, 1) = 4
    S7(2, 2) = 11
    S7(2, 3) = 13
    S7(2, 4) = 12
    S7(2, 5) = 3
    S7(2, 6) = 7
    S7(2, 7) = 14
    S7(2, 8) = 10
    S7(2, 9) = 15
    S7(2, 10) = 6
    S7(2, 11) = 8
    S7(2, 12) = 0
    S7(2, 13) = 5
    S7(2, 14) = 9
    S7(2, 15) = 2
    S7(3, 0) = 6
    S7(3, 1) = 11
    S7(3, 2) = 13
    S7(3, 3) = 8
    S7(3, 4) = 1
    S7(3, 5) = 4
    S7(3, 6) = 10
    S7(3, 7) = 7
    S7(3, 8) = 9
    S7(3, 9) = 5
    S7(3, 10) = 0
    S7(3, 11) = 15
    S7(3, 12) = 14
    S7(3, 13) = 2
    S7(3, 14) = 3
    S7(3, 15) = 12
    
    S8(0, 0) = 13
    S8(0, 1) = 2
    S8(0, 2) = 8
    S8(0, 3) = 4
    S8(0, 4) = 6
    S8(0, 5) = 15
    S8(0, 6) = 11
    S8(0, 7) = 1
    S8(0, 8) = 10
    S8(0, 9) = 9
    S8(0, 10) = 3
    S8(0, 11) = 14
    S8(0, 12) = 5
    S8(0, 13) = 0
    S8(0, 14) = 12
    S8(0, 15) = 7
    S8(1, 0) = 1
    S8(1, 1) = 15
    S8(1, 2) = 13
    S8(1, 3) = 8
    S8(1, 4) = 10
    S8(1, 5) = 3
    S8(1, 6) = 7
    S8(1, 7) = 4
    S8(1, 8) = 12
    S8(1, 9) = 5
    S8(1, 10) = 6
    S8(1, 11) = 11
    S8(1, 12) = 0
    S8(1, 13) = 14
    S8(1, 14) = 9
    S8(1, 15) = 2
    S8(2, 0) = 7
    S8(2, 1) = 11
    S8(2, 2) = 4
    S8(2, 3) = 1
    S8(2, 4) = 9
    S8(2, 5) = 12
    S8(2, 6) = 14
    S8(2, 7) = 2
    S8(2, 8) = 0
    S8(2, 9) = 6
    S8(2, 10) = 10
    S8(2, 11) = 13
    S8(2, 12) = 15
    S8(2, 13) = 3
    S8(2, 14) = 5
    S8(2, 15) = 8
    S8(3, 0) = 2
    S8(3, 1) = 1
    S8(3, 2) = 14
    S8(3, 3) = 7
    S8(3, 4) = 4
    S8(3, 5) = 10
    S8(3, 6) = 8
    S8(3, 7) = 13
    S8(3, 8) = 15
    S8(3, 9) = 12
    S8(3, 10) = 9
    S8(3, 11) = 0
    S8(3, 12) = 3
    S8(3, 13) = 5
    S8(3, 14) = 6
    S8(3, 15) = 11
    
    PC_1(0) = 56 ' 57
    PC_1(1) = 48 ' 49
    PC_1(2) = 40 ' 41
    PC_1(3) = 32 ' 33
    PC_1(4) = 24 ' 25
    PC_1(5) = 16 ' 17
    PC_1(6) = 8 ' 9
    PC_1(7) = 0 ' 1
    PC_1(8) = 57 ' 58
    PC_1(9) = 49 ' 50
    PC_1(10) = 41 ' 42
    PC_1(11) = 33 ' 34
    PC_1(12) = 25 ' 26
    PC_1(13) = 17 ' 18
    PC_1(14) = 9 ' 10
    PC_1(15) = 1 ' 2
    PC_1(16) = 58 ' 59
    PC_1(17) = 50 ' 51
    PC_1(18) = 42 ' 43
    PC_1(19) = 34 ' 35
    PC_1(20) = 26 ' 27
    PC_1(21) = 18 ' 19
    PC_1(22) = 10 ' 11
    PC_1(23) = 2 ' 3
    PC_1(24) = 59 ' 60
    PC_1(25) = 51 ' 52
    PC_1(26) = 43 ' 44
    PC_1(27) = 35 ' 36
    PC_1(28) = 62 ' 63
    PC_1(29) = 54 ' 55
    PC_1(30) = 46 ' 47
    PC_1(31) = 38 ' 39
    PC_1(32) = 30 ' 31
    PC_1(33) = 22 ' 23
    PC_1(34) = 14 ' 15
    PC_1(35) = 6 ' 7
    PC_1(36) = 61 ' 62
    PC_1(37) = 53 ' 54
    PC_1(38) = 45 ' 46
    PC_1(39) = 37 ' 38
    PC_1(40) = 29 ' 30
    PC_1(41) = 21 ' 22
    PC_1(42) = 13 ' 14
    PC_1(43) = 5 ' 6
    PC_1(44) = 60 ' 61
    PC_1(45) = 52 ' 53
    PC_1(46) = 44 ' 45
    PC_1(47) = 36 ' 37
    PC_1(48) = 28 ' 29
    PC_1(49) = 20 ' 21
    PC_1(50) = 12 ' 13
    PC_1(51) = 4 ' 5
    PC_1(52) = 27 ' 28
    PC_1(53) = 19 ' 20
    PC_1(54) = 11 ' 12
    PC_1(55) = 3 ' 4
    
    PC_2(0) = 13 ' 14
    PC_2(1) = 16 ' 17
    PC_2(2) = 10 ' 11
    PC_2(3) = 23 ' 24
    PC_2(4) = 0 ' 1
    PC_2(5) = 4 ' 5
    PC_2(6) = 2 ' 3
    PC_2(7) = 27 ' 28
    PC_2(8) = 14 ' 15
    PC_2(9) = 5 ' 6
    PC_2(10) = 20 ' 21
    PC_2(11) = 9 ' 10
    PC_2(12) = 22 ' 23
    PC_2(13) = 18 ' 19
    PC_2(14) = 11 ' 12
    PC_2(15) = 3 ' 4
    PC_2(16) = 25 ' 26
    PC_2(17) = 7 ' 8
    PC_2(18) = 15 ' 16
    PC_2(19) = 6 ' 7
    PC_2(20) = 26 ' 27
    PC_2(21) = 19 ' 20
    PC_2(22) = 12 ' 13
    PC_2(23) = 1 ' 2
    PC_2(24) = 40 ' 41
    PC_2(25) = 51 ' 52
    PC_2(26) = 30 ' 31
    PC_2(27) = 36 ' 37
    PC_2(28) = 46 ' 47
    PC_2(29) = 54 ' 55
    PC_2(30) = 29 ' 30
    PC_2(31) = 39 ' 40
    PC_2(32) = 50 ' 51
    PC_2(33) = 44 ' 45
    PC_2(34) = 32 ' 33
    PC_2(35) = 47 ' 48
    PC_2(36) = 43 ' 44
    PC_2(37) = 48 ' 49
    PC_2(38) = 38 ' 39
    PC_2(39) = 55 ' 56
    PC_2(40) = 33 ' 34
    PC_2(41) = 52 ' 53
    PC_2(42) = 45 ' 46
    PC_2(43) = 41 ' 42
    PC_2(44) = 49 ' 50
    PC_2(45) = 35 ' 36
    PC_2(46) = 28 ' 29
    PC_2(47) = 31 ' 32
    
    Lsi(1) = 1
    Lsi(2) = 1
    Lsi(3) = 2
    Lsi(4) = 2
    Lsi(5) = 2
    Lsi(6) = 2
    Lsi(7) = 2
    Lsi(8) = 2
    Lsi(9) = 1
    Lsi(10) = 2
    Lsi(11) = 2
    Lsi(12) = 2
    Lsi(13) = 2
    Lsi(14) = 2
    Lsi(15) = 2
    Lsi(16) = 1
End Sub

Public Function FuncByteTo16Code(bytTmp() As Byte) As String
'功能：将byte转成16进制
    Dim i As Long
    
    If UBound(bytTmp) = 0 Then Exit Function
    For i = 0 To UBound(bytTmp)
        FuncByteTo16Code = FuncByteTo16Code & IIf(Len(Hex(bytTmp(i))) = 1, "0", "") & Hex(bytTmp(i))
    Next
End Function

Public Function Func16CodeToByte(strCode As String, bytCode() As Byte) As Boolean
'功能：将16进制转成byte()
    Dim i As Long
    
    If strCode = "" Then Exit Function
    ReDim bytCode(Len(strCode) / 2 - 1) As Byte
    For i = LBound(bytCode) To UBound(bytCode)
        bytCode(i) = CByte("&H" & Mid(strCode, i * 2 + 1, 2))
    Next
    Func16CodeToByte = True
End Function

