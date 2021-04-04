Attribute VB_Name = "mdlGraph_Sysmex"
Option Explicit
'
'调用函数,将图形数据保存为图像的模块,只适用于希森美康的仪器发回的数据.
Public Type typHistGramInfo
    lngstoppos As Long
    lngmaxx As Long
    lngmaxy As Long
    lnglower As Long
    lngupper As Long
    lngresver1 As Long
    lngresver2 As Long
End Type

Private Declare Function strufhistgramprocess Lib "SsmDraw.dll" Alias "strUFHistGramProcess" (ByVal strData As String, ByVal lngdatalen As Long, ByVal lngx As Long, ByVal lngy As Long, ByVal strtempfile As String, ByVal strFilename As String) As Integer
Private Declare Function strkhistgramprocess Lib "SsmDraw.dll" Alias "strKHistGramProcess" (ByVal strData As String, ByVal lngdatalen As Long, ByVal typhisgram As typHistGramInfo, ByVal strtempfile As String, ByVal strFilename As String) As Integer
Private Declare Function strufscatgramprocess Lib "SsmDraw.dll" Alias "strUFScatGramProcess" (ByVal strData As String, ByVal lngdatalen As Long, ByVal lngx As Long, ByVal lngy As Long, ByVal strtempfile As String, ByVal strFilename As String) As Integer
Private Declare Function strsehistgramprocess Lib "SsmDraw.dll" Alias "strSEHistGramProcess" (ByVal strData As String, ByVal lngdatalen As Long, ByVal typhisgram As typHistGramInfo, ByVal strtempfile As String, ByVal strFilename As String) As Integer
Private Declare Function strscatgramprocess Lib "SsmDraw.dll" Alias "strScatGramProcess" (ByVal strData As String, ByVal lngdatalen As Long, ByVal strtempfile As String, ByVal strFilename As String) As Integer
Private Declare Function strhistgramprocess Lib "SsmDraw.dll" Alias "strHistGramProcess" (ByVal strData As String, ByVal lngdatalen As Long, ByRef typhisgram As typHistGramInfo, ByVal strtempfile As String, ByVal strFilename As String) As Integer
Private Declare Function makegif Lib "Ssmg7.dll" Alias "MakeGif" (ByVal strbmp As String, ByVal strgif As String, ByVal strData As String) As Integer
Private Declare Function intsetbackgroundcolor Lib "SsmDraw.dll" Alias "intSetBackgroundColor" (ByVal strgirfilename As String, ByVal inttype As Integer) As Integer

Private Declare Function struf1000histgramprocess Lib "SsmDraw.dll" Alias "strUF1000HistGramProcess" (ByVal strData As String, ByVal lngdatalen As Long, ByVal lngx As Long, ByVal lngy As Long, ByVal strtempfile As String, ByVal strFilename As String) As Integer
Private Declare Function struf1000scatgramprocess Lib "SsmDraw.dll" Alias "strUF1000ScatGramProcess" (ByVal strData As String, ByVal lngdatalen As Long, ByVal lngx As Long, ByVal lngy As Long, ByVal strtempfile As String, ByVal strFilename As String) As Integer
Private Declare Function struf1001scatgramprocess Lib "SsmDraw.dll" Alias "strUF1001ScatGramProcess" (ByVal strData As String, ByVal lngdatalen As Long, ByVal lngx As Long, ByVal lngy As Long, ByVal strtempfile As String, ByVal strFilename As String) As Integer
'---actdiff 2 的DLL
Private Declare Function Bit_And Lib "cm.dll" Alias "BitAnd" (ByVal op_1 As Long, ByVal op_2 As Long) As Long
Private Declare Function Bit_Or Lib "cm.dll" Alias "BitOr" (ByVal op_1 As Long, ByVal op_2 As Long) As Long
Private Declare Function Bit_Not Lib "cm.dll" Alias "BitNot" (ByVal op_1 As Long) As Long
Private Declare Function Bit_Xor Lib "cm.dll" Alias "BitXor" (ByVal op_1 As Long, ByVal op_2 As Long) As Long
Private Declare Function Bit_Mid Lib "cm.dll" Alias "BitMid" (ByVal bits As Long, ByVal right_start As Long, ByVal l_len As Long) As Long
Private Declare Function Left_Shift Lib "cm.dll" Alias "LeftShift" (ByVal bits As Long, ByVal shift As Long) As Long
Private Declare Function Right_Shift Lib "cm.dll" Alias "RightShift" (ByVal bits As Long, ByVal shift As Long) As Long

'Public Function uf_G7Scat(ByVal as_graphdata As String, ByVal as_tempfile As String, ByVal as_file As String) As Integer
'    Dim ls_graphdata As String, ls_tempfile As String, ls_file As String
'    ls_graphdata = as_graphdata
'    ls_tempfile = as_tempfile
'    ls_file = as_file
'
'    uf_G7Scat = makegif(ls_tempfile, ls_file, ls_graphdata)
'End Function

Public Function uf_ufHist(ByVal as_graphdata As String, ByVal al_len As Long, ByVal as_tempfile As String, ByVal as_file As String) As Integer
    'RBC WBC 直方图
    '入参: as_graphdata 图形数据
    '      al_len       图形数据长度
    '      as_tempfile  临时文件名
    '      as_file      产生的图像文件名
    Dim ls_graph As String, ls_tempfile As String, ls_file As String, ll_len As Long
    Dim intReturn As Integer
    
    ls_graph = as_graphdata
    ls_tempfile = as_tempfile
    ls_file = as_file
    ll_len = al_len
    intReturn = strufhistgramprocess(ls_graph, ll_len, 150, 75, ls_tempfile, ls_file)
    'uf_ufHist = intsetbackgroundcolor(ls_file, 0)
    uf_ufHist = intReturn
End Function

Public Function uf_ufscat(ByVal as_graphdata As String, ByVal al_len As Long, ByVal as_tempfile As String, ByVal as_file As String, ByVal al_lngx As Integer, ByVal al_lngy As Integer) As Integer
    '散点图
    '入参: as_graphdata 图形数据
    '      al_len       图形数据长度
    '      as_tempfile  临时文件名
    '      as_file      产生的图像文件名
    Dim ls_graph As String, ls_tempfile As String, ls_file As String, ll_len As Long
    Dim intReturn As Integer
    ls_graph = as_graphdata
    ls_tempfile = as_tempfile
    ls_file = as_file
    ll_len = al_len
    
    intReturn = strufscatgramprocess(as_graphdata, al_len, al_lngx, al_lngy, as_tempfile, as_file)
    'uf_ufscatprocess = intsetbackgroundcolor(ls_file, 1)
    uf_ufscat = intReturn
End Function

Public Function uf_uf1000Hist(ByVal as_graphdata As String, ByVal al_len As Long, ByVal as_tempfile As String, ByVal as_file As String) As Integer
    'RBC WBC 直方图
    '入参: as_graphdata 图形数据
    '      al_len       图形数据长度
    '      as_tempfile  临时文件名
    '      as_file      产生的图像文件名
    Dim ls_graph As String, ls_tempfile As String, ls_file As String, ll_len As Long
    Dim intReturn As Integer
    
    ls_graph = as_graphdata
    ls_tempfile = as_tempfile
    ls_file = as_file
    ll_len = al_len
    intReturn = struf1000histgramprocess(ls_graph, ll_len, 150, 75, ls_tempfile, ls_file)
    'uf_ufHist = intsetbackgroundcolor(ls_file, 0)
    uf_uf1000Hist = intReturn
End Function
Public Function uf_uf1000scat(ByVal as_graphdata As String, ByVal al_len As Long, ByVal as_tempfile As String, ByVal as_file As String, ByVal al_lngx As Integer, ByVal al_lngy As Integer) As Integer
    '散点图
    '入参: as_graphdata 图形数据
    '      al_len       图形数据长度
    '      as_tempfile  临时文件名
    '      as_file      产生的图像文件名
    Dim ls_graph As String, ls_tempfile As String, ls_file As String, ll_len As Long
    Dim intReturn As Integer
    ls_graph = as_graphdata
    ls_tempfile = as_tempfile
    ls_file = as_file
    ll_len = al_len
    
    intReturn = struf1000scatgramprocess(as_graphdata, al_len, al_lngx, al_lngy, as_tempfile, as_file)
    'uf_ufscatprocess = intsetbackgroundcolor(ls_file, 1)
    uf_uf1000scat = intReturn
End Function
Public Function uf_uf1001scat(ByVal as_graphdata As String, ByVal al_len As Long, ByVal as_tempfile As String, ByVal as_file As String, ByVal al_lngx As Integer, ByVal al_lngy As Integer) As Integer
    '散点图
    '入参: as_graphdata 图形数据
    '      al_len       图形数据长度
    '      as_tempfile  临时文件名
    '      as_file      产生的图像文件名
    Dim ls_graph As String, ls_tempfile As String, ls_file As String, ll_len As Long
    Dim intReturn As Integer
    ls_graph = as_graphdata
    ls_tempfile = as_tempfile
    ls_file = as_file
    ll_len = al_len
    
    intReturn = struf1001scatgramprocess(as_graphdata, al_len, al_lngx, al_lngy, as_tempfile, as_file)
    'uf_ufscatprocess = intsetbackgroundcolor(ls_file, 1)
    uf_uf1001scat = intReturn
End Function
Public Function uf_xehist(ByVal as_graphdata As String, ByVal al_len As Long, astr_info As typHistGramInfo, ByVal as_tempfile As String, ByVal as_file As String) As Integer
    '产生血常规直方图
    Dim lstr As typHistGramInfo, ls_graph As String, ls_tempfile As String, ls_file As String, ll_len As Long
    ls_graph = as_graphdata
    ls_tempfile = as_tempfile
    ls_file = as_file
    ll_len = al_len
    lstr.lngstoppos = astr_info.lngstoppos
    lstr.lngmaxx = astr_info.lngmaxx
    lstr.lngmaxy = astr_info.lngmaxy
    lstr.lnglower = astr_info.lnglower
    lstr.lngupper = astr_info.lngupper
    lstr.lngresver1 = astr_info.lngresver1
    lstr.lngresver2 = astr_info.lngresver2
    uf_xehist = strhistgramprocess(ls_graph, ll_len, lstr, ls_tempfile, ls_file)
    Call intsetbackgroundcolor(ls_file, 1) '背景白色
End Function

Public Function uf_xescat(ByVal as_graphdata As String, ByVal al_len As Long, ByVal as_tempfile As String, ByVal as_file As String) As Integer
    '产生血常规散点图
    
    Dim ls_graph As String, ls_tempfile As String, ls_file As String, ll_len As Long
    ls_graph = as_graphdata
    ls_tempfile = as_tempfile
    ls_file = as_file
    ll_len = al_len
    uf_xescat = strscatgramprocess(ls_graph, ll_len, ls_tempfile, ls_file)
    Call intsetbackgroundcolor(ls_file, 1) '背景白色
End Function


Private Function wf_decode(ByVal Code As String, ByVal Kind As Integer) As String
    '解码 act2数据
    Dim ll_i As Long
    Dim ll_val As Long
    Dim ll_result As Long
    Dim ls_result As String
    Dim ls_char(4) As String
    
    Select Case Kind
    Case 2, 3, 4

        For ll_i = 1 To Kind
            ll_val = Asc(Mid(Code, ll_i, 1)) - 48
            ll_val = Left_Shift(Bit_Mid(ll_val, 6, 6), (Kind - ll_i) * 6)
            ll_result = Bit_Or(ll_result, ll_val)
        Next

        For ll_i = Kind - 1 To 1 Step -1
            ls_char(ll_i) = Format(Bit_Mid(ll_result, 8, 8), "000")
            ll_result = Right_Shift(ll_result, 8)
        Next

        For ll_i = 1 To Kind - 1
            ls_result = ls_result & ";" & ls_char(ll_i)
        Next
        
    End Select
    wf_decode = ls_result
End Function

Public Function de_code(ByVal ls_data As String) As String
    Dim ls_block As String
    Dim ls_result  As String
    
    Do While Len(ls_data) >= 4
        ls_block = Left(ls_data, 4)
        ls_result = ls_result & wf_decode(ls_block, 4)
        ls_data = Mid(ls_data, 5)
    Loop
    
    de_code = ls_result
End Function
