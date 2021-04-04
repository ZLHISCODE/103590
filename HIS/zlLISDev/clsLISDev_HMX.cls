VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_HMX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:HMX
'厂商:美国贝克曼库尔特 Beckman coulter
'作者:陈东
'用户:吉化二院
'修改人:李建飞
'修改日期:2010-04-27
'修改内容:1.由于两家用户参数设置不一致，所以采集到的数据格式不一致
'           修改后兼容两种数据格式
'         2.修改获取日期格式方法

Private mstrCode As String

Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '设备标准接口
    '功能：获取设备停止发送命令
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    Dim lng_S As Long, lng_E As Long
    Dim str_Data As String, varline As Variant, str_Line As String
    Dim i As Integer
    Dim str_日期 As String, str_标本号 As String, str_项目 As String
    
    Dim frmGraph_Tmp As Form
    Dim str_RBCH As String, str_PLTH As String, str_PLTF As String, str_DF1 As String, str_DF2 As String
    Dim str_V As String, str_C As String, str_S As String, str_DF1_Line As String
    Dim strImage As String
    
    str_日期 = "": str_标本号 = "": str_项目 = "": strImage = ""
    
    mstrCode = mstrCode & strSource
    If mstrCode Like "*01*2F*--**" Then
        Do While InStr(mstrCode, "") > 0
            lng_S = InStr(mstrCode, "")
            If lng_S > 0 Then
                mstrCode = Mid(mstrCode, 1, lng_S - 5) & Mid(mstrCode, lng_S + 4)
            End If
        Loop
        lng_S = InStr(mstrCode, "01")
        mstrCode = Mid(mstrCode, lng_S)
        
        '----- 取日期,标本号
        lng_S = InStr(mstrCode, "G01")
        lng_E = InStr(mstrCode, "G02")
        If lng_S > 0 And lng_E > 0 Then
            str_Data = Mid(mstrCode, lng_S, lng_E - lng_S)
            mstrCode = Mid(mstrCode, lng_E)
            
            varline = Split(str_Data, vbCrLf)
            For i = LBound(varline) To UBound(varline)
                str_Line = varline(i)
                str_Line = DelCRC(str_Line)
                If str_Line Like "DATE*" Then
                    str_日期 = Format(Mid(str_Line, 12, 2) & "-" & Mid(str_Line, 6, 2) & "-" & Mid(str_Line, 9, 2), "yyyy-MM-dd")
                    '原处理格式有问题
                    'When str_Line="DATE 04/28/10" Then Format(CDate(Mid(str_Line, 5)), "yyyy-MM-dd") = 2010-04-28
                    'When str_Line="DATE 05/02/10" Then Format(CDate(Mid(str_Line, 5)), "yyyy-MM-dd") = 2005-02-10
                    'str_日期 = Format(CDate(Mid(str_Line, 5)), "yyyy-MM-dd")
                ElseIf str_Line Like "CASSPOS*" Then
                ElseIf str_Line Like "ID1*" Then
                    str_标本号 = Val(Mid(str_Line, 4))
                ElseIf str_Line Like "ID2*" Then
                    If str_标本号 = 0 Then str_标本号 = Val(Mid(str_Line, 4))
                ElseIf str_Line Like "SEQUENCE*" Then
                    If Val(str_标本号) <= 0 Then str_标本号 = Val(Mid(str_Line, 9))
                End If
            Next
        End If
        '---- 取检验项目
        lng_S = InStr(mstrCode, "G02")
        lng_E = InStr(mstrCode, "G06")
        If lng_S > 0 And lng_E > 0 Then
            str_Data = Mid(mstrCode, lng_S, lng_E - lng_S)
            mstrCode = Mid(mstrCode, lng_E)
            varline = Split(str_Data, vbCrLf)
            For i = LBound(varline) To UBound(varline)
                str_Line = varline(i)
                str_Line = DelCRC(str_Line)
                If InStr(",WBC ,RBC ,HGB ,HCT ,MCV ,MCH ,MCHC,RDW ,PLT ,PCT ,MPV ,PDW ,", "," & Mid(str_Line, 1, 4) & ",") > 0 Then
                    str_项目 = str_项目 & "|" & Trim(Mid(str_Line, 1, 4)) & "|" & Val(Mid(str_Line, 5))
                End If
            Next
        End If
        '--取RBCH数据

        str_RBCH = Get_PicData(mstrCode, "G0E", "G0F", "RBCH")
        '--- 取PLTH,PLTF
        str_PLTH = Get_PicData(mstrCode, "G0F", "G01", "PLTH,PLTF")
        str_PLTF = Split(str_PLTH, ",")(1)
        str_PLTH = Split(str_PLTH, ",")(0)
        
        '-------- 取检验项目
        lng_S = InStr(mstrCode, "G03")
        lng_E = InStr(mstrCode, "G04")
        If lng_S > 0 And lng_E > 0 Then
            str_Data = Mid(mstrCode, lng_S, lng_E - lng_S)
            mstrCode = Mid(mstrCode, lng_E)
            varline = Split(str_Data, vbCrLf)
            For i = LBound(varline) To UBound(varline)
                str_Line = varline(i)
                str_Line = DelCRC(str_Line)
                If InStr(",LY#,MO#,NE#,EO#,BA#,", "," & Mid(str_Line, 1, 3) & ",") > 0 Then
                    str_项目 = str_项目 & "|" & Trim(Mid(str_Line, 1, 3)) & "|" & Val(Mid(str_Line, 4))
                End If
            Next
        End If
        lng_S = InStr(mstrCode, "G04")
        lng_E = InStr(mstrCode, "G0B")
        If lng_S > 0 And lng_E > 0 Then
            str_Data = Mid(mstrCode, lng_S, lng_E - lng_S)
            mstrCode = Mid(mstrCode, lng_E)
            varline = Split(str_Data, vbCrLf)
            For i = LBound(varline) To UBound(varline)
                str_Line = varline(i)
                str_Line = DelCRC(str_Line)
                If InStr(",LY%,MO%,NE%,EO%,BA%,", "," & Mid(str_Line, 1, 3) & ",") > 0 Then
                    str_项目 = str_项目 & "|" & Trim(Mid(str_Line, 1, 3)) & "|" & Val(Mid(str_Line, 4))
                End If
            Next
        End If
        '-------DF1
        lng_S = InStr(mstrCode, "G0B")
        lng_E = InStr(mstrCode, "G0C")
        If lng_S > 0 And lng_E > 0 Then
            str_Data = Mid(mstrCode, lng_S, lng_E - lng_S)
            varline = Split(str_Data, vbCrLf)
            For i = LBound(varline) To UBound(varline)
                str_Line = varline(i)
                str_Line = DelCRC(str_Line)
                If Mid(str_Line, 1, 3) = "VAL" Then
                    str_DF1_Line = str_DF1_Line & "," & Mid(str_Line, 6)
                End If
            Next
        End If
        If str_DF1_Line <> "" Then str_DF1_Line = Mid(str_DF1_Line, 2)
        
        str_DF1 = Get_PicData(mstrCode, "G0B", "G0C", "DF1")
        '-------DF2
        str_DF2 = Get_PicData(mstrCode, "G0C", "G0D", "DF2")
        '-------V,C,S
        str_V = Get_PicData(mstrCode, "G0D", "--", "V,C,S")
        str_C = Split(str_V, ",")(1)
        str_S = Split(str_V, ",")(2)
        str_V = Split(str_V, ",")(0)
        '-------------------------------------------------------------
        If Val(str_标本号) > 0 And Trim(str_项目) <> "" Then
            strResult = str_日期 & "|" & Val(str_标本号) & "| |血液|0" & str_项目
            If str_RBCH <> "" Then
                strImage = strImage & "^RBC;0;200,262,25-50,50-100,100-200,150-" & Get_ImgData(str_RBCH)
            End If
            If str_V <> "" Then
                strImage = strImage & "^V;0;200,262" & Get_ImgData(str_V)
            End If
            If str_C <> "" Then
                strImage = strImage & "^C;0;200,262" & Get_ImgData(str_C)
            End If
            If str_S <> "" Then
                strImage = strImage & "^S;0;200,262" & Get_ImgData(str_S)
            End If
            If str_PLTF <> "" And str_PLTH <> "" Then
                strImage = strImage & "^PLT;4;200,262;" & Replace(Mid(Get_ImgData(str_PLTF), 2), ";", ",") & "|0,0,0,0,0,0,0,0,0,0,0,0,0,0,0," & Replace(Mid(Get_ImgData(str_PLTH), 2), ";", ",")
            End If
            
            If str_DF1 <> "" Then
                Set frmGraph_Tmp = New frmGraph
                strImage = strImage & "^WBC DF1;100;1;" & frmGraph_Tmp.Draw_HMX_DF1(str_DF1_Line, str_DF1)
                Unload frmGraph_Tmp
            End If
            
            If str_DF2 <> "" Then
                Set frmGraph_Tmp = New frmGraph
                strImage = strImage & "^WBC DF2;100;1;" & frmGraph_Tmp.Draw_HMX_DF2(str_DF2)
                Unload frmGraph_Tmp
            End If
            
            If strImage <> "" Then
                strResult = strResult & vbCrLf & Mid(strImage, 2)
            End If
        End If
        If InStr(mstrCode, "") <= 0 Then mstrCode = ""
    End If
    
    
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", _
       Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String

    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function


Private Function Get_PicData(str_In As String, ByVal str_S As String, ByVal str_E As String, ByVal str_Name As String) As String
        '取图形数据
        Dim lng_S As Long, lng_E As Long, int_Len As Integer, i As Integer, Y As Integer
        Dim str_Return As String, varName As Variant, str_Data As String, varline As Variant, str_Line As String
        lng_S = InStr(str_In, str_S)
        lng_E = InStr(str_In, str_E)
        If lng_S > 0 And lng_E > 0 Then
            str_Data = Mid(str_In, lng_S, lng_E - lng_S)
            str_In = Mid(str_In, lng_E)
            
            
            varline = Split(str_Data, vbCrLf)
            For i = LBound(varline) To UBound(varline)
                str_Line = varline(i)
                varName = Split(str_Name, ",")
                
                For Y = LBound(varName) To UBound(varName)
                    int_Len = Len(varName(Y))
                    If Mid(str_Line, 1, int_Len) = varName(Y) Then
                        str_Return = Mid(str_Line, int_Len + 1)
                        str_Return = Mid(str_Return, 2)
                        Do While InStr(str_Return, "") > 0
                            lng_S = InStr(str_Return, "")
                            If lng_S > 0 Then
                                str_Return = Mid(str_Return, 1, lng_S - 5) & Mid(str_Return, lng_S + 4)
                            End If
                        Loop
                        Get_PicData = Get_PicData & "," & str_Return
                    End If
                Next
            Next
        End If
        If Get_PicData <> "" Then Get_PicData = Mid(Get_PicData, 2)
End Function

Private Function Get_ImgData(ByVal str_In As String) As String
    '转换图形数据
    Dim str_Line As String, str_Item As String, str_R As String
    str_Line = str_In
    
    Do While Len(str_Line) > 0
        str_Item = Left(str_Line, 2)
        str_Line = Mid(str_Line, 3)
        str_R = str_R & ";" & CByte("&H" & str_Item)
    Loop
    If str_R <> "" Then Get_ImgData = str_R
End Function

Private Function DelCRC(ByVal str_In As String) As String
    '取掉CRC校验字符
    Dim str_Line As String
    Dim lng_S As Long
    
    str_Line = str_In
    If InStr(str_Line, "") > 0 Then
        lng_S = InStr(str_Line, "")
        str_Line = Mid(str_Line, 1, lng_S - 5) & Mid(str_Line, lng_S + 4)
        DelCRC = DelCRC(str_Line)
    Else
        DelCRC = str_Line
    End If
    
End Function
