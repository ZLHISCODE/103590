VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Ruby_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Ѫ����
'�����ͺţ�Ruby
'���ң���������
Private mstrCode As String    '��������
Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="


Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸ֹͣ��������
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
                   strCmd As String)
    On Error GoTo errHand '�򿪴�������
    '------------------------------------------------
    '�豸��׼�ӿ�
    '���ܣ���������
    '������
    'strSource��Ҫ������ԭʼ��
    'strResult�����صļ�����(����������������밴���±�׼��֯���)
    '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ�������� ����� ^1 ��ʾ���� ��Ӽ� ^����
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    Dim lng_S As Long, lng_E As Long
    Dim str_Sample As String, str_Day As String
    Dim str_No As String, Str_Item As String, str_Value As String
    Dim INR_Value As String
    Dim str��� As String
    Dim strCode As String
    Dim j As Integer
    Dim lngID As String
    Dim strLine As String
    Dim var_Date As Variant
    Dim str��Ŀ As String
    Dim strVLineLeft  As String '���ֽ���
    Dim strVLineRight As String  '�Ҳ�ֽ���
    Dim strImageName  As String   'ͼ������
    Dim strImage  As String       'ͼ������
    Dim strWBC As String, strWBCImgae As String
    Dim strXY  As String   'ɢ��ͼ��λ��
    '    mstrCode = ""
    mstrCode = mstrCode & strSource

    'Ӧ��
    If InStr(strSource, Chr(&H5)) > 0 Then
        strCmd = Chr(&H6)
    ElseIf InStr(strSource, Chr(&H3)) > 0 Then
        strCmd = Chr(&H6)
    ElseIf InStr(strSource, Chr(&HA)) > 0 Then
        strCmd = Chr(&H6)
    ElseIf InStr(strSource, Chr(&H4)) > 0 Then
        strCmd = Chr(&H6)
    End If

    lng_S = InStr(mstrCode, Chr(&H2))
    If lng_S > 0 Then
        lng_E = InStr(lng_S + 1, mstrCode, "L|1|N")
    Else
        lng_E = InStr(mstrCode, "L|1|N")

    End If

    Do While lng_S > 0 And lng_E > 0

        strLine = ""
        strCode = Mid(mstrCode, lng_S, lng_E)
        mstrCode = Mid(mstrCode, lng_E)

        var_Date = Split(strCode, Chr(&HD))
        For j = LBound(var_Date) To UBound(var_Date)
            strLine = var_Date(j)
            strLine = strLine

            If Left(strLine, 3) = "O|1" Then

                str_Day = "": str_No = "":  str��Ŀ = ""
                str_Day = Format(Now, "yyyy-mm-dd hh:mm:ss")

                '�걾��
                str_No = Split(strLine, "|")(3)

            ElseIf InStr(strLine, "R|") > 0 And InStr(strLine, "F|") > 0 Then
                '����
                If str_Day = "" Then
                    str_Day = Split(strLine, "|")(11)
                    str_Day = Mid(str_Day, 1, 4) & "-" & Mid(str_Day, 5, 2) & "-" & Mid(str_Day, 7, 2) & "��" & _
                              Mid(str_Day, 9, 2) & ":" & Mid(str_Day, 11, 2) & ":" & Mid(str_Day, 13, 2)
                    str_Day = Format(str_Day, "yyyy-mm-dd hh:mm:ss")
                End If

                Str_Item = Split(strLine, "|")(2)
                Str_Item = Split(Str_Item, "^")(6)

                str_Value = Split(strLine, "|")(3)

                If str_Value = "" Then
                    str_Value = 0
                End If

                str��Ŀ = str��Ŀ & "|" & Str_Item & "|" & str_Value
            ElseIf Split(strLine, "|")(0) = "M" Then
                'ֱ��ͼ
                If strWBC <> "" Then
                     strWBC = Replace(Replace(strWBC, vbLf, ""), "", "")
                     Dim objPic  As New frmGraph
                     
                     Call DrawImage(objPic.Picture1, str_No, strXY, strWBC)
                     strWBC = ""
                     Unload objPic
'                    strWBC = Base64DecodeTwo(strWBC, 2)
'                     strWBCImgae = drawDIFF(strWBC, str_No, "WBC")
                End If
                If Split(strLine, "|")(3) = "RBC" Or Split(strLine, "|")(3) = "PLT" Then
                    strVLineLeft = Split(Split(strLine, "|")(5), "\")(0)
                    strVLineRight = Split(Split(strLine, "|")(5), "\")(1)
                    strImageName = IIf(Split(strLine, "|")(3) = "WB1", "WBC", Split(strLine, "|")(3))
                    strLine = Split(strLine, "|")(6)
                    strImage = strImage & "^" & strImageName & ";5;255,255;10,50,50,10;0-0,50-50,100-100,150-150,200-200,250-250|0-0,50-50,100-100,150-150,200-200,250-250;" & _
                               Base64Decode(strLine, False) & ";" & strVLineLeft & "," & strVLineRight
                ElseIf Split(strLine, "|")(3) = "WBC" Then
                    strWBC = Split(strLine, "|")(6)
                    strXY = Split(strLine, "|")(5)
                    strWBC = Mid(strWBC, 1, Len(strWBC) - 3)
'                    strWBC = Base64DecodeTwo(strWBC, 1)
                    
                End If
                
            ElseIf strWBC <> "" And Left(strLine, 2) = vbLf & "" Then
                strWBC = strWBC & Mid(strLine, 4, Len(strLine) - 3 - IIf(Right(strLine, 1) <> "=", 3, 0))
            End If

        Next

        str��� = str��� & "||" & str_Day & "|" & str_No & "|������|ѪҺ|0" & str��Ŀ & vbCrLf & Mid(strImage, 2) & "^WBC;115;1;" & App.Path & "\" & str_No & "_WBC.jpg" & "^RBC1;115;1;" & App.Path & "\" & str_No & "_RBC.jpg"

        lng_S = InStr(mstrCode, Chr(&HD))
        If lng_S > 0 Then
            lng_E = InStr(lng_S + 1, mstrCode, "L|1|N")
        Else
            lng_E = InStr(mstrCode, "L|1|N")
        End If
    Loop

    strResult = Mid(str���, 3)
    '------------------------------------------------
    Exit Sub
    '----------------
errHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\LisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLISDev_Ruby"
    Close #lngErrFileNo
End Sub

Public Function SendSample(ByVal strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    'strSampleInfo:������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ���
    '   ��5��Ԫ�أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    
End Function



'---------------------------ø���ǽӿ��Զ���ָ��--------------------
Public Function CmdAnalyse(strCmdType As String) As String

End Function


Public Function strValue(strItem As String, strC As Integer) As String

Dim str_j As String
str_j = strItem

strValue = Split(strItem, "^")(strC)

End Function


Private Function Base64Decode(strEncoded As String, ByVal blnimageType As Boolean) As String '����

    On Error Resume Next
    Dim arrB() As Byte, bTmp(3) As Byte, bT, bRet() As Byte
    Dim i As Long, j As Long
    Dim strResult   As String
    Dim var()   As String
    '    Dim BASE64CHR As String
    arrB = StrConv(strEncoded, vbFromUnicode)
    j = InStr(strEncoded & "=", "=") - 2
    ReDim bRet(j - j \ 4 - 1)
    Dim bRetTemp(3) As Byte
    For i = 0 To j Step 4
        Erase bTmp
        bTmp(0) = (InStr(BASE64CHR, Chr(arrB(i))) - 1) And 63
        bTmp(1) = (InStr(BASE64CHR, Chr(arrB(i + 1))) - 1) And 63
        bTmp(2) = (InStr(BASE64CHR, Chr(arrB(i + 2))) - 1) And 63
        bTmp(3) = (InStr(BASE64CHR, Chr(arrB(i + 3))) - 1) And 63

        bT = bTmp(0) * 2 ^ 18 + bTmp(1) * 2 ^ 12 + bTmp(2) * 2 ^ 6 + bTmp(3)
        bRet((i \ 4) * 3) = bT \ 65536
        bRet((i \ 4) * 3 + 1) = (bT And 65280) \ 256
        bRet((i \ 4) * 3 + 2) = bT And 255
        If blnimageType = False Then
            If (i \ 4) * 3 Mod 2 = 0 Then
                strResult = strResult & "," & bRet((i \ 4) * 3)
            End If
            If ((i \ 4) * 3 + 1) Mod 2 = 0 Then
                strResult = strResult & "," & bRet((i \ 4) * 3 + 1)
            End If
            If ((i \ 4) * 3 + 2) Mod 2 = 0 Then
                strResult = strResult & "," & bRet((i \ 4) * 3 + 2)
            End If
        Else
            strResult = strResult & ";" & bRet((i \ 4) * 3)
            strResult = strResult & ";" & bRet((i \ 4) * 3 + 1)
            strResult = strResult & ";" & bRet((i \ 4) * 3 + 2)
        End If
    Next
    Base64Decode = strResult
End Function


Private Sub DrawImage(ByVal Picture1 As PictureBox, ByVal strSampleNo As String, ByVal strCoordinate As String, ByVal strImageData As String)
    On Error GoTo hErr '�򿪴�������
    '------------------------------------------------
    Dim intLength  As Integer
    Dim strVar()  As String
    Dim strVarXY() As String    'xy����ϵ
    Dim lngColor  As Long
    Dim strLineXY()  As String   'ͼ�����ߵ�λ�� ��ȡ�������
    Dim i  As Integer
    Dim j As Integer
    Dim k As Integer
    Dim strTmp As String
    With Picture1
        .Appearance = 0
        .BorderStyle = 0
        .ScaleMode = 0
        .ScaleLeft = 0
        .ScaleTop = 2440
        .ScaleHeight = -2440
        .ScaleWidth = 2440
    End With
    strTmp = Base64Decode(strImageData, True)
    strVar = Split(Mid(strTmp, 2), ";")
    Picture1.DrawWidth = 1
    Picture1.Font.Size = 7
    strVarXY = Split(strCoordinate, "\")
    For k = 0 To 1
        Picture1.Line (400, 2440)-(400, 400)
        Picture1.Line (400, 400)-(2440, 400)
        strLineXY = Split(strVarXY(k + 1), "^")
        If k = 0 Then
            Picture1.Line (strLineXY(2) * 8 + 400, strLineXY(3) * 8 + 400)-(strLineXY(4) * 8 + 400, strLineXY(5) * 8 + 400)
            Picture1.CurrentX = 1000
            Picture1.CurrentY = 200
            Picture1.Print "��Ҷ��"
        Else
            Picture1.Line (strLineXY(2) * 8 + 400, strLineXY(3) * 8 + 400)-(strLineXY(4) * 8 + 400, strLineXY(5) * 8 + 400)
            Picture1.Line (strLineXY(4) * 8 + 400, strLineXY(5) * 8 + 400)-(strLineXY(6) * 8 + 400, strLineXY(7) * 8 + 400)
            Picture1.Line (strLineXY(6) * 8 + 400, strLineXY(7) * 8 + 400)-(strLineXY(8) * 8 + 400, strLineXY(9) * 8 + 400)
            Picture1.CurrentX = 1000
            Picture1.CurrentY = 200
            Picture1.Print "���Ӷ�"
        End If
        For i = 1 To 5
            ' ����
            Picture1.Line (350, i * 390 + 400)-(400, i * 390 + 400)
            Picture1.CurrentX = IIf(i <= 1, 100, 50)
            Picture1.CurrentY = i * 390 + 480
            Picture1.Print i * 50
            Picture1.Line (i * 390 + 400, 350)-(i * 390 + 400, 400)
            Picture1.CurrentY = 340
            Picture1.CurrentX = IIf(i <= 1, i * 390 + 220, i * 390 + 190)
            Picture1.Print i * 50
        Next
        Picture1.CurrentX = 250
        Picture1.CurrentY = 390
        Picture1.Print 0
        For i = 0 To UBound(strVar) Step 5
            Select Case strVar(i + 4)
                Case 0, 6, 19, 20, 21
                    lngColor = RGB(160, 160, 160)
                Case 1, 11, 17
                    lngColor = RGB(0, 204, 206)
                Case 2, 12
                    lngColor = RGB(250, 0, 250)
                Case 3, 13, 18
                    lngColor = RGB(0, 0, 0)
                Case 4, 10, 15
                    lngColor = RGB(255, 146, 6)
                Case 16
                    lngColor = RGB(0, 0, 255)
                Case 5
                    lngColor = RGB(0, 220, 21)
                Case 7, 8
                    lngColor = RGB(250, 175, 160)
                Case 9, 14
                    lngColor = RGB(255, 0, 0)
            End Select
            If k = 0 Then
                Picture1.PSet (strVar(i + 2) * 8 + 400, strVar(i + 3) * 8 + 400), lngColor
            Else
                Picture1.PSet (strVar(i + 1) * 8 + 400, strVar(i) * 8 + 400), lngColor
            End If
            lngColor = RGB(255, 255, 255)
            
        Next
        SavePicture Picture1.Image, App.Path & "\" & strSampleNo & IIf(k = 0, "_WBC.jpg", "_RBC.jpg")
        Picture1.Cls
    Next
    '------------------------------------------------
    Exit Sub
    '----------------
hErr:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    
    Open App.Path & "\LisErrMsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",DrawImage,Form1"
    Close #lngErrFileNo
End Sub

