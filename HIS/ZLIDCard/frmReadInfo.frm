VERSION 5.00
Begin VB.Form frmReadInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   4035
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   255
      Top             =   120
   End
End
Attribute VB_Name = "frmReadInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public glngParentHwnd As Long
Private mstrPath As String          '���������ļ�·��

Private mstrID As String
Private mstrName As String
Private mstrSex As String
Private mstrNation As String
Private mdatBirthDay As Date
Private mstrAddress As String
Private mpicPhoto As StdPicture     '���֤��Ƭ��Ϣ
Private mblnAutoRead As Boolean  '�Ƿ��Զ������������� SetEnable ����Ϊ�Զ����������� ReadIDCard ����Ϊ �ֶ�����

Private Enum ReadMode
    Base = 1        '�γ�������Ϣ�ļ�WZ.TXT����Ƭ�ļ�XP.WLT��ZP.BMP
    onlytext = 2    '�γ�������Ϣ�ļ�WZ.TXT����Ƭ�ļ�XP.WLT
    NewAdd = 3      '�γ�����סַ�ļ�NEWADD.TXT
End Enum

Private Const TXTFile = "\wz.txt"
Private Const BMPFile = "\zp.bmp"
Private Const WLTFile = "\zp.wlt"
Private pucManaMsg As String * 4
Private Const IfOpen = 0 '0��ʾ���ڸú����ڲ��򿪺͹رմ��ڣ���ʱȷ��֮ǰ������Syn_OpenPort���򿪶˿ڣ������ڲ���Ҫ��˿�ͨ��ʱ������Syn_ClosePort�رն˿ڣ�
                        '��0��ʾ��API�����ڲ������˴򿪶˿ں͹رն˿ں�����֮ǰ����Ҫ����Syn_OpenPort��Ҳ�����ٵ���Syn_ClosePort?
                        
'Private pucIIN As Integer, pucSN As Integer, puiCHMsgLen As Integer, puiPHMsgLen As Integer, iIfOpen As Integer
Private pucIIN As String * 8
Private pucSN As String * 8
Private puiCHMsgLen As Long
Private puiPHMsgLen As Long
Private iIfOpen As Integer

Private lngReturn As Integer '���巵�ؽ��ֵ
Private mblnCancel As Boolean
Public mobjIDCard As clsIDCard

Private Const GWL_STYLE = (-16)
Private Const WS_DISABLED = &H8000000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long


Private Sub Form_Load()
    iIfOpen = 1
    mblnAutoRead = True         '����Ĭ��Ϊ�Զ�����
    mstrPath = GetSetting("ZLSOFT", "����ȫ��", "����·��", "C:")
    If mstrPath <> "C:" Then mstrPath = Mid(mstrPath, 1, InStrRev(mstrPath, "\") - 1)

    If Dir(mstrPath, vbDirectory) = "" Then
        MsgBox "ZLHISӦ�ó���Ŀ¼��C�̶�������!���ܶ���", vbInformation, App.ProductName
        Timer1.Enabled = False
    End If
End Sub

Private Sub Timer1_Timer()
    Dim i As Integer
    If (GetWindowLong(mobjIDCard.GetParent, GWL_STYLE) And WS_DISABLED) <> WS_DISABLED Then
        '107213:���ϴ�,2017/4/12,GDI��������,Ϊ���⿨�����ڶ�����ѯ�ſ�ʼ����
        If mobjIDCard.GetParent <> 0 Then
            If GetActiveWindow <> glngParentHwnd Then mblnCancel = True: Exit Sub
            If mblnCancel Then mblnCancel = False: Exit Sub '����һ��
        End If
        Select Case glngType
            Case IDCardType.CVR100U, IDCardType.CVR100D, IDCardType.SS_V1, IDCardType.XZX_KDQ, IDCardType.GTICR100, IDCardType.GTICR100_01, _
                IDCardType.GTICR100_1
                If Authenticate = 1 Then Call ReadIDCard
            Case IDCardType.CVR100U_1, IDCardType.CVR100D_1
                If CVR_Authenticate = 1 Then Call ReadIDCard
            Case IDCardType.HX_FDX9
                If SDT_StartFindIDCard(1, "", 1) = CByte(&H9F) Then Call ReadIDCard
            Case IDCardType.DKQ_116D
                lngReturn = Syn_ClosePort(1001)
                lngReturn = Syn_OpenPort(1001)
                If lngReturn = 0 Then Call ReadIDCard
            Case IDCardType.CVR100
                lngReturn = SDT_StartFindIDCard(1001, "", 1)
                If lngReturn = CByte(&H9F) Then Call ReadIDCard
            Case IDCardType.COMMON
                '�ҿ�
                i = SDT_StartFindIDCard(editPort, pucIIN, iIfOpen)
                If i <> CByte(&H9F) Then
                    '���ҿ�
                    i = SDT_StartFindIDCard(editPort, "", 1)
                    If i <> CByte(&H9F) Then
                        i = SDT_ClosePort(editPort)
                    Else
                        Call ReadIDCard
                    End If
                Else
                    Call ReadIDCard
                End If
            Case IDCardType.SS728M01_B01C
                Call ReadIDCard
        End Select
    End If
End Sub


Private Sub ReadIDCard()
    Dim intTmp As Integer, strMSG As String
    Dim strPucManaMsg As String
    Dim i As Integer
    
    Select Case glngType
        Case IDCardType.CVR100U, IDCardType.CVR100D, IDCardType.SS_V1, IDCardType.XZX_KDQ, IDCardType.GTICR100
            intTmp = Read_Content_Path(mstrPath, ReadMode.Base)
            Select Case intTmp
                Case 0
                    strMSG = "�뽫���֤ͣ�����豸��ʶ��λ��,���ٱ���1��!"
                Case 1
                    strMSG = "" '�ɹ�
                Case 2
                    strMSG = "û������סַ��Ϣ!"
                Case -1
                    strMSG = "��Ƭ�������!"
                Case -2
                    strMSG = "wlt�ļ���׺����!"
                Case -3
                    strMSG = "wlt�ļ��򿪴���!"
                Case -4
                    strMSG = "wlt�ļ���ʽ����!"
                Case -5
                    strMSG = "���δ��Ȩ!"
                Case -11
                    strMSG = "��Ч����!"
                Case -12
                    strMSG = "·��̫��!"
                Case Else
                    strMSG = "�豸δ֪����!"
            End Select
        Case IDCardType.GTICR100_1, IDCardType.GTICR100_01
            intTmp = Read_Content(ReadMode.Base)
            Select Case intTmp
                Case 0
                    strMSG = "�뽫���֤ͣ�����豸��ʶ��λ��,���ٱ���1��!"
                Case 1
                    strMSG = "" '�ɹ�
                Case 2
                    strMSG = "û������סַ��Ϣ!"
                Case -1
                    strMSG = "��Ƭ�������!"
                Case -2
                    strMSG = "wlt�ļ���׺����!"
                Case -3
                    strMSG = "wlt�ļ��򿪴���!"
                Case -4
                    strMSG = "wlt�ļ���ʽ����!"
                Case -5
                    strMSG = "���δ��Ȩ!"
                Case -11
                    strMSG = "��Ч����!"
                Case -12
                    strMSG = "·��̫��!"
                Case Else
                    strMSG = "�豸δ֪����!"
            End Select
        Case IDCardType.CVR100U_1, IDCardType.CVR100D_1
            intTmp = CVR_Read_Content(ReadMode.Base)
            Select Case intTmp
                Case 0
                    strMSG = "�뽫���֤ͣ�����豸��ʶ��λ��,���ٱ���1��!"
                Case 1
                    strMSG = "" '�ɹ�
                Case 2
                    strMSG = "û������סַ��Ϣ!"
                Case -1
                    strMSG = "��Ƭ�������!"
                Case -2
                    strMSG = "wlt�ļ���׺����!"
                Case -3
                    strMSG = "wlt�ļ��򿪴���!"
                Case -4
                    strMSG = "wlt�ļ���ʽ����!"
                Case -5
                    strMSG = "���δ��Ȩ!"
                Case -11
                    strMSG = "��Ч����!"
                Case -12
                    strMSG = "·��̫��!"
                Case Else
                    strMSG = "�豸δ֪����!"
            End Select
        Case IDCardType.HX_FDX9
            intTmp = SDT_SelectIDCard(1, strPucManaMsg, 1)
            Select Case intTmp
                Case CByte(&H90)
                    strMSG = "" '�ɹ�
                    intTmp = SDT_ReadBaseMsgToFile(1, mstrPath & TXTFile, puiCHMsgLen, mstrPath & WLTFile, puiPHMsgLen, 1)
                    If intTmp = CByte(&H90) Then
                        intTmp = GetBmp(mstrPath & WLTFile, 1)
                        If intTmp <> 1 Then
                            Timer1.Enabled = False
                            strMSG = "��Ƭ����ʧ�ܣ�"
                            MsgBox strMSG, vbInformation, App.ProductName
                            Timer1.Enabled = True
                        End If
                    Else
                        Timer1.Enabled = False
                        strMSG = "����ʧ�ܣ�"
                        MsgBox strMSG, vbInformation, App.ProductName
                        Timer1.Enabled = True
                    End If
                Case CByte(&H81)
                    strMSG = "ѡ��ʧ�ܣ�"
            End Select
        Case IDCardType.CVR100
            intTmp = SDT_SelectIDCard(1001, strPucManaMsg, 1)
            Select Case intTmp
                Case CByte(&H90)
                    strMSG = "" '�ɹ�
                    intTmp = SDT_ReadBaseMsgToFile(1001, mstrPath & TXTFile, puiCHMsgLen, mstrPath & WLTFile, puiPHMsgLen, 1)
                    If intTmp = CByte(&H90) Then
                        intTmp = GetBmp(mstrPath & WLTFile, 1)
                        If intTmp <> 1 Then
                            Timer1.Enabled = False
                            strMSG = "��Ƭ����ʧ�ܣ�"
                            MsgBox strMSG, vbInformation, App.ProductName
                            Timer1.Enabled = True
                        End If
                    Else
                        Timer1.Enabled = False
                        strMSG = "����ʧ�ܣ�"
                        MsgBox strMSG, vbInformation, App.ProductName
                        Timer1.Enabled = True
                    End If
                Case CByte(&H81)
                    strMSG = "ѡ��ʧ�ܣ�"
            End Select
        Case IDCardType.DKQ_116D
            Call ClearIDCardInfor   '��ʱ��սṹ�壬����Syn_ReadMsg����δ����0ʱ����Ȼ���Զ�ȡ����Ƭ�ֶ�
            intTmp = Syn_StartFindIDCard(1001, pucManaMsg, IfOpen)
            intTmp = Syn_SelectIDCard(1001, pucManaMsg, IfOpen)
            intTmp = Syn_ReadMsg(1001, IfOpen, IDCardInfor)
            Call Syn_ClosePort(1001)   '�رն˿ڣ���ֹ�ര�ڵ���ʱ��ͻ
            Select Case intTmp
                Case 0
                    strMSG = "" '�����ɹ�����Ƭ���������ȷ
                Case -1
                    strMSG = "�˿ڴ�ʧ��/�˿���δ��/�˿ںŲ��Ϸ�"
                Case -2
                    strMSG = "֤/���д���������"
                Case -3
                    strMSG = "PC���ճ�ʱ���ڹ涨��ʱ����δ���յ��涨���ȵ�����"
                Case -4
                    strMSG = "���ݴ������"
                Case -5
                    strMSG = "��SAM_V���ڲ����ã�ֻ��SDT_GetCOMBaudʱ���п��ܷ���"
                Case -6
                    strMSG = "����ҵ���ն����ݵ�У��ʹ�"
                Case -7
                    strMSG = "����ҵ���ն����ݵĳ��ȴ�"
                Case -8
                    strMSG = "����ҵ���ն˵�������󣬰��������еĸ�����ֵ���߼��������"
                Case -9
                    strMSG = "ԽȨ����"
                Case -10
                    strMSG = "�޷�ʶ��Ĵ���"
                Case -11
                    strMSG = "Ѱ��֤/��ʧ��"
                Case -12
                    strMSG = "ѡȡ֤/��ʧ��"
                Case -13
                    strMSG = "����sdtapi.dll����"
                Case -14
                    strMSG = "��Ƭ�������"
                Case -15
                    strMSG = "��Ȩ�ļ�������"
                Case -16
                    strMSG = "�豸���Ӵ���"
            End Select
            If TrimStr(IDCardInfor.Name) <> "" Then strMSG = ""  '��ʱ��������ʧ�ܣ�������Ƭ��Ϣ����Ȼ���Զ�ȡ��
        Case IDCardType.COMMON
            'ѡ��
            i = SDT_SelectIDCard(editPort, pucSN, iIfOpen)
            If i <> CByte(&H90) Then
                strMSG = "ѡ��ʧ�ܣ������·ſ�"
                Call SDT_ClosePort(editPort)
            Else
                '����
                intTmp = SDT_ReadBaseMsgToFile(editPort, mstrPath & TXTFile, puiCHMsgLen, mstrPath & WLTFile, puiPHMsgLen, iIfOpen)
                If intTmp = CByte(&H90) Then
                    intTmp = GetBmp(mstrPath & WLTFile, 1)
                    If intTmp <> 1 Then
                        Timer1.Enabled = False
                        Call SDT_ClosePort(editPort)
                        Select Case intTmp
                            Case 0
                                strMSG = "����sdtapi.dll����"
                            Case 1
                                '����
                            Case -1
                                strMSG = "��Ƭ�������"
                            Case -2
                                strMSG = "wlt�ļ���׺����"
                            Case -3
                                strMSG = "wlt�ļ��򿪴���"
                            Case -4
                                strMSG = "wlt�ļ���ʽ����"
                            Case -5
                                strMSG = "���δ��Ȩ��"
                            Case -6
                                strMSG = "�豸���Ӵ���"
                        End Select
                        MsgBox strMSG, vbInformation, App.ProductName
                        Timer1.Enabled = True
                    End If
                Else
                    Timer1.Enabled = False
                    Call SDT_ClosePort(editPort)
                    strMSG = "����ʧ�ܣ�"
                    MsgBox strMSG, vbInformation, App.ProductName
                    Timer1.Enabled = True
                End If
            End If
        Case IDCardType.SS728M01_B01C
            If Not ReadSS728M01 Then strMSG = "��ȡ���֤ʧ�ܣ�"
    End Select
    If strMSG = "" Then
        Call ReadInfoFromFile
        Call DelIDCardFile
    End If

End Sub

Private Sub ReadInfoFromFile()
    Dim strID As String, strName As String, strSex As String
    Dim strNation As String, datBirthday As Date, strAddress As String

    Dim tmp1 As Byte, tmp2 As Byte, intTmp As Integer
    Dim strData As String, strBirthDay As String

    mstrID = "": mstrName = "": mstrSex = "": mstrNation = "":  mdatBirthDay = datBirthday: mstrAddress = ""
    Set mpicPhoto = Nothing

    Select Case glngType
        Case IDCardType.DKQ_116D
            '��ݺ���
            strID = TrimStr(IDCardInfor.IDcardno)
            '����
            strName = TrimStr(IDCardInfor.Name)
            '�Ա�
            strSex = TrimStr(IDCardInfor.sex)
           '����ת��
            Select Case strSex
                Case "1"
                    strSex = "��"
                Case "2"
                    strSex = "Ů"
                Case Else
                    strSex = "δ֪"
            End Select
            '����
            strNation = TrimStr(IDCardInfor.nation)
            strNation = TranNation(Val(strNation))
            '��������
            datBirthday = CDate(Mid(IDCardInfor.born, 1, 4) & "-" & Mid(IDCardInfor.born, 5, 2) & "-" & Mid(IDCardInfor.born, 7, 2)) 'Format(TrimStr(IDCardInfor.born), "yyyy-MM-dd")
            'סַ
            strAddress = TrimStr(IDCardInfor.address)
            Set mpicPhoto = LoadPicture(TrimStr(IDCardInfor.PhotoFileName))
        Case IDCardType.SS728M01_B01C
            '��ݺ���
            strID = SS728M01.ss_id_query_number
            '����
            strName = SS728M01.ss_id_query_name
            '�Ա�
            strSex = SS728M01.ss_id_query_sex
            '����
            strNation = SS728M01.ss_id_query_folk
            '��������
            If Len(SS728M01.ss_id_query_birth) >= 8 Then datBirthday = CDate(Mid(SS728M01.ss_id_query_birth, 1, 4) & "-" & Mid(SS728M01.ss_id_query_birth, 5, 2) & "-" & Mid(SS728M01.ss_id_query_birth, 7, 2))
            'סַ
            strAddress = TrimStr(IIf(SS728M01.ss_id_query_newaddr = "", SS728M01.ss_id_query_address, SS728M01.ss_id_query_newaddr))
            Set mpicPhoto = LoadPicture(TrimStr(SS728M01.ss_id_query_photofile))
        Case Else
            Open IIf(IDCardType.GTICR100_1 = glngType, App.Path & TXTFile, mstrPath & TXTFile) For Binary As #1
                Do While Not EOF(1)   ' ����ļ�β��
                    Get #1, , tmp1
                    Get #1, , tmp2
                    strData = strData & ChrW(tmp2 * CLng(256) + tmp1)
                Loop
            Close #1

            Open IIf(IDCardType.GTICR100_01 = glngType, App.Path & TXTFile, mstrPath & TXTFile) For Binary As #1
                Do While Not EOF(1)   ' ����ļ�β��
                    Get #1, , tmp1
                    Get #1, , tmp2
                    strData = strData & ChrW(tmp2 * CLng(256) + tmp1)
                Loop
            Close #1



            '��ݺ���
            strID = Trim(Mid(strData, 62, 18))
            '����
            strName = Trim(Mid(strData, 1, 15))
            '�Ա�
            strSex = Mid(strData, 16, 1)
            '����
            strNation = Mid(strData, 17, 2)
            '��������
            strBirthDay = Mid(strData, 19, 8)
            'סַ
            strAddress = Trim(Mid(strData, 27, 35))


            '����ת��
            Select Case strSex
                Case "0"
                    strSex = "δ֪"
                Case "1"
                    strSex = "��"
                Case "2"
                    strSex = "Ů"
                Case Else
                    strSex = "δ˵��"
            End Select
            strNation = GetNation(strNation)
            If IsNumeric(strBirthDay) And Len(strBirthDay) = 8 Then
                datBirthday = CDate(Mid(strBirthDay, 1, 4) & "-" & Mid(strBirthDay, 5, 2) & "-" & Mid(strBirthDay, 7, 2))
            End If

            Set mpicPhoto = LoadPicture(IIf(IDCardType.GTICR100_1 = glngType, App.Path & BMPFile, mstrPath & BMPFile))
    End Select

    If mblnAutoRead = False Then
        mstrID = strID: mstrName = strName: mstrSex = strSex: mstrNation = strNation: mdatBirthDay = datBirthday: mstrAddress = strAddress
    Else
        Call mobjIDCard.ShowIDCardInfo(strID, strName, strSex, strNation, datBirthday, strAddress)
    End If
'    Set mpicPhoto = Nothing
End Sub

Public Sub DelIDCardFile()
    If Dir(mstrPath & TXTFile) <> "" Then Call Kill(mstrPath & TXTFile)
    If Dir(mstrPath & BMPFile) <> "" Then Call Kill(mstrPath & BMPFile)
    '����ɾ���ļ�������Ҫ��� GTICR100
    If Dir(App.Path & TXTFile) <> "" Then Call Kill(App.Path & TXTFile)
    If Dir(App.Path & BMPFile) <> "" Then Call Kill(App.Path & BMPFile)
    '���������
    If Dir(TrimStr(IDCardInfor.PhotoFileName)) <> "" And TrimStr(IDCardInfor.PhotoFileName) <> "" Then Call Kill(TrimStr(IDCardInfor.PhotoFileName))
End Sub

'���������
Public Function GetNation(ByVal strNationcode As String) As String
    Dim strNationArray As Variant

    strNationArray = Array("��", "�ɹ�", "��", "��", "ά���", "��", "��", "׳", "����", "����", _
                        "��", "��", "��", "��", "����", "����", "������", "��", "��", "����", _
                        "��", "�", "��ɽ", "����", "ˮ", "����", "����", "����", "�¶�����", "��", _
                        "���Ӷ�", "����", "Ǽ", "����", "����", "ë��", "����", "����", "����", "����", _
                        "������", "ŭ", "���α��", "����˹", "���¿�", "�°�", "����", "ԣ��", "��", "������", _
                        "����", "���״�", "����", "�Ű�", "���", "��ŵ")

    If Trim(strNationcode) <> "" Then
        If ((CByte(Trim(strNationcode)) - 1) >= 0) And ((CByte(Trim(strNationcode)) - 1) <= 55) Then
            GetNation = strNationArray(CByte(Trim(strNationcode)) - 1)
            '90373:���ϴ���2015/11/6,��������ȫ��
            GetNation = GetNation & "��"
        Else
            GetNation = "����"
        End If
    End If
End Function

Public Sub Read_Card(strID As String, strName As String, strSex As String, _
                             strNation As String, datBirthday As Date, strAddress As String)
    mblnAutoRead = False            '���ô˷���ʱΪ�ֶ�����
    mstrID = "": mstrName = "": mstrSex = "": mstrNation = "": mdatBirthDay = datBirthday: mstrAddress = ""
    Set mpicPhoto = Nothing
    Dim i As Integer
    Select Case glngType
        Case IDCardType.CVR100U, IDCardType.CVR100D, IDCardType.SS_V1, IDCardType.XZX_KDQ, IDCardType.GTICR100, IDCardType.GTICR100_01, IDCardType.GTICR100_1
            If Authenticate = 1 Then Call ReadIDCard
        Case IDCardType.CVR100U_1, IDCardType.CVR100D_1
            If CVR_Authenticate = 1 Then Call ReadIDCard
        Case IDCardType.HX_FDX9
            If SDT_StartFindIDCard(1, "", 1) = CByte(&H9F) Then Call ReadIDCard
        Case IDCardType.DKQ_116D
            lngReturn = Syn_ClosePort(1001)
            lngReturn = Syn_OpenPort(1001)
            If lngReturn = 0 Then Call ReadIDCard
        Case IDCardType.CVR100
            lngReturn = SDT_StartFindIDCard(1001, "", 1)
            If lngReturn = CByte(&H9F) Then Call ReadIDCard
        Case IDCardType.COMMON
            '�ҿ�
            i = SDT_StartFindIDCard(editPort, pucIIN, iIfOpen)
            If i <> CByte(&H9F) Then
                '���ҿ�
                i = SDT_StartFindIDCard(editPort, pucIIN, iIfOpen)
                If i <> CByte(&H9F) Then
                    Call SDT_ClosePort(editPort)
                Else
                    Call ReadIDCard
                End If
            End If
    End Select

    strID = mstrID: strName = mstrName: strSex = mstrSex: strNation = mstrNation: datBirthday = mdatBirthDay: strAddress = mstrAddress

    mblnAutoRead = True             '�ָ�Ϊ�Զ�����
End Sub

Public Function ReadPhotoInfo() As StdPicture
    Set ReadPhotoInfo = mpicPhoto
End Function

Public Function TranNation(ByVal lngNo As Long) As String
    Dim strNation As String
    Select Case lngNo
    Case 1
        strNation = "����"
    Case 2
        strNation = "�ɹ���"
    Case 3
        strNation = "����"
    Case 4
        strNation = "����"
    Case 5
        strNation = "ά�����"
    Case 6
        strNation = "����"
    Case 7
        strNation = "����"
    Case 8
        strNation = "׳��"
    Case 9
        strNation = "������"
    Case 10
        strNation = "������"
    Case 11
        strNation = "����"
    Case 12
        strNation = "����"
    Case 13
        strNation = "����"
    Case 15
        strNation = "������"
    Case 16
        strNation = "������"
    Case 17
        strNation = "��������"
    Case 18
        strNation = "����"
    Case 19
        strNation = "����"
    Case 20
        strNation = "������"
    Case 21
        strNation = "����"
    Case 22
        strNation = "���"
    Case 23
        strNation = "��ɽ��"
    Case 24
        strNation = "������"
    Case 25
        strNation = "ˮ��"
    Case 26
        strNation = "������"
    Case 27
        strNation = "������"
    Case 28
        strNation = "������"
    Case 29
        strNation = "�¶�������"
    Case 30
        strNation = "����"
    Case 31
        strNation = "���Ӷ���"
    Case 32
        strNation = "������"
    Case 33
        strNation = "Ǽ��"
    Case 34
        strNation = "������"
    Case 35
        strNation = "������"
    Case 36
        strNation = "ë����"
    Case 37
        strNation = "������"
    Case 38
        strNation = "������"
    Case 39
        strNation = "������"
    Case 40
        strNation = "������"
    Case 41
        strNation = "��������"
    Case 42
        strNation = "ŭ��"
    Case 43
        strNation = "���α����"
    Case 44
        strNation = "����˹��"
    Case 45
        strNation = "���¿���"
    Case 46
        strNation = "������"
    Case 47
        strNation = "������"
    Case 48
        strNation = "ԣ����"
    Case 49
        strNation = "����"
    Case 50
        strNation = "��������"
    Case 51
        strNation = "������"
    Case 52
        strNation = "���״���"
    Case 53
        strNation = "������"
    Case 54
        strNation = "�Ű���"
    Case 55
        strNation = "�����"
    Case 56
        strNation = "��ŵ��"
    Case Else
        strNation = "����"
    End Select
    TranNation = strNation
End Function



