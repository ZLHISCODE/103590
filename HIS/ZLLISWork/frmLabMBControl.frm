VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLabMBControl 
   BorderStyle     =   0  'None
   Caption         =   "��������"
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraMain 
      Height          =   1605
      Left            =   30
      TabIndex        =   0
      Top             =   -75
      Width           =   6390
      Begin VB.CommandButton cmdCanCel 
         Cancel          =   -1  'True
         Caption         =   "ֹͣ(&S)"
         Height          =   350
         Left            =   2670
         TabIndex        =   3
         Top             =   1080
         Width           =   1100
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   405
         Left            =   150
         TabIndex        =   1
         Top             =   585
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   714
         _Version        =   327682
         Appearance      =   1
      End
      Begin MSCommLib.MSComm MSComm 
         Left            =   5625
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin VB.Label lbl��Ϣ 
         AutoSize        =   -1  'True
         Caption         =   "׼����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   2
         Top             =   225
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmLabMBControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const conCmd_Begin = "Begin"            '�ӽ���������ȡ����������������
Const ConCmd_Out = "Out"                '��������
Const ConCmd_In = "In"                 '��������
Const conCmd_End = "End"
Const conCmd_Revert = "Revert"
Const conCmd_Play = "Play"
Const conCmd_ReadData = "ReadData" '����
Const conCmd_SpecialConnType = "SpecialConnType"

Dim mstrCmdRevert As String              ' ����ͨ��Ӧ������
Dim mobjDevice As Object                '�ӿ�
Public bln_Init As Boolean                 '�����Ƿ��ѳ�ʼ��

Public Function MB_Start(objfrm As Object, ByVal strMachineID As Long) As Boolean
    Dim rsTmp As New adodb.Recordset
    Dim strͨѶ�� As String, str������ As String, str����λ As String, strֹͣλ As String, strУ��λ As String

    On Error GoTo ErrHandle
    
   
    strͨѶ�� = zlDatabase.GetPara("frmLabMB_ͨѶ��", 100, 1208, "COM1")
    str������ = zlDatabase.GetPara("frmLabMB_������", 100, 1208, "9600")
    str����λ = zlDatabase.GetPara("frmLabMB_����λ", 100, 1208, "8")
    strֹͣλ = zlDatabase.GetPara("frmLabMB_ֹͣλ", 100, 1208, "1")
    strУ��λ = zlDatabase.GetPara("frmLabMB_У��λ", 100, 1208, "N")
    
    
    strͨѶ�� = Replace(strͨѶ��, "COM", "")
    strУ��λ = Replace(Replace(Replace(Replace(Replace(strУ��λ, "E-ż��", "E"), "M-���", "M"), "N-ȱʡ", "N"), "O-����", "O"), "S-�ո�", "S")
    strУ��λ = Replace(strУ��λ, "None", "N")
    
    gstrSql = "select ͨѶ������ from �������� where id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strMachineID)
    
    If rsTmp.EOF = True Then MsgBox "û���ҵ�����!", vbInformation, Me.Caption: Exit Function
    
    If Nvl(rsTmp("ͨѶ������")) = "" Then MsgBox "����ͨѶ������Ϊ��,�뵽���������������޸�!", vbInformation: Exit Function
    
    If Not mobjDevice Is Nothing Then Set mobjDevice = Nothing
    Set mobjDevice = CreateObject(rsTmp("ͨѶ������"))
    If mobjDevice Is Nothing Then
        MsgBox "������������ʧ�ܣ�", vbInformation, Me.Caption
        Exit Function
    End If
    
    
    '���Ϳ���ָ��
    If MSComm.PortOpen Then MSComm.PortOpen = False
    MSComm.CommPort = CInt(strͨѶ��)
    MSComm.Settings = str������ & "," & strУ��λ & "," & str����λ & "," & strֹͣλ    '"9600,N,8,1"
    MSComm.InputLen = 0
    MSComm.PortOpen = True
    
    Me.Show , objfrm
    If bln_Init Then
        Call MB_SendCommand(conCmd_End, "�ͷ���������......", 3)
    End If
    
    '=========================================��ʼ����======================================
    If Not MB_SendCommand(conCmd_Begin, "��������...", 2) Then Exit Function
    '========================================================================================
    
    '==================================OUT ��ѡ����=========================================
    If Not MB_SendCommand(ConCmd_Out, "���ڵ���΢�װ�...", 2) Then Exit Function
    
    '========================================================================================
    bln_Init = True
    MB_Start = bln_Init
    Me.Hide
    Exit Function
ErrHandle:
    MsgBox "��������ʱ�����ִ���" & vbNewLine & "[" & Err.Number & "] " & Err.Description, vbInformation, Me.Caption
End Function

Public Sub MB_Stop()
    '����,�Ͽ�����
    Dim strCmd As String
    On Error GoTo errH
    Me.Show
    If bln_Init Then
        Call MB_SendCommand(conCmd_End, "�ͷ���������......", 3)
        If MSComm.PortOpen Then MSComm.PortOpen = False
        bln_Init = False
    End If
    Me.Hide
    Exit Sub
errH:
    MsgBox "�ͷ���������ʱ���ִ���" & vbNewLine & "[" & Err.Number & "] " & Err.Description, vbInformation, Me.Caption
End Sub

Public Sub ShowMe(objfrm As Object, ByVal strControl As String, strResult As String)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����               Objfrm ������
    '                   strMachineID    ����ID
    '                   strControl (1:2:3:4:5:6) (����;���Ƶ��;���ʱ��;���巽ʽ;�հ���ʽ:�ο�����)
    '                   ���ؽ��
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim aItem() As String
    Dim intLoop As Integer             'ѭ��ʱʹ��
    Dim aRow() As String, aResult() As String
    Dim intRow As Integer, intCol As Integer
    Dim TestData(1, 1 To 8, 1 To 12) As String
    
    On Error GoTo errH
    If Not bln_Init Then
        MsgBox "��������������", vbInformation, Me.Caption
        Exit Sub
    End If
    Me.Show , objfrm
    
    '==================================��ʼ׼������==================================================
    Call MB_SendCommand(ConCmd_In, "���ڹر�΢�װ�......", 2)
    
    aItem = Split(strControl, ";")
    For intLoop = 0 To UBound(aItem) - 1
        Select Case intLoop
            Case 0
                If aItem(intLoop) <> "" Then
                    If Not MB_SendCommand(aItem(intLoop), "�������ò���...") Then Exit Sub
                End If
            Case 1
                If aItem(intLoop) <> "" Then
                    If Not MB_SendCommand(aItem(intLoop), "�����������Ƶ��...") Then Exit Sub
                End If
            Case 2
                If aItem(intLoop) <> "" Then
                    If Not MB_SendCommand(aItem(intLoop), "�����������ʱ��...") Then Exit Sub
                End If
            Case 3
                If aItem(intLoop) <> "" Then
                    If Not MB_SendCommand(aItem(intLoop), "�������ý��巽ʽ...") Then Exit Sub
                End If
            Case 4
                If aItem(intLoop) <> "" Then
                    If Not MB_SendCommand(aItem(intLoop), "�������ÿհ���ʽ...") Then Exit Sub
                End If
        End Select
    Next
    '========================================��ʼ����(������)=====================================================
    If Not MB_SendCommand(conCmd_Play, "��ʼ����(������)...") Then Exit Sub
    If Not MB_SendCommand(conCmd_ReadData, "��ʼ��ȡ����(������)...", 15, 1, strResult) Then Exit Sub
    
    '========================================��ʼ����(�ο�����)===================================================
    If aItem(5) <> "" Then
        '--����������
        aRow = Split(strResult, "|")
        For intRow = 1 To 8
            aResult = Split(aRow(intRow - 1), ";")
            For intCol = 1 To 12
                TestData(0, intRow, intCol) = aResult(intCol - 1)
            Next
        Next
        
        '���òο�����
        If Not MB_SendCommand(aItem(5), "�������òο�����...") Then Exit Sub
        
        '��ʼ�����ο�����
        strResult = ""
        If Not MB_SendCommand(conCmd_Play, "��ʼ����(�ο�����)...") Then Exit Sub
        
        If Not MB_SendCommand(conCmd_ReadData, "��ʼ��ȡ����(�ο�����)...", 15, 1, strResult) Then Exit Sub
        
        '����ο�����
        aRow = Split(strResult, "|")
        For intRow = 1 To 8
            aResult = Split(aRow(intRow - 1), ";")
            For intCol = 1 To 12
                TestData(1, intRow, intCol) = aResult(intCol - 1)
            Next
        Next

        '����
        strResult = ""
        For intRow = 1 To 8
            strResult = strResult & "|"
            For intCol = 1 To 12
                strResult = strResult & ";" & TestData(0, intRow, intCol) - TestData(1, intRow, intCol)
            Next
        Next
        strResult = Replace(strResult, "|;", "|")
        strResult = Mid(strResult, 2)
    End If
    
    '=======================================����΢�װ壬��ѡ����=======================================
    If Not MB_SendCommand(ConCmd_Out, "���ڵ���΢�װ�...", 2) Then Exit Sub
    
    Me.Hide
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function MB_SendCommand(ByVal str_Command As String, ByVal str_Info As String, Optional ByVal intOutTime As Integer = 10, Optional ByVal intType As Integer = 0, Optional ByRef str_Data As String) As Boolean
    '�������ø����
    
    'str_Command :Ҫ��������������,���ڲ�����ģ�ÿ��������һ����
    'str_Info   ����ʾ��Ϣ
    'intOutTime ��ִ�д�����ʱ����ʱ���ã�Ĭ��10
    'intType   : 0-���������� 1-��������
    'str_Data  : ���� ���ص�����
    Dim str_R As String                         'ÿ���������������Ӧ��ָ�
    Dim strCmdRevert As String                  'ͨ�õ�Ӧ��ָ��
    Dim strCmd As String                        '��������
    Dim lngBeginDate As Long, strGetCmd As String
    
    If mobjDevice Is Nothing Then Exit Function '��������δ��ʼ�����˳�
    If Not MSComm.PortOpen Then Exit Function '����δ�򿪣��˳�
    
    Dim strReserved As String  '���ý���ʱ�ã�ռλ����
    Dim blnErr As Boolean, int_I As Integer, var_R As Variant
    Dim int��ʱ As Integer
    Dim strSpecialConnType As String
    Dim strgetval As String
    
    On Error GoTo ErrHandle
    
    If mstrCmdRevert = "" Then mstrCmdRevert = mobjDevice.CmdAnalyse(conCmd_Revert)
    
    strCmd = mobjDevice.CmdAnalyse(str_Command)
    int��ʱ = Val(mobjDevice.CmdAnalyse(str_Command & "_TimeOut"))
    If int��ʱ <= 0 Then
        int��ʱ = intOutTime
    End If
    
    '��ȡ����ʱ����Ҫ��ͣ�ķ���ָ��(���������Ż��õ�)=1ʱ��Ч
    strSpecialConnType = mobjDevice.CmdAnalyse(conCmd_SpecialConnType)
    
    '--- ��־
    MbLog "frmLabMBControl", "MB_SendCommand", strCmd, int��ʱ
    
    lngBeginDate = Timer
    If Trim(strCmd) <> "" Then
        
        lbl��Ϣ.Caption = str_Info
        
        If InStr(strCmd, "|") > 0 Then '��ר�ŵ�Ӧ��ָ��
            str_R = Mid(strCmd, InStr(strCmd, "|") + 1)
            strCmd = Mid(strCmd, 1, InStr(strCmd, "|") - 1)
        Else
            str_R = mstrCmdRevert      'ͨ�õ�Ӧ��ָ��
        End If
        MSComm.Output = strCmd
        strGetCmd = ""
        If intType = 0 Then '����������
            Do
                DoEvents
                strGetCmd = strGetCmd & MSComm.Input
                
                Call ShowPbar((CLng(Timer) - lngBeginDate) / int��ʱ * 100)
            Loop Until InStr(strGetCmd, str_R) Or (CLng(Timer) - lngBeginDate > int��ʱ)
                            '--- ��־
            MbLog "frmLabMBControl", "����Ӧ��ָ��", strGetCmd, str_R
            
            If Trim(strGetCmd) = "" Then
                '��ʱ����
                Debug.Print Timer & " " & lngBeginDate
                MsgBox "ִ��" & str_Command & "���ʱ!", vbInformation, Me.Caption
                Exit Function
            Else
                If InStr(str_R, "|") > 0 Then
                    var_R = Split(str_R, "|")
                    blnErr = True
                    For int_I = LBound(var_R) To UBound(var_R)
                        If InStr(strGetCmd, var_R(int_I)) >= 0 Then
                            blnErr = False
                            Exit For
                        End If
                    Next
                Else
                    blnErr = InStr(strGetCmd, str_R) <= 0
                End If
                If blnErr Then
                    MsgBox "ִ��" & str_Command & "����������ص���������!" & vbNewLine & strGetCmd, vbInformation, Me.Caption
                    Exit Function
                End If
            End If
        Else                'Ҫ���ؽ�������
            Do
               DoEvents
               '���������������ͣ�ķ�����ָ��
                If strSpecialConnType = "1" Then
                    MSComm.Output = strCmd
                    Call Sleep(1000)
                End If
               strGetCmd = strGetCmd & MSComm.Input
               Call ShowPbar((CLng(Timer) - lngBeginDate) / int��ʱ * 100)
               mobjDevice.Analyse strGetCmd, str_Data, strReserved, ""
               strgetval = strGetCmd
               strGetCmd = strReserved
            Loop Until str_Data <> "" Or (CLng(Timer) - lngBeginDate > int��ʱ)
            
            '--- ��־
            MbLog "frmLabMBControl", "����ø������", strgetval & "|" & strGetCmd, str_Data

            If Trim(str_Data) = "" And Trim(strGetCmd) = "" Then
                MsgBox "��������ʧ��!", vbInformation, Me.Caption
                Exit Function
            ElseIf Trim(strGetCmd) <> "" And Trim(str_Data) = "" Then
                MsgBox "��������ʧ��!", vbInformation, Me.Caption
                Exit Function
            End If
        End If
        

    End If
    MB_SendCommand = True
    Exit Function
ErrHandle:
    'If MSComm.PortOpen Then MSComm.PortOpen = False
    MsgBox "ִ��" & str_Command & "������ִ���!" & vbNewLine & "[" & Err.Number & "] " & Err.Description
End Function

Private Sub CmdCancel_Click()
    If MsgBox("���ر������������ӣ�����δ��������ݽ����ܽ��գ���ȷ�ϣ�" & vbNewLine & "��[ȷ��]�����ر������������ӣ���[ȡ��]������ԭ���Ĳ�����", vbQuestion + vbOKCancel + vbDefaultButton2) = vbOK Then
        Unload Me
    End If
End Sub

Private Sub ShowPbar(ByVal sinValue As Single)
    On Error Resume Next
    ProgressBar1.Value = sinValue
End Sub


Private Sub MbLog(ByVal strModule As String, ByVal strFunc As String, ByVal strInput As String, ByVal strOutput As String)
    '���ù���������¼��־
    Call zl9Comlib.LogWrite("LIS�ϰ�ͨѶ���������־", strModule, strFunc, strInput & vbCrLf & strOutput)
End Sub

