VERSION 5.00
Begin VB.Form frmIdentify��ɽ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ���������ʶ��"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   Icon            =   "frmIdentify��ɽ.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1170
      Width           =   3765
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   15
      TabIndex        =   4
      Top             =   1605
      Width           =   6150
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   405
      Left            =   4530
      TabIndex        =   3
      Top             =   1890
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   405
      Left            =   2850
      TabIndex        =   2
      Top             =   1890
      Width           =   1305
   End
   Begin VB.TextBox txtPwd 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   645
      Width           =   3765
   End
   Begin VB.Label Label1 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1485
      TabIndex        =   8
      Top             =   1185
      Width           =   660
   End
   Begin VB.Label lblPatiInfo 
      AutoSize        =   -1  'True
      Caption         =   "������Ϣ"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   390
      TabIndex        =   7
      Top             =   945
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   450
      Picture         =   "frmIdentify��ɽ.frx":000C
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblPwd 
      AutoSize        =   -1  'True
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1485
      TabIndex        =   6
      Top             =   705
      Width           =   510
   End
   Begin VB.Label lblNote 
      Caption         =   "���ڲ���IC��֮������������롣"
      Height          =   255
      Left            =   1095
      TabIndex        =   5
      Top             =   180
      Width           =   3645
   End
End
Attribute VB_Name = "frmIdentify��ɽ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�������������ͨ����
Private Declare Function IC_InitComm Lib "DCIC32.DLL" (ByVal Port%) As Long
Private Declare Function IC_ExitComm% Lib "DCIC32.DLL" (ByVal icdev As Long)
Private Declare Function IC_Down% Lib "DCIC32.DLL" (ByVal icdev As Long)
Private Declare Function IC_Pushout% Lib "DCIC32.DLL" (ByVal icdev As Long)
Private Declare Function IC_InitType% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal TypeNo%)
Private Declare Function IC_Status% Lib "DCIC32.DLL" (ByVal icdev As Long)
Private Declare Function IC_Erase% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal offset%, ByVal l%)
Private Declare Function IC_Read% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal offset%, ByVal l%, ByVal Databuffer$)
Private Declare Function IC_Read_Hex% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal offset%, ByVal l%, ByVal Databuffer$)
Private Declare Function IC_Read_Float% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal offset%, fdata As Single)
Private Declare Function IC_Read_Int% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal offset%, fdata As Long)
Private Declare Function IC_Write% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal offset%, ByVal Length%, ByVal Databuffer$)
Private Declare Function IC_Write_Hex% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal offset%, ByVal Length%, ByVal Databuffer$)
Private Declare Function IC_Write_Float% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal offset%, ByVal fdata As Single)
Private Declare Function IC_Write_Int% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal offset%, ByVal fdata As Long)
'ר�Ŵ���4428���ĺ���
Private Declare Function IC_ReadWithProtection% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal offset%, ByVal l%, ByVal ProtBuffer$)
Private Declare Function IC_WriteWithProtection% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal offset%, ByVal l%)
Private Declare Function IC_ReadCount_SLE4428% Lib "DCIC32.DLL" (ByVal icdev As Long)
Private Declare Function IC_CheckPass_SLE4428% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal Password$)
Private Declare Function IC_ChangePass_SLE4428% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal Password$)
Private Declare Function IC_CheckPass_4428hex% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal Password$)
Private Declare Function IC_ChangePass_4428hex% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal Password$)
 


Public mstr��֤�� As String
Public mstrPatiInfo As String
Public mcur��� As Currency
Public mstr���ⲡ
Private mcur������� As Currency        '������ǭ��ҽ��

Private mstrҽ���� As String
Private mstr���� As String
Private mstrRead As String * 25       '���������������ʾ����һ�£�������Ҫ��������
Private blnLoad As Boolean       '�Ƿ��ܹ��ɹ����в���
Private mintst As Long
Private mlngIcdev As Long '��ǰ���ڽ���ͨѶ�Ĵ����豸
Private mstr��� As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
'����߼����ܴ������⣬��Ҫ��ʱ���е���
    Dim strPassWord As String * 20
    Dim intst As Integer, strSQL As String
    Dim strTime As String, rs��ɽ As New ADODB.Recordset
    Dim strTmp As String, lngErrLine As Long
    
    On Error GoTo errHandle
    intst = IC_Status(mlngIcdev): lngErrLine = 1 '��ȡ��������״̬
    If intst < 0 Then
        MsgBox "��������ʼ��ʧ��,���鴮��", vbInformation, gstrSysName
        mstrPatiInfo = ""
        Exit Sub
    End If
    If intst = 0 Then
        lblPatiInfo.Caption = "���ڶ��������Ժ�...."
    End If
    If intst = 1 Then
        MsgBox "��⵽������֮��ȱ��������", vbInformation, gstrSysName
        Exit Sub
    End If
    '�Կ�������صĲ���
    intst = IC_InitType(mlngIcdev, 4): lngErrLine = 2
    If intst <> 0 Then
        MsgBox "IC���ĳ�ʼ��ʧ��,����", vbInformation, gstrSysName
        mstrPatiInfo = ""
        Exit Sub
    End If
    DoEvents
    
    intst = IC_ReadCount_SLE4428(ByVal mlngIcdev&): lngErrLine = 3
    If intst < 0 Then
        MsgBox "�ڶ����Ĺ���֮�з�������", vbInformation, gstrSysName
        mstrPatiInfo = ""
        Exit Sub
    End If
    DoEvents
    '�Կ����������У��
    '�ύʱҪ�ĳ�B518
    strPassWord = mstr��֤��: lngErrLine = 4 ' "B518" '������������ڽ��з�����ʱ��ȷ����������Ҫ�޸�
    intst = IC_CheckPass_4428hex(ByVal mlngIcdev&, ByVal strPassWord$): lngErrLine = 5
    If intst < 0 Then
        MsgBox "����֤�뷢�������뵽���Ļ���", vbInformation, gstrSysName
        mstrPatiInfo = ""
        Exit Sub
    End If
    DoEvents
    '������ǰʹ�õ�IC���Ŀ���
    mstrRead = String(25, " "): lngErrLine = 6
    intst = IC_Read(ByVal mlngIcdev, 80, 25, ByVal mstrRead$): lngErrLine = 7
    If intst <> 0 Then
        MsgBox "����������ʧ�ܣ�������[IC����]", vbInformation, gstrSysName
        mstrPatiInfo = ""
        Exit Sub
    End If
    DoEvents
    'Modified By ���� ���� 06:06:46
    '��ȡIC�������
    If Val(Get���ղ���_��ɽ("���õ���")) = 1 Then   'ǭ������ʹ�õĻ�����ӿ��ڶ�ȡ�ʻ����
        mstr��� = String(6, "0"): lngErrLine = 8
        intst = IC_Read(ByVal mlngIcdev, 105, 6, ByVal mstr���$): lngErrLine = 9
        If intst <> 0 Then
            MsgBox "����������ʧ�ܣ�������[IC�����]", vbInformation, gstrSysName
            mstrPatiInfo = ""
            Exit Sub
        End If
        If IsNumeric(mstr���) Then
            mcur������� = Val(mstr���) / 100: lngErrLine = 10       '���ֱ��棬��Ҫת��
        Else
            mcur������� = 0
        End If
    End If
    DoEvents
    mintst = IC_Down(ByVal mlngIcdev): lngErrLine = 11 '���������µ�
    If mintst < 0 Then
        lblPatiInfo.Caption = "ע�⣺IC���µ�ʧ��"
    End If
    
    
    '�����ݿ�֮�л�ȡ�ֿ����˵���֤��Ϣ
    strTime = CStr(Format(zlDatabase.Currentdate, "yyyymmddhhmmss")) & "00": lngErrLine = 12
'    mstrRead = "1234510226200304250856132"
    strSQL = "insert into Check_doex_interface(Bill_no,App_code," & _
            "Ic_id) values('" & strTime & "','" & Mid(gstrҽԺ����, 1, 4) & _
            "','" & txtPwd.Text & mstrRead & "')": lngErrLine = 13
    gcn��ɽ.Execute strSQL: lngErrLine = 14
    '�������֤��������
    
    strSQL = "insert into Check_bill_request(Bill_no,App_code," & _
            "Request_status) values('" & strTime & "','" & _
            Mid(gstrҽԺ����, 1, 4) & "','4')": lngErrLine = 15
    gcn��ɽ.Execute strSQL: lngErrLine = 16
    
    On Error Resume Next
    If Checkrequest(strTime) = False Then Exit Sub
    DoEvents
    If Requestinfo(strTime) <> "" Then
        DoEvents
        Me.Hide
    End If
    
    On Error GoTo errHandle
    'ɾ����ص�����
    
    strSQL = "delete from Check_bill_request where Bill_no = '" & _
            strTime & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'": lngErrLine = 17
    gcn��ɽ.Execute strSQL: lngErrLine = 18
    strSQL = "delete from Check_doex_interface where Bill_no = '" & _
             strTime & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'": lngErrLine = 19
    gcn��ɽ.Execute strSQL: lngErrLine = 20
    
    'Modified by ZYB 20040921
    '---------------------------------------------------------------------
    '�������ȡ�����ʻ����
    If Val(Get���ղ���_��ɽ("���õ���")) <> 1 Then
        strSQL = "insert into Check_doex_interface(Bill_no,App_code," & _
                "Ic_id) values('" & strTime & "','" & Mid(gstrҽԺ����, 1, 4) & _
                "','" & txtPwd.Text & mstrRead & "')": lngErrLine = 21
        gcn��ɽ.Execute strSQL: lngErrLine = 22
        strSQL = "insert into Check_bill_request(Bill_no,App_code," & _
                "Request_status) values('" & strTime & "','" & _
                Mid(gstrҽԺ����, 1, 4) & "','2')": lngErrLine = 23
        gcn��ɽ.Execute strSQL: lngErrLine = 24
        
        On Error Resume Next
        If Checkrequest(strTime) = False Then Exit Sub
        DoEvents
        strSQL = "select Ps_Bala " & _
                " from Check_Doex_Interface where Bill_no = '" & strTime & "'" & _
                " and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'": lngErrLine = 1
        If rs��ɽ.State = adStateOpen Then rs��ɽ.Close: lngErrLine = 25
        rs��ɽ.Open strSQL, gcn��ɽ
        If rs��ɽ.RecordCount <> 0 Then mcur��� = Nvl(rs��ɽ!Ps_Bala, 0)
        
        On Error GoTo errHandle
        'ɾ����ص�����
        strSQL = "delete from Check_bill_request where Bill_no = '" & _
                strTime & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'": lngErrLine = 26
        gcn��ɽ.Execute strSQL: lngErrLine = 27
        strSQL = "delete from Check_doex_interface where Bill_no = '" & _
                 strTime & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'": lngErrLine = 28
        gcn��ɽ.Execute strSQL: lngErrLine = 29
    End If
    '---------------------------------------------------------------------
    
    
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Or Val(Get���ղ���_��ɽ("���õ���")) = 1 Then
        mstr���ⲡ = Combo1.Text: lngErrLine = 21
        If Combo1.ListIndex = 0 Then
            gbln�������� = False: lngErrLine = 22
        Else
            gbln�������� = True: lngErrLine = 23
        End If
    Else
        mstr���ⲡ = "": lngErrLine = 24
        gbln�������� = False: lngErrLine = 25
    End If
    
    '����Ƿ���Ժ
    Dim rsTmp As New ADODB.Recordset
    gstrSQL = "select nvl(��ǰ״̬,0) as ״̬ from �����ʻ� where ����=[1] and ҽ����=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_�����ɽ, mstrҽ����)
    If rsTmp.RecordCount > 0 Then
        If rsTmp("״̬") > 0 Then
            MsgBox "�ò����Ѿ���Ժ������ͨ�������֤��", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    Exit Sub
errHandle:
    MsgBox "��[�����֤]���壬[cmdOK_Click]�¼��У���" & lngErrLine & "�з�������", vbExclamation, gstrSysName
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOK_Click
End Sub

Private Sub Form_Activate()
    If blnLoad = False Then Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then cmdOK_Click
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rs��ɽ As New ADODB.Recordset
    On Error GoTo errHandle
    
    'Modified by ZYB 20040921
    '---------------------------------------------------------------------
    '��ȡ���ⲡ��
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then
        strSQL = "Select esp_ID||'--'||esp_name ���� From check_esp_interface Order by esp_id"
        If rs��ɽ.State = 1 Then rs��ɽ.Close
        rs��ɽ.CursorLocation = adUseClient
        rs��ɽ.Open strSQL, gcn��ɽ
    End If
    
    With Combo1
        .AddItem "��ͨ��"
        If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then
            Do While Not rs��ɽ.EOF
                .AddItem rs��ɽ!����
                rs��ɽ.MoveNext
            Loop
        Else
            .AddItem "���ⲡ"
        End If
'        .AddItem "01--��֢�������ڵķ��ơ����ơ���ʹ����"
'        .AddItem "02--������˥�߲���͸������"
'        .AddItem "03--������ֲ��Ŀ���������"
'        .AddItem "04--����۲첡�ˣ�3���ڣ�����������"
'        .AddItem "05--80���������˵������ͼ�ͥ������180���ڣ�"
'        .AddItem "06--���򲡡�����Ǵ�"
'        .AddItem "07--���Ը�Ѫѹ�����Ĳ������Ĳ��������к���֢"
'        .AddItem "08--����������֧���������������ס����Ĳ�"
'        .AddItem "09--�����黯������ժ����"
'        .AddItem "10--���Ը�Ӳ��"
'        .AddItem "11--���������ϰ���ƶѪ"
'        .AddItem "12--����"
'        .AddItem "13--��˲�"
    End With
    '---------------------------------------------------------------------
    
    Combo1.ListIndex = 0
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then
        Combo1.Visible = True
    Else
        Combo1.Visible = False
    End If
    '�Դ��ڽ��г�ʼ��
    mstrPatiInfo = ""
'    mintst = IC_ExitComm(mlngIcdev)  '�رմ���
    mlngIcdev = IC_InitComm(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��ǰʹ�õĴ���", "0"))    '��ʼ������ COM1
    If mlngIcdev <= 0 Then
        blnLoad = False
        MsgBox "���ڳ�ʼ��ʧ��,���鴮��", vbInformation, gstrSysName
        Exit Sub
    End If
    mintst = IC_Status(mlngIcdev) '��ȡ������״̬
    If mintst < 0 Then
        blnLoad = False
        MsgBox "���ڳ�ʼ���ɹ������Ƕ�������ʼ��ʧ��", vbInformation, gstrSysName
    Else
        blnLoad = True
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    mintst = IC_ExitComm(mlngIcdev)  'Close COM
    blnLoad = False
End Sub

Private Function Requestinfo(Billno As String) As String
'���ܣ������ݿ���в�ѯ�������ݣ��Ӷ��õ���Ҫ����Ϣ
    Dim strSQL As String, rs��ɽ As New ADODB.Recordset, lngErrLine As Long
    
    On Error GoTo errHandle
    'New:0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����)
    strSQL = "select Ps_Code,Ps_Name,Ps_Sex,Ps_Bdate,Ps_Sfzh,Ep_id,Ps_Bala " & _
            " from Check_Doex_Interface where Bill_no = '" & Billno & "'" & _
            " and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'": lngErrLine = 1
    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close: lngErrLine = 2
    mstrPatiInfo = "": lngErrLine = 3
    rs��ɽ.Open strSQL, gcn��ɽ, adOpenStatic, adLockReadOnly: lngErrLine = 4
    '���쵱ǰ��Ҫ����ʹ�õ�����
    If Not rs��ɽ.BOF Then
        mstrҽ���� = Nvl(rs��ɽ("Ps_Code"), "")
        mstrPatiInfo = mstrRead & ";" & Nvl(rs��ɽ("Ps_Code"), "") & ";" & _
                        txtPwd.Text & ";" & Nvl(rs��ɽ("Ps_Name"), "") & _
                        ";" & Nvl(rs��ɽ("Ps_Sex"), "") & ";" & _
                        CStr(Nvl(rs��ɽ("Ps_Bdate"), "")) & ";" & _
                        Nvl(rs��ɽ("Ps_Sfzh"), "") & ";" & Nvl(rs��ɽ("Ep_id"), ""): lngErrLine = 5
        mcur��� = IIf(IsNull(rs��ɽ("Ps_Bala")), 0, rs��ɽ("Ps_Bala")): lngErrLine = 6
        
        'Modified By ���� ���� 06:07:01
        If Val(Get���ղ���_��ɽ("���õ���")) = 1 Then   'ǭ������ʹ�õĻ�����ӿ��ڶ�ȡ�ʻ����
            mcur��� = mcur�������: lngErrLine = 7
        End If
    End If
    Requestinfo = mstrPatiInfo
    Exit Function
errHandle:
    MsgBox "��[�����֤]���壬[RequestInfo]�¼��е�" & lngErrLine & "�з�������", vbExclamation, "����"
    If ErrCenter() = 1 Then Resume
End Function

Private Sub txtPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Combo1.Visible = True Then
            Combo1.SetFocus
        Else
            cmdOK_Click
        End If
    End If
End Sub
