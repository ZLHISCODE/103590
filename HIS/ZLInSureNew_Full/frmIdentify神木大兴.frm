VERSION 5.00
Begin VB.Form frmIdentify��ľ���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���������֤"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmd�鿨 
      Caption         =   "���»�ȡ(&R)"
      Height          =   350
      Left            =   105
      TabIndex        =   11
      Top             =   3105
      Width           =   1305
   End
   Begin VB.Frame fra 
      Height          =   2745
      Index           =   0
      Left            =   105
      TabIndex        =   10
      Top             =   210
      Width           =   5715
      Begin VB.Frame Frame1 
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   1
         Left            =   345
         TabIndex        =   4
         Top             =   1620
         Width           =   5115
         Begin VB.OptionButton Opt�Ա� 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   840
            TabIndex        =   5
            Top             =   270
            Width           =   885
         End
         Begin VB.OptionButton Opt�Ա� 
            Caption         =   "Ů"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   1890
            TabIndex        =   6
            Top             =   270
            Width           =   885
         End
         Begin VB.OptionButton Opt�Ա� 
            Caption         =   "δ֪"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   2790
            TabIndex        =   7
            Top             =   270
            Value           =   -1  'True
            Width           =   1935
         End
      End
      Begin VB.TextBox txtEdit 
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   1545
         MaxLength       =   20
         TabIndex        =   3
         Top             =   967
         Width           =   3945
      End
      Begin VB.TextBox txtEdit 
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   1545
         MaxLength       =   16
         TabIndex        =   1
         Top             =   420
         Width           =   3945
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   210
         TabIndex        =   2
         Top             =   1005
         Width           =   1260
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ҽ������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   180
         TabIndex        =   0
         Top             =   495
         Width           =   1260
      End
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4605
      TabIndex        =   9
      Top             =   3105
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3450
      TabIndex        =   8
      Top             =   3105
      Width           =   1100
   End
End
Attribute VB_Name = "frmIdentify��ľ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytType As Byte            '0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����,99-�޸�ָ�����˵Ŀ���,88-��סԺ��Ϣ���в�ѯ

Private mlng����ID As Long
Private mstrReturn As String
Private mintPreCol As Integer, mintsort As Integer
Private mblnFirst As Boolean        '��һ����ϵͳʱ����
Private mblnChange As Boolean
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
   
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long


Private Const INFINITE = -1&
Private Const SYNCHRONIZE = &H100000
Private Sub cmd�鿨_Click()
    Call Read������Ϣ
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    cmdȷ��.Enabled = False
    If mbytType = 99 Then
        If ��ȡ�α���Ա��Ϣ = False Then Unload Me: Exit Sub
    Else
        Call Read������Ϣ
    End If
    Call SetCtlEn
End Sub

Private Sub SetOKCtrl(ByVal blnEn As Boolean)
    cmdȷ��.Enabled = blnEn
End Sub

Private Function IsValid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��֤���ݵĺϷ���
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim lng״̬ As Long
    
    IsValid = False
    If Trim(g�������_��ľ����.IC����) = "" Then
        MsgBox "δ����ҽ�����ţ�", vbInformation, gstrSysName
        txtEdit(0).SetFocus
        Exit Function
    End If
    
    If zlCommFun.ActualLen(g�������_��ľ����.IC����) > 16 Then
        MsgBox "ҽ�����ų���,���������16���ַ���8�����֣�", vbInformation, gstrSysName
        txtEdit(0).SetFocus
        Exit Function
    End If
    
    If Trim(g�������_��ľ����.����) = "" Then
        MsgBox "δ���벡��������", vbInformation, gstrSysName
        txtEdit(1).SetFocus
        Exit Function
    End If
    
    If zlCommFun.ActualLen(g�������_��ľ����.����) > 20 Then
        MsgBox "��������,���������20���ַ���10�����֣���", vbInformation, gstrSysName
        txtEdit(1).SetFocus
        Exit Function
    End If
    If mbytType = 99 Then
        IsValid = True
        Exit Function
    End If
      
    If mbytType <> 2 And mbytType <> 88 Then
        If mbytType = 4 Then
            '����鵱ǰ״̬
        Else
            '��鲡��״̬
            gstrSQL = "select nvl(��ǰ״̬,0) as ״̬ from �����ʻ� where ����=[1] and ҽ����=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_��������, g�������_��ľ����.IC����)
            If rsTemp.RecordCount > 0 Then
                If rsTemp("״̬") > 0 Then
                    MsgBox "�ò����Ѿ���Ժ������ͨ�������֤��", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    Else
        '�����������סԺ�ģ�ֻ��ˢ����ʾһ�����ݶ��ѣ�������
         '��ȷ����ǰ״̬,��Ϊ��ǰ״̬�ǲ��ܸı��
        gstrSQL = "Select * from �����ʻ� where ����=" & TYPE_�������� & " and  ҽ����='" & g�������_��ľ����.IC���� & "'"
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            mlng����ID = Nvl(rsTemp!����ID, 0)
        Else
            mlng����ID = 0
        End If
        mstrReturn = mlng����ID
        Unload Me
        Exit Function
    End If
    IsValid = True
End Function
Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    Dim lng����ID As Long
    Dim strTmp As String
    Dim strIdentify As String, strAddition As String
    Dim rsTemp As New ADODB.Recordset
    Dim int��ǰ״̬ As Integer
    Dim i As Byte
    
    With g�������_��ľ����
        .IC���� = Trim(txtEdit(0).Text)
        .���� = Trim(txtEdit(1).Text)
        strTmp = ""
        For i = 0 To 2
            If Opt�Ա�(i).Value = True Then
                strTmp = Decode(i, 0, "��", 1, "Ů", "δ֪")
                Exit For
            End If
        Next
        .�Ա� = strTmp
    End With
    
    
    If IsValid = False Then Exit Sub
    
    If mbytType = 99 Then       '���¿���
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_�������� & ",'����','''" & g�������_��ľ����.IC���� & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����")
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_�������� & ",'ҽ����','''" & g�������_��ľ����.IC���� & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ҽ����")
        
        If gcnOracle_��ľ���� Is Nothing Then
           Open�м��_��ľ����
        ElseIf gcnOracle_��ľ����.State <> 1 Then
           Open�м��_��ľ����
        End If
        gstrSQL = "ZL_סԺ_UPDATE("
        gstrSQL = gstrSQL & "'" & txtEdit(0).Tag & "',"
        gstrSQL = gstrSQL & "'" & g�������_��ľ����.IC���� & "')"
        
        ExecuteProcedure_��ľ���� "�ı俨��"
        mstrReturn = g�������_��ľ����.IC����
        Unload Me
        Exit Sub
    End If
    int��ǰ״̬ = 0
    
    If mbytType = 4 Then
        '��ȷ����ǰ״̬,��Ϊ��ǰ״̬�ǲ��ܸı��
        gstrSQL = "Select * from �����ʻ� where ����=" & TYPE_�������� & " and  ҽ����='" & g�������_��ľ����.IC���� & "'"
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            mlng����ID = Nvl(rsTemp!����ID, 0)
            int��ǰ״̬ = Nvl(rsTemp!��ǰ״̬, 0)
        End If
        rsTemp.Close
    End If
    
    
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    With g�������_��ľ����
        strIdentify = .IC����                        '0����
        strIdentify = strIdentify & ";" & .IC����    '1ҽ����
        strIdentify = strIdentify & ";"              '2����
        strIdentify = strIdentify & ";" & .����      '3����
        strIdentify = strIdentify & ";" & .�Ա�      '4�Ա�
        strIdentify = strIdentify & ";"              '5��������
        strIdentify = strIdentify & ";"              '6���֤
        strIdentify = strIdentify & ";"              '7.��λ����(����)
        strAddition = ";0"                           '8.���Ĵ���
        strAddition = strAddition & ";"              '9.˳���
        strAddition = strAddition & ";"              '10��Ա���
        strAddition = strAddition & ";"              '11�ʻ����

        strAddition = strAddition & ";"              '12��ǰ״̬
        strAddition = strAddition & ";"              '13����ID
        strAddition = strAddition & ";1"             '14��ְ(1,2,3)
        strAddition = strAddition & ";"              '15����֤��
        strAddition = strAddition & ";"              '16�����
        strAddition = strAddition & ";"              '17�Ҷȼ�
        strAddition = strAddition & ";"              '18�ʻ������ۼ�
        strAddition = strAddition & ";0"             '19�ʻ�֧���ۼ�
        strAddition = strAddition & ";0"             '20���깤���ܶ�
        strAddition = strAddition & ";"              '21סԺ�����ۼ�
    End With
    
    mlng����ID = BuildPatiInfo(0, strIdentify & strAddition, mlng����ID, TYPE_��������)
    
    If mlng����ID = 0 Then
        ShowMsgbox "������Ϣ����!"
        Exit Sub
    End If
    '���ظ�ʽ:�м���벡��ID
    If mlng����ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng����ID & strAddition
    End If
    g�������_��ľ����.����ID = mlng����ID
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Public Function GetPatient(Optional bytType As Byte, Optional lng����ID As Long = 0) As String
    mbytType = bytType
    mlng����ID = lng����ID
    mstrReturn = ""
    
    Me.Show 1
    lng����ID = mlng����ID
    GetPatient = mstrReturn
End Function
Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Function ��ȡ�α���Ա��Ϣ() As Boolean
    '��ȡ�α���Ա��Ϣ
    Dim rsTemp As New ADODB.Recordset
    Dim strTmp As String
    
    ��ȡ�α���Ա��Ϣ = False
    Err = 0:    On Error GoTo errHand:
    If mbytType = 99 Then
        gstrSQL = "Select a.*,b.����,b.ҽ���� From ������Ϣ a,�����ʻ� b where a.����id=b.����id and a.����id =" & mlng����ID
    Else
        gstrSQL = "Select * From ������Ϣ where ����id in (Select ����id From �����ʻ� where  ҽ����='" & txtEdit(0).Text & "')"
    End If
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    If rsTemp.EOF Then
        'txtEdit(1).Text = ""
        If mbytType = 99 Then
            ShowMsgbox "����ָ����ҽ������"
        End If
        Exit Function
    End If
    If mbytType = 99 Then
        txtEdit(0).Text = Nvl(rsTemp!ҽ����)
        txtEdit(0).Tag = Nvl(rsTemp!ҽ����)
        
    End If
    txtEdit(1).Text = Nvl(rsTemp!����)
    Opt�Ա�(Decode(Nvl(rsTemp!�Ա�), "��", 0, "Ů", 1, 2)).Value = True
    
    ��ȡ�α���Ա��Ϣ = True
    Exit Function
errHand:
        If ErrCenter = 1 Then
            Resume
        End If
End Function
Private Sub ClearData()
    Dim i As Long
    '��������Ϣ
    With g�������_��ľ����
        .IC���� = ""
        .���� = ""
        .�Ա� = ""
    End With
End Sub

Private Sub Opt�Ա�_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If txtEdit(0).Text = "" Then
        SetOKCtrl False
    Else
        SetOKCtrl True
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Index = 0 And mbytType <> 99 Then
            Call ��ȡ�α���Ա��Ϣ
        End If
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txtEdit(Index), KeyAscii, m�ı�ʽ)
End Sub
Private Sub SetCtlEn()
    If mbytType = 99 Then
        txtEdit(1).BackColor = &H8000000F
        txtEdit(1).Enabled = False
        Frame1(1).Enabled = False
        Me.Caption = "�޸�ҽ������"
    End If
End Sub

Private Sub Read������Ϣ()
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '--��  ��:ͨ�������ļ���ȡ������Ϣ
    '--�����:
    '--������:
    '--��  ��:�ִ�
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    Dim strText As String
    Dim strFile As String
    Dim STRNAME As String
    Dim strҽ��֤�� As String
    Dim int��־ As Integer
    Dim str�Ա� As String
    Dim STR���� As String
    

        
    strFile = Replace(UCase(Replace(InitInfor_��ľ����.����Ŀ¼, "\\", "\")), UCase("ReadYbInfo.INI"), UCase("ReadYbInfo.exe"))
    
    If Not Dir(strFile) <> "" Then
        ShowMsgbox "��Ŀ¼��" & strFile & "���ļ������ڣ����ڲ�������������!"
        Exit Sub
    End If
    Call ExecuteReadCardExe(strFile)
    
    strFile = Replace(InitInfor_��ľ����.����Ŀ¼, "\\", "\")
    If Not Dir(strFile) <> "" Then
        ShowMsgbox "��Ŀ¼��" & strFile & "���ļ������ڣ����ڲ�������������!"
        Exit Sub
    End If
    
    Dim strArr
    
    Err = 0
    On Error GoTo errHand:
    If Dir(strFile) <> "" Then
            Set objText = objFile.OpenTextFile(strFile)
            Do While Not objText.AtEndOfStream
                strText = Trim(objText.ReadLine)
                If UCase("[ReadYlbxIcInfo]") <> UCase(strText) Then
                    If InStr(1, strText, "=") <> 0 Then
                        strArr = Split(strText, "=")
                        Select Case UCase(strArr(0))
                        Case UCase("pcode")
                            strҽ��֤�� = Trim(strArr(1))
                        Case UCase("ycbz")
                            int��־ = Val(strArr(1))
                        Case UCase("xb")
                            str�Ա� = Trim(strArr(1))
                        Case UCase("xm")
                            STR���� = Trim(strArr(1))
                        End Select
                    End If
                End If
            Loop
            objText.Close
    End If
    If int��־ <> 0 Then
        ShowMsgbox "��ҽ��������Ч,����!"
        Exit Sub
    End If
    txtEdit(0) = strҽ��֤��
    txtEdit(1) = STR����
    Select Case str�Ա�
    Case "��", "1"
        Opt�Ա�(0).Value = True
    Case "Ů", "2"
        Opt�Ա�(1).Value = True
    Case Else
        Opt�Ա�(2).Value = True
    End Select
    
    Exit Sub
errHand:
    DebugTool Err.Description
    Exit Sub
End Sub

Private Function ExecuteReadCardExe(ByVal strFile As String) As Boolean
    'ִ��Exe�ļ�
    Dim lngTask As Long, lngRet As Long, lngpHandle As Long
    lngTask = Shell(strFile, vbHide)
    lngpHandle = OpenProcess(SYNCHRONIZE, False, lngTask)
    lngRet = WaitForSingleObject(lngpHandle, INFINITE)
    lngRet = CloseHandle(lngpHandle)
    ExecuteReadCardExe = True
End Function

