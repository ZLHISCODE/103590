VERSION 5.00
Begin VB.Form frmIdentify�ɶ����� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ���������ʶ��"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIdentify�ɶ�����.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Visible         =   0   'False
   Begin VB.CheckBox chk������־ 
      Caption         =   "������־"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1590
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4350
      Width           =   2985
   End
   Begin VB.TextBox txtEdit 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   6
      Left            =   1950
      Locked          =   -1  'True
      MaxLength       =   4
      PasswordChar    =   "*"
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2010
      Width           =   2595
   End
   Begin VB.CheckBox chk�ֹ� 
      Caption         =   "�ֹ����뿨������(&M)"
      Height          =   240
      Left            =   1470
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   660
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.ComboBox cbo������� 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   360
      Left            =   1965
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2490
      Width           =   2595
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "У��(&T)"
      Height          =   405
      Left            =   240
      TabIndex        =   20
      Top             =   4920
      Width           =   1305
   End
   Begin VB.TextBox txtEdit 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   5
      Left            =   1965
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   17
      Top             =   3870
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.TextBox txtEdit 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   1965
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   15
      Top             =   3420
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.TextBox txtEdit 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   1965
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   13
      Top             =   2970
      Width           =   2595
   End
   Begin VB.TextBox txtEdit 
      BackColor       =   &H8000000F&
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1965
      Locked          =   -1  'True
      MaxLength       =   4
      PasswordChar    =   "*"
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1530
      Width           =   2595
   End
   Begin VB.TextBox txtEdit 
      BackColor       =   &H8000000F&
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1965
      Locked          =   -1  'True
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1065
      Width           =   2595
   End
   Begin VB.TextBox txtEdit 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1965
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   615
      Width           =   2595
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   405
      Left            =   1965
      TabIndex        =   21
      Top             =   4920
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   405
      Left            =   3450
      TabIndex        =   22
      Top             =   4920
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   75
      Left            =   -210
      TabIndex        =   19
      Top             =   4665
      Width           =   6660
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "ҽ���չ���Ա"
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
      Index           =   7
      Left            =   330
      TabIndex        =   8
      Top             =   2070
      Width           =   1530
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "�������"
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
      Index           =   6
      Left            =   840
      TabIndex        =   10
      Top             =   2550
      Width           =   1020
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "ȷ������"
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
      Index           =   5
      Left            =   840
      TabIndex        =   16
      Top             =   3930
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "������"
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
      Index           =   4
      Left            =   1095
      TabIndex        =   14
      Top             =   3480
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lblEdit 
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
      Index           =   3
      Left            =   1350
      TabIndex        =   12
      Top             =   3030
      Width           =   510
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "�����ı���"
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
      Index           =   2
      Left            =   585
      TabIndex        =   6
      Top             =   1590
      Width           =   1275
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "���˱���"
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
      Index           =   1
      Left            =   840
      TabIndex        =   4
      Top             =   1125
      Width           =   1020
   End
   Begin VB.Label lblNote 
      Caption         =   "������ȷˢ��֮������������롣"
      Height          =   255
      Left            =   930
      TabIndex        =   0
      Top             =   225
      Width           =   3645
   End
   Begin VB.Label lblEdit 
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
      Index           =   0
      Left            =   1350
      TabIndex        =   2
      Top             =   675
      Width           =   510
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   255
      Picture         =   "frmIdentify�ɶ�����.frx":030A
      Top             =   345
      Width           =   480
   End
End
Attribute VB_Name = "frmIdentify�ɶ�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'<0��ʾ������;0������ȷ�Ҳ忨;1-������ȷ��δ�忨
Private Declare Function IC_Status_1 Lib "JKIC32.DLL" Alias "IC_Status" (ByVal lngICDev As Long) As Integer
'<0����;���ڱ�ʶ��
Private Declare Function IC_InitComm_1 Lib "JKIC32.DLL" Alias "IC_InitComm" (ByVal IntPort As Integer) As Long
'<0δ�������رմ���;>=0�����ر�
Private Declare Function IC_ExitComm_1 Lib "JKIC32.DLL" Alias "IC_ExitComm" (ByVal lngICDev As Long) As Integer
'<0����;=0����
Private Declare Function IC_InitType_1 Lib "JKIC32.DLL" Alias "IC_InitType" (ByVal lngICDev As Long, ByVal intType As Integer) As Integer
'<0����;=0����
Private Declare Function IC_Read_1 Lib "JKIC32.DLL" Alias "IC_Read_Hex" (ByVal lngICDev As Long, ByVal intOffset As Integer, ByVal intLen As Integer, ByVal strData As String) As Integer

'------------------------------------------------------------
Private Declare Function IC_InitComm_2 Lib "ftic_32.dll" Alias "ic_init" (ByVal Port%, ByVal baud As Long) As Long
Private Declare Function IC_Status_2 Lib "ftic_32.dll" Alias "get_status" (ByVal icdev As Long, intCard As Integer) As Integer
Private Declare Function chk_card Lib "ftic_32.dll" (ByVal icdev As Long) As Integer
Private Declare Function IC_Read_2 Lib "ftic_32.dll" Alias "srd_4442" (ByVal icdev As Long, ByVal offset As Long, ByVal Length As Long, ByVal r_string As String) As Integer
Private Declare Function IC_Down_2 Lib "ftic_32.dll" Alias "auto_pull" (ByVal icdev As Long) As Integer
Private Declare Function ic_exit% Lib "ftic_32.dll" (ByVal icdev As Long)
Private Declare Function hex_asc% Lib "ftic_32.dll" (ByVal hex As String, ByVal asc$, ByVal le&)

Private Declare Function srd_4442 Lib "ftic_32.dll" (ByVal icdev As Long, ByVal offset As Long, ByVal Length As Long, ByRef r_string As Byte) As Integer
'------------------------------------------------------------

'˵����������ɵĹ��ܼ������ƣ������������ɹ���ҽ���������֤�⡣������˳ɶ�����ҽ���������֤
Private mstr���� As String
Private mstrҽ���� As String
Private mstr�����ı�� As String
Private mstr���� As String
Private mstr������� As Integer
Private mintInsure As Integer
Private mblnPass As Boolean
Private mblnChangePassword As Boolean
Private mbln������־ As Boolean

Private mint������ As Integer
Private mint�˿ں� As Integer
Private mint���� As Integer

Private mblnOK As Boolean

Private Sub chk�ֹ�_Click()
    Dim lngColor As Long, blnEnable As Boolean
    
    If mintInsure <> TYPE_������ Then Exit Sub
    cmdOK.Enabled = False
    If chk�ֹ�.Value = 1 Then
        '���ֹ�������:�����ı��,�������
        blnEnable = True
        lngColor = &H80000005
        
        txtEdit(2).TabStop = True
    Else
        '�ر��ֹ�������
        blnEnable = False
        lngColor = &H8000000F
        
        txtEdit(2).Text = ""
        cbo�������.ListIndex = 0
    End If
    
    txtEdit(2).Locked = Not blnEnable
    txtEdit(2).BackColor = lngColor
    cbo�������.Enabled = blnEnable
    cbo�������.BackColor = lngColor
    txtEdit(0).SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngIndex As Long
    
    If cmdOK.Enabled = False Then Exit Sub
    For lngIndex = txtEdit.LBound To txtEdit.UBound
        If txtEdit(lngIndex).Visible = True Then
            If zlCommFun.StrIsValid(Trim(txtEdit(lngIndex).Text), IIf(lngIndex = 0, 20, txtEdit(lngIndex).MaxLength)) = False Then
                If txtEdit(lngIndex).Enabled Then txtEdit(lngIndex).SetFocus
                Exit Sub
            End If
        End If
    Next
    If mintInsure = TYPE_������ Then
        If Trim(txtEdit(0).Text) = "" Then
            MsgBox "δ��ȷ��ˢ��,����ͨ����֤��", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    If Trim(txtEdit(1).Text) = "" Then
        MsgBox "δ��ȷ��ˢ��,����ͨ����֤��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mstr���� = Trim(txtEdit(0).Text)
    mstrҽ���� = Trim(txtEdit(1).Text)
    mstr�����ı�� = Trim(txtEdit(2).Text)
    mstr���� = Trim(txtEdit(3).Text)
    mstr������� = cbo�������.ListIndex + 1
    mbln������־ = (chk������־.Value = 1)
    
    If mblnChangePassword = True Then
        '��ǰ�����Ǹ�������
        If mintInsure = TYPE_������ Then
            If txtEdit(4).Text <> "" Or txtEdit(5).Text <> "" Then
                If txtEdit(4).Text <> txtEdit(5).Text Then
                    MsgBox "��������������벻��ͬ�����������롣", vbInformation, gstrSysName
                    If txtEdit(5).Enabled Then txtEdit(5).SetFocus
                    Exit Sub
                End If
            End If
        Else
            If txtEdit(4).Text = "" Then
                MsgBox "�������µ����롣", vbInformation, gstrSysName
                If txtEdit(4).Enabled Then txtEdit(4).SetFocus
                Exit Sub
            End If
            If txtEdit(4).Text <> txtEdit(5).Text Then
                MsgBox "��������������벻��ͬ�����������롣", vbInformation, gstrSysName
                If txtEdit(5).Enabled Then txtEdit(5).SetFocus
                Exit Sub
            End If
        End If
        
        If Trim(txtEdit(4).Text) <> "" Then
            If mintInsure = type_�ɶ����� Then
                If ��������_�ɶ�����(mstr����, mstrҽ����, mstr�����ı��, mstr����, txtEdit(4).Text) = False Then Exit Sub
            ElseIf mintInsure = TYPE_�¶� Then
                If ��������_�¶�(mstr����, mstrҽ����, mstr�����ı��, mstr����, txtEdit(4).Text) = False Then Exit Sub
            ElseIf mintInsure = TYPE_������ Then
                If ��������_������(txtEdit(0).Tag, mstr����, txtEdit(4).Text) = False Then Exit Sub
            End If
            mstr���� = Trim(txtEdit(4).Text)
        End If
    End If
    mblnOK = True
    Unload Me
End Sub

Public Function GetIdentify(ByVal intinsure As Integer, str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String, _
                            Optional ByVal blnPass As Boolean = True, Optional ByVal blnChangePassword As Boolean = False, Optional ByRef bln������־ As Boolean = False) As Boolean
    Dim sinDec As Single
    Dim intAdjust As Integer
    Dim rsTemp As New ADODB.Recordset
    
    mblnOK = False
    mblnPass = blnPass
    mblnChangePassword = blnChangePassword
    mintInsure = intinsure
    
    If intinsure = type_�ɶ����� Or intinsure = TYPE_�¶� Then
        '��������ʾ���ŵ���Ϣ
        txtEdit(0).PasswordChar = "*"
        txtEdit(1).PasswordChar = "*"
        txtEdit(2).PasswordChar = "*"
        
        '���ι���ҽ�����ֿؼ�
        chk�ֹ�.Visible = False
        lblEdit(6).Visible = False
        cbo�������.Visible = False
        lblEdit(7).Visible = False
        txtEdit(6).Visible = False
        cmdTest.Visible = False
        cmdOK.Enabled = True
        '����λ��
        sinDec = txtEdit(0).Top - lblEdit(0).Top
        For intAdjust = 0 To 5
            If intAdjust = 0 Then
                txtEdit(intAdjust).Top = chk�ֹ�.Top
                lblEdit(intAdjust).Top = txtEdit(intAdjust).Top - sinDec
            Else
                txtEdit(intAdjust).Top = txtEdit(intAdjust - 1).Top + 510
                lblEdit(intAdjust).Top = txtEdit(intAdjust).Top - sinDec
            End If
        Next
        Frame1.Top = txtEdit(5).Top + txtEdit(5).Height + 180
        cmdOK.Top = Frame1.Top + 200
        cmdCancel.Top = cmdOK.Top
        cmdTest.Top = cmdOK.Top
        
        '������(2005-12-28):��ȡ���õ���
        gstrSQL = "Select ������,Nvl(����ֵ,0) Value From ���ղ��� Where ������='���õ���' and ����=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ϴ���Ժ��Ϣ����ֵ", intinsure)
        mint���� = rsTemp("value")
        
        'ȡ���ղ���
        mint������ = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "������", 0)
        mint�˿ں� = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "IC�豸�˿�", 1)
    ElseIf intinsure = TYPE_������ Then
        cmdOK.Enabled = False
        
        With cbo�������
            .Clear
            .AddItem "��ҵְ������ҽ�Ʊ���"
            .AddItem "��ҵ����ҽ�Ʊ���"
            .AddItem "������ҵ��λҽ�Ʊ���"
            .ListIndex = 0
        End With
        chk�ֹ�.Value = 0
    End If
    
    lblEdit(4).Visible = blnChangePassword
    lblEdit(5).Visible = blnChangePassword
    txtEdit(4).Visible = blnChangePassword
    txtEdit(5).Visible = blnChangePassword
    
    If blnChangePassword = False Then
        Frame1.Top = txtEdit(4).Top
        cmdOK.Top = Frame1.Top + 200
        cmdCancel.Top = cmdOK.Top
        cmdTest.Top = cmdOK.Top
    End If
    
    frmIdentify�ɶ�����.Height = cmdOK.Top + cmdOK.Height + 500
    frmIdentify�ɶ�����.Show vbModal
    
    GetIdentify = mblnOK
    If mblnOK = True Then
        If mintInsure = TYPE_������ Then
            str���� = mstr���� & "^" & mstr�������
        Else
            str���� = mstr����
        End If
        strҽ���� = mstrҽ����
        str�����ı�� = mstr�����ı��
        str���� = mstr����
    End If
    
    bln������־ = mbln������־
End Function

Private Sub cmdTest_Click()
    Dim str��� As String
    Dim str�Ա� As String
    '�����Ŵ����ӿڷ���,ȡ�䷵�ؽ�������½���
    
    If Trim(txtEdit(0).Text) = "" Then
        MsgBox "��ˢ����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If InitXML = False Then Exit Sub
    If chk�ֹ�.Value = 0 Then
        '�������޸�����
        If mblnChangePassword = True Then
            '��ǰ�����Ǹ�������
            If txtEdit(4).Text <> "" Or txtEdit(5).Text <> "" Then
                If txtEdit(4).Text <> txtEdit(5).Text Then
                    MsgBox "��������������벻��ͬ�����������롣", vbInformation, gstrSysName
                    If txtEdit(5).Enabled Then txtEdit(5).SetFocus
                    Exit Sub
                End If
            End If
            
            If Trim(txtEdit(4).Text) <> "" Then
                If ��������_������(txtEdit(0).Text, txtEdit(3).Text, txtEdit(4).Text) = False Then Exit Sub
                txtEdit(3).Text = txtEdit(4).Text
                txtEdit(4).Text = ""
                txtEdit(5).Text = ""
                mstr���� = Trim(txtEdit(3).Text)
            End If
        End If
    
        If InitXML = False Then Exit Sub
        Call InsertChild(mdomInput.documentElement, "CARDDATA", txtEdit(0).Text)            ' �ſ�����
        Call InsertChild(mdomInput.documentElement, "PASSWORD", txtEdit(3).Text)            ' ����
    Else
        Call InsertChild(mdomInput.documentElement, "CARDID", txtEdit(0).Text)              ' �ſ�����
        Call InsertChild(mdomInput.documentElement, "CENTERCODE", txtEdit(2).Text)          ' �����ı���
        Call InsertChild(mdomInput.documentElement, "INSURETYPE", cbo�������.ListIndex + 1) ' �������
        Call InsertChild(mdomInput.documentElement, "PASSWORD", txtEdit(3).Text)            ' ����
    End If
    
    '���ýӿ�
    If CommServer(IIf(chk�ֹ�.Value = 0, "READCARD", "READCARD_M")) = False Then Exit Sub
    
    'ȡ�÷���ֵ
    txtEdit(0).Tag = txtEdit(0).Text                    '���濨�����ݣ��Ա��������ʱʹ��
    txtEdit(0).Text = GetElemnetValue("CARDID")
    txtEdit(1).Text = GetElemnetValue("PERSONCODE")
    txtEdit(2).Text = GetElemnetValue("CENTERCODE")
    txtEdit(6).Text = IIf(Val(GetElemnetValue("CAREPSNFLAG")) = 0, "��", "��")
    str�Ա� = GetElemnetValue("SEX")
    str�Ա� = Switch(str�Ա� = "1", "��", str�Ա� = "2", "Ů", str�Ա� = "9", "����", True, str�Ա�)
    If str�Ա� = "Ů" Then chk������־.Enabled = True
    cbo�������.ListIndex = GetElemnetValue("INSURETYPE") - 1
    cmdOK.Enabled = True
End Sub

Private Sub Form_Load()
  gblnLED = Val(GetSetting("ZLSOFT", "����ȫ��", "ʹ��", 0)) <> 0
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If mintInsure = TYPE_������ Then
        If chk�ֹ�.Value = 0 Then
            If Index = 0 Then cmdOK.Enabled = False
        Else
            If Index = 0 Or Index = 2 Then cmdOK.Enabled = False
        End If
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    '��ˢ��
    If Index = 0 Then
        If gblnLED And txtEdit(0).Text = "" Then
            zl9LedVoice.Speak "#5"
        End If
    End If
    '����������
    If Index = 3 Then
        If gblnLED And txtEdit(3).Text = "" Then
            zl9LedVoice.Speak "#0"
        End If
    End If
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim varSplit As Variant

    
    If KeyAscii = vbKeyReturn Then
        If Index = 0 Then
            If mintInsure = TYPE_������ Then
                '���Լ���ˢ�����ݽ��н���(����ܽ�ֹ,�ɶԷ��ӿڷ���)
'                txtEdit(0).Text = Replace(txtEdit(0).Text, vbCr, "")
'                txtEdit(0).Text = Replace(txtEdit(0).Text, vbLf, "")
'                If Right(txtEdit(0).Text, 1) = "?" Then
'                    '���Զ�ˢ����Ϣ���зֽ���
'                    '�ſ����ݸ�ʽΪ��;:����=���˱���=�����ı���?
'                    varSplit = Split(txtEdit(0).Text, "=")
'                    txtEdit(0).Text = Mid(varSplit(0), 3)
'                    If UBound(varSplit) > 0 Then txtEdit(1).Text = varSplit(1)
'                    If UBound(varSplit) > 1 Then txtEdit(2).Text = Mid(varSplit(2), 1, Len(varSplit(2)) - 1)
'
'                    If mblnPass = True Then
'                        txtEdit(3).SetFocus
'                    Else
'                        cmdOK_Click
'                    End If
'                End If
                If Trim(txtEdit(0).Text) <> "" Then
                    If mblnPass = True Then
                        txtEdit(3).SetFocus
                    Else
                        Call cmdTest_Click
                    End If
                End If
            Else
                '�ɶ����أ�����ϵͳ�ṩ�ĺ������н���
                Dim lngDev As Long, lngBoud As Long, lngReturn As Long, intReturn As Integer, intCard As Integer
                Dim str���ı��  As String, strҽ���� As String, str���� As String
                Dim strҽ����_IC As String * 256
                Dim str����_IC As String * 256
                Dim strData As String * 256
                
                '������(2006-01-18):���ӱ���
                Dim by����_IC(256) As Byte
                Dim intnum As Integer
                Dim StrHex As String, strTemp As String, str����_IC2 As String, strҽ����_IC2 As String
                
                '�����IC��,��Ҫ�Ƚ��������ڵ�����
                If mint������ <> 0 Then
                    '��ʼ���˿ں�
                    lngBoud = 9600
                    Call DebugTool("׼���򿪶˿�")
                    If mint������ = 1 Then
                        lngDev = IC_InitComm_1(mint�˿ں� - 1)
                        If lngDev < 0 Then
                            Call ShowErr(lngDev)
                            Exit Sub
                        End If
                    Else
                        lngDev = IC_InitComm_2(mint�˿ں� - 1, lngBoud)
                        If lngDev > 0 Then
                        ElseIf lngDev = -149 Then
                            MsgBox "��ǰ�˿��ѱ���������ռ�ã�", vbInformation, gstrSysName
                            Exit Sub
                        Else
                            MsgBox "��ʼ���˿�ʧ�ܣ�", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                    
                    '�ж��Ƿ��Ѳ忨
                    Call DebugTool("�ж��Ƿ��Ѳ忨")
                    If mint������ = 1 Then
                        intReturn = IC_Status_1(lngDev)
                        Select Case intReturn
                        Case Is < 0
                            Call ICErr(intReturn, lngDev)
                            Exit Sub
                        Case 1
                            MsgBox "��忨��", vbInformation, gstrSysName
                            Call CloseCommon(lngDev)
                            Exit Sub
                        End Select
                    Else
                        '��ȡ����״̬
                        intReturn = IC_Status_2(lngDev, intCard)
                        If intReturn < 0 Then
                           MsgBox "���豸״̬����", vbInformation, gstrSysName
                           Exit Sub
                        Else
                           If intCard = 0 Then
                              MsgBox "��忨", vbInformation, gstrSysName
                              Call CloseCommon(lngDev)
                              Exit Sub
                           Else
                              intReturn = chk_card(lngDev)
                              If intReturn < 0 Then
                                MsgBox "��⿨ʧ��", vbInformation, gstrSysName
                                Call CloseCommon(lngDev)
                                Exit Sub
                              Else
                                Select Case intReturn
                                   Case 0
                                      MsgBox "δ֪����", vbInformation, gstrSysName
                                      Call CloseCommon(lngDev)
                                      Exit Sub
                                   Case 21
'                                      MsgBox "�����뿨ΪSLE4432", vbInformation, gstrSysName
'                                      Exit Sub
                                   Case Else
                                      MsgBox "�����뿨��ΪSLE4432", vbInformation, gstrSysName
                                      Call CloseCommon(lngDev)
                                      Exit Sub
                                End Select
                              End If
                           End If
                        End If
                    End If
                    
                    '���ÿ�����
                    Call DebugTool("��ʼ��������Ϊ������4432/4442")
                    If mint������ = 1 Then
                        intReturn = IC_InitType_1(lngDev, 16)
                        If intReturn < 0 Then
                            Call ICErr(intReturn, lngDev)
                            Exit Sub
                        End If
                    End If
                    
                    '��ȡ����
                    Call DebugTool("��ȡ���ţ���59λ��ʼ������ȡ6λ")
                    If mint������ = 1 Then
                        intReturn = IC_Read_1(lngDev, 29, 3, str����_IC)
                        If intReturn < 0 Then
                            Call ICErr(intReturn, lngDev)
                            Exit Sub
                        End If
                        Call DebugTool("��ȡҽ���ţ���17λ��ʼ������ȡ17λ")
                        intReturn = IC_Read_1(lngDev, 8, 9, strҽ����_IC)
                        If intReturn < 0 Then
                            Call ICErr(intReturn, lngDev)
                            Exit Sub
                        End If
                    
                        str���� = TruncZero(str����_IC)
                        strҽ���� = TruncZero(strҽ����_IC)
                        strҽ���� = Mid(strҽ����, 1, 17)
                    Else
'                        intReturn = IC_Read_2(lngDev, 0, 64, str����_IC)
'                        Call hex_asc%(str����_IC, strData, 32)
                     '������(2006-01-18):����hex_asc��������ʱ���������,ʹ���ֹ�����
                     intReturn = srd_4442(lngDev, 8, 32, by����_IC(0))
                     For intnum = 0 To 9
                         If Len(CStr(hex(by����_IC(intnum)))) = 1 Then
                             StrHex = "0" & CStr(hex(by����_IC(intnum)))
                         Else
                             StrHex = CStr(hex(by����_IC(intnum)))
                         End If

                         strTemp = strTemp & Trim(StrHex)
                     Next
                     For intnum = 21 To 23
                         If Len(CStr(hex(by����_IC(intnum)))) = 1 Then
                             StrHex = "0" & CStr(hex(by����_IC(intnum)))
                         Else
                             StrHex = CStr(hex(by����_IC(intnum)))
                         End If
                         str����_IC2 = str����_IC2 & Trim(StrHex)
                     Next
                     strҽ����_IC2 = Mid(strTemp, 1, 16) & Mid(strTemp, 17, 1) & Mid(strTemp, 20, 1)
                        
'                        strҽ����_IC = Mid(strData, 13, 17)
'                        str����_IC = Mid(strData, 55, 6)
                        str���� = Replace(str����_IC2, " ", "")
                        strҽ���� = Replace(strҽ����_IC2, " ", "")
                    End If
                    
                    '�رն˿ں�
                    Call DebugTool("�رն˿ںţ��Ա��´�ʹ��")
                    Call CloseCommon(lngDev)
'                    str���ı�� = "22"
'                    '������(2005-12-28):ʹ��IC��,��Ҫ�Լ��ж����ı��
'                    If mintInsure = TYPE_�¶� And mint���� = 1 Then
'                       str���ı�� = "81"
'                    End If
                     str���ı�� = mintIC��������
                Else
                    If mintInsure = TYPE_�¶� Then
                        If ������_�¶�(txtEdit(0).Text, strҽ����, str����, str���ı��) = False Then Exit Sub
                    Else
                        If ������_�ɶ�����(txtEdit(0).Text, strҽ����, str����, str���ı��) = False Then Exit Sub
                    End If
                End If
                
                txtEdit(0).Text = str����
                txtEdit(1).Text = strҽ����
                txtEdit(2).Text = str���ı��
                txtEdit(3).SetFocus
            End If
        ElseIf Index = 3 Then
            If mintInsure = TYPE_������ Then
                If cmdOK.Enabled = False Then
                    Call cmdTest_Click
                Else
                    Call cmdOK_Click
                End If
            Else
                If cmdOK.Enabled Then Call cmdOK_Click
            End If
        Else
            zlCommFun.PressKey (vbKeyTab)
        End If
        KeyAscii = 0
    End If
End Sub

Private Sub ShowErr(ByVal lngError As Long)
    Dim strMsg As String
    
    lngError = Abs(lngError)
    Select Case lngError
    Case &H80
        strMsg = "������"
    Case &H81
        strMsg = "д����"
    Case &H82
        strMsg = "ͨѶ����"
    Case &H83
        strMsg = "�������"
    Case &H84
        strMsg = "ͨѶ��ʱ��"
    Case &H85
        strMsg = "У��ʹ���"
    Case &H86
        strMsg = "��忨��"
    Case &H87
        strMsg = "����������ʽ����"
    Case Else
        strMsg = "δ֪����"
    End Select
    MsgBox strMsg & "����ţ�" & lngError, vbInformation, gstrSysName
End Sub

Private Sub CloseCommon(ByVal lngDev As Long)
    Dim intReturn As Integer
    
    If mint������ = 0 Then Exit Sub
    
    If mint������ = 1 Then
        intReturn = IC_ExitComm_1(lngDev)
    Else
        intReturn = IC_Down_2(lngDev)
        intReturn = ic_exit%(lngDev)
    End If
    If intReturn < 0 Then
        MsgBox "    �رն˿�ʱ����δ֪������ֻ��ͨ���رղ���ϵͳ���رն˿ڣ�" & vbCrLf & _
                "���ڶ˿��޷��رգ�������վ���޷��Կ���ɶ���д����", vbInformation, gstrSysName
    End If
End Sub

Private Sub ICErr(ByVal lngErr As Long, ByVal lngDev As Long)
    If lngErr < 0 Then
        Call ShowErr(lngErr)
        Call CloseCommon(lngDev)
    End If
End Sub
