VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frm����Set 
   BorderStyle     =   0  'None
   Caption         =   "���ۼۼ�����"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3345
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox PicInput 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2385
      ScaleWidth      =   3315
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3345
      Begin VB.CommandButton CmdNO 
         Caption         =   "ȡ��"
         Height          =   345
         Left            =   2415
         TabIndex        =   2
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton CmdYes 
         Caption         =   "ȷ��"
         Height          =   345
         Left            =   1425
         TabIndex        =   1
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox Txt�Ӽ��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1110
         MaxLength       =   8
         TabIndex        =   3
         Text            =   "15.0000"
         Top             =   1095
         Width           =   2130
      End
      Begin XtremeSuiteControls.ShortcutCaption stcTittle 
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   3990
         _Version        =   589884
         _ExtentX        =   7038
         _ExtentY        =   661
         _StockProps     =   6
         Caption         =   "�ۼۼ�����"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "    ������ӳ��ʣ����ۼ۵ļ��㹫ʽ�����ۼ�=�ɱ���*(1+�ӳ���%)"
         ForeColor       =   &H00400000&
         Height          =   600
         Left            =   60
         TabIndex        =   5
         Top             =   555
         Width           =   3405
      End
      Begin VB.Label Lbl�Ӽ��� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ӳ���(&J)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   4
         Top             =   1155
         Width           =   870
      End
   End
End
Attribute VB_Name = "frm����Set"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mdbl���ۼ� As Double
Private mdbl�ӳ��� As Double
Private mdbl����� As Double '
Private mlng����ID As Long
Private mintUnit As Integer
Private mbln��ǿ�ƿ���ָ���۸� As Boolean

Private msngX As Single, msngY As Single, mlngTxtH As Long
Public Function ShowCalc(ByVal frmMain As Form, _
    ByVal sngX As Single, ByVal sngY As Single, lngTxtH As Long, lng����ID As Long, intUnit As Integer, _
    ByRef dbl���ۼ� As Double, ByRef dbl����� As Double, ByRef dbl�ӳ��� As Double, ByVal bln��ǿ�ƿ���ָ���۸� As Boolean) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���ۼ�����
    '���:
    '����:
    '����:ѡ��,����true,���򷵻�False
    '����:���˺�
    '����:2008-11-07 11:23:35
    '-----------------------------------------------------------------------------------------------------------
    mdbl���ۼ� = dbl���ۼ�: mdbl����� = dbl�����: mdbl�ӳ��� = dbl�ӳ���: mlng����ID = lng����ID: mbln��ǿ�ƿ���ָ���۸� = bln��ǿ�ƿ���ָ���۸�
    msngX = sngX: msngY = sngY: mlngTxtH = lngTxtH: mintUnit = intUnit
    mblnOk = False
    Me.Show 1, frmMain
    dbl���ۼ� = mdbl���ۼ�: dbl����� = mdbl�����: dbl�ӳ��� = mdbl�ӳ���
    ShowCalc = mblnOk
End Function

Private Sub InitData()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-11-07 11:30:47
    '-----------------------------------------------------------------------------------------------------------
    
    If mdbl���ۼ� <> 0 And mdbl����� <> 0 Then
        Txt�Ӽ��� = Format(����ӳ���(), "###0.0000000;-###0.0000000;0;0")
    End If
    Txt�Ӽ���.Tag = Txt�Ӽ���
    
End Sub

Private Function ����ӳ���() As Single
    Dim sinָ�����ۼ� As Single, sin��������� As Single
    Dim rsTemp As New ADODB.Recordset
    '�������ۼ۷���ɱ���,����ʱ�����Ĺ�ʽ�ı仯,����ԭ������ӳ��ʵĹ�ʽ��Ч,�����¼���
    'ԭ��ʽ:(���ۼ�/�ɱ���-1)*100
    '�ֹ�ʽ������:�������ۼ��ǰ��ӳ����������,�ټ������������ǲ��ֽ��,���ʵ�ʰ��ӳ�����������ۼ�=ָ�����ۼ�-(ָ�����ۼ�-���ۼ�)/���������
    '������ԭ��ʽ���ʵ�ʵļӳ���
    ����ӳ��� = 0.15
    
    On Error GoTo ErrHandle
    gstrSQL = "Select a.����ϵ��,a.ָ�����ۼ�,Nvl(a.���������,100) ���������,Nvl(b.�Ƿ���,0) ʱ�� From �������� A, �շ���ĿĿ¼ b Where a.����ID=b.id  and b.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡָ�����ۼ�", mlng����ID)
    If rsTemp.EOF Then Exit Function
    
    sinָ�����ۼ� = rsTemp!ָ�����ۼ�
    sin��������� = rsTemp!���������
    If rsTemp!ʱ�� = 0 Then Exit Function
    
    'ָ�����ۼ�-(ָ�����ۼ�-���ۼ�)/���������
    sinָ�����ۼ� = sinָ�����ۼ� * IIf(mintUnit = 0, 1, Val(NVL(rsTemp!����ϵ��)))
    If sin��������� <> 100 And sin��������� > 0 Then
        mdbl���ۼ� = sinָ�����ۼ� - (sinָ�����ۼ� - mdbl���ۼ�) / sin��������� * 100
    Else
        mdbl���ۼ� = sinָ�����ۼ� - (sinָ�����ۼ� - mdbl���ۼ�)
    End If
    ����ӳ��� = (mdbl���ۼ� / mdbl����� - 1) * 100
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function У�����ۼ�(ByVal dbl���ۼ� As Double) As Double
    '�õ�����ǰ��λϵ�����������ָ�����ۼۣ����ʱ��������ǿ�ƿ���ָ���۸������������ۼ۴���ָ�����ۼۣ���ָ�����ۼ�Ϊ׼
    Dim sinָ�����ۼ� As Single
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select ָ�����ۼ�, " & IIf(mintUnit = 0, 1, " ����ϵ�� ") & "   as ����ϵ��,Nvl(���������,100) ��������� From �������� Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡָ�����ۼ�", mlng����ID)
    If rsTemp.EOF Then Exit Function
    sinָ�����ۼ� = zlStr.NVL(rsTemp!ָ�����ۼ�, 0)
    sinָ�����ۼ� = sinָ�����ۼ� * Val(zlStr.NVL(rsTemp!����ϵ��))
    If sinָ�����ۼ� = 0 Then sinָ�����ۼ� = dbl���ۼ�
    У�����ۼ� = IIf(dbl���ۼ� > sinָ�����ۼ� And Not mbln��ǿ�ƿ���ָ���۸�, sinָ�����ۼ�, dbl���ۼ�)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub CmdNO_Click()
    mblnOk = False
    Unload Me
End Sub
Private Sub CmdYes_Click()
    If Val(Txt�Ӽ���) > 9900 Or Val(Txt�Ӽ���) < 0 Then
        MsgBox "������Ϸ��ļӳ��ʣ���0-9900��", vbInformation, gstrSysName
        Txt�Ӽ���.SetFocus
        Exit Sub
    End If
    
    mdbl�ӳ��� = Val(Txt�Ӽ���)
    '���¼������ۼۡ����
    mdbl���ۼ� = У�����ۼ�(mdbl����� * (1 + (Val(Txt�Ӽ���) / 100)) + _
    ʱ�۲������ۼ�(mlng����ID, mdbl�����, Val(Txt�Ӽ���) / 100))
    mblnOk = True
    Unload Me
End Sub
Private Function ʱ�۲������ۼ�(ByVal lng����ID As Long, ByVal sin�ɹ��� As Single, ByVal sin�ӳ��� As Single, _
    Optional sng�ۼ� As Single = -99999999) As Double
    '------------------------------------------------------------------------------------------------------
    '����:����ָ���۸���۱ȼ����ʱ�۲��ϵĲ���������
    '���:lng����ID-����ID
    '     sin�ɹ���-�ɹ��۸�
    '     sin�ӳ���-�ӳ���(�������0,ͬʱ�ִ���dbl���ۼ�,�򽫰���������ۼ۽��м���)
    '     LngLastRow-���ݵ��к�
    '     sng�ۼ�-��������ۼ�
    '����:
    '����:���ۼ۵��������
    '�޸���:���˺�
    '�޸�ʱ��:2007/2/25
    '------------------------------------------------------------------------------------------------------
    'ʱ�۲������ۼۼ��㹫ʽ:�ɹ���*(1+�ӳ���)
    '��Ϊ:�ɹ���*(1+�ӳ���)+(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
    '���ڲ�������ȵĴ���,��ǰ���а�ָ������ʼ���ĵط�,����Ҫ�������ת���ɼӳ��ʽ��м���,�˺������ڷ��ر��ι�ʽ���ӵĲ��ֽ�(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
    
    Dim sin���ۼ� As Single, sinָ�����ۼ� As Single, sin��������� As Single
    Dim dblϵ�� As Double
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select ָ�����ۼ�,Nvl(���������,100) ���������," & IIf(mintUnit = 0, 1, "����ϵ��") & " As  ����ϵ�� From �������� Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡָ�����ۼ�", lng����ID)
    
    If rsTemp.EOF Then Exit Function
    
    dblϵ�� = Val(zlStr.NVL(rsTemp!����ϵ��))
    sinָ�����ۼ� = rsTemp!ָ�����ۼ�
    sin��������� = rsTemp!���������
    
    ʱ�۲������ۼ� = 0
    If sin��������� = 100 Then Exit Function
    If sinָ�����ۼ� = 0 Then Exit Function
    
    sin���ۼ� = sin�ɹ��� * (1 + sin�ӳ���)
    If sin���ۼ� / dblϵ�� >= sinָ�����ۼ� Then Exit Function
    sinָ�����ۼ� = sinָ�����ۼ� * dblϵ��
    ʱ�۲������ۼ� = (sinָ�����ۼ� - sin���ۼ�) * (1 - sin��������� / 100)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub Form_Activate()
    Call zlControl.ControlSetFocus(Txt�Ӽ���)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call CmdYes_Click
        Exit Sub
    End If
    If KeyCode = vbKeyEscape Then '
        mblnOk = False
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Call InitData
    With Me
        If msngX + .Width > Screen.Width Then
            .Left = Screen.Width - .Width
        Else
            .Left = msngX
        End If
        If msngY + .Height > Screen.Height Then
           .Top = msngY - mlngTxtH - .Height
        Else
            .Top = msngY
        End If
    End With
    
End Sub
Private Sub Txt�Ӽ���_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress Txt�Ӽ���, KeyAscii, m���ʽ
End Sub
