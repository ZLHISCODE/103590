VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{09B13292-AC31-4C5D-B44A-C83E7AAD70E6}#1.1#0"; "zlSubclass.ocx"
Begin VB.Form frmShiftEdit 
   BackColor       =   &H80000004&
   Caption         =   "���˽��Ӱ����ݱ༭"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14940
   Icon            =   "frmShiftEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   14940
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picSplitX 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   0
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6615
      ScaleWidth      =   45
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   50
   End
   Begin VB.PictureBox picMainBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3255
      ScaleWidth      =   8415
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   600
      Width           =   8415
      Begin VB.PictureBox picMain 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   240
         ScaleHeight     =   3225
         ScaleWidth      =   7965
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   480
         Width           =   7995
         Begin VB.PictureBox picEdit 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3975
            Left            =   120
            ScaleHeight     =   3975
            ScaleWidth      =   7815
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   120
            Width           =   7815
            Begin VB.PictureBox picPanel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1260
               Left            =   600
               ScaleHeight     =   1260
               ScaleWidth      =   7125
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   120
               Width           =   7125
               Begin VB.CommandButton cmdType 
                  BackColor       =   &H80000004&
                  Caption         =   "��"
                  Height          =   250
                  Left            =   6600
                  TabIndex        =   21
                  TabStop         =   0   'False
                  ToolTipText     =   "ѡ��(*)"
                  Top             =   0
                  Width           =   270
               End
               Begin VB.CommandButton cmdFind 
                  Caption         =   "��"
                  Height          =   270
                  Left            =   3000
                  TabIndex        =   22
                  TabStop         =   0   'False
                  ToolTipText     =   "���ҵ�ǰ���ҵ�����סԺ�Ĳ���"
                  Top             =   0
                  Width           =   270
               End
               Begin VB.TextBox txtPatiInfo 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Height          =   290
                  Index           =   7
                  Left            =   5025
                  Locked          =   -1  'True
                  TabIndex        =   8
                  TabStop         =   0   'False
                  Top             =   915
                  Width           =   1960
               End
               Begin VB.TextBox txtPatiInfo 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Height          =   290
                  Index           =   6
                  Left            =   3375
                  Locked          =   -1  'True
                  TabIndex        =   7
                  TabStop         =   0   'False
                  Top             =   915
                  Width           =   735
               End
               Begin VB.TextBox txtPatiInfo 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Height          =   290
                  Index           =   5
                  Left            =   855
                  Locked          =   -1  'True
                  TabIndex        =   6
                  TabStop         =   0   'False
                  Top             =   915
                  Width           =   1575
               End
               Begin VB.TextBox txtPatiInfo 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Height          =   290
                  Index           =   4
                  Left            =   5040
                  Locked          =   -1  'True
                  TabIndex        =   5
                  TabStop         =   0   'False
                  Top             =   440
                  Width           =   1920
               End
               Begin VB.TextBox txtPatiInfo 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Height          =   290
                  Index           =   3
                  Left            =   3345
                  Locked          =   -1  'True
                  TabIndex        =   4
                  TabStop         =   0   'False
                  Top             =   440
                  Width           =   735
               End
               Begin VB.TextBox txtPatiInfo 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Height          =   290
                  Index           =   2
                  Left            =   855
                  Locked          =   -1  'True
                  TabIndex        =   3
                  TabStop         =   0   'False
                  Top             =   457
                  Width           =   1560
               End
               Begin VB.TextBox txtPatiInfo 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  Height          =   290
                  Index           =   1
                  Left            =   855
                  TabIndex        =   1
                  Top             =   10
                  Width           =   2400
               End
               Begin VB.TextBox txtPatiInfo 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Height          =   290
                  Index           =   0
                  Left            =   4320
                  Locked          =   -1  'True
                  TabIndex        =   2
                  TabStop         =   0   'False
                  Top             =   10
                  Width           =   2520
               End
               Begin VB.Label lblPatiInfo 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  Caption         =   "��Ժʱ��"
                  ForeColor       =   &H80000008&
                  Height          =   180
                  Index           =   7
                  Left            =   4260
                  TabIndex        =   30
                  Top             =   975
                  Width           =   720
               End
               Begin VB.Label lblPatiInfo 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  Caption         =   "��Ժ;��"
                  ForeColor       =   &H80000008&
                  Height          =   180
                  Index           =   6
                  Left            =   2565
                  TabIndex        =   29
                  Top             =   975
                  Width           =   720
               End
               Begin VB.Label lblPatiInfo 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  Caption         =   "סԺ��"
                  ForeColor       =   &H80000008&
                  Height          =   180
                  Index           =   5
                  Left            =   240
                  TabIndex        =   28
                  Top             =   975
                  Width           =   540
               End
               Begin VB.Label lblPatiInfo 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  Caption         =   "����"
                  ForeColor       =   &H80000008&
                  Height          =   180
                  Index           =   4
                  Left            =   4635
                  TabIndex        =   27
                  Top             =   500
                  Width           =   360
               End
               Begin VB.Label lblPatiInfo 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  Caption         =   "����"
                  ForeColor       =   &H80000008&
                  Height          =   180
                  Index           =   3
                  Left            =   2940
                  TabIndex        =   26
                  Top             =   500
                  Width           =   360
               End
               Begin VB.Label lblPatiInfo 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  Caption         =   "�Ա�"
                  ForeColor       =   &H80000008&
                  Height          =   180
                  Index           =   2
                  Left            =   390
                  TabIndex        =   25
                  Top             =   500
                  Width           =   360
               End
               Begin VB.Label lblPatiInfo 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  Caption         =   "����"
                  ForeColor       =   &H80000008&
                  Height          =   180
                  Index           =   1
                  Left            =   420
                  TabIndex        =   24
                  Top             =   60
                  Width           =   360
               End
               Begin VB.Label lblPatiInfo 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  Caption         =   "��������"
                  ForeColor       =   &H80000008&
                  Height          =   180
                  Index           =   0
                  Left            =   3480
                  TabIndex        =   23
                  Top             =   60
                  Width           =   720
               End
            End
            Begin VB.OptionButton optInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   4920
               TabIndex        =   19
               Top             =   2400
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "��ѡ"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   4560
               TabIndex        =   18
               Top             =   2400
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               Height          =   290
               Index           =   0
               Left            =   2880
               MaxLength       =   250
               MultiLine       =   -1  'True
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   2400
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.PictureBox picTmp 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   290
               Index           =   0
               Left            =   5880
               ScaleHeight     =   285
               ScaleWidth      =   1335
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   2160
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Ŀǰ���"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   0
               Left            =   2040
               TabIndex        =   31
               Top             =   2400
               Visible         =   0   'False
               Width           =   720
            End
         End
         Begin VB.VScrollBar vscBar 
            Height          =   7575
            LargeChange     =   200
            Left            =   7800
            SmallChange     =   200
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   0
            Width           =   255
         End
      End
   End
   Begin VB.PictureBox picInfo 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   8640
      ScaleHeight     =   2775
      ScaleWidth      =   5535
      TabIndex        =   0
      Top             =   480
      Width           =   5535
      Begin VB.Frame fraInfo 
         BackColor       =   &H80000004&
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   2415
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   4575
         Begin RichTextLib.RichTextBox rtbBox 
            Height          =   1935
            Left            =   120
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   240
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   3413
            _Version        =   393217
            BackColor       =   -2147483644
            BorderStyle     =   0
            Enabled         =   -1  'True
            Appearance      =   0
            TextRTF         =   $"frmShiftEdit.frx":6852
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin zlSubclass.Subclass Subclass 
      Left            =   1200
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Label lblWdith 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "��ȼ���"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   9120
      TabIndex        =   11
      Top             =   3960
      Visible         =   0   'False
      Width           =   720
   End
   Begin XtremeCommandBars.CommandBars cbsExec 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmShiftEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const WM_MOUSEWHEEL = &H20A          '������
Private Const con�����X = 470  'SBAR�����X
Private Const con��� = 75  '���

Private Enum MenuType
        ID_���� = 1
        ID_�޸� = 2
        ID_ɾ�� = 3
        ID_���� = 4
        ID_ȡ�� = 5
        ID_���� = 6
        
        ID_���� = 99
End Enum

Private Enum PatiInfo
        idx�������� = 0
        idx���� = 1
        idx�Ա� = 2
        idx���� = 3
        idx���� = 4
        idxסԺ�� = 5
        idx��Ժ;�� = 6
        idx��Ժʱ�� = 7
End Enum

Private Type Ctl����
    'X
    ������1 As Long
    ������2 As Long
    
    'Y
    S������ As Long
    B������ As Long
    A������ As Long
    R������ As Long
    
    'Y
    lngTop As Long '��̬�ؼ�����߶�
    LngY As Long '���ڲ���ʱ��¼��ǰ���ø߶�
End Type

Private Type InfoD
    '������Ϣ
    �������� As String '����,һ��...
    ����ID As Long
    ��ҳID As Long

    '�����¼
    �����¼ID As Long
    �������ID As Long
    ���࿪ʼʱ�� As Date
    �������ʱ�� As Date
    
    '���ݼ�¼
    EditType As Long '�༭��ʽ 0-Ԥ����1-����  2-�޸�
    ����ID As Long
    ������� As Long
End Type


Private mCtl���� As Ctl���� '������Ϣ
Private mInfo As InfoD '�������

Private mobjCtl As Object '�������ý���

Private mbtnNoEdit As Boolean '������༭
Private mblnChange As Boolean '�ؼ���¼�ı�
Private mblnSave As Boolean '�����Ƿ��ѱ���
Private mblnEnter As Boolean '�Ƿ�س�
Private mblnLoad As Boolean '�Ƿ����ڳ�ʼ��
Private mblnԤ�� As Boolean

Public gstrԤ������ As String 'Ԥ���������

'��������
Private mstrLike As String  '��֧��ȫƥ��
Private mblnѪ��ϵͳ As Boolean   '�Ƿ�װѪ��ϵͳ


'���󻺴�
Public gfrmParent As Object             '���������
Private mrsCtlType As ADODB.Recordset   '�������ͼ�¼��
Private mrsCtlInfo As ADODB.Recordset   '�ؼ���Ϣ��¼��

'�ַ�������
Private mstrCtls As String          '����ؼ���Ϣ�ַ���
Private mstrTextCtl As String       '�ı���idx�ַ���
Private mstrPicCtl As String        '���ؼ�idx�ַ���
Private mstrPatiType As String      '���������ַ���


Private Sub chkInfo_Click(Index As Integer)
    If (mInfo.EditType <> 0 And mblnLoad = False) Or mblnԤ�� Then
        mblnChange = True
        Call SetTextVisible(Val(Split(chkInfo(Index).Tag, ",")(0)))
        Call MakeText
    End If
End Sub


Private Sub SetTextVisible(lngIndex As Long)
   '����ѡ��ؼ����ı����Ƿ���ʾ
    Dim i As Long, intType As Integer, blnVisible As Boolean

    On Error GoTo errH
    If lngIndex = 0 Then Exit Sub
    If InStr("," & mstrTextCtl & ",", "," & lngIndex & ",") = 0 Or InStr("," & mstrPicCtl & ",", "," & lngIndex & ",") = 0 Then Exit Sub
    intType = Val(Split(picTmp(lngIndex).Tag, ",")(0))
    
    If intType = 2 Then
            For i = 1 To optInfo.Count - 1
                If Val(Split(optInfo(i).Tag, ",")(0)) = lngIndex Then
                    If optInfo(i).Value = True And Split(optInfo(i).Tag, ",")(1) = "1" Then
                        blnVisible = True
                        Exit For
                    End If
                End If
            Next
    ElseIf intType = 3 Then
        For i = 1 To chkInfo.Count - 1
            If Val(Split(chkInfo(i).Tag, ",")(0)) = lngIndex Then
                If chkInfo(i).Value = 1 And Split(chkInfo(i).Tag, ",")(1) = "1" Then
                    blnVisible = True
                    Exit For
                End If
            End If
        Next
    End If
        
    If txtInfo(lngIndex).Visible <> blnVisible Then
         txtInfo(lngIndex).Visible = blnVisible
         txtInfo(lngIndex).Text = ""
         Call picMain_Resize
         Call DrawLine
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub chkInfo_GotFocus(Index As Integer)
    If mblnEnter Then Call ShowCtl(chkInfo(Index)): mblnEnter = False
End Sub

Private Sub cmdFind_Click()
    Call GetPatiList(1)
End Sub

Private Sub cmdType_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    Dim strTmp As String
    
    On Error GoTo errH

    vPoint = zlcontrol.GetCoordPos(txtPatiInfo(idx��������).Container.hWnd, txtPatiInfo(idx��������).Left, txtPatiInfo(idx��������).Top)
    blnCancel = True
    
    strSQL = "Select a.˳�� As ID, a.���, a.���� ,Decode(b.C2, Null, 0, 1) As �ѹ�ѡcheck" & vbNewLine & _
            "From ҽ�����Ӱಡ������ A, Table(Cast(f_Str2list2([1]) As Zltools.t_Strlist2)) B" & vbNewLine & _
            "Where a.�Ƿ�ͣ�� = 0 And a.��� = b.C2(+)" & vbNewLine & _
            "Order By ˳��"


    Set mobjCtl = txtPatiInfo(idx��������)
    Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "ѡ��������", True, "", "", True, True, True, vPoint.X, vPoint.Y, txtPatiInfo(idx��������).Height, blnCancel, True, True, txtPatiInfo(idx��������).Text)
    
    zlcontrol.ControlSetFocus txtPatiInfo(idx��������)
    If Not blnCancel Then
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                strTmp = strTmp & "," & rsTmp!���
                rsTmp.MoveNext
            Loop
            
            If mInfo.�������� <> Mid(strTmp, 2) Then
                Set mobjCtl = txtPatiInfo(idx��������)
                If mblnChange And txtPatiInfo(idx��������).Text <> "" Then
                    If MsgBox("�л��������ͽ������ǰδ�������Ŀ,��ȷ���Ƿ������", vbInformation + vbDefaultButton2 + vbYesNo, Me.Caption) = vbNo Then
                        zlcontrol.ControlSetFocus txtPatiInfo(idx��������)
                        Exit Sub
                    End If
                End If
                
                txtPatiInfo(idx��������).Text = Mid(strTmp, 2)
                mInfo.�������� = txtPatiInfo(idx��������).Text
                Call UnloadCtl
                Call InitCtl
                Call IntData
                Call MakeText
            End If
        Else
            Set mobjCtl = txtPatiInfo(idx��������)
            MsgBox "δ���ҵ�����ѡ��Ĳ�������!", vbInformation, Me.Caption
            zlcontrol.ControlSetFocus txtPatiInfo(idx��������)
            Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If Not mobjCtl Is Nothing Then
        zlcontrol.ControlSetFocus mobjCtl
        Set mobjCtl = Nothing
    End If
    
    
    '�״ν���Ԥ������ʱ�����Ű���ʳ��ִ��ң������ݴ���
    If gstrԤ������ <> "" Then
        Call picMain_Resize
        Call Form_Resize
        gstrԤ������ = ""
    End If
End Sub

Private Sub Form_Load()

    On Error GoTo errH
    
    If gstrԤ������ <> "" Then mblnԤ�� = True
    
    '��ȡ�ؼ���Ϣ��¼��
    Call GetCtlRs
    
    '����ť
    Call InitExecBar

    '�̶�����
    picEdit.Top = 0: picEdit.Left = 0: picEdit.Width = picMain.Width: picEdit.Height = 10000
    picPanel.Top = 120: picPanel.Left = 600
    mCtl����.lngTop = picPanel.Top + picPanel.Height
    
    '��ʼ������
    picSplitX.BackColor = Me.BackColor
    picSplitX.Left = 8720
    
    '�����¼���ʼ��
    Subclass.hWnd = Me.hWnd
    Subclass.Messages(WM_MOUSEWHEEL) = True
    
    '��ʼ������
    mstrLike = IIf(zlDatabase.GetPara("����ƥ��") = "0", "%", "")
    mblnѪ��ϵͳ = (IsSysSetUp(2200) And Val(zlDatabase.GetPara(236, glngSys)) <> 0)
    
    '�����ʼ��
    Call zlRefresh(0, 0, 0, 0, 0, Now - 1, Now + 1, False, gstrԤ������)

    '����Ԥ������
    If mblnԤ�� Then
        txtPatiInfo(idx��������).Text = gstrԤ������
        Call SetPicEnabled(True)
        cmdType.Visible = False
        Call RestoreWinState(Me, App.ProductName)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Form_Paint()
    Call DrawLine
End Sub

Private Sub Form_Resize()
    Dim lngTop As Long
    
    On Error Resume Next
    
    lngTop = IIf(mblnԤ��, 500, 340)

    picSplitX.Top = lngTop: picSplitX.Height = Me.Height - lngTop - IIf(mblnԤ��, 250, 0)
    picMainBack.Move 0, lngTop, picSplitX.Left, Me.Height - lngTop - IIf(mblnԤ��, 600, 0)
    picMain.Move 0, 85, picMainBack.Width, picMainBack.Height - 85
    picInfo.Move picSplitX.Left + picSplitX.Width, lngTop, Me.Width - (picSplitX.Left + picSplitX.Width) - IIf(mblnԤ��, 270, 0), Me.Height - lngTop - IIf(mblnԤ��, 570, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjCtl = Nothing
    Set mrsCtlType = Nothing
    Set mrsCtlInfo = Nothing
    gstrԤ������ = ""
    
    Subclass.Messages(WM_MOUSEWHEEL) = False

    If mblnԤ�� Then Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub InitExecBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl


    On Error GoTo errH
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsExec.VisualTheme = xtpThemeOfficeXP
    With Me.cbsExec.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .UseFadedIcons = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = False
        If mblnԤ�� = True Then
            .SetIconSize False, 24, 24
        Else
            .SetIconSize False, 16, 16
        End If
    End With
    Set cbsExec.Icons = zlCommFun.GetPubIcons
    cbsExec.EnableCustomization False
    cbsExec.ActiveMenuBar.Visible = False
    
    Set objBar = cbsExec.Add("������", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap '+ xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    With objBar.Controls
        If mblnԤ�� Then
            '������
            Set objControl = .Add(xtpControlButton, ID_����, "�������ͣ�")
            objControl.IconId = 807

            If Not mrsCtlType Is Nothing Then
                mrsCtlType.Filter = ""
                Do While Not mrsCtlType.EOF
                    Set objControl = .Add(xtpControlButton, ID_���� + Val(mrsCtlType!˳�� & ""), mrsCtlType!��� & "")
                    objControl.IconId = IIf(InStr("," & gstrԤ������ & ",", "," & mrsCtlType!��� & ",") > 0, 12, 10)

                    mrsCtlType.MoveNext
                Loop
            End If
        Else
            Set objControl = .Add(xtpControlButton, ID_����, "���Ӳ�������")
                objControl.IconId = 816
            Set objControl = .Add(xtpControlButton, ID_����, "����")
                objControl.IconId = 4112
                objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, ID_�޸�, "�޸�")
                objControl.IconId = 4113
            Set objControl = .Add(xtpControlButton, ID_ɾ��, "ɾ��")
                objControl.IconId = 4114
            Set objControl = .Add(xtpControlButton, ID_����, "����")
                objControl.IconId = 3091
                objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, ID_ȡ��, "ȡ��")
                objControl.IconId = 3014
        End If

    End With
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub UnloadCtl()
    'ж�ؽ���Ŀؼ�
    Dim obj As Object
    
    On Error Resume Next
    For Each obj In lblInfo
        If obj.Index <> 0 Then
            Unload obj
        End If
    Next
    
    For Each obj In txtInfo
        If obj.Index <> 0 Then
            Unload obj
        End If
    Next
    
    For Each obj In optInfo
        If obj.Index <> 0 Then
            Unload obj
        End If
    Next
    
    For Each obj In chkInfo
        If obj.Index <> 0 Then
            Unload obj
        End If
    Next
    
    For Each obj In picTmp
        If obj.Index <> 0 Then
            Unload obj
        End If
    Next
    
End Sub

Private Sub CtlVisible(blnVisible As Boolean)
    '���ƽ���Ŀؼ���ʾ
    Dim obj As Object
    
    On Error Resume Next
    For Each obj In lblInfo
        If obj.Index <> 0 Then
            obj.Visible = blnVisible
        End If
    Next
    
    For Each obj In txtInfo
        If obj.Index <> 0 And InStr("," & mstrPicCtl & ",", "," & obj.Index & ",") = 0 Then
            obj.Visible = blnVisible
        End If
    Next
    
    For Each obj In optInfo
        If obj.Index <> 0 Then
            obj.Visible = blnVisible
        End If
    Next
    
    For Each obj In chkInfo
        If obj.Index <> 0 Then
            obj.Visible = blnVisible
        End If
    Next
    
    For Each obj In picTmp
        If obj.Index <> 0 Then
            obj.Visible = blnVisible
        End If
    Next
End Sub

Private Sub InitCtl()
    '��ʼ������Ķ�̬�ؼ�
    Dim i As Long, j As Long, m As Long, n As Long
    Dim intIndex As Integer, intTabIndex As Integer, intOptIndex As Integer, intChkIndex As Integer
    Dim rsInfoCopy As ADODB.Recordset
    Dim arrTmp As Variant
    Dim blnCheck As Boolean, blnText As Boolean
    
    On Error GoTo errH
    
    '��ջ�����
    mstrCtls = "": mstrTextCtl = "": mstrPicCtl = ""
    
    If mrsCtlType Is Nothing Or mrsCtlInfo Is Nothing Or mInfo.�������� = "" Then
        Call picMain_Resize
        Exit Sub
    End If
    
    '��ԭ��¼��
    mrsCtlType.Filter = ""
    mrsCtlInfo.Filter = ""
    
    If mrsCtlType.EOF Or mrsCtlInfo.EOF Then Exit Sub
    
    mrsCtlType.MoveFirst
    mrsCtlInfo.MoveFirst
    
    Set rsInfoCopy = zlDatabase.CopyNewRec(mrsCtlInfo) '����һ����¼���������ж��Ƿ�����
    
    mblnLoad = True
    intTabIndex = 9
    For m = 1 To 4 '��SBAR��˳�����
        mrsCtlType.MoveFirst
        For i = 1 To mrsCtlType.RecordCount
            If mInfo.�������� = "" Or InStr("," & mInfo.�������� & ",", "," & mrsCtlType!��� & ",") > 0 Then
                    mrsCtlInfo.Filter = "���˼��='" & mrsCtlType!��� & "' And ��Ŀ���='" & Decode(m, 1, "S", 2, "B", 3, "A", 4, "R") & "'"
                    For j = 1 To mrsCtlInfo.RecordCount
                        
                        '������˳��
                        rsInfoCopy.Filter = "��Ŀ���� ='" & mrsCtlInfo!��Ŀ���� & "'"
                        blnCheck = True
                        Do While Not rsInfoCopy.EOF
                            If rsInfoCopy!���˼�� & "" <> mrsCtlInfo!���˼�� & "" And InStr("," & mInfo.�������� & ",", "," & rsInfoCopy!���˼�� & ",") > 0 Then
                                If InStr(mstrPatiType, "," & rsInfoCopy!���˼�� & ",") < InStr(mstrPatiType, "," & mrsCtlInfo!���˼�� & ",") Then
                                    blnCheck = False
                                    Exit Do
                                End If
                            End If
                            rsInfoCopy.MoveNext
                        Loop
                        rsInfoCopy.Filter = ""
                        
                        '�������������
                        If blnCheck Then
                            If ((InStr("," & mInfo.�������� & ",", "����") > 0 Or mInfo.�������� = "") And Val(mrsCtlInfo!���������� & "") = 1) Then
                                blnCheck = False
                            End If
                        End If
                    
                        If InStr(mstrCtls, "," & mrsCtlInfo!��Ŀ����) = 0 And blnCheck Then
                        
                            intIndex = lblInfo.Count
                            Load lblInfo(intIndex)
                            
                            lblInfo(intIndex).Caption = mrsCtlInfo!��Ŀ���� & ""
                            
                             '�����ȹ������»���
                            If lblInfo(intIndex).Width > 1200 Then
                                lblInfo(intIndex).Caption = Mid(mrsCtlInfo!��Ŀ���� & "", 1, Len(lblInfo(intIndex).Caption) / 2) & vbCrLf & Mid(mrsCtlInfo!��Ŀ���� & "", Len(lblInfo(intIndex).Caption) / 2 + 1)
                            End If
                            
                            
                            '����ؼ�һ�����и���
                            rsInfoCopy.Filter = "���˼��='" & mrsCtlType!��� & "' And ��Ŀ���='" & Decode(m, 1, "S", 2, "B", 3, "A", 4, "R") & "' And ���=" & mrsCtlInfo!��� & " And ��Ŀ����<> '" & mrsCtlInfo!��Ŀ���� & "'"
                            If Not rsInfoCopy.EOF Then
                                If InStr(mstrCtls, mrsCtlType!��� & "," & rsInfoCopy!��Ŀ����) = 0 Then
                                    lblInfo(intIndex).Tag = "1," & mrsCtlInfo!���˼�� & "," & mrsCtlInfo!��Ŀ��� & "," & mrsCtlInfo!��������
                                Else
                                    lblInfo(intIndex).Tag = "2," & mrsCtlInfo!���˼�� & "," & mrsCtlInfo!��Ŀ��� & "," & mrsCtlInfo!��������
                                End If
                            Else
                                If Val(mrsCtlInfo!������ʽ & "") = 1 And (Val(mrsCtlInfo!�������� & "") = 1 Or Val(mrsCtlInfo!�������� & "") = 2) Then
                                    lblInfo(intIndex).Tag = "1," & mrsCtlInfo!���˼�� & "," & mrsCtlInfo!��Ŀ��� & "," & mrsCtlInfo!��������
                                Else
                                    lblInfo(intIndex).Tag = "0," & mrsCtlInfo!���˼�� & "," & mrsCtlInfo!��Ŀ��� & "," & mrsCtlInfo!��������
                                End If
                            End If
                            
                            
                            Select Case Val(mrsCtlInfo!������ʽ & "")
                                    Case 1 '�����
                                        Load txtInfo(intIndex)
                                        
                                        '�ؼ�Tab˳����
                                        intTabIndex = intTabIndex + 1
                                        txtInfo(intIndex).TabIndex = intTabIndex
                                        txtInfo(intIndex).Locked = Val(mrsCtlInfo!�Ƿ�ֻ�� & "") = 1
                                        txtInfo(intIndex).BackColor = IIf(txtInfo(intIndex).Locked, &H8000000F, &H80000005)
                                        txtInfo(intIndex).TabStop = Not txtInfo(intIndex).Locked
                                        
                                        '�ؼ��߶�
                                        txtInfo(intIndex).Height = txtInfo(intIndex).Height * IIf(Val(mrsCtlInfo!�������� & "") = 0, 1, Val(mrsCtlInfo!�������� & ""))
                                        
                                        mstrTextCtl = mstrTextCtl & "," & intIndex
                                    Case 2 '����ѡ��
                                        Load picTmp(intIndex) '���������ؼ�
                                        blnText = False
                                        mstrPicCtl = mstrPicCtl & "," & intIndex
                                        
                                        picTmp(intIndex).Tag = "2," & optInfo.Count '������Ŀ��Ӧѡ��ؼ���ʼ
                                        arrTmp = Split(mrsCtlInfo!����ֵ�� & "", ",")
                                        For n = 0 To UBound(arrTmp)
                                            intOptIndex = optInfo.Count
                                            Load optInfo(intOptIndex)
                                            
                                            intTabIndex = intTabIndex + 1
                                            optInfo(intOptIndex).TabStop = True
                                            optInfo(intOptIndex).TabIndex = intTabIndex
                                            optInfo(intOptIndex).Caption = IIf(Mid(arrTmp(n), 1, 1) = "*", Mid(arrTmp(n), 2), arrTmp(n))
                                            optInfo(intOptIndex).Tag = intIndex & "," & IIf(Mid(arrTmp(n), 1, 1) = "*", "1", "")
                                            
                                            '����ѡ����
                                            lblWdith.Caption = arrTmp(n)
                                            optInfo(intOptIndex).Width = lblWdith.Width + 300
                                      
                                            '���ø�������
                                            Set optInfo(intOptIndex).Container = picTmp(intIndex)
                                            
                                            If Mid(arrTmp(n), 1, 1) = "*" Then blnText = True '�����ı���
                                        Next
                                        optInfo(Val(Split(picTmp(intIndex).Tag, ",")(1))).Value = True '��ʱĬ�ϵ�һ��ΪĬ��ѡ��
                                        
                                        '���ô�������Ϣ���ı���
                                        If blnText Then
                                            Load txtInfo(intIndex)
                                            '�ؼ�Tab˳����
                                            intTabIndex = intTabIndex + 1
                                            txtInfo(intIndex).TabIndex = intTabIndex
                                            txtInfo(intIndex).TabStop = True
                                            
                                            txtInfo(intIndex).Visible = Split(optInfo(Val(Split(picTmp(intIndex).Tag, ",")(1))).Tag, ",")(1) = "1"
                                            
                                            '�ؼ��߶�
                                            txtInfo(intIndex).Height = txtInfo(intIndex).Height * IIf(Val(mrsCtlInfo!�������� & "") = 0, 1, Val(mrsCtlInfo!�������� & ""))
                                            
                                            mstrTextCtl = mstrTextCtl & "," & intIndex
                                        End If

                                        picTmp(intIndex).Tag = picTmp(intIndex).Tag & "," & optInfo.Count - 1 '������Ŀ��Ӧѡ��ؼ�����
                                    Case 3 '����ѡ��
                                        blnText = False
                                        Load picTmp(intIndex) '���������ؼ�
                                        
                                        mstrPicCtl = mstrPicCtl & "," & intIndex
                                        
                                        picTmp(intIndex).Tag = "3," & chkInfo.Count '������Ŀ��Ӧѡ��ؼ���ʼ
                                            
                                        arrTmp = Split(mrsCtlInfo!����ֵ�� & "", ",")
                                        For n = 0 To UBound(arrTmp)
                                            intChkIndex = chkInfo.Count
                                            Load chkInfo(intChkIndex)
                                            
                                            intTabIndex = intTabIndex + 1
                                            chkInfo(intChkIndex).TabStop = True
                                            chkInfo(intChkIndex).TabIndex = intTabIndex
                                            chkInfo(intChkIndex).Caption = IIf(Mid(arrTmp(n), 1, 1) = "*", Mid(arrTmp(n), 2), arrTmp(n))
                                            chkInfo(intChkIndex).Tag = intIndex & "," & IIf(Mid(arrTmp(n), 1, 1) = "*", "1", "")
                                            
                                            '����ѡ����
                                            lblWdith.Caption = arrTmp(n)
                                            chkInfo(intChkIndex).Width = lblWdith.Width + 300
                                            
                                            '���ø�������
                                            Set chkInfo(intChkIndex).Container = picTmp(intIndex)
                                            
                                            If Mid(arrTmp(n), 1, 1) = "*" Then blnText = True '�����ı���
                                            
                                            picTmp(intIndex).Refresh
                                        Next
                                        
                                        '���ô�������Ϣ���ı���
                                        If blnText Then
                                            Load txtInfo(intIndex)
                                            '�ؼ�Tab˳����
                                            intTabIndex = intTabIndex + 1
                                            txtInfo(intIndex).TabIndex = intTabIndex
                                            txtInfo(intIndex).TabStop = True
                                            txtInfo(intIndex).Visible = False 'Ĭ��Ϊ����ʾ
                                            
                                            '�ؼ��߶�
                                            txtInfo(intIndex).Height = txtInfo(intIndex).Height * IIf(Val(mrsCtlInfo!�������� & "") = 0, 1, Val(mrsCtlInfo!�������� & ""))
                                            
                                            mstrTextCtl = mstrTextCtl & "," & intIndex
                                        End If

                                        picTmp(intIndex).Tag = picTmp(intIndex).Tag & "," & chkInfo.Count - 1 '������Ŀ��Ӧѡ��ؼ�����
                            End Select
                            
                            mstrCtls = mstrCtls & ";" & mrsCtlType!��� & "," & mrsCtlInfo!��Ŀ����
                            
                        End If
                        mrsCtlInfo.MoveNext
                    Next
            End If
            mrsCtlType.MoveNext
        Next
    Next
    
     
    Call picMain_Resize
    
    Call CtlVisible(True) '�����ʾ�ؼ�
    
    mblnLoad = False
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub CtlResize()
    '��̬�ؼ��Ű�
    Dim i As Long, j As Long
    Dim lngTmp1 As Long, lngTmp2 As Long
    Dim lngTop As Long, lngMaxTop As Long
    Dim lng��� As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim lngLeft As Long

    On Error GoTo errH

    On Error Resume Next
    
    '��ȡ���ı�ǩ������ڶ���
    For i = 1 To lblInfo.Count - 1
        If Val(Mid(lblInfo(i).Tag, 1, 1)) = 2 Then
            If lblInfo(i).Width > lngTmp2 Then lngTmp2 = lblInfo(i).Width
        Else
            If lblInfo(i).Width > lngTmp1 Then lngTmp1 = lblInfo(i).Width
        End If
    Next
    
    '���������
    If lngTmp1 = 0 Then lngTmp1 = lblPatiInfo(idxסԺ��).Width
    If lngTmp2 = 0 Then lngTmp2 = lblPatiInfo(idx��������).Width
    mCtl����.������1 = picPanel.Left + lngTmp1
    mCtl����.������2 = picPanel.Left + picPanel.Width / 2 + lngTmp2 + 200
    
     
    '��������Ϣ�ؼ�����
    lngLeft = picPanel.Width / 3
'    '��һ��
    lblPatiInfo(idx����).Left = mCtl����.������1 - picPanel.Left - lblPatiInfo(idx����).Width
    txtPatiInfo(idx����).Left = mCtl����.������1 - picPanel.Left + con���: txtPatiInfo(idx����).Width = picPanel.Width / 2 - txtPatiInfo(idx����).Left - 10
    lblPatiInfo(idx��������).Left = mCtl����.������2 - picPanel.Left - lblPatiInfo(idx��������).Width
    txtPatiInfo(idx��������).Left = mCtl����.������2 - picPanel.Left + con���: txtPatiInfo(idx��������).Width = picPanel.Width - txtPatiInfo(idx��������).Left - 10
    
    cmdFind.Top = txtPatiInfo(idx����).Top + 20: cmdFind.Left = txtPatiInfo(idx����).Left + txtPatiInfo(idx����).Width - cmdFind.Width - 15
    
    cmdType.Top = txtPatiInfo(idx��������).Top + 20: cmdType.Left = txtPatiInfo(idx��������).Left + txtPatiInfo(idx��������).Width - cmdType.Width - 15
    
    '�ڶ���
    lblPatiInfo(idx�Ա�).Left = mCtl����.������1 - picPanel.Left - lblPatiInfo(idx�Ա�).Width
    txtPatiInfo(idx�Ա�).Left = lblPatiInfo(idx�Ա�).Left + lblPatiInfo(idx�Ա�).Width + con���: txtPatiInfo(idx�Ա�).Width = lngLeft - txtPatiInfo(idx�Ա�).Left - 10
    lblPatiInfo(idx����).Left = lngLeft + lblPatiInfo(idx��Ժ;��).Width + 100 - lblPatiInfo(idx����).Width
    txtPatiInfo(idx����).Left = lblPatiInfo(idx����).Left + lblPatiInfo(idx����).Width + con���: txtPatiInfo(idx����).Width = lngLeft * 2 - txtPatiInfo(idx����).Left - 10
    lblPatiInfo(idx����).Left = lngLeft * 2 + lblPatiInfo(idx��Ժʱ��).Width + 100 - lblPatiInfo(idx����).Width
    txtPatiInfo(idx����).Left = lblPatiInfo(idx����).Left + lblPatiInfo(idx����).Width + con���: txtPatiInfo(idx����).Width = picPanel.Width - txtPatiInfo(idx����).Left - 10

    '������
    lblPatiInfo(idxסԺ��).Left = mCtl����.������1 - picPanel.Left - lblPatiInfo(idxסԺ��).Width
    txtPatiInfo(idxסԺ��).Left = lblPatiInfo(idxסԺ��).Left + lblPatiInfo(idxסԺ��).Width + con���: txtPatiInfo(idxסԺ��).Width = lngLeft - txtPatiInfo(idxסԺ��).Left - 10
    lblPatiInfo(idx��Ժ;��).Left = lngLeft + 100
    txtPatiInfo(idx��Ժ;��).Left = lblPatiInfo(idx��Ժ;��).Left + lblPatiInfo(idx��Ժ;��).Width + con���: txtPatiInfo(idx��Ժ;��).Width = lngLeft * 2 - txtPatiInfo(idx��Ժ;��).Left - 10
    lblPatiInfo(idx��Ժʱ��).Left = lngLeft * 2 + 100
    txtPatiInfo(idx��Ժʱ��).Left = lblPatiInfo(idx��Ժʱ��).Left + lblPatiInfo(idx��Ժʱ��).Width + con���: txtPatiInfo(idx��Ժʱ��).Width = picPanel.Width - txtPatiInfo(idx��Ժʱ��).Left - 10

    'Ĭ�ϲ��ֱ���
    mCtl����.S������ = mCtl����.lngTop + 80
    mCtl����.B������ = 0
    mCtl����.A������ = 0
    mCtl����.R������ = 0
    mCtl����.LngY = 0
    
    '���ñ�ǩλ��
    For i = 1 To lblInfo.Count - 1
        lng��� = 55
        If lblInfo(i).Height > lblPatiInfo(idx����).Height Then
            If InStr("," & mstrPicCtl & ",", "," & i & ",") = 0 And InStr("," & mstrTextCtl & ",", "," & i & ",") > 0 Then
                If txtInfo(i).Height < lblInfo(i).Height Then
                    lng��� = txtInfo(i).Height / 2 - lblInfo(i).Height / 2
                End If
            End If
        End If
        
        If InStr("," & mstrPicCtl & ",", "," & i & ",") > 0 Then
            lng��� = 25
        End If
        

        If Val(Mid(lblInfo(i).Tag, 1, 1)) = 2 Then '���õڶ���
            lblInfo(i).Top = lblInfo(i - 1).Top
            lblInfo(i).Left = mCtl����.������2 - lblInfo(i).Width
            
            '����ѡ���
            If InStr("," & mstrPicCtl & ",", "," & i & ",") > 0 Then
                picTmp(i).Top = lblInfo(i).Top - lng���: picTmp(i).Left = mCtl����.������2 + con���
                picTmp(i).Width = picPanel.Left + picPanel.Width - picTmp(i).Left: picTmp(i).Height = 290
                
                
                lngBegin = Val(Split(picTmp(i).Tag, ",")(1))
                lngEnd = Val(Split(picTmp(i).Tag, ",")(2))
                
                '����Ӧ����ѡ���
                If Val(Mid(picTmp(i).Tag, 1, 1)) = 2 Then
                    For j = lngBegin To lngEnd
                        If j = lngBegin Then
                            optInfo(j).Top = 0: optInfo(j).Left = 0
                        Else
                            '������������ʱ���Զ�����
                            If optInfo(j - 1).Left + optInfo(j - 1).Width + 50 + optInfo(j).Width > picTmp(i).Width Then
                                picTmp(i).Height = picTmp(i).Height + 60 + optInfo(j).Height
                                optInfo(j).Top = optInfo(j - 1).Top + optInfo(j - 1).Height + 50: optInfo(j).Left = 0
                            Else
                                optInfo(j).Top = optInfo(j - 1).Top: optInfo(j).Left = optInfo(j - 1).Left + optInfo(j - 1).Width + 50
                            End If
                        End If
                    Next
                Else
                    For j = lngBegin To lngEnd
                        If j = lngBegin Then
                            chkInfo(j).Top = 0: chkInfo(j).Left = 0
                        Else
                            '������������ʱ���Զ�����
                            If chkInfo(j - 1).Left + chkInfo(j - 1).Width + 50 + chkInfo(j).Width > picTmp(i).Width Then
                                picTmp(i).Height = picTmp(i).Height + 60 + chkInfo(j).Height
                                chkInfo(j).Top = chkInfo(j - 1).Top + chkInfo(j - 1).Height + 50: chkInfo(j).Left = 0
                            Else
                                chkInfo(j).Top = chkInfo(j - 1).Top: chkInfo(j).Left = chkInfo(j - 1).Left + chkInfo(j - 1).Width + 50
                            End If
                        End If
                    Next
                End If
                
                
                 If lngMaxTop < picTmp(i).Top + picTmp(i).Height Then lngMaxTop = picTmp(i).Top + picTmp(i).Height
                 If lngMaxTop < lblInfo(i).Top + lblInfo(i).Height Then lngMaxTop = lblInfo(i).Top + lblInfo(i).Height
                
                '�����ı���
                If InStr("," & mstrTextCtl & ",", "," & i & ",") > 0 And txtInfo(i).Visible Then
                    txtInfo(i).Top = picTmp(i).Top + picTmp(i).Height + 20: txtInfo(i).Left = picTmp(i).Left
                    txtInfo(i).Width = picTmp(i).Width
                    
                    If lngMaxTop < txtInfo(i).Top + txtInfo(i).Height Then lngMaxTop = txtInfo(i).Top + txtInfo(i).Height
                    If lngMaxTop < lblInfo(i).Top + lblInfo(i).Height Then lngMaxTop = lblInfo(i).Top + lblInfo(i).Height
                End If
                
                
            End If
            
            '�����ı���
            If InStr("," & mstrTextCtl & ",", "," & i & ",") > 0 And InStr("," & mstrPicCtl & ",", "," & i & ",") = 0 Then
                txtInfo(i).Top = lblInfo(i).Top - lng���: txtInfo(i).Left = mCtl����.������2 + con���
                txtInfo(i).Width = picPanel.Left + picPanel.Width - txtInfo(i).Left
                
                If lngMaxTop < txtInfo(i).Top + txtInfo(i).Height Then lngMaxTop = txtInfo(i).Top + txtInfo(i).Height
                If lngMaxTop < lblInfo(i).Top + lblInfo(i).Height Then lngMaxTop = lblInfo(i).Top + lblInfo(i).Height
            End If
            
            '����������
            Select Case Split(lblInfo(i).Tag, ",")(2)
                Case "S"
                    mCtl����.S������ = lngMaxTop + 150
                Case "B"
                    mCtl����.B������ = lngMaxTop + 150
                Case "A"
                    mCtl����.A������ = lngMaxTop + 150
                Case "R"
                    mCtl����.R������ = lngMaxTop + 150
            End Select
        Else '���õ�һ��
        
            '��̬���ָ߶ȼ���
            If i = 1 Then
                lngTop = mCtl����.lngTop + IIf(Split(lblInfo(i).Tag, ",")(2) = "S", 200, 300)
            Else
                lngTop = lngMaxTop + 200
            End If
            
            If i <> 1 Then
                If Split(lblInfo(i - 1).Tag, ",")(2) <> Split(lblInfo(i).Tag, ",")(2) Then
                    lngTop = lngMaxTop + 300
                End If
            End If

            lblInfo(i).Top = lngTop
            lblInfo(i).Left = mCtl����.������1 - lblInfo(i).Width
            
            '����ѡ���
            If InStr("," & mstrPicCtl & ",", "," & i & ",") > 0 Then
                picTmp(i).Top = lblInfo(i).Top - lng���: picTmp(i).Left = mCtl����.������1 + con���
                picTmp(i).Width = (picPanel.Left + picPanel.Width) / Decode(Val(Mid(lblInfo(i).Tag, 1, 1)), 0, 1, 2) - picTmp(i).Left
                picTmp(i).Height = 290
                
                lngBegin = Val(Split(picTmp(i).Tag, ",")(1))
                lngEnd = Val(Split(picTmp(i).Tag, ",")(2))
                
                '����Ӧ����ѡ���
                If Val(Mid(picTmp(i).Tag, 1, 1)) = 2 Then
                    For j = lngBegin To lngEnd
                        If j = lngBegin Then
                            optInfo(j).Top = 0: optInfo(j).Left = 0
                        Else
                            '������������ʱ���Զ�����
                            If optInfo(j - 1).Left + optInfo(j - 1).Width + 50 + optInfo(j).Width > picTmp(i).Width Then
                                picTmp(i).Height = picTmp(i).Height + 60 + optInfo(j).Height
                                optInfo(j).Top = optInfo(j - 1).Top + optInfo(j - 1).Height + 50: optInfo(j).Left = 0
                            Else
                                optInfo(j).Top = optInfo(j - 1).Top: optInfo(j).Left = optInfo(j - 1).Left + optInfo(j - 1).Width + 50
                            End If
                        End If
                    Next
                Else
                    For j = lngBegin To lngEnd
                        If j = lngBegin Then
                            chkInfo(j).Top = 0: chkInfo(j).Left = 0
                        Else
                            '������������ʱ���Զ�����
                            If chkInfo(j - 1).Left + chkInfo(j - 1).Width + 50 + chkInfo(j).Width > picTmp(i).Width Then
                                picTmp(i).Height = picTmp(i).Height + 60 + chkInfo(j).Height
                                chkInfo(j).Top = chkInfo(j - 1).Top + chkInfo(j - 1).Height + 50: chkInfo(j).Left = 0
                            Else
                                chkInfo(j).Top = chkInfo(j - 1).Top: chkInfo(j).Left = chkInfo(j - 1).Left + chkInfo(j - 1).Width + 50
                            End If
                        End If
                    Next
                End If
                
                
                 If lngMaxTop < picTmp(i).Top + picTmp(i).Height Then lngMaxTop = picTmp(i).Top + picTmp(i).Height
                 If lngMaxTop < lblInfo(i).Top + lblInfo(i).Height Then lngMaxTop = lblInfo(i).Top + lblInfo(i).Height
                
                '�����ı���
                If InStr("," & mstrTextCtl & ",", "," & i & ",") > 0 And txtInfo(i).Visible Then
                    txtInfo(i).Top = picTmp(i).Top + picTmp(i).Height + 20: txtInfo(i).Left = picTmp(i).Left
                    txtInfo(i).Width = picTmp(i).Width
                    
                    If lngMaxTop < txtInfo(i).Top + txtInfo(i).Height Then lngMaxTop = txtInfo(i).Top + txtInfo(i).Height
                    If lngMaxTop < lblInfo(i).Top + lblInfo(i).Height Then lngMaxTop = lblInfo(i).Top + lblInfo(i).Height
                End If
                
                
            End If
            
            '�����ı���
            If InStr("," & mstrTextCtl & ",", "," & i & ",") > 0 And InStr("," & mstrPicCtl & ",", "," & i & ",") = 0 Then
                txtInfo(i).Top = lblInfo(i).Top - lng���: txtInfo(i).Left = mCtl����.������1 + con���
                txtInfo(i).Width = (picPanel.Left + picPanel.Width) / Decode(Val(Mid(lblInfo(i).Tag, 1, 1)), 0, 1, 2) - txtInfo(i).Left
                
                If lngMaxTop < txtInfo(i).Top + txtInfo(i).Height Then lngMaxTop = txtInfo(i).Top + txtInfo(i).Height
                If lngMaxTop < lblInfo(i).Top + lblInfo(i).Height Then lngMaxTop = lblInfo(i).Top + lblInfo(i).Height
            End If
            
            '����������
            Select Case Split(lblInfo(i).Tag, ",")(2)
                Case "S"
                    mCtl����.S������ = lngMaxTop + 150
                Case "B"
                    mCtl����.B������ = lngMaxTop + 150
                Case "A"
                    mCtl����.A������ = lngMaxTop + 150
                Case "R"
                    mCtl����.R������ = lngMaxTop + 150
            End Select
        End If
    Next
    
    mCtl����.LngY = lngMaxTop + 200

    Call DrawLine
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub optInfo_Click(Index As Integer)
    If (mInfo.EditType <> 0 And mblnLoad = False) Or mblnԤ�� Then
        mblnChange = True
        Call SetTextVisible(Val(Split(optInfo(Index).Tag, ",")(0)))
        Call MakeText
    End If
End Sub

Private Sub optInfo_GotFocus(Index As Integer)
    If mblnEnter Then Call ShowCtl(optInfo(Index)): mblnEnter = False
End Sub

Private Sub picEdit_Resize()
    On Error Resume Next
    picPanel.Width = picEdit.Width - picPanel.Left - 150
End Sub

Private Sub picMain_Resize()
    Call RefreshResize
End Sub

Public Sub RefreshResize()
    On Error Resume Next
    vscBar.Top = 0: vscBar.Height = picMain.Height
    vscBar.Left = picMain.Width - vscBar.Width
    

    '���������߶�
    If picEdit.Height < picMain.Height Then picEdit.Height = picMain.Height
    vscBar.Visible = picMain.Height < mCtl����.LngY
    picEdit.Width = IIf(vscBar.Visible, picMain.Width - vscBar.Width, picMain.Width)
    Call CtlResize
    picEdit.Height = mCtl����.LngY
    If picEdit.Height < picMain.Height Then picEdit.Height = picMain.Height
    
    '�л���������ʾʱ����ˢ��
    If IIf(vscBar.Visible, 1, 0) <> IIf(picMain.Height < mCtl����.LngY, 1, 0) Then
        vscBar.Visible = picMain.Height < picEdit.Height
        picEdit.Width = IIf(vscBar.Visible, picMain.Width - vscBar.Width, picMain.Width)
        Call CtlResize
        picEdit.Height = mCtl����.LngY
        If picEdit.Height < picMain.Height Then picEdit.Height = picMain.Height
    End If
    
    vscBar.Max = picEdit.Height - picMain.Height
End Sub


Private Sub picInfo_GotFocus()
    On Error Resume Next
    If zlcontrol.IsCtrlSetFocus(txtPatiInfo(idx����)) Then
        txtPatiInfo(idx����).SetFocus
    Else
        Call SeekNextCtl
    End If
End Sub

Private Sub rtbBox_Change()
    If rtbBox.Visible And rtbBox.Locked = False And mInfo.EditType <> 0 And mblnLoad = False Then mblnChange = True
End Sub

Private Sub rtbBox_KeyPress(KeyAscii As Integer)
    If InStr("&'<>", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub


Private Sub rtbBox_LostFocus()
    If rtbBox.Locked = False And rtbBox.Text <> rtbBox.Tag Then
        With rtbBox
            .SelStart = 0
            .SelLength = Len(rtbBox.Text)
            .SelColor = RGB(30, 144, 255)
            .SelStart = Len(rtbBox.Text)
        End With
    End If
End Sub

Private Sub txtPatiInfo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index <> idx���� Then
        Call zlCommFun.ShowTipInfo(txtPatiInfo(Index).hWnd, txtPatiInfo(Index).Text, True, True)
    End If
End Sub

Private Sub txtPatiInfo_Validate(Index As Integer, Cancel As Boolean)
    On Error GoTo errH
    If Index = idx���� Then
        If txtPatiInfo(Index).Visible And txtPatiInfo(Index).Locked = False Then
            txtPatiInfo(Index).Text = txtPatiInfo(Index).Tag
        End If
    End If
    
    If txtPatiInfo(Index).Enabled = True And txtPatiInfo(Index).Locked = False Then
        Call MakeText
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vscBar_Change()
    On Error Resume Next
    picEdit.Top = -vscBar.Value
    picEdit.SetFocus
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    fraInfo.Top = 0: fraInfo.Left = 0: fraInfo.Height = picInfo.Height - 25: fraInfo.Width = picInfo.Width - 25
    rtbBox.Height = fraInfo.Height - 300
    rtbBox.Width = fraInfo.Width - 200
End Sub


Private Sub picEdit_Paint()
    Call DrawLine
End Sub

Private Sub DrawLine()
    Dim lngUp As Long
    '�����
    On Error Resume Next

    picEdit.Cls
    picEdit.Line (con�����X, 0)-(con�����X, picEdit.Height)
    picEdit.ForeColor = RGB(105, 105, 105)
    
    '��������Ϊ��ʱ��SBAR��
    If mInfo.�������� = "" And mCtl����.S������ > 0 Then
        mCtl����.B������ = mCtl����.S������ + (picEdit.Height - mCtl����.S������) / 3
        mCtl����.A������ = mCtl����.S������ + ((picEdit.Height - mCtl����.S������) / 3) * 2
        mCtl����.R������ = mCtl����.S������ + ((picEdit.Height - mCtl����.S������) / 3) * 2
    End If
    
    If mCtl����.S������ > 0 Then
        If mCtl����.R������ = 0 And mCtl����.A������ = 0 And mCtl����.B������ = 0 Then
            picEdit.FontName = "����": picEdit.FontSize = 13: picEdit.FontBold = True
            picEdit.CurrentX = 150
            picEdit.CurrentY = lngUp + (picEdit.Height - lngUp) / 2 - 100
            picEdit.Print "S"
        Else
            lngUp = mCtl����.S������
            picEdit.Line (0, mCtl����.S������)-(picEdit.Width, mCtl����.S������)
            picEdit.FontName = "����": picEdit.FontSize = 13: picEdit.FontBold = True
            picEdit.CurrentX = 150
            picEdit.CurrentY = (mCtl����.S������) / 2 - 100
            picEdit.Print "S"
        End If
    End If
    
    If mCtl����.B������ > 0 Then
        If mCtl����.R������ = 0 And mCtl����.A������ = 0 Then
            picEdit.FontName = "����": picEdit.FontSize = 13: picEdit.FontBold = True
            picEdit.CurrentX = 150
            picEdit.CurrentY = lngUp + (picEdit.Height - lngUp) / 2 - 100
            picEdit.Print "B"
        Else
            picEdit.Line (0, mCtl����.B������)-(picEdit.Width, mCtl����.B������)
            
            picEdit.FontName = "����": picEdit.FontSize = 13: picEdit.FontBold = True
            picEdit.CurrentX = 150
            picEdit.CurrentY = mCtl����.B������ - (mCtl����.B������ - lngUp) / 2 - 100
            picEdit.Print "B"
            
            lngUp = mCtl����.B������
        End If
    End If
    
    If mCtl����.A������ > 0 Then
        If mCtl����.R������ = 0 Then
            picEdit.FontName = "����": picEdit.FontSize = 13: picEdit.FontBold = True
            picEdit.CurrentX = 150
            picEdit.CurrentY = lngUp + (picEdit.Height - lngUp) / 2 - 100
            picEdit.Print "A"
            
        Else
            picEdit.Line (0, mCtl����.A������)-(picEdit.Width, mCtl����.A������)
            
            picEdit.FontName = "����": picEdit.FontSize = 13: picEdit.FontBold = True
            picEdit.CurrentX = 150
            picEdit.CurrentY = mCtl����.A������ - (mCtl����.A������ - lngUp) / 2 - 100
            picEdit.Print "A"
            lngUp = mCtl����.A������
        End If
    End If
    
    If mCtl����.R������ > 0 Then
        picEdit.FontName = "����": picEdit.FontSize = 13: picEdit.FontBold = True
        picEdit.CurrentX = 150
        picEdit.CurrentY = lngUp + (picEdit.Height - lngUp) / 2 - 100
        picEdit.Print "R"
    End If
    
    picEdit.Refresh
End Sub


Private Sub GetCtlRs()
    '���ܣ���ȡ�ؼ���Ϣ��¼��
    Dim strSQL As String

    On Error GoTo errH
    '�������ͼ�¼��
    strSQL = "Select a.���, a.����, a.˳��, a.��ʼ����, a.��ȡsql From ҽ�����Ӱಡ������ A Where a.�Ƿ�ͣ�� = 0 order by A.˳��"
    Set mrsCtlType = zlDatabase.OpenSQLRecord(strSQL, "GetCtlRs")
    
    '���没�������ַ���
    If Not mrsCtlType Is Nothing Then
        mstrPatiType = ""
        Do While Not mrsCtlType.EOF
            mstrPatiType = mstrPatiType & "," & mrsCtlType!���
            mrsCtlType.MoveNext
        Loop
        mstrPatiType = mstrPatiType & ","
        mrsCtlType.MoveFirst
    End If
        
    '�ؼ���Ϣ��¼��
    strSQL = "select ���˼��, ��Ŀ����, ���, ��Ŀ���, ������ʽ, ��������, �����ʽ, ����ֵ��, ��������, ��ȡ��Դ, ��ȡ����, ��ȡSQL, ��������, �Ƿ�ֻ��, ���������� from ҽ�����Ӱಡ����Ŀ  order by ���˼��,��Ŀ���,���,Rownum"
    Set mrsCtlInfo = zlDatabase.OpenSQLRecord(strSQL, "GetCtlRs")
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtInfo_Change(Index As Integer)
    If txtInfo(Index).Visible And txtInfo(Index).Locked = False And mInfo.EditType <> 0 And mblnLoad = False Then mblnChange = True
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    zlcontrol.TxtSelAll txtInfo(Index)
    If mblnEnter Then Call ShowCtl(txtInfo(Index)): mblnEnter = False
End Sub


Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strMask As String
    On Error GoTo errH
    
    If InStr("&'<>", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call SeekNextCtl
    ElseIf Not (KeyAscii >= 0 And KeyAscii < 32) Then
        Select Case Val(Split(lblInfo(Index).Tag, ",")(3))
            Case 2
                strMask = "1234567890."
        End Select

        If InStr(strMask, Chr(KeyAscii)) = 0 And strMask <> "" Then
            KeyAscii = 0: Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
    Dim strMsg As String
    
    On Error GoTo errH
    
    If txtInfo(Index).Text <> "" Then
        Select Case Val(Split(lblInfo(Index).Tag, ",")(3))
            Case 1 '����
                 txtInfo(Index).Text = Format(zlStr.FullDate(txtInfo(Index).Text), "yyyy-MM-dd HH:mm")
                 If Not IsDate(txtInfo(Index).Text) Then
                     strMsg = "ʱ���ʽ����ȷ,������¼�롣"
                 End If
            Case 2 '����
                 If Not IsNumeric(txtInfo(Index).Text) Then
                     strMsg = "���ָ�ʽ����ȷ,������¼�롣"
                 End If
        End Select
        
        If strMsg <> "" Then
            Set mobjCtl = txtInfo(Index)
            MsgBox strMsg, vbInformation, Me.Caption
            zlcontrol.TxtSelAll txtInfo(Index)
            zlcontrol.ControlSetFocus txtInfo(Index)
            Cancel = True
            Exit Sub
        End If
    End If
    
    '�������
    If InStr(txtInfo(Index).Text, vbCrLf) > 0 Then txtInfo(Index).Text = Replace(txtInfo(Index).Text, vbCrLf, "")

    If txtInfo(Index).Tag <> txtInfo(Index).Text And txtInfo(Index).Enabled = True And txtInfo(Index).Locked = False Then
        Call MakeText
    End If
    txtInfo(Index).Tag = txtInfo(Index).Text
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtPatiInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo errH
    
    If InStr("&'<>", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii = 13 Then
        Select Case Index
            Case idx����
                If txtPatiInfo(Index).Visible And txtPatiInfo(Index).Locked = False And txtPatiInfo(Index).Text <> txtPatiInfo(Index).Tag And txtPatiInfo(Index).Text <> "" Then
                    Call GetPatiList(0)
                Else
                     KeyAscii = 0
                     Call SeekNextCtl
                End If
            Case Else
                KeyAscii = 0
                Call SeekNextCtl
        End Select
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtPatiInfo_GotFocus(Index As Integer)
    If mblnEnter Then Call ShowCtl(txtPatiInfo(Index)): mblnEnter = False
    zlcontrol.TxtSelAll txtPatiInfo(Index)
End Sub

Private Sub optInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SeekNextCtl
    End If
End Sub

Private Sub chkInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SeekNextCtl
    End If
End Sub

Private Function SeekNextCtl() As Boolean
'���ܣ���λ����һ������Ŀؼ���
    Call zlCommFun.PressKey(vbKeyTab)
    mblnEnter = True
    SeekNextCtl = True
End Function

Private Sub Subclass_WndProc(Msg As Long, wParam As Long, lParam As Long, Result As Long)
    '�Զ������Ϣ������
    Dim tP As POINTAPI
    Dim sngX As Single, sngY As Single   '�������
    Dim intShift As Integer              '��갴��
    Dim bWay As Boolean                  '��귽��
    Dim bMouseFlag As Boolean            '����¼������־
    Dim wzDelta, wKeys As Integer
    Select Case Msg
        Case WM_MOUSEWHEEL   '����
            wzDelta = (wParam And &HFFFF0000) \ &H10000 'ȡ��32λֵ�ĸ�16λ
            If wzDelta > 0 Then
                vscBar.Value = IIf(vscBar.Value - vscBar.LargeChange < 0, 0, vscBar.Value - vscBar.LargeChange)
            Else
                vscBar.Value = IIf(vscBar.Value + vscBar.LargeChange > vscBar.Max, vscBar.Max, vscBar.Value + vscBar.LargeChange)
            End If
    End Select
End Sub

Private Sub picSplitX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglNew As Single
    
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    
    If picSplitX.Tag <> "Draging" Then
        picSplitX.Tag = "Draging"
        picSplitX.BackColor = 0
    End If
    
    sglNew = picSplitX.Left + X
    
    picSplitX.Left = sglNew
End Sub

Private Sub picSplitX_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    If picSplitX.Tag = "Draging" Then
        Call Form_Resize
        picSplitX.BackColor = Me.BackColor
        picSplitX.Tag = ""
    End If
End Sub

Private Sub SetPicEnabled(ByVal blnEnabled As Boolean)
    '���ÿؼ��Ƿ���ñ���ɫ
    Dim obj As Object
    
    picEdit.Enabled = blnEnabled
    For Each obj In txtInfo
        If obj.Index <> 0 And obj.Locked = False Then
            obj.BackColor = IIf(blnEnabled, &H80000005, &H80000004)
        End If
    Next
    
    picMainBack.BackColor = IIf(blnEnabled, &H80000005, &H80000004)
    
    '������������
    fraInfo.BackColor = IIf(blnEnabled, &H80000005, &H80000004)
    picSplitX.BackColor = IIf(blnEnabled, &H80000005, &H80000004)
    Me.BackColor = IIf(blnEnabled, &H80000005, &H80000004)
    rtbBox.Locked = Not blnEnabled
    rtbBox.BackColor = IIf(blnEnabled, &H80000005, &H80000004)
End Sub


Public Sub zlRefresh(ByVal lngPatiID As Long, ByVal lngPageID As Long, _
    ByVal lngDeptID As Long, ByVal lngDataID As Long, ByVal lng��¼ID As Long, dtBegin As Date, dtEnd As Date, btnNoEdit As Boolean, ByVal str�������� As String)
    On Error GoTo errH
    '��������
    mInfo.EditType = 0
    mInfo.����ID = lngPatiID
    mInfo.��ҳID = lngPageID
    mInfo.�������ID = lngDeptID
    mInfo.����ID = lngDataID
    mInfo.�����¼ID = lng��¼ID
    mInfo.���࿪ʼʱ�� = dtBegin
    mInfo.�������ʱ�� = dtEnd
    mInfo.�������� = str��������
    
    '��ʱ�������
    mbtnNoEdit = btnNoEdit
    mblnSave = False
    mblnChange = False
    mblnEnter = False
    Set mobjCtl = Nothing
    
    '���¼��ؿؼ�
    picEdit.Enabled = True
    Call ClearData
    Call UnloadCtl
    
    '�Զ��������
    If mInfo.����ID <> 0 Then Call LoadData
    
    Call InitCtl
    Call CtlEnabled
    Call SetPicEnabled(False)
    Call DrawLine '����ˢ�½���
    
    
    If mInfo.����ID <> 0 Then 'Ԥ��

        'δ�������Զ���ȡ����
        If rtbBox.Text = "" Then
            Call IntData
            Call MakeText
        Else
            '��ȡ�ѱ��������
           Call ReadData
        End If
    Else
        Call MakeText
    End If
    
    cmdType.Visible = False
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadData()
    '��ʼ����ȡ����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String

    
    On Error GoTo errH
    If mInfo.����ID = 0 Then Exit Sub
    '��ȡ�ѱ��������
    strSQL = "select ����Id, ��¼id, ���, ��������, ����id, ��ҳid, ����, �Ա�, ����, ����, ��ʶ��, ��Ժʱ��, ��Ժ��ʽ, �������� from ҽ�����Ӱ����� where ����id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mInfo.����ID)
    
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            mInfo.������� = Val(rsTmp!��� & "")
            mInfo.����ID = Val(rsTmp!����ID & "")
            mInfo.��ҳID = Val(rsTmp!��ҳID & "")
            
            txtPatiInfo(idx����).Text = rsTmp!���� & ""
            txtPatiInfo(idx�Ա�).Text = rsTmp!�Ա� & ""
            txtPatiInfo(idx����).Text = rsTmp!���� & ""
            txtPatiInfo(idx����).Text = rsTmp!���� & ""
            txtPatiInfo(idxסԺ��).Text = rsTmp!��ʶ�� & ""
            txtPatiInfo(idx��Ժ;��).Text = rsTmp!��Ժ��ʽ & ""
            txtPatiInfo(idx��Ժʱ��).Text = Format(rsTmp!��Ժʱ�� & "", "yyyy-MM-dd HH:mm")
            
            txtPatiInfo(idx��������).Text = rsTmp!�������� & ""
            mInfo.�������� = rsTmp!�������� & ""

            rtbBox.Text = rsTmp!�������� & ""
            rtbBox.Tag = rsTmp!�������� & ""
            
            With rtbBox
                .SelStart = 0
                .SelLength = Len(rtbBox.Text)
                .SelColor = vbBlack
                .SelStart = Len(rtbBox.Text)
            End With
            
            mblnSave = rsTmp!�������� & "" <> ""
 
            'δ����ʱ֧���޸Ĳ�������
            If mInfo.EditType = 2 And rsTmp!�������� & "" = "" Then cmdType.Visible = True
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ClearData()
'���ܣ���ս�������
    Dim obj As Object

    On Error GoTo errH
    For Each obj In txtPatiInfo
        obj.Text = ""
        obj.Tag = ""
    Next
    
    rtbBox.Text = ""
    rtbBox.Tag = ""
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub GetPatiList(ByVal intType As Integer)
'���ܣ���ȡ��ǰ���ҵĲ���
'������0 �ı��򰴻س���1 �㰴ť
    Dim strSQL As String, rsTmp As Recordset, rsTmp1 As Recordset
    Dim strInput As String, vRect As RECT
    Dim blnCancel As Boolean
    Dim blnDo As Boolean
    
    On Error GoTo errH
    
    If intType = 0 Then
        If txtPatiInfo(idx����).Tag = txtPatiInfo(idx����).Text And txtPatiInfo(idx����).Text <> "" Then
            Call SeekNextCtl
            Exit Sub
        End If
        
        '¼����Ϊ��ʱ������ȫ��
        If txtPatiInfo(idx����).Text = "" Then intType = 1
    End If
            
    strInput = Trim(UCase(txtPatiInfo(idx����).Text))   '�����ֵ����ǰ׺�ո�

    strSQL = "Select b.����ID as ID,A.��ҳID,b.����, b.�Ա�, b.����, b.��ǰ���� As ����, a.סԺ��, a.��Ժ��ʽ, a.��Ժ����" & vbNewLine & _
            "From ������ҳ A, ������Ϣ B, ��Ժ���� C" & vbNewLine & _
            "Where c.����id = a.����id And c.��ҳid = a.��ҳid And a.����id = b.����id And C.����id = [1] And a.�������� In(0,2)" & _
            IIf(intType = 0, " And (A.סԺ�� = [2] Or A.���� Like [3] or b.��ǰ���� like [3])", "") & " ORDER BY a.��Ժ���� desc"
        
    vRect = zlcontrol.GetControlRect(txtPatiInfo(idx����).hWnd)
    Set mobjCtl = txtPatiInfo(idx����)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "�����Ҳ����б�", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtPatiInfo(idx����).Height, blnCancel, False, True, mInfo.�������ID, Val(strInput), mstrLike & strInput & "%", CDate(mInfo.���࿪ʼʱ��), CDate(mInfo.�������ʱ��))
    
    zlcontrol.ControlSetFocus txtPatiInfo(idx����)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            Set mobjCtl = txtPatiInfo(idx����)
            Call MsgBox("û���ڵ�ǰ��������ҵ�ƥ��Ĳ���!", vbInformation, gstrSysName)
        End If
        txtPatiInfo(idx����).Text = txtPatiInfo(idx����).Tag
        blnDo = False
    Else
        If Not rsTmp.EOF Then
            blnDo = True
        Else
            txtPatiInfo(idx����).Text = txtPatiInfo(idx����).Tag
            blnDo = False
        End If
    End If
    
    If blnDo Then

        '�жϵ�ǰ�����Ƿ����
        strSQL = "select ����Id,��������,���� from ҽ�����Ӱ����� where ��¼id=[1] and ����ID=[2] AND ��ҳID=[3]"
        Set rsTmp1 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mInfo.�����¼ID, Val(rsTmp!id & ""), Val(rsTmp!��ҳID & ""))

        If Not rsTmp1 Is Nothing Then
            If Not rsTmp1.EOF Then
                 Set mobjCtl = txtPatiInfo(idx����)
                 Call MsgBox("�ڵ�ǰ�����¼���Ѵ��ڵ�ǰѡ��Ĳ���,������ѡ��!", vbInformation, gstrSysName)
                 zlcontrol.ControlSetFocus txtPatiInfo(idx����)
                 txtPatiInfo(idx����).Text = txtPatiInfo(idx����).Tag
                 Call zlcontrol.TxtSelAll(txtPatiInfo(idx����))
                 Exit Sub
            End If
        End If

        '�����������
        Call ClearData

        '���ز�����Ϣ
        mInfo.����ID = rsTmp!id & ""
        mInfo.��ҳID = rsTmp!��ҳID & ""
        mInfo.�������� = GetPatiType(Val(rsTmp!id & ""), Val(rsTmp!��ҳID & ""))
        txtPatiInfo(idx��������).Text = mInfo.��������
        txtPatiInfo(idx����).Text = rsTmp!���� & ""
        txtPatiInfo(idx����).Tag = rsTmp!���� & ""
        txtPatiInfo(idx�Ա�).Text = rsTmp!�Ա� & ""
        txtPatiInfo(idx����).Text = rsTmp!���� & ""
        txtPatiInfo(idx����).Text = rsTmp!���� & ""
        txtPatiInfo(idxסԺ��).Text = rsTmp!סԺ�� & ""
        txtPatiInfo(idx��Ժ;��).Text = rsTmp!��Ժ��ʽ & ""
        txtPatiInfo(idx��Ժʱ��).Text = Format(rsTmp!��Ժ���� & "", "yyyy-MM-dd HH:mm")
        
        Call UnloadCtl
        Call InitCtl
        Call IntData '�Զ��������
        Call MakeText

        zlcontrol.ControlSetFocus txtPatiInfo(idx����)
        Call SeekNextCtl
    Else
        zlcontrol.ControlSetFocus txtPatiInfo(idx����)
        Call zlcontrol.TxtSelAll(txtPatiInfo(idx����))
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub cbsExec_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim str���� As String
    Dim objControl As CommandBarControl
    On Error GoTo errH
    Select Case Control.id
        Case ID_����
            If Not gobjPublicAdvice Is Nothing Then
                If mInfo.����ID <> 0 And mInfo.��ҳID <> 0 Then
                    Call gobjPublicAdvice.ShowArchive(Me, mInfo.����ID, mInfo.��ҳID)
                End If
            End If
        Case ID_����
            '��Ŀ�ı���δ����MakeTextʱ
            If Me.ActiveControl.Name = "txtInfo" Then Call MakeText
            
            If CheckData Then
                Call SaveData
                Call LoadData
                Call ReadData
                'Ԥ��
                mInfo.EditType = 0
                Call SetPicEnabled(False)
                If Not gfrmParent Is Nothing Then
                    Call gfrmParent.SetEnable
                    Call gfrmParent.RefreshEdit(mInfo.����ID)
                End If
                mblnChange = False
                cmdType.Visible = False
            End If
        Case ID_ȡ��
            If mblnChange And mInfo.����ID <> 0 Then
                If MsgBox("��ǰ���������ѷ����ı䣬��ȷ���Ƿ�ȡ���༭��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
            Call SetPicEnabled(False)
            If Not gfrmParent Is Nothing Then
                Call gfrmParent.SetEnable
                If mInfo.EditType = 1 Then
                    Call gfrmParent.RefreshEdit(mInfo.����ID)
                ElseIf mInfo.EditType = 2 And mblnChange Then
                    Call zlRefresh(mInfo.����ID, mInfo.��ҳID, mInfo.�������ID, mInfo.����ID, mInfo.�����¼ID, mInfo.���࿪ʼʱ��, mInfo.�������ʱ��, mbtnNoEdit, mInfo.��������)
                End If
            End If
            mInfo.EditType = 0
            
            Call CtlEnabled
            cmdType.Visible = False
            
            mblnChange = False
        Case ID_����
            If mbtnNoEdit Then Exit Sub
            If mInfo.�����¼ID = 0 Then Exit Sub
            If Not gfrmParent Is Nothing Then Call gfrmParent.SetEnable(1)
            Call SetPicEnabled(True)
            cmdType.Visible = True
            Call ClearData
            mInfo.EditType = 1: mInfo.����ID = 0: mInfo.��ҳID = 0: mInfo.����ID = 0:  mInfo.������� = 0: mInfo.�������� = ""
            '���ÿؼ�״̬
            Call UnloadCtl
            Call InitCtl
            Call CtlEnabled
            
            mblnChange = False
            cmdType.Visible = True
            Call MakeText
        Case ID_�޸�
            Call EditState
            cmdType.Visible = Not mblnSave
            mblnChange = False
        Case ID_ɾ��
            If mbtnNoEdit Then Exit Sub
            If mInfo.�����¼ID = 0 Or mInfo.����ID = 0 Then Exit Sub
            If DelEdit Then
                Call gfrmParent.RefreshEdit(mInfo.����ID)
            End If
            mInfo.EditType = 0
            cmdType.Visible = False
            mblnChange = False
        Case Else
            If Control.id <> ID_���� Then 'Ԥ����������
                Call txtPatiInfo(idx����).SetFocus
                
                
                For Each objControl In cbsExec(2).Controls
                    If objControl.id = Control.id Then
                        objControl.IconId = IIf(objControl.IconId = 10, 12, 10)
                    End If
                    
                    If objControl.id > ID_���� And objControl.IconId = 12 Then
                        str���� = str���� & "," & objControl.Caption
                    End If
                Next


                txtPatiInfo(idx��������).Text = Mid(str����, 2)
                mInfo.�������� = txtPatiInfo(idx��������).Text
                Call UnloadCtl
                Call InitCtl
                Call IntData
                Call MakeText
            End If
            
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsExec_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
        Case ID_����
            Control.Visible = (Not mblnԤ��)
        Case ID_����
            Control.Visible = (Not mblnԤ��) And mInfo.EditType <> 0
        Case ID_ȡ��
            Control.Visible = (Not mblnԤ��) And mInfo.EditType <> 0
        Case ID_����
            Control.Visible = (Not mblnԤ��) And mInfo.EditType = 0 And (Not mbtnNoEdit)
        Case ID_�޸�
            Control.Visible = (Not mblnԤ��) And mInfo.EditType = 0 And mInfo.����ID <> 0 And (Not mbtnNoEdit)
        Case ID_ɾ��
            Control.Visible = (Not mblnԤ��) And mInfo.EditType = 0 And mInfo.����ID <> 0 And (Not mbtnNoEdit)
        Case Else
            Control.Visible = mblnԤ��
    End Select
End Sub


Public Sub EditState()
'���ܣ��޸�״̬
    Dim i As Long
    On Error GoTo errH
    If mInfo.�����¼ID = 0 Or mInfo.����ID = 0 Or mbtnNoEdit Then Exit Sub
    If Not gfrmParent Is Nothing Then Call gfrmParent.SetEnable(1)
    Call SetPicEnabled(True)
    '����ѡ�ؼ�������ת������
    mblnLoad = True
    
    For i = 1 To txtInfo.Count - 1
        If txtInfo(i).Enabled And txtInfo(i).Locked = False Then
            txtInfo(i).SetFocus
            Exit For
        End If
    Next
    
    mblnLoad = False
    mInfo.EditType = 2
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function DelEdit() As Boolean
    Dim strSQL As String
    Dim blnTran As Boolean
    
    On Error GoTo errH
    If MsgBox("ȷ��Ҫɾ��ѡ�еĲ��˽����¼��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
    strSQL = "Zl_ҽ�����Ӱ�����_Edit(2," & mInfo.����ID & ")"
    Screen.MousePointer = 11
    gcnOracle.BeginTrans: blnTran = True
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    gcnOracle.CommitTrans: blnTran = False
    Screen.MousePointer = 0
    DelEdit = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Function



Private Sub CtlEnabled()
    On Error GoTo errH

    '��ʾѡ���˰�ť
    txtPatiInfo(idx����).Locked = mInfo.EditType <> 1 Or mblnԤ�� = True
    txtPatiInfo(idx����).TabStop = mInfo.EditType = 1 And mblnԤ�� = False
    txtPatiInfo(idx����).BackColor = IIf(mInfo.EditType = 1 And mblnԤ�� = False, &H80000005, &H8000000F)
    cmdFind.Visible = mInfo.EditType = 1 And mblnԤ�� = False
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetPatiType(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
    '��ȡ�������
    Dim strSQL As String
    Dim strType As String
    Dim strSqlTmp As String, lngCount As Long, i As Long, lngBegin As Long, lngEnd As Long, lng��ʼ As Long, strReplace As String
    Dim rsPatiType As ADODB.Recordset
    
    On Error GoTo errH
    If mrsCtlType Is Nothing Then Exit Function
    mrsCtlType.Filter = ""
    mrsCtlType.MoveFirst
    If mrsCtlType.EOF Then Exit Function
    If lng����ID = 0 Or mInfo.�����¼ID = 0 Then Exit Function

    '������ǰ���ಡ�����͵�SQL
    Do While Not mrsCtlType.EOF
        If mrsCtlType!��ȡSQL & "" <> "" Then
        
            '�滻����ID
            strSqlTmp = Replace(mrsCtlType!��ȡSQL & "", "����id", "����ID")
            strSqlTmp = Replace(strSqlTmp, "����iD", "����ID")
            strSqlTmp = Replace(strSqlTmp, "����Id", "����ID")
            
            '��ȡ����ID���ִ���
            lngCount = (Len(strSqlTmp) - Len(Replace(strSqlTmp, "����ID", ""))) / Len("����ID")
            
            lngBegin = 0: lngEnd = 0: strReplace = "": lng��ʼ = 1 '��1��ʼ
            'ѭ������
            For i = 1 To lngCount
                lng��ʼ = InStr(lng��ʼ, strSqlTmp, "����ID") + 1
                
                lngBegin = lng��ʼ - 1
                lngEnd = InStr(lngBegin, strSqlTmp, "-1")

                If lngBegin + 4 < lngEnd And lngBegin <> 0 And lngEnd <> 0 Then
                    strReplace = Replace(Mid(strSqlTmp, lngBegin + 4, lngEnd - (lngBegin + 4)), " ", "")
                    strReplace = Replace(strReplace, vbCrLf, "")
                    If strReplace = "<>" Then
                        strSqlTmp = Replace(strSqlTmp, Mid(strSqlTmp, lngBegin, lngEnd + 2 - lngBegin), "����ID=[4]")
                        lng��ʼ = lngEnd + 2
                    End If
                End If
            Next

        
            '�滻��ҳID
            strSqlTmp = Replace(strSqlTmp, "��ҳid", "��ҳID")
            strSqlTmp = Replace(strSqlTmp, "��ҳiD", "��ҳID")
            strSqlTmp = Replace(strSqlTmp, "��ҳId", "��ҳID")
            
            '��ȡ��ҳID���ִ���
            lngCount = (Len(strSqlTmp) - Len(Replace(strSqlTmp, "��ҳID", ""))) / Len("��ҳID")
            
            lngBegin = 0: lngEnd = 0: strReplace = "": lng��ʼ = 1 '��1��ʼ
            'ѭ������
            For i = 1 To lngCount
                If i = 1 Then lng��ʼ = 1 '��1��ʼ
                
                lng��ʼ = InStr(lng��ʼ, strSqlTmp, "��ҳID") + 1
                
                lngBegin = lng��ʼ - 1
                lngEnd = InStr(lngBegin, strSqlTmp, "-1")

                If lngBegin + 4 < lngEnd And lngBegin <> 0 And lngEnd <> 0 Then
                    strReplace = Replace(Mid(strSqlTmp, lngBegin + 4, lngEnd - (lngBegin + 4)), " ", "")
                    strReplace = Replace(strReplace, vbCrLf, "")
                    If strReplace = "<>" Then
                        strSqlTmp = Replace(strSqlTmp, Mid(strSqlTmp, lngBegin, lngEnd + 2 - lngBegin), "��ҳID=[5]")
                        lng��ʼ = lngEnd + 2
                    End If
                End If
            Next
            
            strSQL = strSQL & " Union All " & strSqlTmp
        End If
        mrsCtlType.MoveNext
    Loop
    
    If strSQL <> "" Then
        strSQL = Mid(strSQL, 12)
        strSQL = Replace(strSQL, "[��ʼʱ��]", "[1]")
        strSQL = Replace(strSQL, "[����ʱ��]", "[2]")
        strSQL = Replace(strSQL, "[����ID]", "[3]")
        
        '�ݴ���
        strSQL = Replace(strSQL, "����ID<>-1", "����ID=[4]")
        strSQL = Replace(strSQL, "����ID <> -1", "����ID=[4]")
        strSQL = Replace(strSQL, "��ҳID<>-1", "��ҳID=[5]")
        strSQL = Replace(strSQL, "��ҳID <> -1", "��ҳID=[5]")
        Set rsPatiType = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(mInfo.���࿪ʼʱ��), CDate(mInfo.�������ʱ��), mInfo.�������ID & "", lng����ID, lng��ҳID)
        rsPatiType.Filter = "����ID =" & lng����ID & " And ��ҳID =" & lng��ҳID
    End If
    
    If Not rsPatiType Is Nothing Then
        Do While Not rsPatiType.EOF
            If InStr("," & strType & ",", "," & rsPatiType!���� & ",") = 0 And rsPatiType!���� & "" <> "" Then
                strType = strType & "," & rsPatiType!����
            End If
            rsPatiType.MoveNext
        Loop
        strType = Mid(strType, 2)
    End If
    
    GetPatiType = strType
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function





Private Function ReadData() As Boolean
    '��ȡ�ѱ���Ľ�������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim obj As Object
    Dim strValue As String, strTmp As String
    Dim lngBegin As Long, lngEnd As Long, j As Long

    
    On Error GoTo errH
    If mInfo.����ID = 0 Then Exit Function
    
    strSQL = "Select ����id, ���, ��Ŀ, ���� From ҽ�����Ӱ����� Where ����id = [1] Order By ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mInfo.����ID)
    
    If rsTmp.EOF Then Exit Function
    
    mblnLoad = True
    For Each obj In lblInfo
        If obj.Index <> 0 Then
            rsTmp.Filter = "��Ŀ ='" & Replace(obj.Caption, vbCrLf, "") & "'"
            
            If Not rsTmp.EOF Then
                 strValue = rsTmp!���� & ""
                 
                 '����ѡ���
                 If InStr("," & mstrPicCtl & ",", "," & obj.Index & ",") > 0 And strValue <> "" Then
                    
                    If InStr(strValue, ";") > 0 Then
                        strTmp = Split(strValue, ";")(0)
                        strValue = Mid(strValue, Len(strTmp) + 3)
                    End If
                    
                    lngBegin = Val(Split(picTmp(obj.Index).Tag, ",")(1))
                    lngEnd = Val(Split(picTmp(obj.Index).Tag, ",")(2))
                 
                    If Val(Mid(picTmp(obj.Index).Tag, 1, 1)) = 2 Then
                        For j = lngBegin To lngEnd
                            If InStr("," & strTmp & ",", "," & optInfo(j).Caption & ",") > 0 Then
                                optInfo(j).Value = True
                            End If
                        Next
                    Else
                        For j = lngBegin To lngEnd
                            If InStr("," & strTmp & ",", "," & chkInfo(j).Caption & ",") > 0 Then
                                chkInfo(j).Value = 1
                            End If
                        Next
                    End If
                    
                    Call SetTextVisible(obj.Index)
                 End If
                 
                 '�����ı���
                 If InStr("," & mstrTextCtl & ",", "," & obj.Index & ",") > 0 And strValue <> "" Then
                    txtInfo(obj.Index) = strValue
                    txtInfo(obj.Index).Tag = strValue
                 End If
            End If
        End If
    Next
    
    mblnLoad = False
    ReadData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function CheckData()
    '����ǰ���
    Dim obj As Object
    Dim blnCheck As Boolean
    Dim lngBegin As Long, lngEnd As Long, j As Long
    
    On Error GoTo errH
    
    If mInfo.����ID = 0 Then
        Set mobjCtl = txtPatiInfo(idx����)
        Call MsgBox("��ѡ����Ҫ��д�����¼�Ĳ���!", vbInformation, gstrSysName)
        zlcontrol.ControlSetFocus txtPatiInfo(idx����)
        Call zlcontrol.TxtSelAll(txtPatiInfo(idx����))
        Exit Function
    End If
    
    If txtPatiInfo(idx��������).Text = "" Then
        Set mobjCtl = txtPatiInfo(idx��������)
        Call MsgBox("��ѡ��ǰ���˵Ĳ�������!", vbInformation, gstrSysName)
        zlcontrol.ControlSetFocus txtPatiInfo(idx��������)
        Call zlcontrol.TxtSelAll(txtPatiInfo(idx��������))
        Exit Function
    End If
    
    
    For Each obj In lblInfo
        If obj.Index <> 0 Then
                '����ѡ���
                If InStr("," & mstrPicCtl & ",", "," & obj.Index & ",") > 0 Then
                   lngBegin = Val(Split(picTmp(obj.Index).Tag, ",")(1))
                   lngEnd = Val(Split(picTmp(obj.Index).Tag, ",")(2))
                    blnCheck = False
                   If Val(Mid(picTmp(obj.Index).Tag, 1, 1)) = 3 And lngEnd <> 0 And lngBegin <> lngEnd Then
                       For j = lngBegin To lngEnd
                           If chkInfo(j).Value = 1 Then
                               blnCheck = True
                               Exit For
                           End If
                       Next
                       
                        If blnCheck = False Then
                            If InStr("," & mstrTextCtl & ",", "," & obj.Index & ",") = 0 Then
                                 Call ShowMsg(obj.Caption & "û��ѡ���κ�ѡ��,���ܱ���!", chkInfo(lngBegin))
                                 Exit Function
                            Else
                                If txtInfo(obj.Index) = "" Then
                                    Call ShowMsg(obj.Caption & "û��ѡ���κ�ѡ��,���ܱ���!", chkInfo(lngBegin))
                                    Exit Function
                                End If
                            End If
                        End If
                   End If
                End If
                
                '����ı����Ƿ�Ϊ��
                If InStr("," & mstrTextCtl & ",", "," & obj.Index & ",") > 0 Then
                    If txtInfo(obj.Index).Text = "" And txtInfo(obj.Index).Locked = False And txtInfo(obj.Index).Visible = True And InStr("," & mstrPicCtl & ",", "," & obj.Index & ",") = 0 Then
                        Call ShowMsg(obj.Caption & "����Ϊ�գ�������¼��!", txtInfo(obj.Index))
                        Exit Function
                    End If
                    
                    '������������Ƿ���ȷ
                    Select Case Val(Split(obj.Tag, ",")(3))
                            Case 1
                                If (Not IsDate(txtInfo(obj.Index).Text)) And txtInfo(obj.Index).Locked = False And txtInfo(obj.Index).Visible = True Then
                                    Call ShowMsg(obj.Caption & "������Ч�����ڣ�������¼��!", txtInfo(obj.Index))
                                    Exit Function
                                End If
                            Case 2
                                If (Not IsNumeric(txtInfo(obj.Index).Text)) And txtInfo(obj.Index).Locked = False And txtInfo(obj.Index).Visible = True Then
                                    Call ShowMsg(obj.Caption & "������Ч�����֣�������¼��!", txtInfo(obj.Index))
                                    Exit Function
                                End If
                    End Select
                End If
        End If
    Next
    
    CheckData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Function ShowMsg(strMsg As String, obj As Object)
    '���ڼ����ʾʱ��λ�ؼ�
    On Error Resume Next
    '��Ϣ��ʾ
    Set mobjCtl = obj
    Call MsgBox(strMsg, vbInformation, gstrSysName)
    zlcontrol.ControlSetFocus obj
    If obj.Name = "txtInfo" Then
        Call zlcontrol.TxtSelAll(txtInfo(obj.Index))
    End If

    '�ؼ���λ
    Call ShowCtl(obj)
End Function


Private Function ShowCtl(obj As Object)
    '���ڼ����ʾʱ��λ�ؼ�
    Dim lngTop1 As Long, lngTop2 As Long

    On Error Resume Next
    '�ؼ���λ
    If vscBar.Visible Then
        Select Case obj.Name
                Case "txtInfo"
                    lngTop1 = obj.Top + obj.Height + 100
                    lngTop2 = obj.Top
                Case "chkInfo", "optInfo"
                    lngTop1 = obj.Container.Top + obj.Container.Height + 100
                    lngTop2 = obj.Container.Top
        End Select

        If lngTop1 < picMain.Height Then
            vscBar.Value = vscBar.Min
        ElseIf picEdit.Height - lngTop2 < picMain.Height Then
            vscBar.Value = vscBar.Max
        Else
            vscBar.Value = lngTop1 - picMain.Height
        End If
        zlcontrol.ControlSetFocus obj
    End If
End Function


Private Function GetDateStr(ByVal intIndex As Integer) As String
    '��ȡ���ڸ�ʽ�ַ���
    Dim strType As String, strName As String
    Dim strValue As String
    On Error GoTo errH
    If intIndex = 0 Then Exit Function
    
    strType = Split(lblInfo(intIndex).Tag, ",")(1)
    strName = Replace(lblInfo(intIndex).Caption, vbCrLf, "")

    If Not mrsCtlInfo Is Nothing Then
        mrsCtlInfo.Filter = "���˼�� ='" & strType & "' And ��Ŀ���� ='" & strName & "'"
        If Not mrsCtlInfo.EOF Then strValue = mrsCtlInfo!�����ʽ & ""
        mrsCtlInfo.Filter = ""
    End If
    
    If strValue = "" Then strValue = "yyyy-MM-dd HH:mm" 'Ĭ�ϸ�ʽ
    
    GetDateStr = strValue
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitRecordset(rsTmp As Recordset)
'���ܣ���ʼ��������ȡ��¼��
    Set rsTmp = New ADODB.Recordset
    
    rsTmp.Fields.Append "��������", adVarChar, 5000
    rsTmp.Fields.Append "�������", adVarChar, 5000
    rsTmp.Fields.Append "���ȡֵ", adVarChar, 5000
    
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
End Sub

Private Function IntData() As Boolean
    '��ʼ����ȡ����
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset, rsValue As ADODB.Recordset, rs������ȡ As ADODB.Recordset
    Dim obj As Object
    Dim strValue As String
    Dim str��� As String, str���� As String, str��Ѫ��� As String
    Dim arrTmp As Variant, i As Long
    
    On Error GoTo errH
    If mrsCtlInfo Is Nothing Then Exit Function
    
    mrsCtlInfo.Filter = ""
    If mInfo.����ID = 0 Or mrsCtlInfo.EOF Or mInfo.�������� = "" Then Exit Function

    Call InitRecordset(rs������ȡ)

    '���Ȱ��Զ�������Դ��SQLһ�ζ�ȡ����
    For Each obj In lblInfo
        If obj.Index <> 0 Then
            If InStr("," & mstrTextCtl & ",", "," & obj.Index & ",") > 0 Then
                mrsCtlInfo.Filter = "��Ŀ����='" & Replace(obj.Caption, vbCrLf, "") & "' And ���˼�� ='" & Split(obj.Tag, ",")(1) & "'"
                If Not mrsCtlInfo.EOF Then
                    '�Զ���sqlƴ��
                    If Val(mrsCtlInfo!������ʽ & "") = 1 And Val(mrsCtlInfo!��ȡ��Դ & "") = 99 And mrsCtlInfo!��ȡSQL <> "" Then
                            strSQL = strSQL & " Union All Select '" & Replace(obj.Caption, vbCrLf, "") & "' As ��Ŀ����, a.* From (" & mrsCtlInfo!��ȡSQL & ") A "
                    
                    '�������ƴ��
                    ElseIf Val(mrsCtlInfo!������ʽ & "") = 1 And Val(mrsCtlInfo!��ȡ��Դ & "") = 4 And mrsCtlInfo!��ȡ���� & "" <> "" And InStr(mrsCtlInfo!��ȡ���� & "", ":") > 0 Then
                            rs������ȡ.Filter = "�������� ='" & Split(mrsCtlInfo!��ȡ���� & "", ":")(0) & "'"
                            If rs������ȡ.EOF Then
                                rs������ȡ.AddNew
                                rs������ȡ!�������� = Split(mrsCtlInfo!��ȡ���� & "", ":")(0)
                                rs������ȡ!������� = Split(mrsCtlInfo!��ȡ���� & "", ":")(1)
                            Else
                                arrTmp = Array()
                                arrTmp = Split(Split(mrsCtlInfo!��ȡ���� & "", ":")(1), ";")
                                For i = 0 To UBound(arrTmp)
                                    If arrTmp(i) <> "" And InStr(";" & rs������ȡ!������� & ";", ";" & arrTmp(i) & ";") = 0 Then
                                        rs������ȡ!������� = rs������ȡ!������� & ";" & arrTmp(i)
                                    End If
                                Next
                            End If
                    End If
                End If
            End If
        End If
    Next
    
    'һ�ζ�ȡ������ȡ
    rs������ȡ.Filter = ""
    Do While Not rs������ȡ.EOF
        If rs������ȡ!�������� & "" <> "" And rs������ȡ!������� <> "" Then
            rs������ȡ!���ȡֵ = GetOPSByEmr(rs������ȡ!�������� & "", mInfo.����ID, mInfo.��ҳID, Replace(rs������ȡ!������� & "", ";", "|"))
            rs������ȡ!���ȡֵ = Replace(rs������ȡ!���ȡֵ & "", vbCrLf, "")
            rs������ȡ!���ȡֵ = Replace(rs������ȡ!���ȡֵ & "", Chr(10), "")
            rs������ȡ!���ȡֵ = Replace(rs������ȡ!���ȡֵ & "", Chr(13), "")
        End If
        rs������ȡ.MoveNext
    Loop
    
    
    If strSQL <> "" Then
        strSQL = Mid(strSQL, 12)
        strSQL = Replace(strSQL, "[����ID]", "[1]")
        strSQL = Replace(strSQL, "[��ҳID]", "[2]")
        strSQL = Replace(strSQL, "[��ʼʱ��]", "[3]")
        strSQL = Replace(strSQL, "[����ʱ��]", "[4]")
        strSQL = Replace(strSQL, "[����ID]", "[5]")
        
        Set rsValue = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mInfo.����ID, mInfo.��ҳID, CDate(mInfo.���࿪ʼʱ��), CDate(mInfo.�������ʱ��), mInfo.�������ID & "")
    End If


    For Each obj In lblInfo
        If obj.Index <> 0 Then
            'ֻ��ȡ�ı����ʼ����
            If InStr("," & mstrTextCtl & ",", "," & obj.Index & ",") > 0 Then
                mrsCtlInfo.Filter = "��Ŀ����='" & Replace(obj.Caption, vbCrLf, "") & "' And ���˼�� ='" & Split(obj.Tag, ",")(1) & "'"
                If Not mrsCtlInfo.EOF Then
                    If Val(mrsCtlInfo!������ʽ & "") = 1 Then
                            Select Case Val(mrsCtlInfo!��ȡ��Դ & "")
                                Case 1 '��ȡ�������
                                    If str��� = "" Then
                                        strSQL = "Select Zl_Fun_Getpatishift(1, [1], [2], [3], [4]) as ������� From Dual"
                                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mInfo.����ID, mInfo.��ҳID, CDate(mInfo.���࿪ʼʱ��), CDate(mInfo.�������ʱ��))
                                        If Not rsTmp Is Nothing Then str��� = rsTmp!������� & ""
                                     End If
                                    txtInfo(obj.Index).Text = str���
                                Case 2 '��ȡ��������
                                    If str���� = "" Then
                                        strSQL = "Select Zl_Fun_Getpatishift(2, [1], [2], [3], [4]) as �������� From Dual"
                                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mInfo.����ID, mInfo.��ҳID, CDate(mInfo.���࿪ʼʱ��), CDate(mInfo.�������ʱ��))
                                        If Not rsTmp Is Nothing Then str���� = rsTmp!�������� & ""
                                    End If
                                    txtInfo(obj.Index).Text = str����
                                Case 3 '��ȡ��Ѫ���
                                    If str��Ѫ��� = "" Then
                                        strSQL = "Select Zl_Fun_Getpatishift(3, [1], [2], [3], [4]) as ��Ѫ��� From Dual"
                                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mInfo.����ID, mInfo.��ҳID, CDate(mInfo.���࿪ʼʱ��), CDate(mInfo.�������ʱ��))
                                        If Not rsTmp Is Nothing Then str��Ѫ��� = rsTmp!��Ѫ��� & ""
                                    End If
                                    txtInfo(obj.Index).Text = str��Ѫ���
                                Case 4 '��ȡ�°没������
                                    strValue = Get�������(mrsCtlInfo!��ȡ���� & "", rs������ȡ)
                                    If strValue & "" <> "" Then
                                       If Val(mrsCtlInfo!�������� & "") = 0 Then
                                            txtInfo(obj.Index).Text = strValue
                                       ElseIf Val(mrsCtlInfo!�������� & "") = 1 Then
                                            txtInfo(obj.Index).Text = Format(strValue, "yyyy-MM-dd HH:mm")
                                       ElseIf Val(mrsCtlInfo!�������� & "") = 2 Then
                                            txtInfo(obj.Index).Text = Val(strValue)
                                       End If
                                    End If
                                Case 99 'ͨ��SQL��ȡ
                                    If Not rsValue Is Nothing Then
                                        rsValue.Filter = "��Ŀ���� = '" & Replace(obj.Caption, vbCrLf, "") & "'"
                                        If Not rsValue.EOF Then
                                            If rsValue.Fields(1).Value & "" <> "" Then
                                               If Val(mrsCtlInfo!�������� & "") = 0 Then
                                                    txtInfo(obj.Index).Text = rsValue.Fields(1).Value & ""
                                               ElseIf Val(mrsCtlInfo!�������� & "") = 1 Then
                                                    txtInfo(obj.Index).Text = Format(rsValue.Fields(1).Value & "", "yyyy-MM-dd HH:mm")
                                               ElseIf Val(mrsCtlInfo!�������� & "") = 2 Then
                                                    txtInfo(obj.Index).Text = Val(rsValue.Fields(1).Value & "")
                                               End If
                                            End If
                                        End If
                                    End If
                            End Select
                    End If
                End If
                txtInfo(obj.Index).Tag = txtInfo(obj.Index).Text
            End If
        End If
    Next
    
    mblnChange = False
    
    IntData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function Get�������(ByVal str��ʽ As String, rsTmp As ADODB.Recordset) As String
    '��ȡ�°没�����
    Dim arrTmp As Variant, arr���� As Variant, arrȡֵ As Variant, i As Long, j As Long
    Dim strTmp As String
    Dim strValue As String
    
    On Error GoTo errH
    
    '�ݴ���
    If InStr(str��ʽ, ":") = 0 Or mInfo.����ID = 0 Then Exit Function
    If Split(str��ʽ, ":")(1) = "" Or Split(str��ʽ, ":")(0) = "" Then Exit Function
    If rsTmp Is Nothing Then Exit Function
    rsTmp.Filter = "�������� ='" & Split(str��ʽ, ":")(0) & "'"
    If rsTmp.EOF Then Exit Function
    If rsTmp!���ȡֵ & "" = "" Then Exit Function
    arrTmp = Array()
    arrTmp = Split(Split(str��ʽ, ":")(1), ";")
    arr���� = Array()
    arr���� = Split(rsTmp!������� & "", ";")
    arrȡֵ = Array()
    arrȡֵ = Split(rsTmp!���ȡֵ & "", "|")
    
    If UBound(arr����) <> UBound(arrȡֵ) Then Exit Function

    For i = 0 To UBound(arrTmp)
        For j = 0 To UBound(arr����)
            If arrTmp(i) = arr����(j) Then
                strValue = arrȡֵ(j)
                strValue = Replace(strValue, arr����(j) & " : ", "")
                strValue = Replace(strValue, arr����(j) & ": ", "")
                strValue = Replace(strValue, arr����(j) & ":", "")
                strValue = Replace(strValue, arr����(j) & " :", "")
                
                strValue = Replace(strValue, arr����(j) & " �� ", "")
                strValue = Replace(strValue, arr����(j) & "�� ", "")
                strValue = Replace(strValue, arr����(j) & "��", "")
                strValue = Replace(strValue, arr����(j) & " ��", "")

                strTmp = strTmp & ";" & strValue
                Exit For
            End If
        Next
    Next
    
    strTmp = Mid(strTmp, 2)

    Get������� = strTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetOPSByEmr(ByVal strDocKind As String, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str��� As String) As String
'���ܣ���ȡָ�����˵�ָ������ڲ�����д����Ϣ�����磺���ߣ���ϵȡ��Ӳ����л�ȡ����ֵ
    Dim strText As String
    
    On Error Resume Next
    
    If gobjEmr Is Nothing Then Exit Function
    If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then Set gobjEmr = Nothing: Exit Function
 
    If Not gobjEmr Is Nothing Then
        strText = gobjEmr.GetContentOfSpecifyDoc(strDocKind, lng����ID, lng��ҳID, str���)
    End If
    
    Err.Clear
    GetOPSByEmr = strText
End Function

Private Sub MakeText()
     '���ɽ�������
    Dim str���� As String, strS���� As String, str���� As String, str��� As String, strB���� As String, strA���� As String, strR���� As String
    Dim strType As String, strTmp As String, strValue As String
    Dim obj As Object, i As Long, j As Long, lngBegin As Long, lngEnd As Long
    Dim dtNow As Date, dtTmp As Date
        
    On Error GoTo errH
    
    dtNow = zlDatabase.Currentdate
    
    With rtbBox
        .SelStart = 0
        .SelLength = Len(rtbBox.Text)
        .SelColor = vbBlack
        .SelStart = Len(rtbBox.Text)
    End With
    
    If mInfo.����ID = 0 Then
        str���� = "[����]����[����]��[�Ա�]��[����]��סԺ�ţ�[סԺ��]��[��Ժʱ��]��[����]Ϊ����[��Ժ;��]��Ժ��"
        rtbBox.Text = str����
        rtbBox.Tag = str����
        Exit Sub
    End If
    '���ɲ�����Ϣ��S����Ŀ������
    str���� = "[����]����[����][�Ա�][����]��סԺ�ţ�[סԺ��][strS����]��[��Ժʱ��][str����][��Ժ;��]��Ժ[str���]��[strB����][strA����][strR����]"
    
    For Each obj In lblPatiInfo
        If obj.Caption = "�Ա�" Or obj.Caption = "����" Then
            str���� = Replace(str����, "[" & obj.Caption & "]", IIf(txtPatiInfo(obj.Index).Text = "", "", "��" & txtPatiInfo(obj.Index).Text))
        ElseIf obj.Caption = "����" Then
            str���� = Replace(str����, "[" & obj.Caption & "]", IIf(txtPatiInfo(obj.Index).Text = "", "", txtPatiInfo(obj.Index).Text & "��"))
        ElseIf obj.Caption = "��Ժʱ��" And IsDate(txtPatiInfo(obj.Index).Text) And txtPatiInfo(obj.Index).Text <> "" Then
            dtTmp = CDate(txtPatiInfo(obj.Index).Text)
            If Year(dtTmp) = Year(dtNow) And Month(dtTmp) = Month(dtNow) And Day(dtTmp) = Day(dtNow) Then
                strTmp = "����" & Format(txtPatiInfo(obj.Index).Text, "HHʱmm��")
            ElseIf Year(dtTmp) = Year(dtNow) And Month(dtTmp) = Month(dtNow) And Day(dtTmp) = Day(dtNow) + 1 Then
                strTmp = "����" & Format(txtPatiInfo(obj.Index).Text, "HHʱmm��")
            Else
                strTmp = IIf(txtPatiInfo(obj.Index).Text = "", "", Format(txtPatiInfo(obj.Index).Text, "yyyy��MM��dd��HHʱmm��"))
            End If
            str���� = Replace(str����, "[" & obj.Caption & "]", strTmp)
        Else
            str���� = Replace(str����, "[" & obj.Caption & "]", txtPatiInfo(obj.Index).Text)
        End If
    Next
    

    For i = 1 To lblInfo.Count - 1
        If lblInfo(i).Caption = "����" Then '�������ߵ�����
            str���� = IIf(txtInfo(i).Text = "", "", "�ԡ�" & txtInfo(i).Text & "��Ϊ����")
        ElseIf InStr(lblInfo(i).Caption, "���") > 0 Then '������ϵ�����
            If txtInfo(i).Text <> "" Then
                str��� = "��" & lblInfo(i).Caption & "��" & txtInfo(i).Text
            End If
        Else
            mrsCtlInfo.Filter = "��Ŀ����='" & Replace(lblInfo(i).Caption, vbCrLf, "") & "' And ���˼�� ='" & Split(lblInfo(i).Tag, ",")(1) & "'"
            If Not mrsCtlInfo.EOF Then
                strValue = "": strTmp = ""
                If mrsCtlInfo!�������� & "" = "" Then
                    strValue = lblInfo(i).Caption & "��"
                Else
                    If mrsCtlInfo!�������� & "" = "-" Then
                        strValue = ""
                    Else
                        strValue = mrsCtlInfo!�������� & "��"
                    End If
                End If
                
                If InStr("," & mstrPicCtl & ",", "," & i & ",") > 0 Then
                   lngBegin = Val(Split(picTmp(i).Tag, ",")(1))
                   lngEnd = Val(Split(picTmp(i).Tag, ",")(2))
                   
                   If Val(Mid(picTmp(i).Tag, 1, 1)) = 3 And lngEnd <> 0 Then
                       For j = lngBegin To lngEnd
                           If chkInfo(j).Value = 1 Then
                                strTmp = strTmp & IIf(strTmp = "", "", "��") & chkInfo(j).Caption
                           End If
                       Next
                   ElseIf Val(Mid(picTmp(i).Tag, 1, 1)) = 2 And lngEnd <> 0 Then
                        For j = lngBegin To lngEnd
                           If optInfo(j).Value = True Then
                                strTmp = strTmp & IIf(strTmp = "", "", "��") & optInfo(j).Caption
                                Exit For
                           End If
                        Next
                   End If
                End If
                
                If InStr("," & mstrTextCtl & ",", "," & i & ",") > 0 Then
                   If Val(mrsCtlInfo!�������� & "") = 1 And Val(mrsCtlInfo!������ʽ & "") = 1 And txtInfo(i).Text <> "" And IsDate(txtInfo(i).Text) Then
                        dtTmp = CDate(txtInfo(i).Text)
                        If Year(dtTmp) = Year(dtNow) And Month(dtTmp) = Month(dtNow) And Day(dtTmp) = Day(dtNow) Then
                            strTmp = strTmp & IIf(strTmp = "", "", "��") & "����" & Format(txtInfo(i).Text, "HHʱmm��")
                        ElseIf Year(dtTmp) = Year(dtNow) And Month(dtTmp) = Month(dtNow) And Day(dtTmp) = Day(dtNow) + 1 Then
                            strTmp = strTmp & IIf(strTmp = "", "", "��") & "����" & Format(txtInfo(i).Text, "HHʱmm��")
                        Else
                            strTmp = strTmp & IIf(strTmp = "", "", "��") & IIf(txtInfo(i).Text = "", "", Format(txtInfo(i).Text, GetDateStr(i)))
                        End If
                    
                   ElseIf txtInfo(i).Text <> "" Then
                       strTmp = strTmp & IIf(strTmp = "", "", "��") & txtInfo(i).Text
                   End If
                End If
                
                '����B���͵���ʼ����
                If Split(lblInfo(i).Tag, ",")(2) = "B" Then
                    If InStr("," & strType & ",", "," & Split(lblInfo(i).Tag, ",")(1) & ",") = 0 Then
                        mrsCtlType.Filter = "��� ='" & Split(lblInfo(i).Tag, ",")(1) & "'"
                        
                        If Not mrsCtlType.EOF Then
                            If mrsCtlType!��ʼ���� & "" <> "" Then
                                strB���� = strB���� & IIf(Right(strB����, 1) <> "��" And Right(strB����, 1) <> "��" And strB���� <> "", "��", "") & mrsCtlType!��ʼ����
                            End If
                        End If
                        strType = strType & "," & Split(lblInfo(i).Tag, ",")(1)
                    End If
                End If
                
                If strTmp <> "" Then
                    Select Case Split(lblInfo(i).Tag, ",")(2)
                            Case "S"
                                strS���� = strS���� & IIf(strS���� = "", "", "��") & strValue & strTmp
                            Case "B"
                                If InStr(strB����, "[" & Replace(lblInfo(i).Caption, vbCrLf, "") & "]") > 0 Then
                                    strB���� = Replace(strB����, "[" & Replace(lblInfo(i).Caption, vbCrLf, "") & "]", strTmp)
                                Else
                                    strB���� = strB���� & IIf(strB���� = "" Or Right(strB���� & "", 1) = "��" Or Right(strB���� & "", 1) = "��" Or Right(strB���� & "", 1) = "��", "", "��") & strValue & strTmp
                                End If
                            Case "A"
                                strA���� = strA���� & IIf(strS���� = "", "", "��") & strValue & strTmp
                            Case "R"
                                strR���� = strR���� & IIf(strS���� = "", "", "��") & strValue & strTmp
                    End Select
                End If
            End If
        End If
    Next

    str���� = Replace(str����, "[str����]", str����)
    str���� = Replace(str����, "[str���]", str���)
    str���� = Replace(str����, "[strS����]", strS����)
    str���� = Replace(str����, "[strA����]", strA���� & IIf(strA���� = "", "", "��"))
    str���� = Replace(str����, "[strR����]", strR���� & IIf(strR���� = "", "", "��"))
    str���� = Replace(str����, "[strB����]", strB���� & IIf(strB���� = "", "", "��"))

    '�����ַ�����
    str���� = Replace(str����, "&", "")
    str���� = Replace(str����, "'", "")
    str���� = Replace(str����, "<", "")
    str���� = Replace(str����, ">", "")
    rtbBox.Text = str����
    rtbBox.Tag = str����
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function SaveData()
    '��������
    Dim i As Long, j As Long, lngBegin As Long, lngEnd As Long
    Dim strValue As String, blnTran As Boolean
    Dim arrSQL As Variant

    On Error GoTo errH

    arrSQL = Array()
    '�������ݼ�¼
    If mInfo.EditType = 1 Then  '����
        mInfo.����ID = GetNextId("ҽ�����Ӱ�����", "����ID")
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_ҽ�����Ӱ�����_Edit(0," & mInfo.����ID & "," & mInfo.�����¼ID & "," & mInfo.������� & ",'" & mInfo.�������� & "'," & mInfo.����ID & "," & mInfo.��ҳID & ",'" & _
                            txtPatiInfo(idx����).Text & "','" & txtPatiInfo(idx�Ա�).Text & "','" & txtPatiInfo(idx����).Text & "','" & txtPatiInfo(idx����).Text & "'," & ZVal(txtPatiInfo(idxסԺ��).Text) & _
                            ",to_date('" & Format(txtPatiInfo(idx��Ժʱ��).Text, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')" & ",'" & _
                            txtPatiInfo(idx��Ժ;��).Text & "','" & rtbBox.Text & "')"
    ElseIf mInfo.EditType = 2 Then  '�޸�
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_ҽ�����Ӱ�����_Edit(1," & mInfo.����ID & "," & mInfo.�����¼ID & "," & mInfo.������� & ",'" & mInfo.�������� & "'," & mInfo.����ID & "," & mInfo.��ҳID & ",'" & _
                            txtPatiInfo(idx����).Text & "','" & txtPatiInfo(idx�Ա�).Text & "','" & txtPatiInfo(idx����).Text & "','" & txtPatiInfo(idx����).Text & "'," & ZVal(txtPatiInfo(idxסԺ��).Text) & _
                            ",to_date('" & Format(txtPatiInfo(idx��Ժʱ��).Text, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')" & ",'" & _
                            txtPatiInfo(idx��Ժ;��).Text & "','" & rtbBox.Text & "')"
    End If
    
    '������������
    If mInfo.EditType = 2 Then  '�޸�ʱ��ɾ���������ύ
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_ҽ�����Ӱ�����_Edit(2," & mInfo.����ID & ")"
    End If
    
    For i = 1 To lblInfo.Count - 1
        '��ȡ�ؼ�����
        strValue = ""
        
         '����ѡ���
         If InStr("," & mstrPicCtl & ",", "," & i & ",") > 0 Then
            
            lngBegin = Val(Split(picTmp(i).Tag, ",")(1))
            lngEnd = Val(Split(picTmp(i).Tag, ",")(2))
         
            If Val(Mid(picTmp(i).Tag, 1, 1)) = 2 Then
                For j = lngBegin To lngEnd
                    If optInfo(j).Value Then
                        strValue = strValue & "," & optInfo(j).Caption
                    End If
                Next
            Else
                For j = lngBegin To lngEnd
                    If chkInfo(j).Value = 1 Then
                        strValue = strValue & "," & chkInfo(j).Caption
                    End If
                Next
            End If
            strValue = Mid(strValue, 2)
            strValue = IIf(strValue <> "", strValue & ";", "")
         End If
         
         '�����ı���
         If InStr("," & mstrTextCtl & ",", "," & i & ",") > 0 Then
            strValue = strValue & txtInfo(i).Text
         End If
         
        strValue = Replace(strValue, "'", "")
        '���SQL
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_ҽ�����Ӱ�����_Edit(1," & mInfo.����ID & "," & i & ",'" & Replace(lblInfo(i).Caption, vbCrLf, "") & "','" & strValue & "')"
    Next
    
    Screen.MousePointer = 11
    gcnOracle.BeginTrans: blnTran = True
    For i = 0 To UBound(arrSQL)
        Debug.Print CStr(arrSQL(i))
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTran = False
    
    
    If mInfo.EditType = 1 Then  'ͬ����¼���
        mInfo.������� = Val(Sys.RowValue("ҽ�����Ӱ�����", mInfo.����ID, "���", "����ID") & "")
    End If
    
    On Error GoTo 0
    Screen.MousePointer = 0
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Function
