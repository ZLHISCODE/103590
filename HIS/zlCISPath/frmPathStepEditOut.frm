VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPathStepEditOut 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�׶�����"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "frmPathStepEditOut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txt���� 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2835
      MaxLength       =   3
      TabIndex        =   10
      Top             =   1020
      Width           =   420
   End
   Begin VB.TextBox txt���� 
      Alignment       =   2  'Center
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1545
      MaxLength       =   3
      TabIndex        =   9
      Top             =   1005
      Width           =   420
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2040
      TabIndex        =   7
      Top             =   3690
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3135
      TabIndex        =   6
      Top             =   3690
      Width           =   1100
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   4755
      TabIndex        =   3
      Top             =   0
      Width           =   4755
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʱ��׶�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1065
         TabIndex        =   5
         Top             =   120
         Width           =   780
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "  ����·�����е�һ��ʱ��׶Σ������Ǿ���ĳ��������Ҳ������һ��������Χ��"
         Height          =   360
         Left            =   1065
         TabIndex        =   4
         Top             =   360
         Width           =   3240
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   105
         Picture         =   "frmPathStepEditOut.frx":058A
         Top             =   45
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   10000
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   10000
         Y1              =   840
         Y2              =   840
      End
   End
   Begin VB.TextBox txt���� 
      Alignment       =   2  'Center
      Height          =   660
      Left            =   1200
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "���У�Ctrl+Enter"
      Top             =   1560
      Width           =   2880
   End
   Begin VB.TextBox txt˵�� 
      Height          =   660
      Left            =   1200
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2760
      Width           =   2880
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   1200
      TabIndex        =   0
      Top             =   2340
      Width           =   2880
   End
   Begin MSComCtl2.UpDown ud���� 
      Height          =   300
      Index           =   1
      Left            =   3255
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1020
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txt����(1)"
      BuddyDispid     =   196627
      BuddyIndex      =   1
      OrigLeft        =   2265
      OrigTop         =   1815
      OrigRight       =   2520
      OrigBottom      =   2010
      Max             =   999
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   0   'False
   End
   Begin MSComCtl2.UpDown ud���� 
      Height          =   300
      Index           =   0
      Left            =   1965
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1005
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txt����(0)"
      BuddyDispid     =   196627
      BuddyIndex      =   0
      OrigLeft        =   2265
      OrigTop         =   1815
      OrigRight       =   2520
      OrigBottom      =   2010
      Max             =   999
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      Height          =   180
      Left            =   720
      TabIndex        =   17
      Top             =   1065
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   180
      Index           =   0
      Left            =   1290
      TabIndex        =   16
      Top             =   1065
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��  -         ��"
      Height          =   180
      Index           =   1
      Left            =   2310
      TabIndex        =   15
      Top             =   1065
      Width           =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   0
      X2              =   10000
      Y1              =   3555
      Y2              =   3555
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   0
      X2              =   10000
      Y1              =   3570
      Y2              =   3570
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Left            =   720
      TabIndex        =   14
      Top             =   1620
      Width           =   360
   End
   Begin VB.Label lbl˵�� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "˵��"
      Height          =   180
      Left            =   720
      TabIndex        =   13
      Top             =   2820
      Width           =   360
   End
   Begin VB.Label lbl���� 
      Caption         =   "����"
      Height          =   180
      Left            =   720
      TabIndex        =   12
      Top             =   2400
      Width           =   540
   End
End
Attribute VB_Name = "frmPathStepEditOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event CheckDataValid(TimeStep As TYPE_PATH_STEP, Cancel As Boolean)

Private mvStep      As TYPE_PATH_STEP
Private mvPreStep   As TYPE_PATH_STEP
Private mvNextStep  As TYPE_PATH_STEP
Private mstr����s   As String
Private mblnOK      As Boolean

Public Function ShowEdit(frmParent As Object, vStep As TYPE_PATH_STEP, _
    vPreStep As TYPE_PATH_STEP, vNextStep As TYPE_PATH_STEP, ByVal str����s As String) As Boolean
'���ܣ����õ�ǰѡ��ʱ��׶ε���ϸ����
'������vStep=��Ҫ���޸�ʱ��ǰʱ��׶ε����ݣ������е�"��ID<>0"��ʾ���÷�֧
'      mvPreStep,mvNextStep=ǰ�����ڵ�һ��ʱ��׶ε����ݣ���������ʱ�ο�
'      str����s=��ǰ·�����У�ǰ��׶α��÷�֧�ķ�����������"|"���
    mvStep = vStep
    mvPreStep = vPreStep
    mvNextStep = vNextStep
    mstr����s = str����s
    
    Me.Show 1, frmParent
    
    If mblnOK Then vStep = mvStep
    ShowEdit = mblnOK
End Function

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "|" Then KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim blnCancel As Boolean
    Dim strTmp As String, i As Integer
    
    '�������
    If txt����(0).Text <> "" And Val(txt����(0).Text) <= 0 Then
        MsgBox "������һ����Ч�Ŀ�ʼ����ֵ��", vbInformation, gstrSysName
        txt����(0).SetFocus: Exit Sub
    End If
    If txt����(1).Text <> "" And Val(txt����(1).Text) <= 0 Then
        MsgBox "������һ����Ч�Ľ�������ֵ��", vbInformation, gstrSysName
        txt����(0).SetFocus: Exit Sub
    End If
    If txt����(0).Text <> "" And txt����(1).Text <> "" Then
        If Val(txt����(1).Text) < Val(txt����(0).Text) Then
            MsgBox "��������Ӧ�ô��ڿ�ʼ������", vbInformation, gstrSysName
            txt����(1).SetFocus: Exit Sub
        ElseIf Val(txt����(0).Text) = Val(txt����(1).Text) Then
            MsgBox "ָ��Ϊĳһ������ʱ������Ҫ�������������", vbInformation, gstrSysName
            txt����(1).SetFocus: Exit Sub
        End If
    End If
    If txt����(1).Text <> "" And txt����(0).Text = "" Then
        MsgBox "�����뿪ʼ������", vbInformation, gstrSysName
        txt����(0).SetFocus: Exit Sub
    End If
    
    If Trim(txt����.Text) = "" Then
        MsgBox "������ʱ��׶ε����ơ�", vbInformation, gstrSysName
        txt����.SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(txt����.Text) > txt����.MaxLength Then
        MsgBox "��������̫����������� " & txt����.MaxLength \ 2 & " �����ֻ��� " & txt����.MaxLength & " ���ַ���", vbInformation, gstrSysName
        txt����.SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(cbo����.Text) > 50 Then
        MsgBox "��������̫����������� 25 �����ֻ��� 50 ���ַ���", vbInformation, gstrSysName
        cbo����.SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(txt˵��.Text) > txt˵��.MaxLength Then
        MsgBox "˵������̫����������� " & txt˵��.MaxLength \ 2 & " �����ֻ��� " & txt˵��.MaxLength & " ���ַ���", vbInformation, gstrSysName
        txt˵��.SetFocus: Exit Sub
    End If
    
    '�ռ�����
    mvStep.���� = txt����.Text
    mvStep.˵�� = txt˵��.Text
    mvStep.��ʼ���� = Val(txt����(0).Text)
    mvStep.�������� = Val(txt����(1).Text)
    mvStep.���� = cbo����.Text
        
    '��������
    If mvStep.��ID = 0 Then
        RaiseEvent CheckDataValid(mvStep, blnCancel)
        If blnCancel Then Exit Sub
        '����ָ��������Χ
        If txt����(0).Text = "" And txt����(1).Text = "" And txt����(0).Enabled Then
            If MsgBox("û��ȷ����ʱ��׶�����Ӧ��������Χ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    mblnOK = False
    
    txt����.Text = mvStep.����
    cbo����.Text = mvStep.����
    txt˵��.Text = mvStep.˵��
    txt����(0).Text = IIf(mvStep.��ʼ���� = 0, "", mvStep.��ʼ����)
    txt����(1).Text = IIf(mvStep.�������� = 0, "", mvStep.��������)
    
    '������ʱ������ǰһ���׶ε�������Χ����ȱʡ
    If mvStep.���� = "" Then
        If mvPreStep.���� <> "" Then
            If mvPreStep.�������� <> 0 Then
                txt����(0).Text = mvPreStep.�������� + 1
            ElseIf mvPreStep.��ʼ���� <> 0 Then
                txt����(0).Text = mvPreStep.��ʼ���� + 1
            End If
        Else
            txt����(0).Text = "1"
        End If
        If mvNextStep.���� <> "" And txt����(0).Text <> "" Then
            If mvNextStep.��ʼ���� <> 0 And mvNextStep.��ʼ���� - 1 > Val(txt����(0).Text) Then
                txt����(1).Text = mvNextStep.��ʼ���� - 1
            End If
        End If
        If txt����(0).Text <> "" Then
            Call MakeStepName
        End If
    End If
    
    '���÷�ֻ֧�����޸�˵��
    If mvStep.��ID <> 0 Then
        Me.Caption = "��֧����"
        txt����.Enabled = False
        txt����.BackColor = Me.BackColor
        For i = 0 To txt����.UBound
            txt����(i).Enabled = False
            txt����(i).BackColor = Me.BackColor
        Next
        For i = 0 To ud����.UBound
            ud����(i).Enabled = False
        Next
    End If
    
    '���÷�֧�����÷���
    If mvStep.��ID = 0 Then
        lbl˵��.Top = lbl˵��.Top - cbo����.Height - (cbo����.Top - txt����.Top - txt����.Height)
        txt˵��.Top = txt˵��.Top - cbo����.Height - (cbo����.Top - txt����.Top - txt����.Height)
        cmdOK.Top = cmdOK.Top - cbo����.Height - (cbo����.Top - txt����.Top - txt����.Height)
        cmdCancel.Top = cmdCancel.Top - cbo����.Height - (cbo����.Top - txt����.Top - txt����.Height)
        
        Line1(0).Y1 = Line1(0).Y1 - cbo����.Height - (cbo����.Top - txt����.Top - txt����.Height)
        Line1(0).Y2 = Line1(0).Y1
        Line1(1).Y1 = Line1(1).Y1 - cbo����.Height - (cbo����.Top - txt����.Top - txt����.Height)
        Line1(1).Y2 = Line1(1).Y1
        
        Me.Height = Me.Height - cbo����.Height - (cbo����.Top - txt����.Top - txt����.Height)
    
        lbl����.Visible = False
        cbo����.Visible = False
    Else
        For i = 0 To UBound(Split(mstr����s, "|"))
            cbo����.AddItem Split(mstr����s, "|")(i)
        Next
    End If
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt˵��_GotFocus()
    Call zlControl.TxtSelAll(txt˵��)
End Sub

Private Sub txt˵��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txt����_Change(Index As Integer)
    txt����(1).Enabled = txt����(0).Text <> ""
    ud����(1).Enabled = txt����(1).Enabled
    If Not txt����(1).Enabled Then
        txt����(1).Text = ""
        txt����(1).BackColor = Me.BackColor
    Else
        txt����(1).BackColor = txt����(0).BackColor
    End If
    
    If Visible Then Call MakeStepName
End Sub

Private Sub MakeStepName()
    Dim str���� As String
    Dim i As Long
    
    If txt����(1).Text <> "" Then
        str���� = "�����" & txt����(0).Text & "-" & txt����(1).Text & "��"
    Else
        str���� = "�����" & txt����(0).Text & "��"
    End If
    
    txt����.Text = str����
End Sub

Private Sub txt����_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txt����(Index))
End Sub

Private Sub txt����_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
