VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmBatPrint 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "·����������ӡ"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5745
   Icon            =   "frmBatPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdPrintSetup 
      Caption         =   "��ӡ����"
      Height          =   300
      Left            =   240
      TabIndex        =   7
      Top             =   1720
      Width           =   1100
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "��ʼ��ӡ"
      Default         =   -1  'True
      Height          =   300
      Left            =   2040
      TabIndex        =   6
      Top             =   1720
      Width           =   1100
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Ԥ��(&V)"
      Height          =   300
      Left            =   3240
      TabIndex        =   5
      Top             =   1720
      Width           =   1100
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5160
      Top             =   360
   End
   Begin VB.CommandButton cmdCancle 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&E)"
      Height          =   300
      Left            =   4440
      TabIndex        =   2
      Top             =   1720
      Width           =   1100
   End
   Begin VB.PictureBox picTime 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   240
      Picture         =   "frmBatPrint.frx":6852
      ScaleHeight     =   255
      ScaleWidth      =   5325
      TabIndex        =   1
      Top             =   975
      Width           =   5320
   End
   Begin VB.PictureBox picTime 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   240
      Picture         =   "frmBatPrint.frx":719E
      ScaleHeight     =   255
      ScaleWidth      =   5325
      TabIndex        =   0
      Top             =   960
      Width           =   5320
   End
   Begin XtremeSuiteControls.TabControl tbcPath 
      Height          =   3090
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   5475
      _Version        =   589884
      _ExtentX        =   9657
      _ExtentY        =   5450
      _StockProps     =   64
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmBatPrint.frx":7A54
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblMsg 
      Caption         =   "����ӡ    �����ˣ����ڴ�ӡ��    �����ˡ�"
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmBatPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmPath As frmPathTable
Private mfrmPathOut As frmPathTableOut
Private mrsTmp As Recordset
Private mlngI As Long
Private mlngCount As Long
Private mblnOk As Boolean
Private mblnPrinted As Boolean
Private mbytFunc As Byte

Public Sub ShowMe(ByVal frmParent As Object, ByVal rsTmp As ADODB.Recordset, Optional ByVal bytFunc As Byte)
'����:bytFunc=0 0-�ٴ�·������;1-�����ٴ�·������
    On Error Resume Next
    Set mrsTmp = rsTmp
    mbytFunc = bytFunc
    Me.Show , frmParent
End Sub

Public Sub BatPrint()
'���ܣ�������ӡ
    mblnPrinted = False
    If mlngCount = 0 Then Unload Me: Exit Sub
    
    If Not mrsTmp.EOF Then
        lblMsg.Caption = "����ӡ " & mlngCount & " ������," & "��ǰ���ڴ�ӡ�� " & mlngI & " �����ˡ�" & mrsTmp!���� & "����"
        If mbytFunc = 0 Then
            Call mfrmPath.zlRefresh(Val(mrsTmp!����ID & ""), Val(mrsTmp!��ҳID & ""), Val(mrsTmp!����ID & ""), Val(mrsTmp!����ID & ""), Val(mrsTmp!����״̬ & ""), Val(mrsTmp!����ת�� & "") = 1)
            Call mfrmPath.zlPrintOutPut(1, True)
        Else
            Call mfrmPathOut.zlRefresh(Val(mrsTmp!����ID & ""), Val(mrsTmp!�Һ�ID & ""), mrsTmp!NO & "", Val(mrsTmp!����ID & ""), Val(mrsTmp!����״̬ & ""), Val(mrsTmp!����ת�� & "") = 1)
            Call mfrmPathOut.zlPrintOutPut(1, True)
        End If
        picTime(1).Width = picTime(1).Width + (picTime(0).Width / mlngCount)
        Me.Refresh
        mlngI = mlngI + 1
    Else
        Unload Me
    End If
    mblnPrinted = True
End Sub

Private Sub cmdCancle_Click()
    '���������Կ��ǣ����ڴ�ӡʱ����ESC���˳���ӡ����������ͣ��ӡ
    If cmdStart.Tag = "Stop" Then
        Call cmdStart_Click
    Else
        If mlngI > 1 And mblnOk Then
            If MsgBox("�Ѿ���ʼ��ӡ��ȡ����ֹͣ��ӡ����Ĳ��ˣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                Unload Me
            End If
        Else
            Unload Me
        End If
    End If
End Sub

Private Sub cmdPreview_Click()
'Ԥ����ǰ��¼���Ĳ���·����
    If Not mrsTmp.EOF Then
        If mbytFunc = 0 Then
            Call mfrmPath.zlRefresh(Val(mrsTmp!����ID & ""), Val(mrsTmp!��ҳID & ""), Val(mrsTmp!����ID & ""), Val(mrsTmp!����ID & ""), Val(mrsTmp!����״̬ & ""), Val(mrsTmp!����ת�� & "") = 1)
            Call mfrmPath.zlPrintOutPut(2, True)
        Else
            Call mfrmPathOut.zlRefresh(Val(mrsTmp!����ID & ""), Val(mrsTmp!�Һ�ID & ""), mrsTmp!NO & "", Val(mrsTmp!����ID & ""), Val(mrsTmp!����״̬ & ""), Val(mrsTmp!����ת�� & "") = 1)
            Call mfrmPathOut.zlPrintOutPut(2, True)
        End If
    Else
        MsgBox "û�п�����Ĳ���·����", vbInformation, Me.Caption
    End If
End Sub

Private Sub cmdPrintSetup_Click()
    '��ӡ����
    Call zlPrintSet
End Sub

Private Sub cmdStart_Click()
     '��ʼ��ӡ��ť����Ϊ��ͣ��ӡ
    If cmdStart.Tag = "Start" Then
        Timer1.Enabled = True
        cmdStart.Tag = "Stop"
        cmdStart.Caption = "��ͣ��ӡ"
        mblnOk = True
    Else
        Timer1.Enabled = False
        cmdStart.Tag = "Start"
        cmdStart.Caption = "��ʼ��ӡ"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And cmdStart.Tag = "Stop" Then
        Call cmdStart_Click
    End If
End Sub

Private Sub Form_Load()
    Dim tabItem As TabControlItem
    If mbytFunc = 0 Then
        Set mfrmPath = New frmPathTable
    Else
        Set mfrmPathOut = New frmPathTableOut
    End If
    picTime(1).Width = 0
    mlngI = 1
    mlngCount = mrsTmp.RecordCount
    mrsTmp.MoveFirst
    cmdStart.Tag = "Start"
    mblnOk = False
    mblnPrinted = True
    lblMsg.Caption = "����ӡ " & mlngCount & " ������," & "�Ƿ�Ҫ��ʼ��ӡ��Щ���˵�·����"
    
    'TabControl
    '-----------------------------------------------------
    With Me.tbcPath
        With .PaintManager
            .Appearance = xtpTabAppearanceVisio
            .Color = xtpTabColorOffice2003
        End With
        If mbytFunc = 0 Then
            Set tabItem = .InsertItem(0, "�����ٴ�·��", mfrmPath.Hwnd, 0)
        Else
            Set tabItem = .InsertItem(0, "��������·��", mfrmPathOut.Hwnd, 0)
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mbytFunc = 0 Then
        Unload mfrmPath
    Else
        Unload mfrmPathOut
    End If
    Set mfrmPath = Nothing
    Set mrsTmp = Nothing
End Sub

Private Sub Timer1_Timer()
    '������ϴδ�ӡ��ɺ��ٿ�ʼ��һ�����˵Ĵ�ӡ
    If mblnPrinted Then
        Call BatPrint
        If Not mrsTmp Is Nothing Then mrsTmp.MoveNext
    End If
End Sub
