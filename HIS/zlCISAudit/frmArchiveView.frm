VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmArchiveView 
   AutoRedraw      =   -1  'True
   Caption         =   "���Ӳ�������"
   ClientHeight    =   8460
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11760
   Icon            =   "frmArchiveView.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   11760
   Begin VB.ComboBox cboDept 
      Height          =   300
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   0
      Width           =   3165
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   3765
      ScaleHeight     =   975
      ScaleWidth      =   7695
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   135
      Width           =   7695
      Begin VB.Frame fraInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " ����������Ϣ "
         ForeColor       =   &H80000008&
         Height          =   840
         Left            =   60
         TabIndex        =   3
         Top             =   75
         Width           =   7500
         Begin VB.Frame fraIn 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   195
            TabIndex        =   22
            Top             =   255
            Visible         =   0   'False
            Width           =   7170
            Begin VB.Label lbl����zy 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   4770
               TabIndex        =   40
               Top             =   0
               Width           =   1080
            End
            Begin VB.Label lbl����zy 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����:"
               Height          =   180
               Index           =   0
               Left            =   4305
               TabIndex        =   39
               Top             =   0
               Width           =   450
            End
            Begin VB.Label lblסԺ��zy 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "סԺ��:"
               Height          =   180
               Index           =   0
               Left            =   1560
               TabIndex        =   38
               Top             =   0
               Width           =   630
            End
            Begin VB.Label lbl����zy 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����:"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   37
               Top             =   0
               Width           =   450
            End
            Begin VB.Label lbl����zy 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����:"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   36
               Top             =   255
               Width           =   450
            End
            Begin VB.Label lbl����zy 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����:"
               Height          =   180
               Index           =   0
               Left            =   3150
               TabIndex        =   35
               Top             =   0
               Width           =   450
            End
            Begin VB.Label lblҽ����zy 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ҽ����:"
               Height          =   180
               Index           =   0
               Left            =   5940
               TabIndex        =   34
               Top             =   0
               Width           =   630
            End
            Begin VB.Label lbl��Ժzy 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժ:"
               Height          =   180
               Index           =   0
               Left            =   4305
               TabIndex        =   33
               Top             =   255
               Width           =   450
            End
            Begin VB.Label lbl����zy 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����:"
               Height          =   180
               Index           =   0
               Left            =   3150
               TabIndex        =   32
               Top             =   255
               Width           =   450
            End
            Begin VB.Label lbl����zy 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��  ��:"
               Height          =   180
               Index           =   0
               Left            =   1560
               TabIndex        =   31
               Top             =   255
               Width           =   630
            End
            Begin VB.Label lbl����zy 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   2190
               TabIndex        =   30
               Top             =   255
               Width           =   900
            End
            Begin VB.Label lbl����zy 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H000000FF&
               Height          =   180
               Index           =   1
               Left            =   3585
               TabIndex        =   29
               Top             =   255
               Width           =   675
            End
            Begin VB.Label lbl��Ժzy 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   4770
               TabIndex        =   28
               Top             =   255
               Width           =   90
            End
            Begin VB.Label lblҽ����zy 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00008000&
               Height          =   180
               Index           =   1
               Left            =   6600
               TabIndex        =   27
               Top             =   0
               Width           =   90
            End
            Begin VB.Label lbl����zy 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   3585
               TabIndex        =   26
               Top             =   0
               Width           =   675
            End
            Begin VB.Label lbl����zy 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   435
               TabIndex        =   25
               Top             =   255
               Width           =   1080
            End
            Begin VB.Label lbl����zy 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   435
               TabIndex        =   24
               Top             =   0
               Width           =   1080
            End
            Begin VB.Label lblסԺ��zy 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   2190
               TabIndex        =   23
               Top             =   0
               Width           =   900
            End
         End
         Begin VB.Frame fraOut 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   195
            TabIndex        =   4
            Top             =   255
            Visible         =   0   'False
            Width           =   7170
            Begin VB.Label lbl�� 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   21.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   435
               Left            =   6705
               TabIndex        =   21
               Top             =   0
               Visible         =   0   'False
               Width           =   435
            End
            Begin VB.Label lbl�Һŵ�mz 
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   3870
               TabIndex        =   20
               Top             =   0
               Width           =   1065
            End
            Begin VB.Label lbl�Һŵ�mz 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�Һŵ�:"
               Height          =   180
               Index           =   0
               Left            =   3255
               TabIndex        =   19
               Top             =   0
               Width           =   630
            End
            Begin VB.Label lblҽ��mz 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   2385
               TabIndex        =   18
               Top             =   0
               Width           =   780
            End
            Begin VB.Label lblҽ��mz 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ҽ��:"
               Height          =   180
               Index           =   0
               Left            =   1935
               TabIndex        =   17
               Top             =   0
               Width           =   450
            End
            Begin VB.Label lbl������mz 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00008000&
               Height          =   180
               Index           =   1
               Left            =   5655
               TabIndex        =   16
               Top             =   255
               Width           =   90
            End
            Begin VB.Label lbl������mz 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������:"
               Height          =   180
               Index           =   0
               Left            =   5025
               TabIndex        =   15
               Top             =   255
               Width           =   630
            End
            Begin VB.Label lbl�����mz 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�����:"
               Height          =   180
               Index           =   0
               Left            =   3240
               TabIndex        =   14
               Top             =   255
               Width           =   630
            End
            Begin VB.Label lbl����mz 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����:"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   13
               Top             =   0
               Width           =   450
            End
            Begin VB.Label lbl�ѱ�mz 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ѱ�:"
               Height          =   180
               Index           =   0
               Left            =   1935
               TabIndex        =   12
               Top             =   255
               Width           =   450
            End
            Begin VB.Label lblҽ����mz 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ҽ����:"
               Height          =   180
               Index           =   0
               Left            =   5025
               TabIndex        =   11
               Top             =   0
               Width           =   630
            End
            Begin VB.Label lblҽ����mz 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00008000&
               Height          =   180
               Index           =   1
               Left            =   5655
               TabIndex        =   10
               Top             =   0
               Width           =   90
            End
            Begin VB.Label lbl�ѱ�mz 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   2385
               TabIndex        =   9
               Top             =   255
               Width           =   765
            End
            Begin VB.Label lbl����mz 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   450
               TabIndex        =   8
               Top             =   0
               Width           =   1425
            End
            Begin VB.Label lbl�����mz 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   3870
               TabIndex        =   7
               Top             =   255
               Width           =   1095
            End
            Begin VB.Label lbl����mz 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����:"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   6
               Top             =   255
               Width           =   450
            End
            Begin VB.Label lbl����mz 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   450
               TabIndex        =   5
               Top             =   255
               Width           =   1455
            End
         End
      End
   End
   Begin VB.Frame fraLR 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2745
      Left            =   3660
      MousePointer    =   9  'Size W E
      TabIndex        =   1
      Top             =   1515
      Width           =   45
   End
   Begin XtremeSuiteControls.TabControl tbcArchive 
      Height          =   6315
      Left            =   3900
      TabIndex        =   0
      Top             =   1605
      Width           =   7365
      _Version        =   589884
      _ExtentX        =   12991
      _ExtentY        =   11139
      _StockProps     =   64
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":058A
            Key             =   "סԺ"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":6DEC
            Key             =   "����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":7386
            Key             =   "object_report"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":7920
            Key             =   "object_case"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":7EBA
            Key             =   "object_tend"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":8454
            Key             =   "object_first"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":89EE
            Key             =   "object_advice"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":8F88
            Key             =   "object_file"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":9522
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":FD84
            Key             =   "Path"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwArchive 
      Height          =   2775
      Left            =   615
      TabIndex        =   41
      Top             =   3585
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   4895
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   0
   End
   Begin XtremeSuiteControls.TabControl tbcHistory 
      Height          =   4155
      Left            =   495
      TabIndex        =   42
      Top             =   2925
      Width           =   3210
      _Version        =   589884
      _ExtentX        =   5662
      _ExtentY        =   7329
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmArchiveView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobjRichEMR As Object
Private mobjPACSDoc As Object
Private mfrmOutMedRec As frmArchiveOutMedRec
Private mfrmInMedRec As frmArchiveInMedRec

'������ҳ
Private mclsArchiveMedRec As zlMedRecPage.clsArchive
Private mfrmArchiveMedRec As Object
Private mclsOutAdvices As clsDockOutAdvices
Private mclsInAdvices As clsDockInAdvices
Private mclsDockAduits As clsDockAduits
Private mclsPath As clsDockPath
Private WithEvents mclsTendsNew As zl9TendFile.clsTendFile    '�°滤ʿ����վ
Attribute mclsTendsNew.VB_VarHelpID = -1

Private mlng����ID  As Long
Private mlng����ID As Long '���˵�ǰ�������ľ���ID������Ϊ�Һ�ID,סԺ����ҳID
Private mstr�Һŵ� As String
Private mlng����ID As Long
Private mlng����ID As Long
Private mblnMoved As Boolean
Private mblnNewTends As Boolean
Private mrsData         As ADODB.Recordset

Public Sub ShowArchive(ByVal frmParent As Object, ByVal lng����ID As Long, ByVal lng����ID As Long, Optional ByVal blnModal As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lngIdx As Long
    Dim objTab As TabControlItem

    mlng����ID = lng����ID
    mlng����ID = lng����ID
    mstr�Һŵ� = ""
    mblnMoved = False
    mblnNewTends = False

    With tbcHistory
        Screen.MousePointer = 11
        LockWindowUpdate Me.hWnd
        Me.cboDept.Clear
        .RemoveAll

        On Error GoTo ErrH

        strSQL = _
            " Select ID as ����ID,NO,����ʱ�� as ��ʼʱ��,Null as ����ʱ��,ִ�в���ID as ����ID,0 as ����ת�� From ���˹Һż�¼ Where ����ID=[1]" & _
            " Union ALL" & _
            " Select ID as ����ID,NO,����ʱ�� as ��ʼʱ��,Null as ����ʱ��,ִ�в���ID as ����ID,1 as ����ת�� From H���˹Һż�¼ Where ����ID=[1]" & _
            " Union ALL" & _
            " Select ��ҳID as ����ID,Null,��Ժ���� as ��ʼʱ��,��Ժ����,��Ժ����ID,����ת�� From ������ҳ Where ����ID=[1] And Nvl(��ҳID,0)<>0"
        strSQL = "Select A.����ID,A.NO,A.��ʼʱ��,A.����ʱ��,B.���� as ����,A.����ת�� From (" & strSQL & ") A,���ű� B Where A.����ID=B.ID Order by ��ʼʱ�� Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
        Set mrsData = rsTmp
        Do While Not rsTmp.EOF
            If rsTmp.AbsolutePosition = 1 Then
                Set objTab = .InsertItem(.ItemCount, _
                    IIf(IsNull(rsTmp!NO), "��" & rsTmp!����ID & "��סԺ", "�������") & ":" & _
                    rsTmp!���� & "," & Format(rsTmp!��ʼʱ��, "yyyy-MM-dd HH:mm") & _
                    IIf(Not IsNull(rsTmp!����ʱ��), "��" & Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm"), ""), _
                    tvwArchive.hWnd, IIf(IsNull(rsTmp!NO), 0, 1))

                objTab.Tag = rsTmp!����ID & "," & NVL(rsTmp!NO) & "," & NVL(rsTmp!����ת��, 0)
            End If
               cboDept.AddItem IIf(IsNull(rsTmp!NO), "��" & rsTmp!����ID & "��סԺ", "�������") & ":" & _
                    rsTmp!���� & "," & Format(rsTmp!��ʼʱ��, "yyyy-MM-dd HH:mm") & _
                    IIf(Not IsNull(rsTmp!����ʱ��), "��" & Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm"), "")
               cboDept.ItemData(cboDept.NewIndex) = Val(rsTmp!����ID)
            rsTmp.MoveNext
        Loop
        cboDept.ListIndex = 0
        LockWindowUpdate 0
        Screen.MousePointer = 0
        '��Ϊǿ�Ƽ����¼�:ֻ��һ��ʱҲ�����Զ�����
        Call cboDept_Click
    End With

    If Me.WindowState = vbMinimized Then
        Me.WindowState = vbNormal
    End If

    Me.Show IIf(blnModal, 1, 0), frmParent

    Exit Sub
ErrH:
    LockWindowUpdate 0
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboDept_Click()
    Dim strTemp As String
    If cboDept.Text = "" Then Exit Sub
    strTemp = cboDept.ItemData(cboDept.ListIndex)
    mlng����ID = strTemp
    mrsData.Filter = "����ID='" & strTemp & "'"
    If Not mrsData.EOF Then
        mstr�Һŵ� = NVL(mrsData!NO, "")
        mblnMoved = Val(NVL(mrsData!����ת��, "")) = 1
    End If
    '��ʾ������Ϣ
    If mstr�Һŵ� <> "" Then
        Call ShowOutPatiInfo
    Else
        Call ShowInPatiInfo
    End If
    fraOut.Visible = mstr�Һŵ� <> ""
    fraIn.Visible = mstr�Һŵ� = ""

    '��ʾ����Ŀ¼
    Me.tbcHistory(0).Caption = cboDept.Text
    Call ShowArchiveTree
    If tvwArchive.Visible And tvwArchive.Enabled Then tvwArchive.SetFocus
    Call Form_Resize
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim objTab As TabControlItem
    Dim frmTendBody As Object

    picInfo.BackColor = fraLR.BackColor
    fraInfo.BackColor = picInfo.BackColor
    fraIn.BackColor = picInfo.BackColor
    fraOut.BackColor = picInfo.BackColor
    '�Ӵ���
    '-----------------------------------------------------
    Set mfrmOutMedRec = New frmArchiveOutMedRec
    Set mfrmInMedRec = New frmArchiveInMedRec

    '��ʼCISJOB��ҳ�ӿ�
    Set mclsArchiveMedRec = New zlMedRecPage.clsArchive
    Call mclsArchiveMedRec.InitArchiveMedRec(gcnOracle, glngSys)
    Set mfrmArchiveMedRec = mclsArchiveMedRec.zlGetForm(1)

    '��ʼ����
    '------------------------------------------------------------------------------------------------------------------
    If Not gobjEmr Is Nothing Then
        If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then
            Set gobjEmr = Nothing
        Else
            Set mobjRichEMR = DynamicCreate("zlRichEMR.clsDockContent", "�°没��", False)
            If Not mobjRichEMR Is Nothing Then Call mobjRichEMR.Init(gobjEmr, gcnOracle, glngSys, 0)
        End If
    End If
    Set mobjPACSDoc = DynamicCreate("zlPublicPACS.clsPublicPacs", "�°�PACS�༭��", False)
    If Not mobjPACSDoc Is Nothing Then
        Call mobjPACSDoc.InitInterface(gcnOracle, gstrDBUser)
    End If

    Set mclsOutAdvices = New clsDockOutAdvices
    Set mclsInAdvices = New clsDockInAdvices
    Set mclsDockAduits = New clsDockAduits
    Set mclsPath = New clsDockPath

    Set mclsTendsNew = New zl9TendFile.clsTendFile
    Call mclsTendsNew.InitTendFile(gcnOracle, glngSys)

    Set frmTendBody = mclsDockAduits.zlGetFormTendBody
    Call FormSetCaption(frmTendBody, False, False)

    With tbcArchive
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .COLOR = xtpTabColorOffice2003
            .Layout = xtpTabLayoutAutoSize
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
        End With
        Set objTab = .InsertItem(0, "������ҳ", mfrmOutMedRec.hWnd, 0): objTab.Tag = objTab.Caption: objTab.Visible = False
        Set objTab = .InsertItem(1, "סԺ��ҳ", mfrmArchiveMedRec.hWnd, 0): objTab.Tag = objTab.Caption: objTab.Visible = False
        Set objTab = .InsertItem(2, "������Ϣ", mclsDockAduits.zlGetFormEPR.hWnd, 0): objTab.Tag = objTab.Caption: objTab.Visible = False
        Set objTab = .InsertItem(3, "����ҽ��", mclsOutAdvices.zlGetForm.hWnd, 0): objTab.Tag = objTab.Caption: objTab.Visible = False
        Set objTab = .InsertItem(4, "סԺҽ��", mclsInAdvices.zlGetForm.hWnd, 0): objTab.Tag = objTab.Caption: objTab.Visible = False
        Set objTab = .InsertItem(5, "���¼�¼��", frmTendBody.hWnd, 0): objTab.Tag = objTab.Caption: objTab.Visible = False
        Set objTab = .InsertItem(6, "�����¼��", mclsDockAduits.zlGetFormTendFile.hWnd, 0): objTab.Tag = objTab.Caption: objTab.Visible = False
        Set objTab = .InsertItem(7, "�ٴ�·��", mclsPath.zlGetForm.hWnd, 0): objTab.Tag = objTab.Caption: objTab.Visible = False
        Set objTab = .InsertItem(8, "�°滤��", mclsTendsNew.zlGetfrmInTendFile.hWnd, 0): objTab.Tag = objTab.Caption: objTab.Visible = False
        If Not mobjRichEMR Is Nothing Then
            Set objTab = .InsertItem(9, "���Ӳ���", mobjRichEMR.zlGetForm.hWnd, 0): objTab.Tag = objTab.Caption: objTab.Visible = False
        End If
        If Not mobjPACSDoc Is Nothing Then
            Set objTab = .InsertItem(10, "��鱨��", mobjPACSDoc.zlDocGetForm.hWnd, 0): objTab.Tag = objTab.Caption: objTab.Visible = False
        End If
    End With

    '������ʷ
    '-----------------------------------------------------
    With tbcHistory
        With .PaintManager
            .Appearance = xtpTabAppearanceVisio
            .COLOR = xtpTabColorOffice2003
            .DisableLunaColors = False
            .BoldSelected = True
            .HotTracking = True
            .ShowIcons = True
        End With
        .SetImageList ils16
    End With

    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
   On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    Me.cboDept.Width = tbcHistory.Width
    Me.tbcHistory.Top = cboDept.Height
    Me.tbcHistory.Left = 0
    
    Me.tbcHistory.Height = Me.ScaleHeight - cboDept.Height

    Me.fraLR.Top = 0
    Me.fraLR.Left = Me.tbcHistory.Width
    Me.fraLR.Height = Me.ScaleHeight

    Me.picInfo.Top = 0
    Me.picInfo.Left = Me.fraLR.Left + Me.fraLR.Width
    Me.picInfo.Width = Me.ScaleWidth - Me.tbcHistory.Width - Me.fraLR.Width

    Me.tbcArchive.Left = Me.fraLR.Left + Me.fraLR.Width
    Me.tbcArchive.Top = Me.picInfo.Top + Me.picInfo.Height
    Me.tbcArchive.Width = Me.ScaleWidth - Me.tbcHistory.Width - Me.fraLR.Width
    Me.tbcArchive.Height = Me.ScaleHeight - Me.picInfo.Height

    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Call SaveWinState(Me, App.ProductName)

    Unload mfrmOutMedRec: Set mfrmOutMedRec = Nothing
    Unload mfrmInMedRec:  Set mfrmInMedRec = Nothing
    Unload mobjRichEMR.zlGetForm: Set mobjRichEMR.zlGetForm = Nothing
    Unload mclsOutAdvices.zlGetForm: Set mclsOutAdvices.zlGetForm = Nothing
    Unload mobjPACSDoc.zlDocGetForm: Set mobjPACSDoc.zlDocGetForm = Nothing

    Set mclsOutAdvices = Nothing
    Unload mclsInAdvices.zlGetForm: Set mclsInAdvices.zlGetForm = Nothing
    Set mclsInAdvices = Nothing
    Unload mclsDockAduits.zlGetFormEPR: Set mclsDockAduits.zlGetFormEPR = Nothing
    Unload mclsDockAduits.zlGetFormTendFile: Set mclsDockAduits.zlGetFormTendFile = Nothing
    Unload mclsDockAduits.zlGetFormTendBody: Set mclsDockAduits.zlGetFormTendBody = Nothing
    Set mclsDockAduits = Nothing
    Unload mclsTendsNew.zlGetfrmInTendFile: Set mclsTendsNew.zlGetfrmInTendFile = Nothing
    Set mclsTendsNew = Nothing
    Set mobjRichEMR = Nothing
    Set mobjPACSDoc = Nothing
    Unload mclsPath.zlGetForm:  Set mclsPath.zlGetForm = Nothing
    Set mclsPath = Nothing

    Unload mfrmArchiveMedRec: Set mfrmArchiveMedRec = Nothing
    Set mclsArchiveMedRec = Nothing
    Set mrsData = Nothing
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next

    If Button = 1 Then
        If fraLR.Left + X < 1000 Or fraLR.Left + X > Me.ScaleWidth - 3000 Then
        Exit Sub
        End If
        Me.tbcHistory.Width = tbcHistory.Width + X
        Call Form_Resize
    End If
End Sub



Private Sub picInfo_Resize()
    On Error Resume Next
    fraInfo.Width = picInfo.ScaleWidth - fraInfo.Left * 3

    fraIn.Width = fraInfo.Width - fraIn.Left * 2
    fraOut.Width = fraIn.Width
    lbl��.Left = fraOut.Width - lbl��.Width - 60
End Sub

Private Sub tbcHistory_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If tbcHistory.Tag = "don't refresh" Then Exit Sub
    If Item.Tag = "" Then Exit Sub

    mlng����ID = Val(Split(Item.Tag, ",")(0))
    mstr�Һŵ� = Split(Item.Tag, ",")(1)
    mblnMoved = Val(Split(Item.Tag, ",")(2)) = 1

    '��ʾ������Ϣ
    If mstr�Һŵ� <> "" Then
        Call ShowOutPatiInfo
    Else
        Call ShowInPatiInfo
    End If
    fraOut.Visible = mstr�Һŵ� <> ""
    fraIn.Visible = mstr�Һŵ� = ""

    '��ʾ����Ŀ¼
    Call ShowArchiveTree
    If tvwArchive.Visible And tvwArchive.Enabled Then tvwArchive.SetFocus
'    Dim i As Integer
'    For i = 0 To tbcHistory.ItemCount - 1
'        tbcHistory.Height = tbcHistory.Height + 225
'    Next
    Call Form_Resize
End Sub

Private Sub ShowArchiveTab(ByVal strShow As String, ByVal strCaption As String)
'���ܣ��л���ʾ��ͬ�ĵ���ҳ��
    Dim i As Long

    For i = 0 To tbcArchive.ItemCount - 1
        If tbcArchive(i).Tag = strShow Then
            tbcArchive(i).Caption = strCaption
            If Not tbcArchive(i).Visible Then
                tbcArchive(i).Visible = True
                tbcArchive(i).Selected = True
            End If
        End If
    Next
    For i = 0 To tbcArchive.ItemCount - 1
        If tbcArchive(i).Tag <> strShow Then
            If tbcArchive(i).Visible Then
                tbcArchive(i).Visible = False
            End If
        End If
    Next
End Sub

Private Sub tvwArchive_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim arrPar As Variant
    Dim intSel As Integer

    If tvwArchive.Tag = Node.Key Then Exit Sub

    LockWindowUpdate Me.hWnd

    arrPar = Split(Node.Tag, ";")

    If Node.Key = "R11" Then
        If mstr�Һŵ� <> "" Then '��ҳ��Ϣ
            Call mfrmOutMedRec.zlRefresh(mlng����ID, mlng����ID, mblnMoved)
            Call ShowArchiveTab("������ҳ", tbcHistory.Selected.Caption)
        Else
'            Call mfrmInMedRec.zlRefresh(mlng����ID, mlng����ID, 0, 0, False, mblnMoved)
'            Call ShowArchiveTab("סԺ��ҳ", tbcHistory.Selected.Caption)

            Call mclsArchiveMedRec.zlRefresh(1, mlng����ID, mlng����ID, mblnMoved)
            Call ShowArchiveTab("סԺ��ҳ", tbcHistory.Selected.Caption)


        End If
    ElseIf Node.Key = "R12" Then 'ҽ����¼
        If mstr�Һŵ� <> "" Then
            Call mclsOutAdvices.zlRefresh(mlng����ID, mstr�Һŵ�, False, mblnMoved)
            Call ShowArchiveTab("����ҽ��", tbcHistory.Selected.Caption)
        Else
            Call mclsInAdvices.zlRefresh(mlng����ID, mlng����ID, mlng����ID, mlng����ID, 0, mblnMoved)
            Call ShowArchiveTab("סԺҽ��", tbcHistory.Selected.Caption)
        End If
    ElseIf Node.Key Like "R1K*" Then '���ﲡ��
        Call mclsDockAduits.zlRefresh(1, Val(arrPar(0)))
        Call ShowArchiveTab("������Ϣ", Node.Text)
    ElseIf Node.Key Like "R2K*" Then 'סԺ����
        Call mclsDockAduits.zlRefresh(2, Val(arrPar(0)))
        Call ShowArchiveTab("������Ϣ", Node.Text)
    ElseIf Node.Key Like "R3K*" Then '�����¼
        If UBound(arrPar) >= 1 Then
            If mblnNewTends = False Then
                If Val(arrPar(1)) = -1 Then
                    Call mclsDockAduits.zlRefreshTendBody(mlng����ID, mlng����ID, Val(Split(arrPar(0), "_")(0)), Val(arrPar(4)))
                    Call ShowArchiveTab("���¼�¼��", Node.Text)
                Else
                    Call mclsDockAduits.zlRefresh(3, Val(arrPar(3)), mlng����ID, mlng����ID, Val(Split(arrPar(0), "_")(0)), CStr(arrPar(2)), , Val(arrPar(4)))
                    Call ShowArchiveTab("�����¼��", Node.Text)
                End If
            Else
                '�˲������� ����
                Select Case Val(arrPar(1))
                    Case -1 '���µ�
                        intSel = 0
                    Case 1  '����ͼ
                        intSel = 2
                    Case Else '��¼��
                        intSel = 1
                End Select
                Call mclsTendsNew.zlRefreshTendFile(mlng����ID, mlng����ID, Val(arrPar(4)), Val(arrPar(0)), False, False, intSel, Val(arrPar(3)), 1)
                Call ShowArchiveTab("�°滤��", Node.Text)
            End If
        End If
    ElseIf Node.Key Like "R4K*" Then '������
        Call mclsDockAduits.zlRefresh(4, Val(arrPar(0)))
        Call ShowArchiveTab("������Ϣ", Node.Text)
    ElseIf Node.Key Like "R5K*" Then '����֤��
        Call mclsDockAduits.zlRefresh(5, Val(arrPar(0)))
        Call ShowArchiveTab("������Ϣ", Node.Text)
    ElseIf Node.Key Like "R6K*" Then '֪���ļ�
        Call mclsDockAduits.zlRefresh(6, Val(arrPar(0)))
        Call ShowArchiveTab("������Ϣ", Node.Text)
    ElseIf Node.Key Like "R7K*" Then '���Ʊ���
        If Val(arrPar(0)) <> 0 Then
			Call mclsDockAduits.zlRefresh(7, Val(arrPar(0)))
			Call ShowArchiveTab("������Ϣ", Node.Text)
        Else                    '�°�PACS���� ����=0;ҽ��ID;��鱨��ID
            Call mobjPACSDoc.zlDocRefresh(arrPar(2))
            Call ShowArchiveTab("��鱨��", Node.Text)
        End If
    ElseIf Node.Key = "R8" Then
        If mstr�Һŵ� = "" Then
            Call mclsPath.zlRefreshReadOnly(mlng����ID, mlng����ID)
            Call ShowArchiveTab("�ٴ�·��", Node.Text)
        End If
    ElseIf InStr(Node.Key, "R") = 0 And Len(Node.Tag) >= 32 Then
        'EMR����Ԥ��
        If Not mobjRichEMR Is Nothing Then
            If InStr(Node.Tag, "|") > 0 Then
                Call mobjRichEMR.zlShowDoc(Split(Node.Tag, "|")(0), Split(Node.Tag, "|")(1))
            Else
                Call mobjRichEMR.zlShowDoc(Node.Tag, "")
            End If
            Call ShowArchiveTab("���Ӳ���", Node.Text)
        End If
    Else
        LockWindowUpdate 0
        Exit Sub
    End If

    tvwArchive.Tag = Node.Key

    LockWindowUpdate 0
    If tvwArchive.Visible And tvwArchive.Enabled Then tvwArchive.SetFocus
End Sub

Private Function ShowArchiveTree() As Boolean
'���ܣ���ʾ���˵�������Ŀ¼
    Dim rsTmp As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, objNode As Node, strSQL1 As String
    Dim strSel As String, blnPath As Boolean

    Screen.MousePointer = 11

    On Error GoTo ErrH

    If Not tvwArchive.SelectedItem Is Nothing Then
        If tvwArchive.SelectedItem.Key = "R11" Or tvwArchive.SelectedItem.Key = "R12" Then
            strSel = Split(tvwArchive.SelectedItem.Key, "K")(0)
        End If
    End If

    '���˿��Ҵ��ڿ��õ��ٴ�·��ʱ����ʾ�ٴ�·����¼
    '���ж��Ƿ���"�ٴ�·��Ӧ��" ���=1256
    If mstr�Һŵ� = "" Then
        If GetPrivFunc(glngSys, 1256) <> "" Then
            blnPath = gclsPackage.GetHavePath(mlng����ID)
        End If
    End If

    '1-���ﲡ��;2-סԺ����;3-�����¼;4-������;5-����֤��;6-֪���ļ�;7-���Ʊ���,11-��ҳ��Ϣ,12-ҽ����¼
    strSQL = _
        " Select * From (" & _
            " Select 'R11' As ID, '' As �ϼ�id, '��ҳ��Ϣ' As ����, '' As ����,1 As ĩ��,'object_first' As ͼ��,'01' As ���� From Dual Union All" & _
            " Select 'R12' As ID, '' As �ϼ�id, 'ҽ����¼' As ����, '' As ����,1 As ĩ��,'object_advice' As ͼ��,'02' As ���� From Dual Union All" & _
            " Select 'R1' As ID, '' As �ϼ�id, '���ﲡ��' As ����, '' As ����,0 As ĩ��,'Folder' As ͼ��,'03' As ���� From Dual Where [3]=0 Union All" & _
            " Select 'R2' As ID, '' As �ϼ�id, 'סԺ����' As ����, '' As ����,0 As ĩ��,'Folder' As ͼ��,'04' As ���� From Dual Where [3]=1 Union All" & _
            " Select 'R3' As ID, '' As �ϼ�id, '�����¼' As ����, '' As ����,0 As ĩ��,'Folder' As ͼ��,'05' As ���� From Dual Where [3]=1 Union All" & _
            " Select 'R4' As ID, '' As �ϼ�id, '������' As ����, '' As ����,0 As ĩ��,'Folder' As ͼ��,'06' As ���� From Dual Where [3]=1 Union All" & _
            " Select 'R7' As ID, '' As �ϼ�id, '���Ʊ���' As ����, '' As ����,0 As ĩ��,'Folder' As ͼ��,'07' As ���� From Dual Union All" & _
            " Select 'R5' As ID, '' As �ϼ�id, '����֤��' As ����, '' As ����,0 As ĩ��,'Folder' As ͼ��,'08' As ���� From Dual Union All" & _
            " Select 'R6' As ID, '' As �ϼ�id, '֪���ļ�' As ����, '' As ����,0 As ĩ��,'Folder' As ͼ��,'09' As ���� From Dual" & _
            IIf(blnPath, " Union All Select 'R8' As ID, '' As �ϼ�id, '�ٴ�·��' As ����, '' As ����,0 As ĩ��,'Path' As ͼ��,'10' As ���� From Dual", "")
    '��������
    'ID=�ϼ�ID+K����ID,ҽ��ID,0
    '����=����ID;ҽ��ID
    strSQL = strSQL & " Union All" & _
        " Select A.�ϼ�id||'K'||Trim(To_Char(A.ID))||','||Trim(To_Char(Nvl(A.ҽ��id,0)))||',0' As ID,A.�ϼ�id," & _
        "       Decode(A.ҽ��id,Null,A.����||'('||To_Char(A.����ʱ��, 'YYYY-MM-DD')||')',A.����||'��'||B.ҽ������||'('||To_Char(A.����ʱ��, 'YYYY-MM-DD')||')') As ����," & _
        "       Trim(To_Char(A.ID))||';'||Decode(A.ҽ��id,Null,'0',Trim(To_Char(A.ҽ��id))) As ����," & _
        "       1 As ĩ��,Decode(��������,1,'object_case',2,'object_case',4,'object_case',7,'object_report','object_file') As ͼ��,���� " & _
        " From (Select A.ID, 'R'||A.�������� As �ϼ�id, A.�������� As ����,C.ҽ��id,A.��������,A.����ʱ��,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') As ����" & _
        "       From ���Ӳ�����¼ A,����ҽ������ C " & _
        "       Where A.����id = [1] And A.��ҳid = [2] And (A.������Դ=2 And [3]=1 Or Nvl(A.������Դ,0)<>2 And [3]=0)" & _
        "           And C.����id(+)=A.ID And A.�������� In (1, 2, 3, 4, 5, 6, 7)" & _
        "       ) A,����ҽ����¼ B Where A.ҽ��id=B.Id(+)"

     strSQL = strSQL & " Union All " & vbNewLine & _
                        "Select ID, �ϼ�id, ����, ����,ĩ��, ͼ��, ����" & vbNewLine & _
                        "From (Select 'R7K' || RawtoHex(b.��鱨��id) ID, 'R7' �ϼ�id, '<' || a.ҽ������ || '>' ����, '0;' || a.Id || ';' || b.��鱨��id ����,1 ĩ��, 'object_report' ͼ��," & vbNewLine & _
                        "       To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����" & vbNewLine & _
                            "From ����ҽ����¼ A, ����ҽ������ B" & vbNewLine & _
                            "Where a.����id = [1] And a.��ҳid = [2] And a.Id = b.ҽ��id And RawtoHex(b.��鱨��id) Is Not Null" & vbNewLine & _
                            "Order By a.����ʱ��)"

    '������
    'ID=�ϼ�ID+K�ļ�ID,0,����ID
    '����=����ID;����;��ʼ����ֹ;�ļ�ID
    strSQL1 = "Select 1 From ���˻����¼ A Where a.����id = [1] And a.��ҳid = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL1, "����Ƿ�����ϰ�����", mlng����ID, mlng����ID)
    If rsTemp.RecordCount > 0 Then
        mblnNewTends = False
        strSQL = strSQL & " Union All" & _
            " Select 'R3K'||Trim(To_Char(A.ID))||',0,'||Trim(To_Char(A.����Id)) || ',' || Nvl(a.Ӥ��,0) As ID,'R3' As �ϼ�id," & _
            "       A.����||'('||B.����||'��'||To_Char(A.��ʼ, 'YYYY-MM-DD HH24:MI') || '��' ||To_Char(A.��ֹ, 'YYYY-MM-DD HH24:MI') || ')' As ����," & _
            "       Trim(To_Char(B.ID))||';'||Trim(To_Char(Nvl(����,0)))||';'||To_Char(A.��ʼ, 'YYYY-MM-DD HH24:MI')||'��'||To_Char(A.��ֹ, 'YYYY-MM-DD HH24:MI')||';'||Trim(To_Char(A.ID)) ||';'|| Trim(To_Char(A.Ӥ��)) As ����," & _
            "       1 As ĩ��,'object_tend' As ͼ��,To_Char(a.��ʼ,'YYYY-MM-DD HH24:MI:SS') As ����" & _
            " From (" & _
            "   Select F.ID, F.���, F.����, R.��ʼ, R.��ֹ, R.����id, f.����,R.Ӥ��" & _
            "   From (" & _
            "       Select ID, ���, ����, 3 As ������, ͨ��, 0 As ����id, ����" & _
            "          From �����ļ��б� Where ���� = 3 And ���� < 0" & _
            "       Union All" & _
            "       Select L.ID, L.���, L.����, F.���� As ������, L.ͨ��, A.����id, L.����" & _
            "          From ����ҳ���ʽ F, �����ļ��б� L, ����Ӧ�ÿ��� A" & _
            "          Where L.���� = 3 And L.���� = 0 And L.���� = F.���� And L.��� = F.��� And L.ID = A.�ļ�id(+)" & _
            "       ) F,(" & _
            "       Select R.����id, R.Ӥ��,Nvl(Min(R.������), 3) As ������, Min(R.����ʱ��) As ��ʼ, Max(R.����ʱ��) As ��ֹ" & _
            "          From ���˻����¼ R" & _
            "          Where R.������Դ = 2 And R.����id = [1] And Nvl(R.��ҳid, 0) = [2]" & _
            "          Group By R.����id,R.Ӥ��" & _
            "       ) R" & _
            "       Where (F.ͨ�� = 1 Or F.ͨ�� = 2 And F.����id = R.����id) And F.������ >= R.������" & _
            "   ) A, ���ű� B Where A.����id = B.ID)" & _
            "Order By Decode(�ϼ�id,Null,' ',�ϼ�id),����"
    Else
        mblnNewTends = True
        strSQL = strSQL & " Union All" & _
                " Select 'R3K'||Trim(To_Char(A.ID))||',0,'||Trim(To_Char(A.����Id)) As ID,'R3' As �ϼ�id," & vbNewLine & _
                "     A.����||'('||B.����||'��'||To_Char(A.��ʼ, 'YYYY-MM-DD HH24:MI') || '��' ||To_Char(A.��ֹ, 'YYYY-MM-DD HH24:MI') || ')' As ����," & vbNewLine & _
                "      Trim(To_Char(B.ID))||';'||Trim(To_Char(Nvl(����,0)))||';'||To_Char(A.��ʼ, 'YYYY-MM-DD HH24:MI')||'��'||To_Char(A.��ֹ, 'YYYY-MM-DD HH24:MI')||';'||Trim(To_Char(A.ID))||';'||Trim(To_Char(A.Ӥ��)) As ����," & vbNewLine & _
                "       1 As ĩ��,'object_tend' As ͼ��,To_Char(a.��ʼ,'YYYY-MM-DD HH24:MI:SS') As ����" & vbNewLine & _
                " From (" & vbNewLine & _
                "   Select R.ID, F.���, R.����,R.Ӥ��, R.��ʼ, NVL(R.��ֹ,nvl(R.ʱ��,R.��ʼ)) ��ֹ, R.����id, ����" & vbNewLine & _
                "   From (" & vbNewLine & _
                "       Select L.ID, L.���, L.����, F.���� As ������, L.ͨ��, L.����" & vbNewLine & _
                "          From ����ҳ���ʽ F, �����ļ��б� L" & vbNewLine & _
                "          Where L.���� = 3 And L.���� = F.���� And L.��� = F.��� And (L.ͨ��=1 OR L.ͨ��=2)" & vbNewLine & _
                "" & vbNewLine & _
                "       ) F,(" & vbNewLine & _
                "       Select R.ID,R.����id,R.�ļ����� ����,R.��ʽID,nvl(R.Ӥ��,0) Ӥ��,Min(R.��ʼʱ��) As ��ʼ, Max(R.����ʱ��) As ��ֹ,MAX(T.����ʱ��) ʱ��" & vbNewLine & _
                "          From ���˻����ļ� R,���˻������� T" & vbNewLine & _
                "          Where R.ID=T.�ļ�ID(+) And R.����id = [1] And Nvl(R.��ҳid, 0) = [2]" & vbNewLine & _
                "          Group By R.ID,R.�ļ�����,R.����id,R.��ʽID,R.Ӥ��" & vbNewLine & _
                "       ) R" & vbNewLine & _
                "       Where F.ID=R.��ʽID" & vbNewLine & _
                "   ) A, ���ű� B Where A.����id = B.ID And DECODE(A.����,-1,0,A.Ӥ��)=A.Ӥ��)" & vbNewLine & _
                " Order By Decode(�ϼ�id,Null,' ',�ϼ�id),����"
    End If

    If mblnMoved Then
        strSQL = Replace(strSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
        strSQL = Replace(strSQL, "���˻����¼", "H���˻����¼")
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "���˻����ļ�", "H���˻����ļ�")
        strSQL = Replace(strSQL, "���˻�������", "H���˻�������")
    End If

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng����ID, IIf(mstr�Һŵ� = "", 1, 0))

    tvwArchive.Tag = ""
    tvwArchive.Nodes.Clear

    Do While Not rsTmp.EOF
        If NVL(rsTmp!�ϼ�ID) = "" Then
            Set objNode = tvwArchive.Nodes.Add(, , CStr(rsTmp!ID), rsTmp!����, NVL(rsTmp!ͼ��))
        Else
            Set objNode = tvwArchive.Nodes.Add(CStr(rsTmp!�ϼ�ID), tvwChild, CStr(rsTmp!ID), rsTmp!����, NVL(rsTmp!ͼ��))
        End If
        objNode.Tag = NVL(rsTmp!����)
        objNode.Expanded = True

        If tvwArchive.Nodes.count = 1 Then
            objNode.Selected = True
        ElseIf objNode.Key = strSel Then
            objNode.Selected = True
        End If
        rsTmp.MoveNext
    Loop

    Set rsTmp = New ADODB.Recordset
    Set rsTmp = GetEmrCISStruct(mlng����ID, mlng����ID)
    If Not rsTmp Is Nothing Then
        If rsTmp.State = ADODB.adStateOpen Then
        If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        Do Until rsTmp.EOF
            Set objNode = tvwArchive.Nodes.Add(rsTmp!�ϼ�ID.Value, tvwChild, rsTmp!ID.Value, rsTmp!����.Value, rsTmp!ͼ��.Value, rsTmp!ͼ��.Value)
            objNode.Tag = NVL(rsTmp!����) '�ĵ�ID[|���ĵ�ID]
            rsTmp.MoveNext
        Loop
        End If
        End If
    End If

    If Not tvwArchive.SelectedItem Is Nothing Then
        tvwArchive.SelectedItem.EnsureVisible
        Call tvwArchive_NodeClick(tvwArchive.SelectedItem)
    End If

    Screen.MousePointer = 0
    Exit Function
ErrH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowOutPatiInfo() As Boolean
'���ܣ�ѡ�����ﲡ��ĳ����ʷ�����¼ʱ����ȡ��صĲ�����Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo ErrH

    strSQL = "Select B.Id,B.NO,B.�����,B.����,B.�Ա�,B.����,A.ҽ�Ƹ��ʽ," & _
        " A.�ѱ�,A.����,A.ҽ����,B.����,B.����ʱ��,B.ִ����,B.ִ��״̬,B.ִ��ʱ��," & _
        " B.ִ�в���ID as ����ID,B.����,B.����,D.������,C.���� as ����" & _
        " From ������Ϣ A,���˹Һż�¼ B,���ű� C,����������Ϣ D" & _
        " Where A.����ID=B.����ID And B.ID=[1] And B.ִ�в���ID=C.ID" & _
        " And B.����ID=D.����ID(+) And B.����=D.����(+)"
    If mblnMoved Then
        strSQL = Replace(strSQL, "���˹Һż�¼", "H���˹Һż�¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    With rsTmp
        '���ղ���������ɫ��ʾ
        lbl����mz(1).Caption = NVL(!����)
        If Not IsNull(!����) Then
            lbl����mz(1).ForeColor = vbRed
        Else
            lbl����mz(1).ForeColor = lbl�����mz(1).ForeColor
        End If
        lblҽ��mz(1).Caption = NVL(!ִ����)
        lbl�Һŵ�mz(1).Caption = !NO
        lbl�����mz(1).Caption = NVL(!�����)
        lbl����mz(1).Caption = NVL(!ҽ�Ƹ��ʽ)
        lbl�ѱ�mz(1).Caption = NVL(!�ѱ�)
        lblҽ����mz(1).Caption = NVL(!ҽ����)
        lbl������mz(1).Caption = NVL(!������)
        lbl��.Visible = NVL(!����, 0) <> 0

        mlng����ID = NVL(!����ID, 0)
        mlng����ID = 0
    End With

    ShowOutPatiInfo = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowInPatiInfo() As Boolean
'���ܣ�ѡ��ĳ��סԺ��¼ʱ����ȡ��صĲ�����Ϣ
'���أ�blnMoved=����סԺ�����Ƿ�ת����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo ErrH

    strSQL = "Select A.����,A.�Ա�,A.����,B.סԺ��,B.��Ժ����,B.ҽ�Ƹ��ʽ," & _
        " D.��Ϣֵ as ҽ����,B.����,B.��ǰ����,C.���� as ����ȼ�,B.��Ժ����," & _
        " B.��Ժ����,B.��������,B.״̬,B.��Ժ����ID,B.��ǰ����ID,A.סԺ����" & _
        " From ������Ϣ A,������ҳ B,�շ���ĿĿ¼ C,������ҳ�ӱ� D" & _
        " Where A.����ID=B.����ID And A.����ID=[1] And B.��ҳID=[2] And B.����ȼ�ID=C.ID(+)" & _
        " And B.����ID=D.����ID(+) And B.��ҳID=D.��ҳID(+) And D.��Ϣ��(+)='ҽ����'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng����ID)

    With rsTmp
        '���ղ�����ɫ������ʾ
        lbl����zy(1).Caption = NVL(!����)
        lbl����zy(1).ForeColor = zlDatabase.GetPatiColor(NVL(!��������))

        lblסԺ��zy(1).Caption = NVL(!סԺ��)
        lbl����zy(1).Caption = NVL(!��Ժ����)
        lblҽ����zy(1).Caption = NVL(!ҽ����)
        lbl����zy(1).Caption = NVL(!����ȼ�)
        lbl����zy(1).Caption = NVL(!ҽ�Ƹ��ʽ)

        'Σ�ز��˲�����ɫ��ʾ
        lbl����zy(1).Caption = NVL(!��ǰ����)
        If NVL(!��ǰ����) = "Σ" Or NVL(!��ǰ����) = "��" Or NVL(!��ǰ����) = "��" Then
            lbl����zy(1).ForeColor = vbRed
        Else
            lbl����zy(1).ForeColor = lblסԺ��zy(1).ForeColor
        End If

        lbl��Ժzy(1).Caption = Format(!��Ժ����, "yyyy-MM-dd HH:mm")
        If Not IsNull(!��Ժ����) Then
            lbl��Ժzy(1).Caption = lbl��Ժzy(1).Caption & "��" & Format(!��Ժ����, "yyyy-MM-dd HH:mm")
        End If

        lbl����zy(1).Caption = NVL(!��������)

        mlng����ID = NVL(!��Ժ����ID, 0)
        mlng����ID = NVL(!��ǰ����ID, 0)
    End With

    ShowInPatiInfo = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetEmrCISStruct(ByVal lngPatiID As Long, ByVal lngPageID As Long) As ADODB.Recordset
Dim rsTemp As New ADODB.Recordset, strExtendTag As String, strReturn As String
    If gobjEmr Is Nothing Then Set GetEmrCISStruct = Nothing: Exit Function
    strExtendTag = GetEMRIn_Tag(lngPatiID, lngPageID)
    If strExtendTag = "" Then Set GetEmrCISStruct = Nothing: Exit Function
    
    '�б���clsPackage.GetEmrCISStruct ��Ҫ���ϼ�ID��Ӧ��ϵ��һ��
    gstrSQL = "Select Decode(e.Kind,'01','R1', '02', 'R2', '03', 'R4', '04', 'R5', '05', 'R6', 'R2') �ϼ�id," & vbNewLine & _
                "       Nvl(d.Subdoc_Id, Rawtohex(b.Id)) As ID, d.Subdoc_Id As ���ĵ�id," & vbNewLine & _
                "       Nvl(d.Subdoc_Title, b.Title) ||" & vbNewLine & _
                "        Decode(d.Completor, Null, ''," & vbNewLine & _
                "               '�� ' || d.Completor || ' ��' || To_Char(d.Complete_Time, 'yyyy-mm-dd hh24:mi') || 'ǩ����') As ����," & vbNewLine & _
                "       Rawtohex(b.Id) || Decode(d.Subdoc_Id, Null, Null, '|' || d.Subdoc_Id) As ����, 'object_case' As ͼ��" & vbNewLine & _
                "From Bz_Doc_Log B," & vbNewLine & _
                "     (Select Distinct a.Id, c.Antetype_Id, c.Subdoc_Id, c.Subdoc_Title, c.Real_Doc_Id, c.Complete_Time, c.Completor" & vbNewLine & _
                "       From (Select ID" & vbNewLine & _
                "              From Bz_Act_Log" & vbNewLine & _
                "              Where Extend_Tag = :etag" & vbNewLine & _
                "              Union" & vbNewLine & _
                "              Select b.Id" & vbNewLine & _
                "              From Bz_Act_Log F, Bz_Act_Log B" & vbNewLine & _
                "              Where f.Extend_Tag = :etag And b.Basiclog_Id = f.Id) A, Bz_Doc_Tasks C" & vbNewLine & _
                "       Where a.Id = c.Actlog_Id And c.Real_Doc_Id Is Not Null) D, Antetype_List E" & vbNewLine & _
                "Where b.Actlog_Id = d.Id And d.Real_Doc_Id = b.Id And d.Antetype_Id = e.Id And" & vbNewLine & _
                "      Decode(d.Subdoc_Id, Null, d.Antetype_Id, b.Antetype_Id) = b.Antetype_Id And" & vbNewLine & _
                "      Decode(d.Subdoc_Title, Null, e.Title, d.Subdoc_Title) = e.Title" & vbNewLine & _
                "Order By e.Code, b.Creat_Time, d.Complete_Time"
    strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, strExtendTag & "^" & DbType.T_String & "^etag", rsTemp)
    If strReturn <> "" Then
        MsgBox strReturn, vbCritical, ParamInfo.��Ʒ����
        Set GetEmrCISStruct = Nothing: Exit Function
    End If
    
    Set GetEmrCISStruct = rsTemp
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function