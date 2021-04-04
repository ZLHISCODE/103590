VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPatholConsultation_Feedback 
   Caption         =   "���ﷴ��"
   ClientHeight    =   7245
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   10635
   Icon            =   "frmPatholConsultation_Feedback.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   10635
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picFeedback 
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   4320
      ScaleHeight     =   6975
      ScaleWidth      =   6015
      TabIndex        =   4
      Top             =   120
      Width           =   6015
      Begin VB.Frame framFeedback 
         Height          =   6375
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   6015
         Begin RichTextLib.RichTextBox txtAdvice 
            Height          =   2175
            Left            =   240
            TabIndex        =   11
            Top             =   3360
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   3836
            _Version        =   393217
            BorderStyle     =   0
            ScrollBars      =   2
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"frmPatholConsultation_Feedback.frx":179A
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
         Begin RichTextLib.RichTextBox txtResult 
            Height          =   2295
            Left            =   240
            TabIndex        =   10
            Top             =   480
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   4048
            _Version        =   393217
            BorderStyle     =   0
            ScrollBars      =   2
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"frmPatholConsultation_Feedback.frx":1837
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
         Begin VB.TextBox txtMemo 
            Height          =   300
            Left            =   1200
            TabIndex        =   0
            Top             =   5880
            Width           =   4695
         End
         Begin VB.Label labMemo 
            Caption         =   "��ע˵����"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   5880
            Width           =   975
         End
         Begin VB.Label labAdvice 
            Caption         =   "��������"
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Label labResult 
            Caption         =   "��Ͻ����"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "�� ��(&E)"
         Height          =   400
         Left            =   4800
         TabIndex        =   2
         Top             =   6480
         Width           =   1215
      End
      Begin VB.CommandButton cmdSure 
         Caption         =   "ȷ ��(&S)"
         Height          =   400
         Left            =   3360
         TabIndex        =   1
         Top             =   6480
         Width           =   1215
      End
   End
   Begin VB.PictureBox picWord 
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   120
      ScaleHeight     =   6855
      ScaleWidth      =   3855
      TabIndex        =   3
      Top             =   120
      Width           =   3855
      Begin zl9PACSWork.WordInputModule wimWord 
         Height          =   6615
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   11668
         CurDepartId     =   0
      End
   End
   Begin XtremeDockingPane.DockingPane dkpFeedback 
      Left            =   4080
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPatholConsultation_Feedback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mufgParentGrid As ucFlexGrid
Private mlngConsultationId As Long
Private mlngCurDepartId As Long

Private mCurEditText As RichTextBox

Private mParentForm As Form

Public blnIsOk As Boolean


Public Function ShowFeedbackWindow(ufgParentGrid As ucFlexGrid, ByVal lngConsultationId As Long, _
    ByVal lngCurDepartId As Long, owner As Form) As Boolean
'��ʾ���ﷴ������
    Dim blnIsReadOnly As Boolean
    
    Set mParentForm = owner
    Set mufgParentGrid = ufgParentGrid
    
    mlngConsultationId = lngConsultationId
    mlngCurDepartId = lngCurDepartId
    
    blnIsOk = False
    
    Call LoadReportModule
    
    Call LoadFeedbackContext

    blnIsReadOnly = IIf(mufgParentGrid.Text(mufgParentGrid.SelectionRow, gstrConsultation_��ǰ״̬) = "�Ѳ���", True, False)
    Call ConfigReadOnly(blnIsReadOnly)
    
    Call Me.Show(0, owner)
End Function





Public Function ShowFeedbackViewWindow(ufgParentGrid As ucFlexGrid, owner As Form)
'��ʾ���ﷴ��Ԥ������
    Set mParentForm = owner
    Set mufgParentGrid = ufgParentGrid
    
    blnIsOk = False
    
    Call LoadFeedbackContext
    Call ConfigReadOnly(True)
    
    Call Me.Show(0, owner)
End Function


Private Sub LoadFeedbackContext()
'��ȡ��������ϼ�¼
    txtAdvice.Text = mufgParentGrid.Text(mufgParentGrid.SelectionRow, gstrConsultation_������)
    txtResult.Text = mufgParentGrid.Text(mufgParentGrid.SelectionRow, gstrConsultation_��Ͻ��)
    txtMemo.Text = mufgParentGrid.Text(mufgParentGrid.SelectionRow, gstrConsultation_��ע)
End Sub



Private Sub ConfigReadOnly(ByVal blnIsReadOnly As Boolean)
'���û���ı༭״̬
    txtAdvice.Locked = blnIsReadOnly
    txtResult.Locked = blnIsReadOnly
    txtMemo.Locked = blnIsReadOnly
    
    txtAdvice.BackColor = IIf(Not blnIsReadOnly, vbWhite, Me.BackColor)
    txtResult.BackColor = IIf(Not blnIsReadOnly, vbWhite, Me.BackColor)
    txtMemo.BackColor = IIf(Not blnIsReadOnly, vbWhite, Me.BackColor)
    

    cmdSure.Enabled = Not blnIsReadOnly
End Sub


Private Sub InitFace()
'��ʼ�����沼��
    Dim Pane1 As Pane, Pane2 As Pane
    
    With dkpFeedback
        .CloseAll
        .Options.HideClient = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    Set Pane1 = dkpFeedback.CreatePane(1, Round(Width / 3), Me.Height, DockLeftOf, Nothing)
    Pane1.Title = "�ʾ�ģ��"
    Pane1.Handle = picWord.hWnd
    Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Pane1.MinTrackSize.Width = 50
    Pane1.MinTrackSize.Height = 50

    Set Pane2 = dkpFeedback.CreatePane(2, Round(Width / 3 * 2), Me.Height, DockRightOf, Pane1)
    Pane2.Title = "���ﷴ��"
    Pane2.Handle = picFeedback.hWnd
    Pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Pane2.MinTrackSize.Width = 50
    Pane2.MinTrackSize.Height = 50
End Sub


Private Sub cmdExit_Click()
    blnIsOk = False
    
    Call Unload(Me)
End Sub


Private Sub SaveFeedbackData()
'������ﷴ��
    Dim strSql As String
    Dim lngFindIndex As Long
    Dim dtServicesTime As Date
    
    
    dtServicesTime = zlDatabase.Currentdate
    strSql = "Zl_�������_����(" & mlngConsultationId & ",'" & _
                                txtResult.Text & "','" & _
                                txtAdvice.Text & "'," & _
                                To_Date(dtServicesTime) & ",'" & _
                                txtMemo.Text & "')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    lngFindIndex = mufgParentGrid.FindRowIndex(mlngConsultationId, gstrConsultation_ID)
    
    If lngFindIndex <= 0 Then Exit Sub
    
    mufgParentGrid.Text(lngFindIndex, gstrConsultation_��ǰ״̬) = "�ѷ���"
    mufgParentGrid.Text(lngFindIndex, gstrConsultation_��Ͻ��) = txtResult.Text
    mufgParentGrid.Text(lngFindIndex, gstrConsultation_������) = txtAdvice.Text
    mufgParentGrid.Text(lngFindIndex, gstrConsultation_��ע) = txtMemo.Text
    mufgParentGrid.Text(lngFindIndex, gstrConsultation_���ʱ��) = dtServicesTime
    
    mParentForm.txtResult.Text = txtResult.Text
    mParentForm.txtAdvice.Text = txtAdvice.Text
End Sub



Private Sub cmdSure_Click()
'���淴����Ϣ
On Error GoTo errHandle
    Call SaveFeedbackData
    
    blnIsOk = True
'    Call Me.Hide
    Call Unload(Me)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Initialize()
    Set mCurEditText = txtResult
End Sub


Private Sub LoadReportModule()
'����ʾ�ģ��
    Dim strLinkClassName As String
    
    If mlngCurDepartId = wimWord.CurDepartId Then Exit Sub
    
    strLinkClassName = zlDatabase.GetPara("���汨��ģ��", glngSys, glngModul, "")
    
    wimWord.ModuleName = strLinkClassName
    wimWord.CurDepartId = mlngCurDepartId
    
    Call wimWord.LoadWordModel
End Sub


Private Sub Form_Load()
On Error GoTo errHandle
    Call RestoreWinState(Me, App.ProductName)
    
    '����д���ﱨ���ʱ����Ҫ��ʾ����ǰ��
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '�������ö�
    
    '��ʼ�����沼��
    Call InitFace
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub picFeedback_Resize()
On Error Resume Next
    Dim lngAvgHeight As Long
    
    framFeedback.Left = 0
    framFeedback.Top = 60
    framFeedback.Width = picFeedback.Width - 120
    framFeedback.Height = picFeedback.Height - cmdSure.Height - 360
    
    
    lngAvgHeight = Fix((framFeedback.Height - txtMemo.Height - labResult.Height - labAdvice.Height - 120 * 9) / 2)
    
    labResult.Left = 120
    labResult.Top = 240
    
    txtResult.Left = 120
    txtResult.Top = labResult.Top + labResult.Height + 60
    txtResult.Width = framFeedback.Width - 240
    txtResult.Height = lngAvgHeight
    
    labAdvice.Left = 120
    labAdvice.Top = txtResult.Top + txtResult.Height + 240
    
    txtAdvice.Left = 120
    txtAdvice.Top = labAdvice.Top + labAdvice.Height + 60
    txtAdvice.Width = framFeedback.Width - 240
    txtAdvice.Height = lngAvgHeight
    
    txtMemo.Width = framFeedback.Width - labMemo.Width - 480
    txtMemo.Left = framFeedback.Width - txtMemo.Width - 120
    txtMemo.Top = txtAdvice.Top + txtAdvice.Height + 240
    
    labMemo.Left = 120
    labMemo.Top = txtMemo.Top + 60
    
    cmdExit.Left = picFeedback.Width - cmdExit.Width - 120
    cmdExit.Top = picFeedback.Height - cmdExit.Height - 120
    
    cmdSure.Left = cmdExit.Left - cmdSure.Width - 120
    cmdSure.Top = cmdExit.Top
    
End Sub


Private Sub picWord_Resize()
On Error Resume Next
    wimWord.Left = 120
    wimWord.Top = 120
    wimWord.Width = picWord.Width - 240
    wimWord.Height = picWord.Height - 240
End Sub


Private Sub txtAdvice_GotFocus()
    Set mCurEditText = txtAdvice
End Sub

Private Sub txtResult_GotFocus()
    Set mCurEditText = txtResult
End Sub

Private Sub wimWord_OnWordDbClickEvent(ByVal strWord As String)
'����ʾ�
On Error GoTo errHandle
    If Not mCurEditText.Locked Then mCurEditText.SelText = strWord
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
