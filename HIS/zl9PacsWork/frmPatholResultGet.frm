VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPatholResultGet 
   Caption         =   "�ؼ���¼��"
   ClientHeight    =   7365
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   8625
   Icon            =   "frmPatholResultGet.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   8625
   StartUpPosition =   3  '����ȱʡ
   Begin RichTextLib.RichTextBox txtResult 
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   4920
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   2778
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmPatholResultGet.frx":179A
   End
   Begin VB.CheckBox chkWay 
      Caption         =   "���䷽ʽ"
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   4560
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.ComboBox cbxColCount 
      Height          =   300
      ItemData        =   "frmPatholResultGet.frx":1837
      Left            =   1920
      List            =   "frmPatholResultGet.frx":1856
      TabIndex        =   2
      Text            =   "3"
      Top             =   4530
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "ȡ ��(&C)"
      Height          =   400
      Left            =   7320
      TabIndex        =   4
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "ȷ ��(&O)"
      Height          =   400
      Left            =   6000
      TabIndex        =   3
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "�ؼ��¼��"
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   8415
      Begin VB.ComboBox cbxSpeexamDetails 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3720
         Width           =   2055
      End
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   3255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   6588
         IsKeepRows      =   0   'False
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
      End
      Begin VB.Label labSpeexamDetails 
         Caption         =   "�ؼ�ϸĿ��"
         Height          =   255
         Left            =   5280
         TabIndex        =   9
         Top             =   3800
         Width           =   975
      End
   End
   Begin VB.Label labCol 
      Caption         =   "������"
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label labStyle 
      Caption         =   "��ʽԤ����"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Width           =   1095
   End
End
Attribute VB_Name = "frmPatholResultGet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngPatholAdviceId As String
Private mlngCurSpeexamType As Long
Private mstrPrivs As String


Public IsOk As Boolean


Public Sub ShowResultGetWind(ByVal lngPatholAdviceId As Long, _
    ByVal lngCurSpeExamType As Long, ByVal strPrivs As String, owner As Form)
    
    mlngPatholAdviceId = lngPatholAdviceId
    mlngCurSpeexamType = lngCurSpeExamType
    mstrPrivs = strPrivs
    
    Call SetCaption(lngCurSpeExamType)
    Call LoadResultData(lngCurSpeExamType)
    Call LoadSpeExamDetails(lngCurSpeExamType)
    
    Call Me.Show(1, owner)
    
End Sub


Private Sub AdjustFace()
    Frame1.Left = 120
    Frame1.Top = 120
    Frame1.Width = Me.Width - 360
    Frame1.Height = Me.Height - cmdSure.Height - txtResult.Height - cbxColCount.Height - 1100
    
    ufgData.Left = 120
    ufgData.Top = 240
    ufgData.Width = Frame1.Width - 240
    ufgData.Height = Frame1.Height - cbxSpeexamDetails.Height - 480

    cbxSpeexamDetails.Left = ufgData.Width - cbxSpeexamDetails.Width
    cbxSpeexamDetails.Top = Frame1.Height - cbxSpeexamDetails.Height - 120
    
    labSpeexamDetails.Left = cbxSpeexamDetails.Left - 60 - labSpeexamDetails.Width
    labSpeexamDetails.Top = cbxSpeexamDetails.Top + 60
    
    cbxColCount.Top = Frame1.Top + Frame1.Height + 120
    
    labStyle.Left = 120
    labStyle.Top = cbxColCount.Top + 30
    
    
    labCol.Left = labStyle.Left + labStyle.Width + 360
    labCol.Top = labStyle.Top
    
    cbxColCount.Left = labCol.Left + labCol.Width
    
    chkWay.Left = cbxColCount.Left + cbxColCount.Width + 240
    chkWay.Top = labStyle.Top
    
    txtResult.Left = 120
    txtResult.Top = cbxColCount.Top + cbxColCount.Height + 120
    txtResult.Width = Me.Width - 360
    
    cmdExit.Left = Me.Width - cmdExit.Width - 240
    cmdExit.Top = Me.Height - cmdExit.Height - 600
    
    cmdSure.Left = cmdExit.Left - cmdSure.Width - 120
    cmdSure.Top = cmdExit.Top
    
    
End Sub



Private Sub LoadSpeExamDetails(ByVal lngSpeExamType As Long)
'�����ؼ���ϸ
    cbxSpeexamDetails.Clear
    
    Call cbxSpeexamDetails.AddItem("")
    
    If lngSpeExamType = TSpeexamType.stMianyi Then
        Call cbxSpeexamDetails.AddItem("1-����(����)")
        Call cbxSpeexamDetails.AddItem("2-����(��ҩ��ҩ)")
    ElseIf lngSpeExamType = TSpeexamType.stFenzi Then
        Call cbxSpeexamDetails.AddItem("1-����(ӫ��)")  '��Ӧ 3
        Call cbxSpeexamDetails.AddItem("2-����(��ͨ)")  '��Ӧ 4
    End If
    
    If cbxSpeexamDetails.ListCount > 0 Then cbxSpeexamDetails.ListIndex = 0
End Sub



Private Sub SetCaption(ByVal lngSpeExamType As Long)
    Select Case lngSpeExamType
        Case 0
            Me.Caption = "�����黯���¼��"
            Frame1.Caption = "�����黯��¼"
        Case 1
            Me.Caption = "����Ⱦɫ���¼��"
            Frame1.Caption = "����Ⱦɫ��¼"
        Case 2
            Me.Caption = "���Ӳ�����¼��"
            Frame1.Caption = "���Ӳ����¼"
    End Select
End Sub

Private Sub InitResultList()
'��ʼ�������ʾ�б�
    Dim strTemp As String
    
            
    '��������
    ufgData.GridRows = glngStandardRowCount
    '�����и�
    ufgData.RowHeightMin = glngStandardRowHeight
    
    '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
    strTemp = zlDatabase.GetPara("�ؼ����б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
    ufgData.DefaultColNames = gstrSpeExamResultGetCols
     
    If strTemp = "" Then
        ufgData.ColNames = gstrSpeExamResultGetCols
    Else
        ufgData.ColNames = strTemp
    End If
    '��ֹ�Ҽ������б����ô���
    ufgData.IsEjectConfig = False
    ufgData.ColConvertFormat = gstrSpeExamResultGetConvertFormat
End Sub

Private Sub ufgData_OnCheckAllChanged()
On Error GoTo ErrHandle
    Call ConfigResult(Val(cbxColCount.Text))
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnCheckChanged(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrHandle
    Call ConfigResult(Val(cbxColCount.Text))
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnColFormartChange()
'�رմ���ʱ�����б�����
     zlDatabase.SetPara "�ؼ����б�����", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
End Sub


Private Sub LoadResultData(ByVal lngSpeExamType As Long)
'�����ؼ�����
    Dim strSQL As String
    '�Ŀ����|�걾����|��������|��Ŀ���
    strSQL = "select a.Id, b.���,b.�걾����,c.��������,a.��Ŀ���,�ؼ�ϸĿ,a.��������, a.��Ŀ˳�� " & _
        " from �����ؼ���Ϣ a, ����ȡ����Ϣ b, ��������Ϣ c" & _
        " where a.�Ŀ�ID=b.�Ŀ�ID and a.����ID=c.����ID and a.��ǰ״̬=2 and a.����ҽ��ID=[1] and a.�ؼ�����=[2] order by a.�ؼ�ϸĿ,a.��������, b.���,a.��Ŀ˳��"
    'If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngPatholAdviceId, lngSpeExamType)
    
    Call ufgData.RefreshData
End Sub


Private Sub ConfigResult(ByVal lngColCount As Long)
'�����ؼ���
    Dim i As Long
    Dim lngCol As Long
    Dim lngMaxNameLen As Long
    Dim lngMaxResultLen As Long
    Dim lngCurNameLen As Long
    Dim lngCurResultLen As Long
    
    
    lngCol = 0
    lngMaxNameLen = 0
    lngMaxResultLen = 0
    txtResult.Text = ""
    
    For i = 1 To ufgData.GridRows - 1
        If ufgData.GetRowCheck(i) Then
            If lngMaxNameLen < Len(ufgData.Text(i, gstrSpeExamResultGet_��������)) Then
                lngMaxNameLen = Len(ufgData.Text(i, gstrSpeExamResultGet_��������))
            End If
            
            If lngMaxResultLen < Len(ufgData.Text(i, gstrSpeExamResultGet_��Ŀ���)) Then
                lngMaxResultLen = Len(ufgData.Text(i, gstrSpeExamResultGet_��Ŀ���))
            End If
        End If
    Next i
    
    For i = 1 To ufgData.GridRows - 1
        If ufgData.GetRowCheck(i) Then
            If chkWay.value <> 0 Then
                If Trim(txtResult.Text) <> "" Then txtResult.Text = txtResult.Text & "��"
            Else
                lngCol = lngCol + 1
                If lngCol > lngColCount Then
                    txtResult.Text = txtResult.Text & vbCrLf
                    lngCol = 1
                End If
            End If
            
            lngCurNameLen = Len(ufgData.Text(i, gstrSpeExamResultGet_��������))
            lngCurResultLen = Len(ufgData.Text(i, gstrSpeExamResultGet_��Ŀ���))
            
            '��check��Ϊ0����ʹ�ö���ʽ
            If chkWay.value <> 0 Then
                lngCurNameLen = lngMaxNameLen
                lngCurResultLen = lngMaxResultLen
            Else
                txtResult.Text = txtResult.Text & "  "
            End If
'
'            txtResult.Text = txtResult.Text & String(lngMaxNameLen - lngCurNameLen, " ") & ufgData.Text(i, gstrSpeExamResultGet_��������) & _
'                                                    "��" & ufgData.Text(i, gstrSpeExamResultGet_��Ŀ���) & String(lngMaxResultLen - lngCurResultLen, " ")

            txtResult.Text = txtResult.Text & String(lngMaxNameLen - lngCurNameLen, " ") & ufgData.Text(i, gstrSpeExamResultGet_��������) & _
                                                    "(" & ufgData.Text(i, gstrSpeExamResultGet_��Ŀ���) & String(lngMaxResultLen - lngCurResultLen, " ") & ")"

        End If
    Next i
    
    If chkWay.value <> 0 Then
        If Trim(txtResult.Text) <> "" Then txtResult.Text = txtResult.Text & "��"
    End If
End Sub



Private Sub cbxColCount_Change()
On Error GoTo ErrHandle
    Call ConfigResult(Val(cbxColCount.Text))
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbxColCount_Click()
On Error GoTo ErrHandle
    Call ConfigResult(Val(cbxColCount.Text))
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub cbxSpeexamDetails_Click()
On Error GoTo ErrHandle
    Call SpeExamined_Save
    
    If Trim(cbxSpeexamDetails.Text) = "" Then
        ufgData.AdoData.Filter = ""
        Call ufgData.RefreshData
        
        Exit Sub
    End If
    
    
    
    ufgData.AdoData.Filter = "�ؼ�ϸĿ=" & IIf(mlngCurSpeexamType = 0, Val(cbxSpeexamDetails.Text), Val(cbxSpeexamDetails.Text) + 2)
    
    Call ufgData.RefreshData
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkWay_Click()
On Error GoTo ErrHandle
    Call ConfigResult(Val(cbxColCount.Text))
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

'Private Sub cmdClearSelect_Click()
'On Error GoTo errHandle
'    Call ufgData.ClearSelect
'
'    Call ConfigResult(Val(cbxColCount.Text))
'Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'End Sub

Private Sub cmdExit_Click()
    IsOk = False
    
    Call Unload(Me)
End Sub



Private Sub SpeExamined_Save()
'�����ؼ���Ŀ
    Dim i As Long
    Dim strSQL As String
    
    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.IsEmptyKey(i) Then
            If ufgData.RowState(i) = TDataRowState.Update Then
                strSQL = "Zl_�����ؼ�_��Ŀ¼��(" & ufgData.KeyValue(i) & ",'" & ufgData.Text(i, gstrSpeExamResultGet_��Ŀ���) & "')"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            End If
        End If
    Next i
    
End Sub


Private Sub cmdSure_Click()
On Error GoTo ErrHandle
    Call SpeExamined_Save
    
    IsOk = True
    
    Call Me.Hide
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Initialize()
    IsOk = False
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitResultList
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Call AdjustFace
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    Call SaveWinState(Me, App.ProductName)
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrHandle
    Call ConfigResult(Val(cbxColCount.Text))
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


