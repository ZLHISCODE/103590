VERSION 5.00
Begin VB.Form frmPatholDecalcification 
   Caption         =   "�Ѹ��������"
   ClientHeight    =   6495
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   9240
   Icon            =   "frmPatholDecalcification.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   9240
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer timeDate 
      Interval        =   1000
      Left            =   4920
      Top             =   5880
   End
   Begin VB.Timer timeDecalin 
      Interval        =   30000
      Left            =   3960
      Top             =   5880
   End
   Begin VB.CommandButton cmdSucceed 
      Caption         =   "�� ��(&F)"
      Height          =   400
      Left            =   7920
      TabIndex        =   2
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "�� ��(&H)"
      Height          =   400
      Left            =   6600
      TabIndex        =   1
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Frame framDecalin 
      Caption         =   "�ѸƼ�¼"
      Height          =   5655
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9015
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   5295
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   9340
         DefaultCols     =   ""
         GridRows        =   21
         IsKeepRows      =   0   'False
         BackColor       =   12648447
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         Editable        =   0
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
      End
   End
   Begin VB.Label labTime 
      Caption         =   "��ǰʱ��: 2011-11-11 11:11:11"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   6000
      Width           =   3015
   End
End
Attribute VB_Name = "frmPatholDecalcification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrPrivs As String
Private mblnMoved As Boolean
Private mlngModul As Long

Private mblnIsSoundHint As Boolean
Private mlngHintTime As Long

Private mfrmParent As Form

Private mblnPlaySound As Boolean


Public Sub ShowDecalinTaskWind(ByVal strPrivs As String, ByVal blnMoved As Boolean, ByVal lngModul As Long, owner As Form)
'��ʾ�Ѹ����񴰿�
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '�������ö�
    
    mstrPrivs = strPrivs
    mblnMoved = blnMoved
    mlngModul = lngModul
    
    Set mfrmParent = owner
    
    
    '��ʼ������
    Call InitParameter
    
    
    If Not owner.Visible Then Exit Sub
    
    Call Me.Show(0, owner)
End Sub

Private Sub InitDecalinList()
'��ʼ���Ѹ������б�
    Dim strTemp As String
    


     '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
    strTemp = zlDatabase.GetPara("�Ѹ������б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
    ufgData.DefaultColNames = gstrDecalinTaskCols
	
    If strTemp = "" Then
        ufgData.ColNames = gstrDecalinTaskCols
    Else
        ufgData.ColNames = strTemp
    End If
        '��������
    ufgData.GridRows = glngStandardRowCount
    '�����и�
    ufgData.RowHeightMin = glngStandardRowHeight
    ufgData.ColConvertFormat = gstrDecalinConvertFormat
End Sub


Private Sub ufgData_OnColFormartChange()
  '�����б����
    zlDatabase.SetPara "�Ѹ������б�����", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
End Sub


Private Sub LoadDecalinData()
'�����Ѹ���Ϣ
    Dim strSQL As String
    
    strSQL = "select a.ID,a.�걾ID,c.�����, b.�걾����,a.��ʼʱ��,case when a.����ʱ�� / 60 < 1 then '0' else '' end || to_char(a.����ʱ�� / 60) as ����ʱ��, (case when a.����ʱ�� - ((sysdate - a.��ʼʱ��) * 24 * 60 ) < 0 then 0 else trunc(a.����ʱ�� - ((sysdate - a.��ʼʱ��) * 24 * 60 )) end) as ʣ��ʱ��, (a.��ʼʱ�� + a.����ʱ��/60/24) as ����ʱ��, a.��ǰ�״�,a.���״̬,a.����Ա" & _
                " from �����Ѹ���Ϣ a, ����걾��Ϣ b, ��������Ϣ c" & _
                " where a.�걾id = b.�걾id and b.ҽ��ID = c.ҽ��ID and a.����Ա=[1] and a.���״̬<>1 and a.��ʼʱ��>sysdate - 30 order by ���״̬,ʣ��ʱ��,��ʼʱ��,ID"
    
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.����)
    
    Call ufgData.RefreshData
End Sub


Private Sub Decalin_Change(ByVal dtStart As Date, ByVal lngTimeLen As Double)
'�Ѹƻ���
    Dim strSQL As String
    Dim lngDecalinId As Long
    
    lngDecalinId = ufgData.KeyValue(ufgData.SelectionRow)
    
    strSQL = "Zl_�����Ѹ�_����(" & lngDecalinId & "," & zlStr.To_Date(dtStart) & "," & Fix(lngTimeLen * 60) & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    '�����Ѹ���ʾ�б�
    ufgData.Text(ufgData.SelectionRow, gstrDecalin_��ǰ�״�) = Val(ufgData.Text(ufgData.SelectionRow, gstrDecalin_��ǰ�״�)) + 1
    ufgData.Text(ufgData.SelectionRow, gstrDecalin_��ʼʱ��) = dtStart
    ufgData.Text(ufgData.SelectionRow, gstrDecalin_����ʱ��) = Format$(lngTimeLen, "0.0")
    ufgData.Text(ufgData.SelectionRow, gstrDecalin_ʣ��ʱ��) = Fix(lngTimeLen * 60)
    ufgData.Text(ufgData.SelectionRow, gstrDecalin_����ʱ��) = DateAdd("n", lngTimeLen * 60, dtStart)
End Sub


Private Sub cmdChange_Click()
On Error GoTo ErrHandle
    Dim frmChangeInput As frmPatholMaterials_Change

    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ���׵ļ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsNullRow(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ҫ���׵ļ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '�жϵ�ǰ��¼�Ƿ��Ѿ���ʼ�Ѹ�
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "�ñ걾��δ��ʼ�Ѹƣ�����ִ�л��ײ���������ִ���Ѹơ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.Text(ufgData.SelectionRow, gstrDecalin_��ǰ״̬) = "�����" Then
        Call MsgBoxD(Me, "�Ѹ���������ɣ����ܽ��л��ײ�����", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Set frmChangeInput = New frmPatholMaterials_Change
    On Error GoTo errFree
    
        Call frmChangeInput.ShowChangeWindow(Me)
            
        If Not frmChangeInput.IsSure Then Exit Sub
        
        '����
        Call Decalin_Change(frmChangeInput.StartTime, frmChangeInput.TimeLen)
errFree:
    Unload frmChangeInput
    Set frmChangeInput = Nothing
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Decalin_Succed()
'����Ѹ�
    Dim strSQL As String
    Dim lngDecalinId As Long
    
    lngDecalinId = ufgData.KeyValue(ufgData.SelectionRow)
    
    strSQL = "Zl_�����Ѹ�_���(" & lngDecalinId & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    '�����Ѹ���ʾ�б�
    ufgData.Text(ufgData.SelectionRow, gstrDecalin_��ǰ״̬) = "�����"
End Sub


Private Sub cmdSucceed_Click()
On Error GoTo ErrHandle


    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ����ѸƵļ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsNullRow(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ҫ����ѸƵļ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '�жϵ�ǰ��¼�Ƿ��Ѿ���ʼ�Ѹ�
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "�ñ걾��δ��ʼ�Ѹƣ�����ִ�иò���������ִ���Ѹơ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call Decalin_Succed
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitDecalinList
    
    Call LoadDecalinData
    
    Call CheckListState
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub InitParameter()
On Error Resume Next
    mblnIsSoundHint = Val(zlDatabase.GetPara("�Ѹ���������", glngSys, mlngModul, 1))
    mlngHintTime = Val(zlDatabase.GetPara("���Ѽ��ʱ��", glngSys, mlngModul, "30"))
    
    timeDecalin.Interval = mlngHintTime * 1000
End Sub


Private Sub AdjustFace()
    framDecalin.Left = 120
    framDecalin.Top = 120
    framDecalin.Width = Me.Width - 360
    framDecalin.Height = Me.Height - cmdSucceed.Height - 900
    
    ufgData.Left = 120
    ufgData.Top = 240
    ufgData.Width = framDecalin.Width - 240
    ufgData.Height = framDecalin.Height - 360
    
    cmdSucceed.Left = Me.Width - cmdSucceed.Width - 240
    cmdSucceed.Top = Me.Height - cmdSucceed.Height - 620
    
    cmdChange.Left = cmdSucceed.Left - cmdChange.Width - 120
    cmdChange.Top = cmdSucceed.Top
    
    labTime.Left = 120
    labTime.Top = cmdSucceed.Top + 60
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call Me.Hide
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Call AdjustFace
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub


Private Sub timeDate_Timer()
On Error Resume Next
    labTime.Caption = "��ǰʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    
    '������ʾ
    If mblnPlaySound Then
        If Not (mfrmParent Is Nothing) Then
            If InStr(mfrmParent.Caption, "        ") <= 0 Then mfrmParent.Caption = mfrmParent.Caption & "        "
            
            If mfrmParent.Caption Like "*�Ѹ����������*" Then
                mfrmParent.Caption = Replace(mfrmParent.Caption, "        �Ѹ���������ɣ�����", "        ")
            Else
                mfrmParent.Caption = Replace(mfrmParent.Caption, "        ", "        �Ѹ���������ɣ�����")
            End If
        End If
    End If
End Sub

Private Sub timeDecalin_Timer()
On Error Resume Next
    Call LoadDecalinData
    
    Call CheckListState
End Sub


Private Sub PalyHintSound()
'������ʾ����
    Call Beep(2000, 100)
    Call Beep(1000, 100)
    Call Beep(2000, 100)
    Call Beep(1000, 100)
    Call Beep(2000, 100)
    Call Beep(1000, 100)
End Sub


Private Sub CheckListState()
'����Ѹ������б�״̬
    Dim i As Long
    
    
    mblnPlaySound = False
    For i = 1 To ufgData.GridRows - 1
        If Val(ufgData.Text(i, gstrDecalin_ʣ��ʱ��)) = 0 Then
            Call ufgData.SetRowColor(i, &H80FF80)
            
            mblnPlaySound = True
            
        ElseIf Val(ufgData.Text(i, gstrDecalin_ʣ��ʱ��)) < 5 Then
            Call ufgData.SetRowColor(i, &H80FFFF)
        Else
            Call ufgData.SetRowColor(i, ufgData.BackColor)
        End If
    Next i
    
    
    '������ʾ
    If mblnPlaySound And mblnIsSoundHint Then Call PalyHintSound
End Sub




Private Sub ufgData_OnColsNameReSet()
On Error GoTo ErrHandle

    Call LoadDecalinData

Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub
