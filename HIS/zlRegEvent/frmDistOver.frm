VERSION 5.00
Begin VB.Form frmDistOver 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ɾ���"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5415
      TabIndex        =   8
      Top             =   3855
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4065
      TabIndex        =   7
      Top             =   3870
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   165
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3825
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   120
      Left            =   -15
      TabIndex        =   10
      Top             =   3600
      Width           =   6675
   End
   Begin VB.CommandButton cmdDoct 
      Caption         =   "��"
      Height          =   240
      Left            =   6195
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3225
      Width           =   270
   End
   Begin VB.TextBox txtDoct 
      Height          =   300
      Left            =   4500
      TabIndex        =   5
      Top             =   3195
      Width           =   1995
   End
   Begin VB.TextBox txtMain 
      Height          =   2670
      Left            =   120
      TabIndex        =   1
      Top             =   375
      Width           =   6420
   End
   Begin VB.ComboBox CboRoom 
      Height          =   300
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3210
      Width           =   1740
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ҽ��(&D)"
      Height          =   180
      Left            =   3840
      TabIndex        =   4
      Top             =   3270
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��������(&R)"
      Height          =   180
      Left            =   150
      TabIndex        =   2
      Top             =   3270
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����ժҪ"
      Height          =   180
      Left            =   195
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmDistOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrSQL As String
Private mobjQueue As Object, mbytQueueType As Byte, mstrPrivs As String, mstrQueuePrivs As String
Private mlngModule As Long, mlng����ID As Long, mstrNo As String, mblnOk As Boolean
Private mlng�Һ�ID As Long

Private Declare Function ClientToScreen Lib "user32" (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long

Public Function zlShowEdit(ByVal frmMain As Form, ByVal strPrivs As String, ByVal strQueuePrivs As String, _
        ByVal objQueue As Object, ByVal lngModule As Long, _
        strNO As String, lng����ID As Long, strȱʡ���� As String, strȱʡҽ�� As String, bytQueueType As Byte, _
        Optional lng�Һ�ID As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ��������
    '��Σ�frmMain-���õ�������
    '         objQueue-�ŶӽкŶ���
    '         lngModule-ģ���
    '         bytQueueType-�ŶӺź�ģʽ
    '���Σ�
    '���أ��ɹ�,����true,���򷴻�False
    '���ƣ����˺�
    '���ڣ�2010-06-03 17:22:29
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    mlng����ID = lng����ID: mlngModule = lngModule: mstrNo = strNO: mblnOk = False
    mlng�Һ�ID = lng�Һ�ID
    Set mobjQueue = objQueue: mbytQueueType = bytQueueType: mstrPrivs = strPrivs: mstrQueuePrivs = strQueuePrivs
    Err = 0: On Error GoTo Errhand:

    '  �������еĸú�������ҹ�ѡ��
    If gbytRegistMode = 0 Then
        strSQL = _
            " Select b.����,b.����,b.λ��" & vbCrLf & _
            " From �ҺŰ������� a,�������� b,�ҺŰ��� c,���˹Һż�¼ d" & vbCrLf & _
            " Where a.��������=b.���� And a.�ű�ID=c.id And c.����=d.�ű� And d.NO=[1] AND d.��¼����=1 and d.��¼״̬=1 "
    Else
        If Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss") < Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") Then
            strSQL = _
                " Select b.����,b.����,b.λ��" & vbCrLf & _
                " From �ҺŰ������� a,�������� b,�ҺŰ��� c,���˹Һż�¼ d" & vbCrLf & _
                " Where a.��������=b.���� And a.�ű�ID=c.id And c.����=d.�ű� And d.NO=[1] AND d.��¼����=1 and d.��¼״̬=1 "
        Else
            strSQL = _
                " Select b.����,b.����,b.λ��" & vbCrLf & _
                " From �ٴ��������Ҽ�¼ a,�������� b,�ٴ������¼ c,���˹Һż�¼ d" & vbCrLf & _
                " Where a.����id=b.id And a.��¼ID=c.id And c.id=d.�����¼id And d.NO=[1] AND d.��¼����=1 and d.��¼״̬=1 "
        End If
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    With rsTemp
        CboRoom.Clear
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
                CboRoom.AddItem zlCommFun.Nvl(!����)
                If CboRoom.ListIndex < 0 Then CboRoom.ListIndex = CboRoom.NewIndex
                 If zlCommFun.Nvl(!����) = strȱʡ���� Then CboRoom.ListIndex = CboRoom.NewIndex
                .MoveNext
        Loop
    End With
    Me.Show 1, frmMain
    zlShowEdit = mblnOk
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Private Function QueueStauteUpdate() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ��Ŷӵ�ִ��״̬����
    '���ƣ����˺�
    '���ڣ�2010-06-03 18:11:42
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strQueueName As String, lngID As Long
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim i As Byte

    If mbytQueueType = 0 Then
        QueueStauteUpdate = True
        Exit Function
    End If
    If mobjQueue Is Nothing Then Exit Function
    If Not (InStr(mstrQueuePrivs, ";����;") > 0) Then Exit Function

    strSQL = "SELECT ID,ִ�в���ID,����,ִ���� From ���˹Һż�¼  where NO=[1]  AND ��¼����=1 and ��¼״̬=1"
    On Error GoTo Hd
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrNo)
    If rsTemp.EOF Then Exit Function
    
    strQueueName = Nvl(rsTemp!ִ�в���id)
    lngID = Val(Nvl(rsTemp!ID))
    '��ɾ���
    mobjQueue.zlQueueExec strQueueName, 0, lngID, 4
    QueueStauteUpdate = True
    Exit Function
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function
Private Sub CboRoom_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDoct_Click()
    On Error GoTo errHandle
    Dim rsTmp As ADODB.Recordset
    
    mstrSQL = "Select ִ�в���ID From ���˹Һż�¼ Where NO=[1]  AND ��¼����=1 And ��¼״̬=1"
    mstrSQL = _
        " Select c.���,c.����,c.����,c.id From ��Ա����˵�� a, ������Ա b ,��Ա�� c" & vbCrLf & _
        " Where b.��Աid=c.id And b.��Աid=a.��Աid  And  a.��Ա����='ҽ��' And (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null) " & vbCrLf & _
        " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & vbNewLine & _
        " And b.����id in (" & mstrSQL & ") "
    'ҽ��
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mstrNo)
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        mstrSQL = frmSelCurr.ShowCurrSel(Me, rsTmp, "���,1200,0,2;����,1500,0,1;����,1500,0,2;id,1,0,2", 1, "ҽ��ѡ��", , Me.txtDoct.Text, 1, , 6000)
        If mstrSQL <> "" Then
            Me.txtDoct.Text = mstrSQL
        End If
    Else
        MsgBox "���κ�ҽ������ѡ��", vbInformation, gstrSysName
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub cmdOk_Click()
    If Trim(Me.txtDoct.Text) = "" Then
        MsgBox "���������ҽ����", vbInformation, gstrSysName
        Me.txtDoct.SetFocus
        Exit Sub
    End If
    
    If Trim(Me.CboRoom.Text) = "" Then
        MsgBox "��ѡ��������ң�", vbInformation, gstrSysName
        Me.CboRoom.SetFocus
        Exit Sub
    End If
    
    If ExcPlugInFun(2, mlng�Һ�ID, Me.txtDoct.Text, Me.CboRoom.Text) = False Then Exit Sub
    
    mstrSQL = Replace(Me.txtMain.Text, "'", "''")
    '����ID_IN ������Ϣ.����ID%TYPE,
    'NO_IN     ���˹Һż�¼.NO%TYPE,
    '����_IN   ���˹Һż�¼.����%TYPE:=NULL,
    'ִ����_IN ���˹Һż�¼.ִ����%TYPE:=NULL,
    'ժҪ_IN   ���˹Һż�¼.ժҪ%TYPE:=NULL
    mstrSQL = "ZL_���˽������(" & mlng����ID & ",'" & mstrNo & "','" & Me.CboRoom.Text & "','" & Me.txtDoct.Text & "','" & mstrSQL & "',1)"
    On Error GoTo errHandle
    gcnOracle.BeginTrans
    Call zlDatabase.ExecuteProcedure(mstrSQL, Me.Caption)
    If QueueStauteUpdate = False Then
            gcnOracle.RollbackTrans: Exit Sub
    End If
    gcnOracle.CommitTrans
    Unload Me
    Exit Sub
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtDoct_GotFocus()
    zlControl.TxtSelAll Me.txtDoct
End Sub

Private Sub txtDoct_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandle
    Dim CurPoint As POINTAPI
    Dim rsTmp As ADODB.Recordset
    Dim strWidth As String
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        mstrSQL = _
            " (c.��� like '" & UCase(txtDoct.Text) & "%' or " & _
            "  c.���� like '" & gstrLike & UCase(txtDoct.Text) & "%' or " & _
            "  c.���� like '" & gstrLike & UCase(txtDoct.Text) & "%' ) "
        
        mstrSQL = _
            "Select c.���,c.����,c.����,c.id From ��Ա����˵�� a, ������Ա b ,��Ա�� c" & vbCrLf & _
            " Where b.��Աid=c.id And b.��Աid=a.��Աid  And  a.��Ա����='ҽ��' And (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null) And " & mstrSQL & vbCrLf & _
            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & vbNewLine & _
            " And b.����id in (Select ִ�в���ID From ���˹Һż�¼ Where NO=[1] And ��¼����=1 And ��¼״̬=1) "
        'ҽ��
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mstrNo)
        If rsTmp.RecordCount > 1 Then
            rsTmp.MoveFirst
            '��λѡ����
            CurPoint.X = (txtDoct.Left) / Screen.TwipsPerPixelX
            CurPoint.Y = (txtDoct.Top + txtDoct.Height + Screen.TwipsPerPixelY) / Screen.TwipsPerPixelY
            ClientToScreen Me.Hwnd, CurPoint
            '��ʼѡ����
            strWidth = "1000;1200;1200;0"
            strWidth = frmSelectChild.ShowSelectChild(Me, CurPoint.X * Screen.TwipsPerPixelX, CurPoint.Y * Screen.TwipsPerPixelY, 3400 + 30 * Screen.TwipsPerPixelX, Screen.TwipsPerPixelY * 200, rsTmp, strWidth)
            If Trim(strWidth) = "" Or Trim(strWidth) = ";;" Then
                Exit Sub
            End If
            '������صĲ���
            txtDoct.Text = Split(strWidth, ";")(1)
            zlCommFun.PressKey vbKeyTab
        ElseIf rsTmp.RecordCount = 1 Then
            txtDoct.Text = zlCommFun.Nvl(rsTmp!����)
            zlCommFun.PressKey vbKeyTab
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetDefineSize()
    '���ܣ��õ����ݿ�ı��ֶεĳ���
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset
    mstrSQL = "Select ժҪ,ִ���� From ���˹Һż�¼ Where Rownum<0"
    Call zlDatabase.OpenRecordset(rsTmp, mstrSQL, Me.Caption)
    
    txtMain.MaxLength = rsTmp.Fields("ժҪ").DefinedSize
    txtDoct.MaxLength = rsTmp.Fields("ִ����").DefinedSize
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
