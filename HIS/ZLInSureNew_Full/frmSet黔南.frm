VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BILLEDIT.OCX"
Begin VB.Form frmSetǭ�� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chk������ 
      Caption         =   "��վ���ж�����."
      Height          =   345
      Left            =   1560
      TabIndex        =   21
      Top             =   4395
      Width           =   2190
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   105
      TabIndex        =   3
      Top             =   210
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "����������(&S)"
      TabPicture(0)   =   "frmSetǭ��.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraҽ��������"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "�շ����(&T)"
      TabPicture(1)   =   "frmSetǭ��.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "mshBill"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fraҽ�������� 
         Caption         =   "ҽԺǰ��ҽ��������"
         Height          =   1605
         Left            =   180
         TabIndex        =   12
         Top             =   510
         Width           =   4155
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   0
            Left            =   1260
            MaxLength       =   40
            TabIndex        =   16
            Top             =   330
            Width           =   1635
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   1260
            MaxLength       =   40
            PasswordChar    =   "*"
            TabIndex        =   15
            Top             =   720
            Width           =   1635
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   1260
            MaxLength       =   40
            TabIndex        =   14
            Top             =   1110
            Width           =   1635
         End
         Begin VB.CommandButton cmdTest 
            Caption         =   "����(&T)"
            Height          =   1095
            Left            =   3000
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   330
            Width           =   1005
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "�û���(&U)"
            Height          =   180
            Index           =   0
            Left            =   390
            TabIndex        =   19
            Top             =   390
            Width           =   810
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "����(&P)"
            Height          =   180
            Index           =   1
            Left            =   570
            TabIndex        =   18
            Top             =   780
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "������(&S)"
            Height          =   180
            Index           =   2
            Left            =   390
            TabIndex        =   17
            Top             =   1170
            Width           =   810
         End
      End
      Begin VB.Frame fra 
         Caption         =   "ҽ������ǰ��ҽ��������"
         Height          =   1605
         Left            =   180
         TabIndex        =   4
         Top             =   2280
         Width           =   4155
         Begin VB.CommandButton cmd���� 
            Caption         =   "����(&4)"
            Height          =   1095
            Left            =   3000
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   330
            Width           =   1005
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   5
            Left            =   1260
            MaxLength       =   40
            TabIndex        =   7
            Top             =   1110
            Width           =   1635
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   4
            Left            =   1260
            MaxLength       =   40
            PasswordChar    =   "*"
            TabIndex        =   6
            Top             =   720
            Width           =   1635
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   3
            Left            =   1260
            MaxLength       =   40
            TabIndex        =   5
            Top             =   330
            Width           =   1635
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "������(&3)"
            Height          =   180
            Index           =   3
            Left            =   390
            TabIndex        =   11
            Top             =   1170
            Width           =   810
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "����(&2)"
            Height          =   180
            Index           =   4
            Left            =   570
            TabIndex        =   10
            Top             =   780
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "�û���(&1)"
            Height          =   180
            Index           =   5
            Left            =   390
            TabIndex        =   9
            Top             =   390
            Width           =   810
         End
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   3495
         Left            =   -74940
         TabIndex        =   20
         Top             =   465
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   6165
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
   End
   Begin VB.CommandButton cmdCardPara 
      Caption         =   "���ö���������"
      Height          =   660
      Left            =   4905
      TabIndex        =   2
      Top             =   3630
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4965
      TabIndex        =   1
      Top             =   540
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4965
      TabIndex        =   0
      Top             =   930
      Width           =   1100
   End
End
Attribute VB_Name = "frmSetǭ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mcnTest As New ADODB.Connection
Private mcnZxTest   As New ADODB.Connection
Private mblnChange As Boolean
Private mblnFirst As Boolean
Private Enum enum�ı�
    textҽ���û� = 0
    Textҽ������ = 1
    Textҽ�������� = 2
    text�����û� = 3
    Text�������� = 4
    Text���ķ����� = 5
End Enum
Private Enum mColHead
    �շ���� = 0
    ������Ŀ
End Enum
Private Function LoadCbo() As Boolean
    Dim rsTemp As New ADODB.Recordset
    If gcnOracle_ǭ�� Is Nothing Then Exit Function
    gstrSQL = " select * from ҽ���շ���� "
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, gcnOracle_ǭ��
    mshBill.Clear
    Do While Not rsTemp.EOF
        mshBill.AddItem rsTemp!���� & "-" & rsTemp!����
        rsTemp.MoveNext
    Loop
    LoadCbo = True
End Function
 Private Function iniData() As Boolean
    '��ʼ����
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim lngRow As Long
    Dim strTmp As String
    
    '����ҳͷ
    Err = 0
    On Error Resume Next
    '���ñ���ͷ
    Call initGrid
    strSQL = "" & _
        "   Select A.���,b.����ֵ From �շ���� a,(Select * From ���ղ��� where ����=" & TYPE_ǭ�� & ") b " & _
        "   Where A.���=b.������(+) " & _
        "   order by A.���� "
        
    zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption
    With mshBill
        .ClearBill
        If rsTmp.RecordCount = 0 Then
            .Rows = 2
        Else
            .Rows = rsTmp.RecordCount + 1
        End If
        lngRow = 1
        Do While Not rsTmp.EOF
            .TextMatrix(lngRow, mColHead.�շ����) = Nvl(rsTmp!���)
            strTmp = Nvl(rsTmp!����ֵ)
            
            If Trim(strTmp) <> "" Then
                .TextMatrix(lngRow, mColHead.������Ŀ) = strTmp
            End If
            lngRow = lngRow + 1
            rsTmp.MoveNext
        Loop
        
    End With
    
End Function
Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = 2
        
        .msfObj.FixedCols = 1
        .AllowAddRow = False
        
        .TextMatrix(0, mColHead.�շ����) = "�շ����"
        .TextMatrix(0, mColHead.������Ŀ) = "������Ŀ"
        
        
        .ColWidth(mColHead.�շ����) = 1500
        .ColWidth(mColHead.������Ŀ) = 2000
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��

        .ColData(mColHead.�շ����) = 5
        .ColData(mColHead.������Ŀ) = 3
        
        .ColAlignment(mColHead.�շ����) = flexAlignLeftCenter
        .ColAlignment(mColHead.������Ŀ) = flexAlignLeftCenter
        .PrimaryCol = mColHead.������Ŀ
        .LocateCol = mColHead.������Ŀ
    End With
End Sub


Public Function ��������() As Boolean
    mblnChange = False
    Dim rsTemp As New ADODB.Recordset
    frmSetǭ��.Show vbModal, frmҽ�����
    �������� = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdTest_Click()
    Dim rsTemp As New ADODB.Recordset
    If mcnTest.State = adStateOpen Then mcnTest.Close
    
    If OraDataOpen(mcnTest, txtEdit(Textҽ��������).Text, txtEdit(textҽ���û�).Text, txtEdit(Textҽ������).Text) = False Then
        Exit Sub
    End If
    MsgBox "���ӳɹ���", vbInformation, gstrSysName
End Sub

Private Sub cmdCardPara_Click()
    If sCard_SetupCardOption_ǭ�� = False Then Exit Sub
End Sub

Private Sub cmd����_Click()
    Dim rsTemp As New ADODB.Recordset
    If mcnZxTest.State = adStateOpen Then mcnZxTest.Close
    
    If OraDataOpen(mcnZxTest, txtEdit(Text���ķ�����).Text, txtEdit(text�����û�).Text, txtEdit(Text��������).Text) = False Then
        Exit Sub
    End If
    MsgBox "���ӳɹ���", vbInformation, gstrSysName
End Sub

Private Sub Form_Activate()
    Dim rsTemp As New ADODB.Recordset
    Dim strReg As String
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
     
    gstrSQL = "Select * From ���ղ��� where ����=" & TYPE_ǭ��
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    With rsTemp
        Do While Not .EOF
            Select Case Nvl(!������)
            Case "ҽ���û���"
                  txtEdit(textҽ���û�).Text = Nvl(!����ֵ)
            Case "ҽ���û�����"
                  txtEdit(Textҽ������).Text = Nvl(!����ֵ)
            Case "ҽ��������"
                  txtEdit(Textҽ��������).Text = Nvl(!����ֵ)
            Case "�����û���"
                  txtEdit(text�����û�).Text = Nvl(!����ֵ)
            Case "�����û�����"
                  txtEdit(Text��������).Text = Nvl(!����ֵ)
            Case "���ķ�����"
                  txtEdit(Text���ķ�����).Text = Nvl(!����ֵ)
            End Select
            .MoveNext
        Loop
    End With
    GetRegInFor g����ȫ��, "ҽ��", "������", strReg
    If Val(strReg) = 1 Then
        chk������.Value = 1
    Else
        chk������.Value = 0
    End If
    Call iniData
 End Sub

Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 1 Then
        If gcnOracle_ǭ�� Is Nothing Then
            If Open�м��() = False Then Exit Sub
        End If
        LoadCbo
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Textҽ������ Then
        txtEdit(Index).Tag = txtEdit(Index).Text
    End If
    
    If Index = Textҽ�������� Or Index = Textҽ������ Or Index = textҽ���û� Then
        '�رն�ҽ�������������ӣ���Ϊ�ڲ����������ʱ��Ҫ���´�
        If mcnTest.State = adStateOpen Then mcnTest.Close
    End If
    If Index = text�����û� Or Index = Text�������� Or Index = Text���ķ����� Then
        If mcnZxTest.State = adStateOpen Then mcnZxTest.Close
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub cmdOK_Click()
    If IsValid = False Then Exit Sub
    If SaveData = False Then Exit Sub
    
    mblnOK = True
    mblnChange = False
    Unload Me
End Sub

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    Dim rsTemp As New ADODB.Recordset
    
    
    For lngCount = txtEdit.LBound To txtEdit.UBound
        If zlCommFun.StrIsValid(txtEdit(lngCount).Text, txtEdit(lngCount).MaxLength) = False Then
            zlControl.TxtSelAll txtEdit(lngCount)
            txtEdit(lngCount).SetFocus
            Exit Function
        End If
    Next
    
    If mcnTest.State = adStateClosed Then
        If OraDataOpen(mcnTest, txtEdit(Textҽ��������).Text, txtEdit(textҽ���û�).Text, txtEdit(Textҽ������).Text, False) = False Then
            If MsgBox("ҽ�������������������ӣ��Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    
    If mcnZxTest.State = adStateClosed Then
        If OraDataOpen(mcnZxTest, txtEdit(Text���ķ�����).Text, txtEdit(text�����û�).Text, txtEdit(Text��������).Text, False) = False Then
            If MsgBox("ҽ�������������������ӣ��Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
        
    IsValid = True
End Function

Private Function SaveData() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lst As ListItem
    Dim lngRow As Long
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    'ɾ���Ѿ�����
    gstrSQL = "zl_���ղ���_Delete(" & TYPE_ǭ�� & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    With mshBill
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, mColHead.�շ����) <> "" Then
                '������������
                gstrSQL = "zl_���ղ���_Insert(" & TYPE_ǭ�� & ",null,'" & .TextMatrix(lngRow, mColHead.�շ����) & "' ,'" & .TextMatrix(lngRow, mColHead.������Ŀ) & "'," & lngRow + 2 & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
        Next
    End With
    
    '������������
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_ǭ�� & ",null,'ҽ���û���','" & txtEdit(textҽ���û�).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_ǭ�� & ",null,'ҽ���û�����','" & txtEdit(Textҽ������).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_ǭ�� & ",null,'ҽ��������','" & txtEdit(Textҽ��������).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_ǭ�� & ",null,'�����û���','" & txtEdit(text�����û�).Text & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_ǭ�� & ",null,'�����û�����','" & txtEdit(Text��������).Text & "',5)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_ǭ�� & ",null,'���ķ�����','" & txtEdit(Text���ķ�����).Text & "',6)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    SaveRegInFor g����ȫ��, "ҽ��", "������", IIf(chk������.Value = 1, 1, 0)
    gcnOracle.CommitTrans
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Function
Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub


Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m�ı�ʽ
End Sub
Private Function Open�м��() As Boolean
    '�����м��
    '�м������
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strServer As String, strPass As String
    
    Err = 0
    On Error GoTo errHand:
    
    gstrSQL = "select ������,����ֵ from ���ղ��� where ������ like 'ҽ��%' and ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�山ҽ��", TYPE_ǭ��)
    Do Until rsTemp.EOF
        Select Case rsTemp("������")
            Case "ҽ���û���"
                strUser = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ��������"
                strServer = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ���û�����"
                strPass = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        End Select
        rsTemp.MoveNext
    Loop
    Set gcnOracle_ǭ�� = New ADODB.Connection

    If OraDataOpen(gcnOracle_ǭ��, strServer, strUser, strPass, False) = False Then
        MsgBox "�޷����ӵ�ҽ���м�⣬���鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
        Exit Function
    End If


    Open�м�� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
