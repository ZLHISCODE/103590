VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmHomePage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ҳ����"
   ClientHeight    =   4005
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5880
   Icon            =   "frmHomePage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3240
      TabIndex        =   15
      Top             =   3585
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4470
      TabIndex        =   16
      Top             =   3585
      Width           =   1100
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   75
      TabIndex        =   17
      Top             =   3570
      Width           =   1100
   End
   Begin TabDlg.SSTab tbs 
      Height          =   3510
      Left            =   30
      TabIndex        =   18
      Top             =   15
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   6191
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "��ҳ��Ϣ"
      TabPicture(0)   =   "frmHomePage.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdTest"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cbo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "��ҳ����"
      TabPicture(1)   =   "frmHomePage.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblSize(1)"
      Tab(1).Control(1)=   "UsrPicture(1)"
      Tab(1).Control(2)=   "cmdOpen(1)"
      Tab(1).Control(3)=   "cmdClear(1)"
      Tab(1).Control(4)=   "cmdPos(1)"
      Tab(1).Control(5)=   "cmbMode"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "��������"
      TabPicture(2)   =   "frmHomePage.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdPos(2)"
      Tab(2).Control(1)=   "cmdClear(2)"
      Tab(2).Control(2)=   "cmdOpen(2)"
      Tab(2).Control(3)=   "UsrPicture(2)"
      Tab(2).Control(4)=   "lblSize(2)"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "ҽԺ��־"
      TabPicture(3)   =   "frmHomePage.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdPos(3)"
      Tab(3).Control(1)=   "cmdClear(3)"
      Tab(3).Control(2)=   "cmdOpen(3)"
      Tab(3).Control(3)=   "UsrPicture(3)"
      Tab(3).Control(4)=   "lblSize(3)"
      Tab(3).ControlCount=   5
      Begin VB.ComboBox cmbMode 
         Height          =   300
         Left            =   -70980
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   2010
         Width           =   1110
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Left            =   1245
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3015
         Width           =   2610
      End
      Begin VB.CommandButton cmdTest 
         Cancel          =   -1  'True
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   3960
         TabIndex        =   5
         Top             =   3000
         Width           =   1100
      End
      Begin VB.CommandButton cmdPos 
         Caption         =   "λ��(&P)"
         Height          =   350
         Index           =   3
         Left            =   -70980
         TabIndex        =   14
         Top             =   1605
         Width           =   1100
      End
      Begin VB.CommandButton cmdPos 
         Caption         =   "λ��(&P)"
         Height          =   350
         Index           =   2
         Left            =   -70980
         TabIndex        =   11
         Top             =   1590
         Width           =   1100
      End
      Begin VB.CommandButton cmdPos 
         Caption         =   "λ��(&P)"
         Height          =   350
         Index           =   1
         Left            =   -70980
         TabIndex        =   8
         Top             =   1515
         Width           =   1100
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "���(&L)"
         Height          =   350
         Index           =   3
         Left            =   -70980
         TabIndex        =   13
         Top             =   1020
         Width           =   1100
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "���(&L)"
         Height          =   350
         Index           =   2
         Left            =   -70980
         TabIndex        =   10
         Top             =   1005
         Width           =   1100
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "���(&L)"
         Height          =   350
         Index           =   1
         Left            =   -70980
         TabIndex        =   7
         Top             =   1020
         Width           =   1100
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "ͼƬ(&B)"
         Height          =   350
         Index           =   3
         Left            =   -70980
         TabIndex        =   12
         Top             =   600
         Width           =   1100
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "ͼƬ(&B)"
         Height          =   350
         Index           =   2
         Left            =   -70980
         TabIndex        =   9
         Top             =   600
         Width           =   1100
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "ͼƬ(&B)"
         Height          =   350
         Index           =   1
         Left            =   -70980
         TabIndex        =   6
         Top             =   600
         Width           =   1100
      End
      Begin VB.Frame Frame2 
         Caption         =   "��ҳͼƬ"
         Height          =   2535
         Left            =   150
         TabIndex        =   19
         Top             =   405
         Width           =   5085
         Begin VB.CommandButton cmdPos 
            Caption         =   "λ��(&P)"
            Height          =   350
            Index           =   0
            Left            =   3750
            TabIndex        =   2
            Top             =   1290
            Width           =   1100
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "���(&L)"
            Height          =   350
            Index           =   0
            Left            =   3750
            TabIndex        =   1
            Top             =   630
            Width           =   1100
         End
         Begin VB.CommandButton cmdOpen 
            Caption         =   "ͼƬ(&A)"
            Height          =   350
            Index           =   0
            Left            =   3750
            TabIndex        =   0
            Top             =   225
            Width           =   1100
         End
         Begin zl9NewQuery.ctlPicture UsrPicture 
            Height          =   2205
            Index           =   0
            Left            =   105
            TabIndex        =   20
            Top             =   240
            Width           =   3450
            _ExtentX        =   6085
            _ExtentY        =   3889
         End
         Begin VB.Label lblSize 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "800 X 600"
            Height          =   180
            Index           =   0
            Left            =   3675
            TabIndex        =   21
            Top             =   2175
            Width           =   1260
         End
      End
      Begin zl9NewQuery.ctlPicture UsrPicture 
         Height          =   2565
         Index           =   1
         Left            =   -74805
         TabIndex        =   22
         Top             =   600
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   4524
      End
      Begin zl9NewQuery.ctlPicture UsrPicture 
         Height          =   2565
         Index           =   2
         Left            =   -74805
         TabIndex        =   24
         Top             =   600
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   4524
      End
      Begin zl9NewQuery.ctlPicture UsrPicture 
         Height          =   2565
         Index           =   3
         Left            =   -74805
         TabIndex        =   26
         Top             =   600
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   4524
      End
      Begin VB.Label Label3 
         Caption         =   "��������(&M)"
         Height          =   270
         Left            =   135
         TabIndex        =   3
         Top             =   3075
         Width           =   1185
      End
      Begin VB.Label lblSize 
         Alignment       =   2  'Center
         Caption         =   "800 X 600"
         Height          =   180
         Index           =   3
         Left            =   -71175
         TabIndex        =   27
         Top             =   2580
         Width           =   1500
      End
      Begin VB.Label lblSize 
         Alignment       =   2  'Center
         Caption         =   "800 X 600"
         Height          =   180
         Index           =   2
         Left            =   -71175
         TabIndex        =   25
         Top             =   2625
         Width           =   1470
      End
      Begin VB.Label lblSize 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "800 X 600"
         Height          =   180
         Index           =   1
         Left            =   -71235
         TabIndex        =   23
         Top             =   2745
         Width           =   1605
      End
   End
End
Attribute VB_Name = "frmHomePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvarSvrPicRange As String           '��������ͼƬ�ķ�Χ
Private mvarSvrPicType As String            '��������ͼƬ������
Private mstrHomeCode As String

Private Sub cbo_Click()
    cmdOK.Tag = "1"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click(Index As Integer)
    UsrPicture(Index).Tag = ""
    UsrPicture(Index).Cls
    cmdOK.Tag = "1"
End Sub

Private Sub cmdOK_Click()
    Dim strSQL(1 To 4) As String
    Dim i As Long
    
    If cmdOK.Tag = "1" Then
        On Error GoTo errHand
        gcnOracle.BeginTrans
        strSQL(1) = "zl_��ѯҳ��Ŀ¼_delete(0)"
        strSQL(2) = "zl_��ѯҳ��Ŀ¼_insert(0,'��ҳ',1,0," & IIf(Val(UsrPicture(2).Tag) = 0, "NULL", Val(UsrPicture(2).Tag)) & "," & IIf(Val(UsrPicture(1).Tag) = 0, "NULL", Val(UsrPicture(1).Tag)) & "," & cbo.ItemData(cbo.ListIndex) & ",NULL,1,'" & mstrHomeCode & "','ZY')"
        strSQL(3) = "zl_��ѯ����Ŀ¼_insert(0,1,'',NULL,0,0,'����;12;0;0;0',0,NULL,0," & IIf(Val(UsrPicture(0).Tag) = 0, "NULL", Val(UsrPicture(0).Tag)) & ",0)"
        strSQL(4) = "zl_��ѯ����Ŀ¼_insert(0,2,'',NULL,0,0,'����;12;0;0;0',0,NULL,0," & IIf(Val(UsrPicture(3).Tag) = 0, "NULL", Val(UsrPicture(3).Tag)) & ",0)"
        For i = 1 To 4
            'gcnOracle.Execute strSQL(i), , adCmdStoredProc
            Call zlDatabase.ExecuteProcedure(strSQL(i), Me.Caption)
        Next
        gcnOracle.CommitTrans
    End If
    Unload Me
    Exit Sub
errHand:
    
    gcnOracle.RollbackTrans
    
    If ErrCenter() = -1 Then Resume
    
End Sub

Private Sub cmdOpen_Click(Index As Integer)
    Dim lngKey As Long
    Dim strFilter As String
    Dim strTitle As String
    
    Select Case Index
    Case 0
        strFilter = "9;0;1;2;3;4"
        strTitle = "�����ҳ����"
    Case 1
        strFilter = "4;0;1;2;3;9"
        strTitle = "�����ҳ����"
    Case 2
        strFilter = "1;0;2;3;4;9"
        strTitle = "�����ҳ��������"
    Case 3
        strFilter = "0;1;2;3;4;9"
        strTitle = "���ҽԺ��־ͼƬ"
    End Select
    If frmPicSelect.OpenPictureBox(Me, strTitle, strFilter, lngKey, mvarSvrPicRange, mvarSvrPicType) Then
        '����ͼƬ��ʾ
        UsrPicture(Index).Tag = lngKey
        Call ShowPicture(lngKey, Index)
        cmdOK.Tag = "1"
    End If
End Sub

Private Sub cmdPos_Click(Index As Integer)
    Select Case Index
    Case 0
        Call frmPosSample.ShowPageSample("��ҳͼƬ")
    Case 1
        Call frmPosSample.ShowPageSample("��ҳ����")
    Case 2
        Call frmPosSample.ShowPageSample("��������")
    Case 3
        Call frmPosSample.ShowPageSample("��־ͼƬ")
    End Select
End Sub

Private Sub cmdTest_Click()
    Dim vFileData As New FileSystemObject
    Dim strFile As String
    
    Call MusicClose
    
    
    If cbo.ListIndex < 0 Then Exit Sub
    If cbo.ItemData(cbo.ListIndex) <= 0 Then Exit Sub
    
    '1.���ͼ��Ŀ¼�Ƿ����
    On Error Resume Next
    vFileData.CreateFolder App.Path & "\ͼ��"
    
    '2.��鱾ϵͳ�п���ʹ�õ���ͼƬ
    gstrSQL = "select ���,����,����,�޸����� from ��ѯͼƬԪ�� where ���=[1]"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(cbo.ItemData(cbo.ListIndex)))
    If gRs.BOF Then Exit Sub
    
    strFile = IIf(IsNull(gRs!����), "", gRs!����)
    If strFile <> "" Then Call CheckFileNew(strFile, IIf(IsNull(gRs!����), 0, gRs!����), gRs!���, gRs!�޸�����, vFileData)
            
    Call MusicPlay(strFile)

End Sub

Private Sub Command1_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    mvarSvrPicRange = ""
    mvarSvrPicType = ""
    
    For i = 0 To lblSize.UBound
        lblSize(i).Caption = ""
    Next
    
    
    cmbMode.AddItem "ƽ��"
    cmbMode.AddItem "����"
    cmbMode.AddItem "����"
    Select Case GetPara("������ʾģʽ", "ƽ��")
        Case "����"
            cmbMode.ListIndex = 1
        Case "����"
            cmbMode.ListIndex = 2
        Case Else
            cmbMode.ListIndex = 0
    End Select
    
    cbo.AddItem "[��]"
    gstrSQL = "select ���,���� from ��ѯͼƬԪ�� where ����=3"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If gRs.BOF = False Then
        While Not gRs.EOF
            cbo.AddItem IIf(IsNull(gRs!����), "", gRs!����)
            cbo.ItemData(cbo.NewIndex) = IIf(IsNull(gRs!���), 0, gRs!���)
            gRs.MoveNext
        Wend
    End If
    cbo.ListIndex = 0
    
    '��ȡ��ҳ��Ϣ
    On Error GoTo errHand
    gstrSQL = "select ��������,ҳ�汳��,��������,���� from ��ѯҳ��Ŀ¼ where ҳ�����=0"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If gRs.BOF = False Then
        mstrHomeCode = IIf(IsNull(gRs!����), "", gRs!����)
        UsrPicture(1).Tag = IIf(IsNull(gRs!ҳ�汳��), 0, gRs!ҳ�汳��)
        UsrPicture(2).Tag = IIf(IsNull(gRs!��������), 0, gRs!��������)
        cbo.ListIndex = FindCboIndex(cbo, IIf(IsNull(gRs!��������), 0, gRs!��������))
        
        gstrSQL = "select A.��ͼ��� from ��ѯ����Ŀ¼ A where A.ҳ�����=0 and A.�������=1"
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If gRs.BOF = False Then UsrPicture(0).Tag = IIf(IsNull(gRs!��ͼ���), 0, gRs!��ͼ���)
        
        gstrSQL = "select A.��ͼ��� from ��ѯ����Ŀ¼ A where A.ҳ�����=0 and A.�������=2"
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If gRs.BOF = False Then UsrPicture(3).Tag = IIf(IsNull(gRs!��ͼ���), 0, gRs!��ͼ���)
        
        Call ShowPicture(Val(UsrPicture(0).Tag), 0)
        Call ShowPicture(Val(UsrPicture(1).Tag), 1)
        Call ShowPicture(Val(UsrPicture(2).Tag), 2)
        Call ShowPicture(Val(UsrPicture(3).Tag), 3)
    End If
            
    cmdOK.Tag = ""
    Exit Sub
errHand:
    cmdOK.Tag = ""
    If ErrCenter() = -1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowPicture(ByVal PicNo As Long, ByVal Index As Long)
    Dim rs As New ADODB.Recordset
    
    gstrSQL = "select ���,���,�߶�,���� from ��ѯͼƬԪ�� where ���=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PicNo)
    If rs.BOF = False Then
        Call UsrPicture(Index).ShowPictureByFieldNew(rs!���, rs!��� * Screen.TwipsPerPixelX, rs!�߶� * Screen.TwipsPerPixelY, IIf(IsNull(rs!����), 0, rs!����))
        If Index = 0 Then lblSize(Index).Caption = "���:" & Format(rs!��� * Screen.TwipsPerPixelX / 567, "0.0(����)") & vbCrLf & "�߶�:" & Format(rs!�߶� * Screen.TwipsPerPixelY / 567, "0.0(����)")
    End If
    CloseRecord rs
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Select Case cmbMode.ListIndex
        Case 1
            SetPara "������ʾģʽ", "����"
        Case 2
            SetPara "������ʾģʽ", "����"
        Case Else
            SetPara "������ʾģʽ", "ƽ��"
    End Select
End Sub

