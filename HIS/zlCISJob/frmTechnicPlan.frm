VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTechnicPlan 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ִ�б���"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   Icon            =   "frmTechnicPlan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboRoom 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   945
      Width           =   4620
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4725
      TabIndex        =   13
      Top             =   3600
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3540
      TabIndex        =   12
      Top             =   3600
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   285
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3585
      Width           =   1100
   End
   Begin VB.Frame fraDetail 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   -75
      TabIndex        =   23
      Top             =   1290
      Width           =   6250
      Begin VB.TextBox txtItem 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   1185
      End
      Begin VB.ComboBox cboSex 
         Height          =   300
         IMEMode         =   3  'DISABLE
         ItemData        =   "frmTechnicPlan.frx":000C
         Left            =   3570
         List            =   "frmTechnicPlan.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   915
         Width           =   930
      End
      Begin MSComCtl2.DTPicker dtpBirth 
         Height          =   300
         Left            =   1245
         TabIndex        =   5
         Top             =   915
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   98762755
         CurrentDate     =   38156
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   1
         Left            =   3570
         MaxLength       =   30
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   120
         Width           =   2295
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   4
         Left            =   5310
         MaxLength       =   10
         TabIndex        =   7
         Top             =   915
         Width           =   555
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   2
         Left            =   1245
         MaxLength       =   20
         TabIndex        =   3
         Top             =   510
         Width           =   1185
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   3570
         MaxLength       =   30
         TabIndex        =   4
         Top             =   525
         Width           =   2295
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   3570
         MaxLength       =   3
         TabIndex        =   9
         Top             =   1320
         Width           =   915
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   1245
         MaxLength       =   3
         TabIndex        =   8
         Top             =   1320
         Width           =   1185
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "������(&C)"
         Height          =   225
         Left            =   1245
         TabIndex        =   10
         Top             =   1740
         Width           =   1290
      End
      Begin VB.CheckBox chk��Ƭ 
         Caption         =   "���Ž�Ƭ(&F)"
         Height          =   225
         Left            =   3570
         TabIndex        =   11
         Top             =   1740
         Width           =   1290
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&H)"
         Height          =   180
         Index           =   0
         Left            =   420
         TabIndex        =   31
         Top             =   180
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "kg"
         Height          =   180
         Left            =   4560
         TabIndex        =   25
         Top             =   1380
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cm"
         Height          =   180
         Left            =   2475
         TabIndex        =   24
         Top             =   1380
         Width           =   180
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����豸(&D)"
         Height          =   180
         Index           =   8
         Left            =   2535
         TabIndex        =   15
         Top             =   180
         Width           =   990
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&A)"
         Height          =   180
         Index           =   6
         Left            =   4635
         TabIndex        =   20
         Top             =   975
         Width           =   630
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   1
         Left            =   600
         TabIndex        =   16
         Top             =   570
         Width           =   630
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ӣ����(&E)"
         Height          =   180
         Index           =   4
         Left            =   2715
         TabIndex        =   18
         Top             =   555
         Width           =   810
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�(&S)"
         Height          =   180
         Index           =   5
         Left            =   2895
         TabIndex        =   19
         Top             =   975
         Width           =   630
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&B)"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   975
         Width           =   990
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&W)"
         Height          =   180
         Index           =   3
         Left            =   2895
         TabIndex        =   22
         Top             =   1380
         Width           =   630
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���(&H)"
         Height          =   180
         Index           =   7
         Left            =   600
         TabIndex        =   21
         Top             =   1380
         Width           =   630
      End
   End
   Begin VB.Frame fraSplit2 
      Height          =   120
      Left            =   0
      TabIndex        =   30
      Top             =   3345
      Width           =   6420
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   8000
      Y1              =   765
      Y2              =   765
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   8000
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Label lblItemDetail 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C00000&
      Height          =   345
      Left            =   1185
      TabIndex        =   29
      Top             =   405
      Width           =   4575
   End
   Begin VB.Label lblRoom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ִ�м�(&R)"
      Height          =   180
      Left            =   330
      TabIndex        =   28
      Top             =   1005
      Width           =   810
   End
   Begin VB.Label lblPati 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ˣ�"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   645
      TabIndex        =   27
      Top             =   165
      Width           =   540
   End
   Begin VB.Label lblItemTit 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ŀ��"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   645
      TabIndex        =   26
      Top             =   420
      Width           =   540
   End
End
Attribute VB_Name = "frmTechnicPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngҽ��ID As Long
Private mlng���ͺ� As Long
Private mlngִ�п���ID As Long
Private mrsRoom As ADODB.Recordset
Private mblnOK As Boolean
Private mblnִ�б���ʱ���� As Boolean
Private mlng�����ID  As Long
Private mlng����ID As Long
Private mstrPrivs As String
Private mobjSquareCard As Object      '���������

Public Function ShowMe(objParent As Object, ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, Optional ByVal lngִ�п���ID As Long, Optional ByVal lng�����ID As Long, Optional ByVal lng����ID As Long, Optional ByVal strPrivs As String, Optional ByRef objSquareCard As Object) As Boolean
    mlngҽ��ID = lngҽ��ID
    mlng���ͺ� = lng���ͺ�
    mlngִ�п���ID = lngִ�п���ID
    mlng�����ID = lng�����ID
    mlng����ID = lng����ID
    mstrPrivs = strPrivs
    Set mobjSquareCard = objSquareCard
    
    On Local Error Resume Next
    Me.Show 1, objParent
    On Error GoTo 0
    
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    Dim blnTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strҽ��IDs As String
    Dim i As Long
     
    '�����������
    If fraDetail.Visible Then
        If zlCommFun.ActualLen(txtItem(1).Text) > txtItem(1).MaxLength Then
            MsgBox "����豸����������� " & txtItem(1).MaxLength \ 2 & " �����ֻ� " & txtItem(1).MaxLength & " ���ַ������顣", vbInformation, gstrSysName
            txtItem(1).SetFocus: Exit Sub
        End If
        If zlCommFun.ActualLen(txtItem(2).Text) > txtItem(2).MaxLength Then
            MsgBox "��������������� " & txtItem(2).MaxLength \ 2 & " �����ֻ� " & txtItem(2).MaxLength & " ���ַ������顣", vbInformation, gstrSysName
            txtItem(2).SetFocus: Exit Sub
        End If
        If zlCommFun.ActualLen(txtItem(4).Text) > txtItem(4).MaxLength Then
            MsgBox "��������������� " & txtItem(4).MaxLength \ 2 & " �����ֻ� " & txtItem(4).MaxLength & " ���ַ������顣", vbInformation, gstrSysName
            txtItem(4).SetFocus: Exit Sub
        End If
        If Trim(txtItem(1).Text) = "" Then
            MsgBox "���������豸��", vbInformation, gstrSysName
            txtItem(1).SetFocus: Exit Sub
        End If
        If Trim(txtItem(2).Text) = "" Then
            MsgBox "�����벡��������", vbInformation, gstrSysName
            txtItem(2).SetFocus: Exit Sub
        End If
    End If
    
    On Error GoTo errH
    '����һ��ͨ,����ִ�б���ǰ�������շѻ��ȼ������,�������ݺţ�����ҽ��ID��ȡ����δ�շѵ��ݻ�δ��˵ļ��ʵ�
    If mblnִ�б���ʱ���� Then
        '��ȡ����ҽ����ID��
        strSQL = "select a.ID from ����ҽ����¼ a where a.���id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҽ��ID)
        strҽ��IDs = mlngҽ��ID
        For i = 1 To rsTmp.RecordCount
            strҽ��IDs = strҽ��IDs & "," & rsTmp!ID
            rsTmp.MoveNext
        Next
        If Not mobjSquareCard Is Nothing Then
            If mobjSquareCard.zlSquareAffirm(Me, pҽ������վ, mstrPrivs, mlng����ID, mlng�����ID, False, , , strҽ��IDs) = False Then
                Exit Sub
            End If
        Else
            MsgBox "һ��ͨ������ʼ��ʧ�ܣ����鲿����", vbInformation, Me.Caption
            Exit Sub
        End If
    End If
    gcnOracle.BeginTrans: blnTrans = True
    If mlngִ�п���ID <> 0 Then
        strSQL = "Zl_����ҽ������_���ұ��(" & mlngҽ��ID & "," & mlng���ͺ� & "," & mlngִ�п���ID & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
    
    If fraDetail.Visible And Not IsNull(dtpBirth.Value) Then
        strSQL = "To_Date('" & Format(dtpBirth.Value, "yyyy-MM-dd") & "','YYYY-MM-DD')"
    Else
        strSQL = "NULL"
    End If
    strSQL = "ZL_����ҽ��ִ��_Plan(" & mlngҽ��ID & "," & mlng���ͺ� & ",1," & _
        "'" & cboRoom.Text & "','" & lblItemDetail.Tag & "'," & ZVal(txtItem(0).Text) & ",'" & txtItem(2).Text & "'," & _
        "'" & txtItem(3).Text & "','" & zlCommFun.GetNeedName(cboSex.Text) & "','" & txtItem(4).Text & "'," & _
        strSQL & "," & ZVal(txtItem(5).Text) & "," & ZVal(txtItem(6).Text) & "," & _
        chk����.Value & "," & chk��Ƭ.Value & ",'" & txtItem(1).Text & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0
    
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dtpBirth_Change()
    If Not IsNull(dtpBirth.Value) And IsNumeric(txtItem(4).Text) Then
        txtItem(4).Text = CInt(Format(zlDatabase.Currentdate, "yyyy")) - CInt(Format(dtpBirth.Value, "yyyy"))
        If Format(zlDatabase.Currentdate, "MMdd") < Format(dtpBirth.Value, "MMdd") Then
            txtItem(4).Text = CInt(txtItem(4).Text) - 1
        End If
        If CInt(txtItem(4).Text) < 0 Then txtItem(4).Text = ""
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Call cmdHelp_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        Call ZLCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    mblnOK = False
    On Error GoTo errH
    '�Ա��ֵ�
    cboSex.AddItem " "
    cboSex.ListIndex = 0
    strSQL = "Select ����,����,����,ȱʡ��־ From �Ա� Order by ����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        cboSex.AddItem rsTmp!���� & "-" & rsTmp!����
        If Nvl(rsTmp!ȱʡ��־, 0) = 1 Then
            cboSex.ListIndex = cboSex.NewIndex
        End If
        rsTmp.MoveNext
    Next
    
    'ִ����Ŀ����
    strSQL = _
        "Select A.ִ�в���ID,A.ִ�м�,B.ҽ������ as ����,D.����," & _
        " Nvl(D.����,C.����) as ����,Nvl(D.�Ա�,C.�Ա�) as �Ա�,Nvl(D.����,C.����) as ����," & _
        " Nvl(D.��������,C.��������) as ��������,D.Ӣ����,D.���,D.����," & _
        " D.������,D.���Ž�Ƭ,D.����豸," & _
        " F.���� as �������,E.Ӱ�����,E.���в���,E.�ɷ���Ƭ" & _
        " From ����ҽ������ A,����ҽ����¼ B,������Ϣ C," & _
            " Ӱ�����¼ D,Ӱ������Ŀ E,Ӱ������� F" & _
        " Where A.ҽ��ID=B.ID And B.����ID=C.����ID" & _
        " And B.������ĿID=E.������ĿID(+) And E.Ӱ�����=F.����(+)" & _
        " And A.ҽ��ID=D.ҽ��ID(+) And A.���ͺ�=D.���ͺ�(+)" & _
        " And A.ҽ��ID=[1] And A.���ͺ�=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҽ��ID, mlng���ͺ�)
    If rsTmp.EOF Then
        MsgBox "������ȷ��ȡִ����Ŀ��Ϣ��", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    If IsNull(rsTmp!Ӱ�����) Then
        '��ִͨ����Ŀ
        fraDetail.Visible = False
        Me.Height = Me.Height - fraDetail.Height + 300
        fraSplit2.Top = fraSplit2.Top - fraDetail.Height + 300
        cmdHelp.Top = cmdHelp.Top - fraDetail.Height + 300
        cmdOK.Top = cmdOK.Top - fraDetail.Height + 300
        cmdCancel.Top = cmdCancel.Top - fraDetail.Height + 300
        
        lblRoom.Top = lblRoom.Top + 100
        cboRoom.Top = cboRoom.Top + 100
        lblRoom.Left = lblRoom.Left + 500
        cboRoom.Left = cboRoom.Left + 500
        cboRoom.Width = cboRoom.Width - 1000
        
        lblPati.Caption = "���ˣ�" & rsTmp!����
        lblItemDetail.Caption = Nvl(rsTmp!����)
    Else
        'Ӱ������Ŀ
        lblPati.Caption = "���ˣ�" & rsTmp!���� & "  Ӱ�����" & rsTmp!Ӱ����� & "-" & rsTmp!�������
        lblItemTit.Caption = "Ӱ���飺"
        lblItemDetail.Caption = Nvl(rsTmp!����)
        lblItemDetail.Tag = rsTmp!Ӱ�����
        
        If Not IsNull(rsTmp!����) Then
            txtItem(0).Text = rsTmp!����
        Else
            txtItem(0).Text = Next����(rsTmp!Ӱ�����)
        End If
        txtItem(1).Text = Nvl(rsTmp!����豸)
        txtItem(2).Text = rsTmp!����
        If IsNull(rsTmp!Ӣ����) Then
            txtItem(3).Text = ZLCommFun.SpellCode(rsTmp!����)
        Else
            txtItem(3).Text = rsTmp!Ӣ����
        End If
        If IsNull(rsTmp!��������) Then
            dtpBirth.Value = Empty
        Else
            dtpBirth.Value = rsTmp!��������
        End If
        If Not IsNull(rsTmp!�Ա�) Then
             Cbo.SeekIndex cboSex, rsTmp!�Ա�, True
        End If
        txtItem(4).Text = Nvl(rsTmp!����)
        txtItem(5).Text = Nvl(rsTmp!���)
        txtItem(6).Text = Nvl(rsTmp!����)
        
        chk����.Value = IIf(Nvl(rsTmp!������, 0) = 0, 0, 1)
        If Nvl(rsTmp!���в���, 0) = 0 Then
            chk����.Enabled = False
            chk����.Value = 0
        ElseIf Nvl(rsTmp!���в���, 0) = 1 Then
            chk����.Enabled = False
            chk����.Value = 1
        End If
        
        chk��Ƭ.Value = IIf(Nvl(rsTmp!���Ž�Ƭ, 0) = 0, 0, 1)
        If Nvl(rsTmp!�ɷ���Ƭ, 0) = 0 Then
            chk��Ƭ.Enabled = False
            chk��Ƭ.Value = 0
        ElseIf Nvl(rsTmp!�ɷ���Ƭ, 0) = 1 Then
            chk��Ƭ.Enabled = False
            chk��Ƭ.Value = 1
        End If
    End If
    mblnִ�б���ʱ���� = Val(zlDatabase.GetPara("ִ�б���ʱ�շѻ�������", glngSys, pҽ������վ)) = 1
    'ִ�м�����
    strSQL = "Select ����ID,ִ�м�,��ǰ����,����豸,���� From ҽ��ִ�з��� Where ����ID=[1]"
    'Set mrsRoom = New ADODB.Recordset
    Set mrsRoom = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(mlngִ�п���ID <> 0, mlngִ�п���ID, Val(rsTmp!ִ�в���ID)))
    If mrsRoom.EOF And IsNull(rsTmp!Ӱ�����) Then
        MsgBox "��ǰ���һ�û������ִ�м䣬�������á�", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    For i = 1 To mrsRoom.RecordCount
        cboRoom.AddItem mrsRoom!ִ�м�
        mrsRoom.MoveNext
    Next
    If cboRoom.ListCount > 0 Then
        cboRoom.ListIndex = 0
    End If
    If Not IsNull(rsTmp!ִ�м�) Then
        Call Cbo.SeekIndex(cboRoom, rsTmp!ִ�м�, True)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlngҽ��ID = 0
    mlng���ͺ� = 0
    Set mrsRoom = Nothing
    Set mobjSquareCard = Nothing
End Sub

Private Sub txtItem_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtItem(Index))
End Sub

Private Sub txtItem_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 5 Or Index = 6 Then
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    ElseIf Index = 3 Then
        If Not Between(KeyAscii, 32, 128) Then KeyAscii = 0
    End If
End Sub

Private Sub txtItem_Validate(Index As Integer, Cancel As Boolean)
    If Index = 5 Or Index = 6 Then
        If Val(txtItem(Index).Text) = 0 Then
            txtItem(Index).Text = ""
        End If
    End If
End Sub

Public Function Next����(str��� As String) As Double
    Dim rsCtrl As New ADODB.Recordset
    Dim strSQL As String, dblNO As Double

ReStart:
    err = 0
    On Error GoTo errH
    With rsCtrl
        If .State = 1 Then .Close
        strSQL = "Select ����,������,����,����,���� From Ӱ������� Where ����='" & str��� & "'"
        Call SQLTest(App.ProductName, Me.Caption, strSQL)
        .CursorLocation = adUseClient
        .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
        Call SQLTest
        If .EOF Then Exit Function
        
        dblNO = Val(Nvl(!������, 0)) + 1
        
        On Error Resume Next
        .Update "������", dblNO
        If err <> 0 Then
            .CancelUpdate
            GoTo ReStart
        End If
        Next���� = dblNO
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
