VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChargeWholeSetClassEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����շ���Ŀ����༭"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6150
   Icon            =   "frmChargeWholeSetClassEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   375
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3135
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   1950
      MaxLength       =   10
      TabIndex        =   9
      Tag             =   "����"
      Text            =   "000000"
      Top             =   1545
      Width           =   1380
   End
   Begin VB.TextBox txtSymbol 
      Height          =   300
      Left            =   1815
      MaxLength       =   10
      TabIndex        =   8
      Tag             =   "����"
      Top             =   2265
      Width           =   1425
   End
   Begin VB.CheckBox chkCodeLen 
      Caption         =   "������ı��볤�ȣ������˵�����ͬ������(&L)"
      Height          =   285
      Left            =   1095
      TabIndex        =   7
      Top             =   2685
      Width           =   4290
   End
   Begin VB.Frame fraBottom 
      Height          =   75
      Left            =   -30
      TabIndex        =   6
      Top             =   2970
      Width           =   6345
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3345
      TabIndex        =   5
      Top             =   3150
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4575
      TabIndex        =   4
      Top             =   3150
      Width           =   1100
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1815
      MaxLength       =   40
      TabIndex        =   3
      Tag             =   "����"
      Top             =   1905
      Width           =   3795
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&P"
      Height          =   300
      Left            =   5325
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   885
      Width           =   285
   End
   Begin VB.TextBox txtParent 
      Height          =   300
      Left            =   1815
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "(��)"
      ToolTipText     =   "��Del����ϼ������ó�������"
      Top             =   885
      Width           =   3495
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   2310
      Left            =   -3720
      TabIndex        =   0
      Tag             =   "1000"
      Top             =   2100
      Visible         =   0   'False
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   4075
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   135
      Top             =   1275
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeWholeSetClassEdit.frx":0442
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeWholeSetClassEdit.frx":09DC
            Key             =   "expend"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtUpCode 
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   1815
      MaxLength       =   10
      TabIndex        =   11
      TabStop         =   0   'False
      Tag             =   "����"
      Text            =   "0000"
      Top             =   1500
      Width           =   1620
   End
   Begin VB.Label lblSymbol 
      AutoSize        =   -1  'True
      Caption         =   "����(&S)"
      Height          =   180
      Left            =   1095
      TabIndex        =   18
      Top             =   2325
      Width           =   630
   End
   Begin VB.Label lblKind 
      AutoSize        =   -1  'True
      Caption         =   "���׷���"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   120
      TabIndex        =   17
      Top             =   735
      Width           =   720
   End
   Begin VB.Label lblNote 
      Caption         =   "    �����շ���Ŀ�ɸ����ٴ���ҽ������Ӧ�ò������ص����ͳһ�������á�"
      Height          =   435
      Left            =   1065
      TabIndex        =   16
      Top             =   255
      Width           =   4755
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   285
      Picture         =   "frmChargeWholeSetClassEdit.frx":0F76
      Top             =   225
      Width           =   480
   End
   Begin VB.Label lblCode 
      AutoSize        =   -1  'True
      Caption         =   "����(&D)"
      Height          =   180
      Left            =   1095
      TabIndex        =   15
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "����(&N)"
      Height          =   180
      Left            =   1095
      TabIndex        =   14
      Top             =   1950
      Width           =   630
   End
   Begin VB.Label lblParent 
      AutoSize        =   -1  'True
      Caption         =   "�ϼ�(&U)"
      Height          =   180
      Left            =   1095
      TabIndex        =   13
      Top             =   945
      Width           =   630
   End
   Begin VB.Label lblHint 
      Caption         =   "(��ʾ����Del����ϼ������ó�������)"
      Height          =   210
      Left            =   1785
      TabIndex        =   12
      Top             =   1245
      Width           =   3330
   End
End
Attribute VB_Name = "frmChargeWholeSetClassEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum EditWhileSetType
    Ed_���� = 1
    Ed_�޸� = 2
End Enum
Private mEditType As EditWhileSetType
Private mstrPrivs As String, mlngModule As Long
Private mstrID As String
Private mintMaxLen As Integer
Private mObjNode As Node
Private mblnChanged As Boolean
Private mintSucces As Integer
Private mlng�ϼ�ID As Long
Private mstrLike  As String
Private mblnFirst As Boolean
Public Function EditCard(ByVal frmMain As Form, ByVal EditType As EditWhileSetType, _
    ByVal strPrivs As String, ByVal lngModule As Long, ByVal lng�ϼ�ID As Long, ByVal strID As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������༭�ӿ�
    '���:
    '����:
    '����:�༭�ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2010-08-26 13:30:38
    '˵��:27327
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mEditType = EditType: mstrPrivs = strPrivs: mlngModule = lngModule: mstrID = strID: mintSucces = 0
    mlng�ϼ�ID = lng�ϼ�ID
    Me.Show 1, frmMain
    EditCard = mintSucces > 0
End Function
Private Sub chkCodeLen_Click()
    On Error GoTo ErrHandle
    If Me.chkCodeLen.value = 1 Then
        Me.txtCode.MaxLength = mintMaxLen - Len(Me.txtUpCode.Text)
    Else
        Me.txtCode.MaxLength = Me.txtCode.Tag
        Me.txtCode.Text = Mid(Me.txtCode.Text, 1, Me.txtCode.MaxLength)
    End If
    Me.txtCode.SetFocus
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chkCodeLen_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim lngItemID As Long
    
    If Me.txtCode.MaxLength = 0 Then
        MsgBox "�ϼ������Ѿ��ﵽ��󳤶ȣ����������¼���", vbExclamation, gstrSysName
        Me.cmdCancel.SetFocus
        Exit Sub
    End If
    If Trim(Me.txtCode.Text) = "" Then
        MsgBox "�����������", vbExclamation, gstrSysName
        Me.txtCode.SetFocus
        Exit Sub
    End If
    If Me.chkCodeLen.value = 0 And Len(Trim(Me.txtCode.Text)) <> Me.txtCode.MaxLength Then
        MsgBox "���볤�ȱ���Ϊ" & Me.txtCode.MaxLength & "λ��������ѡ����ĳ���ѡ��", vbExclamation, gstrSysName
        Me.txtCode.SetFocus
        Exit Sub
    End If
    If Trim(Me.txtName.Text) = "" Then
        MsgBox "���Ʊ�������", vbExclamation, gstrSysName
        Me.txtName.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txtName.Text), vbFromUnicode)) > Me.txtName.MaxLength Then
        MsgBox "���Ƴ���" & Me.txtName.MaxLength & "�ĳ�������", vbExclamation, gstrSysName
        Me.txtName.SetFocus
        Exit Sub
    End If
    
    Err = 0: On Error GoTo ErrHand
    If mEditType = Ed_���� Then
        lngItemID = sys.NextId("������Ŀ����")
        'Zl_������Ŀ����_Insert
        gstrSQL = "ZL_������Ŀ����_INSERT("
    Else
        lngItemID = Val(mstrID)
        'Zl_������Ŀ����_Update
        gstrSQL = "ZL_������Ŀ����_UPDATE("
    End If
    '  Id_In      ������Ŀ����.ID%Type,
    gstrSQL = gstrSQL & "" & lngItemID & ","
    '  �ϼ�id_In  ������Ŀ����.�ϼ�id%Type,
    gstrSQL = gstrSQL & "" & IIF(Val(Me.txtParent.Tag) = 0, "NULL", Val(Me.txtParent.Tag)) & ","
    '  ����_In    ������Ŀ����.����%Type,
    gstrSQL = gstrSQL & "'" & Me.txtUpCode.Text & Trim(Me.txtCode.Text) & "',"
    '  ����_In    ������Ŀ����.����%Type,
    gstrSQL = gstrSQL & "'" & Trim(Me.txtName.Text) & "',"
    '  ����_In    ������Ŀ����.����%Type,
    gstrSQL = gstrSQL & "'" & Trim(Me.txtSymbol.Text) & "',"
    '  v_Brethren Number
    '  --�Ƿ��ͬ��������г��ȴ���,0-��,1-��
    gstrSQL = gstrSQL & "" & Me.chkCodeLen.value & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    mintSucces = mintSucces + 1
    mblnChanged = False
    If mEditType = Ed_�޸� Then Unload Me: Exit Sub
    txtName.Text = ""
    Call zlDefaultCode
    mblnChanged = False
    txtName.SetFocus
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdSelect_Click()
    Call SearchPreLevel("")
End Sub
Private Function SearchPreLevel(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ���ϼ�����
    '����:
    '����:���˺�
    '����:2010-08-26 13:39:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strKey As String, strWhere As String
    Dim vRect As RECT, bytStyle As Byte
    Dim blnCancel As Boolean
    
    On Error GoTo ErrHandle
    strKey = mstrLike & strInput & "%"
    If strInput <> "" Then
        If IsNumeric(strInput) Then
            strWhere = " ���� Like [1]"
        ElseIf zlStr.IsCharAlpha(strInput) Then
            strWhere = " ���� Like upper([1])"
        Else
            strWhere = " ���� Like [1] or ���� Like upper([1]) or ���� like [1]"
        End If
        If mEditType = Ed_�޸� Then
            strWhere = "( " & strWhere & ")  and ID not in (select id from ������Ŀ���� start with ID = [2] connect by prior id=�ϼ�id )"
        End If
        gstrSQL = "" & _
        " Select ID,�ϼ�ID,����,����,����" & _
        " From ������Ŀ����" & _
        " Where " & strWhere
        bytStyle = 0
    Else
        If mEditType = Ed_�޸� Then
            gstrSQL = "" & _
            " Select ID,�ϼ�ID,����,����,����" & _
            " From ������Ŀ����" & _
            " Where id not in (select id from ������Ŀ���� start with ID = [2] connect by prior id=�ϼ�id ) " & _
            " Start with �ϼ�ID is null  connect by prior ID=�ϼ�ID"
        Else
            gstrSQL = "" & _
            " Select ID,�ϼ�ID,����,����,����" & _
            " From ������Ŀ����" & _
            " Start with �ϼ�ID is null Connect by prior ID=�ϼ�ID"
        End If
        bytStyle = 1
    End If
    
    vRect = zlControl.GetControlRect(txtParent.hwnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, bytStyle, "�����շ���Ŀ����", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, txtParent.Height, blnCancel, False, True, strKey, mlng�ϼ�ID)
    
    If blnCancel = True Then
        If txtParent.Enabled And txtParent.Visible Then txtParent.SetFocus
        Exit Function
    End If
    If rsTemp Is Nothing Then
        MsgBox "δ�ҵ�ƥ��ķ�����Ϣ,����!", vbInformation + vbDefaultButton1, gstrSysName
        If txtParent.Enabled And txtParent.Visible Then txtParent.SetFocus
        Exit Function
    End If
    If rsTemp.State <> 1 Then
        MsgBox "δ�ҵ�ƥ��ķ�����Ϣ,����!", vbInformation + vbDefaultButton1, gstrSysName
        If txtParent.Enabled And txtParent.Visible Then txtParent.SetFocus
        Exit Function
    End If
    txtParent.Text = Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����)
    txtParent.Tag = Nvl(rsTemp!ID)
    Call zlDefaultCode
    If txtName.Enabled And txtName.Visible Then txtName.SetFocus
    SearchPreLevel = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ReadCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-08-26 13:35:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrHandle
    Select Case mEditType
    Case Ed_����
        gstrSQL = "" & _
        " Select A.ID,A.�ϼ�ID,A.����,A.����,A.����" & _
        " From ������Ŀ���� A" & _
        " Where id=0"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID))
        mintMaxLen = rsTemp.Fields("����").DefinedSize
        Me.txtName.MaxLength = rsTemp.Fields("����").DefinedSize
        Me.txtSymbol.MaxLength = rsTemp.Fields("����").DefinedSize
        Me.txtParent.Tag = mlng�ϼ�ID
        Call zlDefaultCode
    Case Else
        gstrSQL = "" & _
        " Select A.ID,A.�ϼ�ID,A.����,A.����,A.����,B.���� as �ϼ�����,B.���� as �ϼ�����" & _
        " From ������Ŀ���� A,������Ŀ���� B" & _
        " Where A.ID=[1] and A.�ϼ�id=b.id(+)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID))
        If rsTemp.EOF Then
            MsgBox "�÷�������Ѿ�������ɾ��,���ܽ����޸�!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
        mlng�ϼ�ID = Val(Nvl(rsTemp!�ϼ�id))
        txtParent.Text = IIF(mlng�ϼ�ID = 0, "��", Nvl(rsTemp!�ϼ�����) & "-" & Nvl(rsTemp!�ϼ�����))
        txtParent.Tag = mlng�ϼ�ID
        txtUpCode.Text = Nvl(rsTemp!�ϼ�����)
        txtCode.Text = Mid(Nvl(rsTemp!����), Len(txtUpCode.Text) + 1)
        txtCode.MaxLength = Len(txtCode.Text)
        txtName.Text = Nvl(rsTemp!����)
        txtSymbol.Text = Nvl(rsTemp!����)
        mintMaxLen = rsTemp.Fields("����").DefinedSize
        Me.txtName.MaxLength = rsTemp.Fields("����").DefinedSize
        Me.txtSymbol.MaxLength = rsTemp.Fields("����").DefinedSize
    End Select
    ReadCard = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Call ReadCard
    Me.txtCode.ZOrder
    mblnChanged = False
    If txtName.Enabled And txtName.Visible Then txtName.SetFocus
End Sub

Private Sub Form_Load()
    mstrLike = IIF(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    mblnChanged = False: mblnFirst = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnChanged = True Then
        If MsgBox("�����Ѿ��ı䣬��ȷ��Ҫ�˳���", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub lblCode_Click()
    Me.txtCode.SetFocus
End Sub

Private Sub lblName_Click()
    Me.txtName.SetFocus
End Sub

Private Sub lblParent_Click()
    Me.txtParent.SetFocus
End Sub

Private Sub lblSymbol_Click()
    Me.txtSymbol.SetFocus
End Sub

Private Sub tvwClass_DblClick()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Me.txtParent.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    Me.txtParent.SetFocus
    Call zlDefaultCode
End Sub

Private Sub tvwClass_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
        If Me.tvwClass.SelectedItem.Children > 0 Then Exit Sub
        Call tvwClass_DblClick
    Case vbKeySpace
        Call tvwClass_DblClick
    Case vbKeyEscape
        Call tvwClass_LostFocus
    End Select
End Sub

Private Sub tvwClass_LostFocus()
    If cmdSelect Is ActiveControl Then Exit Sub
    Me.tvwClass.Visible = False
End Sub

Private Sub txtCode_Change()
    mblnChanged = True
End Sub

Private Sub txtCode_GotFocus()
    Me.txtCode.SelStart = 0: Me.txtCode.SelLength = 100
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtName_Change()
    mblnChanged = True
End Sub

Private Sub txtName_GotFocus()
    Call OS.OpenIme(True)
    Me.txtName.SelStart = 0: Me.txtName.SelLength = 100
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*_+|=`;'"":/.,?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub txtName_KeyUp(KeyCode As Integer, Shift As Integer)
    Me.txtSymbol.Text = Mid(zlStr.GetCodeByORCL(Me.txtName.Text), 1, 10)
End Sub

Private Sub txtName_LostFocus()
    Call OS.OpenIme(False)
End Sub

Private Sub txtParent_Change()
    mblnChanged = True
End Sub

Private Sub txtParent_GotFocus()
    Me.txtParent.SelStart = 0: Me.txtParent.SelLength = 100
End Sub

Private Sub txtParent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Me.txtParent.Tag = 0
        Call zlDefaultCode
    End If
    Me.txtParent.SelStart = 0: Me.txtParent.SelLength = 100
End Sub

Private Sub txtParent_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtSymbol_Change()
    mblnChanged = True
End Sub

Private Sub txtSymbol_GotFocus()
    Me.txtSymbol.SelStart = 0: Me.txtSymbol.SelLength = 100
End Sub

Private Sub txtSymbol_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtUpCode_Change()
    Me.txtCode.Width = txtUpCode.Width - TextWidth(txtUpCode.Text) - 120
    Me.txtCode.Left = txtUpCode.Left + TextWidth(txtUpCode.Text) + 60
    mblnChanged = True
End Sub

Private Sub zlDefaultCode()
    '-----------------------------------------------------
    '���ܣ�����ѡ����ϼ�ID(�����txtParent.Tag))���������ñ����ȱʡֵ
    '-----------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Err = 0: On Error GoTo ErrHand
    Me.chkCodeLen.value = 0
    Me.chkCodeLen.Enabled = True
NotPreID:
    If Val(txtParent.Tag) = 0 Then
        Me.txtParent.Text = "(��)"
        Me.txtUpCode.Text = ""
        gstrSQL = "select max(����) as ���� From ������Ŀ���� Where �ϼ�ID is null "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        With rsTemp
            If IIF(IsNull(!����), "", !����) = "" Then
                Me.txtCode.Text = "01"
                Me.txtCode.MaxLength = mintMaxLen
                Me.txtCode.Tag = Me.txtCode.MaxLength
                Me.chkCodeLen.value = 1
                Me.chkCodeLen.Enabled = False
            Else
                Me.txtCode.MaxLength = Len(Trim(Nvl(!����)))
                Me.txtCode.Tag = Me.txtCode.MaxLength
                If Nvl(!����) = String(Me.txtCode.MaxLength, "9") Then
                    If Me.txtCode.MaxLength >= mintMaxLen Then
                        MsgBox "������ͱ��볤���Ѿ��ﵽ������ƣ��޷���������", vbExclamation, gstrSysName
                        Me.txtCode.Text = Space(Me.txtCode.MaxLength)
                        Me.chkCodeLen.value = 0
                        Me.chkCodeLen.Enabled = False
                    Else
                        MsgBox "�������Ѿ��ﵽ�������ƣ������������볤����������Ҫ", vbExclamation, gstrSysName
                        Me.txtCode.Text = "1" & String(Me.txtCode.MaxLength, "0")
                        Me.txtCode.MaxLength = Me.txtCode.MaxLength + 1
                        Me.txtCode.Tag = Me.txtCode.MaxLength
                        Me.chkCodeLen.value = 1
                    End If
                Else
                    Me.txtCode.Text = Format(Mid(Nvl(!����), Len(Me.txtUpCode.Text) + 1) + 1, String(Me.txtCode.MaxLength, "0"))
                End If
            End If
        End With
    Else
        gstrSQL = "select ����,���� From ������Ŀ���� Where ID =[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(txtParent.Tag))
        If rsTemp.EOF Then
            MsgBox "�ϼ�������ܱ�����ɾ��,����!", vbInformation + vbDefaultButton1, gstrSysName
            mlng�ϼ�ID = 0: GoTo NotPreID:
        End If
        Me.txtParent.Text = Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����)
        Me.txtUpCode.Text = Nvl(rsTemp!����)
        Me.txtCode.MaxLength = IIF(mintMaxLen - Len(Me.txtUpCode.Text) > 0, mintMaxLen - Len(Me.txtUpCode.Text), 1)
        Me.txtCode.Tag = Me.txtCode.MaxLength
        
        gstrSQL = "select nvl(����,'') as ����  From ������Ŀ���� Where �ϼ�ID=[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(txtParent.Tag))
        If rsTemp.EOF Then
            'û������,
            If Me.txtCode.MaxLength > 1 Then
                Me.txtCode.Text = "01"
            Else
                Me.txtCode.Text = "1"
            End If
            Me.chkCodeLen.value = 1
            Me.chkCodeLen.Enabled = False
        Else
            With rsTemp
                Me.txtCode.MaxLength = IIF(Len(Nvl(!����)) - Len(Me.txtUpCode.Text) > 0, Len(Nvl(!����)) - Len(Me.txtUpCode.Text), 1)
                Me.txtCode.Tag = Me.txtCode.MaxLength
                If Mid(Nvl(!����), Len(Me.txtUpCode.Text) + 1) = String(Me.txtCode.MaxLength, "9") Then
                    If Len(Me.txtUpCode.Text) + Me.txtCode.MaxLength >= mintMaxLen Then
                        MsgBox "�÷����¼�������ͱ��볤���Ѿ��ﵽ������ƣ��޷���������", vbExclamation, gstrSysName
                        Me.txtCode.Text = Space(Me.txtCode.MaxLength)
                        Me.chkCodeLen.value = 0
                        Me.chkCodeLen.Enabled = False
                    Else
                        MsgBox "�÷����¼��������Ѿ��ﵽ�������ƣ������������볤����������Ҫ", vbExclamation, gstrSysName
                        Me.txtCode.Text = "1" & String(Me.txtCode.MaxLength, "0")
                        Me.txtCode.MaxLength = Me.txtCode.MaxLength + 1
                        Me.txtCode.Tag = Me.txtCode.MaxLength
                        Me.chkCodeLen.value = 1
                    End If
                Else
                    If Len(Nvl(!����)) >= Len(Me.txtUpCode.Text) + 1 Then
                        Me.txtCode.Text = Format(Mid(Nvl(!����), Len(Me.txtUpCode.Text) + 1) + 1, String(Me.txtCode.MaxLength, "0"))
                    End If
                End If
            End With
        End If
    End If
    Me.txtParent.SetFocus
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub




