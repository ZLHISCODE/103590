VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAuditItemEdit 
   BorderStyle     =   0  'None
   ClientHeight    =   5190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraAuditItem 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5145
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   7965
      Begin VB.Frame fraSource 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   2445
         TabIndex        =   29
         Top             =   60
         Visible         =   0   'False
         Width           =   2970
         Begin VB.OptionButton optSource 
            Caption         =   "��׼��"
            Height          =   195
            Index           =   0
            Left            =   1050
            TabIndex        =   31
            Top             =   188
            Value           =   -1  'True
            Width           =   840
         End
         Begin VB.OptionButton optSource 
            Caption         =   "EMR��"
            Height          =   180
            Index           =   1
            Left            =   2070
            TabIndex        =   30
            Top             =   195
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "����Դ(&S)"
            Height          =   165
            Left            =   120
            TabIndex        =   32
            Top             =   203
            Width           =   825
         End
      End
      Begin VB.ComboBox CboPalValue 
         Height          =   300
         Left            =   5265
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   555
         Width           =   960
      End
      Begin VB.TextBox txtNumValue 
         DataField       =   "����"
         Height          =   300
         Left            =   6990
         MaxLength       =   4
         TabIndex        =   25
         ToolTipText     =   "���۷�ֵ"
         Top             =   570
         Width           =   975
      End
      Begin VB.TextBox txtFileID 
         DataField       =   "�ļ�ID"
         Height          =   300
         Left            =   4710
         TabIndex        =   24
         Top             =   5310
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.TextBox txtCode 
         DataField       =   "����"
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1065
         MaxLength       =   10
         TabIndex        =   22
         Top             =   960
         Width           =   1710
      End
      Begin VB.TextBox txtMnemonicCode 
         DataField       =   "����"
         Height          =   300
         Left            =   3585
         MaxLength       =   100
         TabIndex        =   21
         Top             =   555
         Width           =   885
      End
      Begin VB.CommandButton cmdSelectFile 
         Height          =   300
         Left            =   6210
         Picture         =   "frmAuditItemEdit.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1965
         Width           =   300
      End
      Begin VB.ComboBox cboLink 
         DataField       =   "���ö���"
         Height          =   300
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1950
         Width           =   1740
      End
      Begin VB.CommandButton cmdCheck 
         Height          =   300
         Left            =   2865
         Picture         =   "frmAuditItemEdit.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3345
         Width           =   300
      End
      Begin VB.ComboBox cboUsed 
         DataField       =   "���ö���"
         Height          =   300
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1455
         Width           =   1740
      End
      Begin VB.TextBox txtAudit_NotCheck 
         DataField       =   "�������"
         Height          =   780
         Left            =   1050
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         ToolTipText     =   "�������"
         Top             =   3345
         Width           =   1740
      End
      Begin VB.TextBox txtName 
         DataField       =   "����"
         Height          =   300
         Left            =   1065
         MaxLength       =   100
         TabIndex        =   4
         Top             =   570
         Width           =   1710
      End
      Begin VB.TextBox txtDescription 
         DataField       =   "˵��"
         Height          =   855
         Left            =   1050
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         ToolTipText     =   "˵��"
         Top             =   2385
         Width           =   6915
      End
      Begin VB.TextBox txtTypeID 
         DataField       =   "����"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1065
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   210
         Width           =   6450
      End
      Begin VB.CommandButton cmdSelect 
         Height          =   300
         Left            =   7560
         Picture         =   "frmAuditItemEdit.frx":6930
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   210
         Width           =   300
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfFiles 
         Height          =   1320
         Left            =   3840
         TabIndex        =   19
         Top             =   915
         Width           =   4200
         _cx             =   7408
         _cy             =   2328
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAuditItemEdit.frx":D182
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&M)"
         Height          =   180
         Index           =   8
         Left            =   4560
         TabIndex        =   27
         Top             =   615
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ֵ(&N)"
         Height          =   180
         Index           =   4
         Left            =   6285
         TabIndex        =   26
         Top             =   615
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&S)"
         Height          =   180
         Index           =   3
         Left            =   2850
         TabIndex        =   23
         Top             =   615
         Width           =   630
      End
      Begin VB.Label lab����ID 
         AutoSize        =   -1  'True
         Caption         =   "[����ID]��[��ҳID]Ϊϵͳ�������ֱ����ϵͳ�еĲ���ID����ҳID��"
         Height          =   180
         Left            =   1515
         TabIndex        =   18
         Top             =   4485
         Width           =   5550
      End
      Begin VB.Label labNote 
         AutoSize        =   -1  'True
         Caption         =   "ע��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   1065
         TabIndex        =   17
         Top             =   4485
         Width           =   390
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���û���(&L)"
         Height          =   180
         Index           =   7
         Left            =   30
         TabIndex        =   8
         Top             =   2010
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ö���(&D)"
         Height          =   180
         Index           =   5
         Left            =   30
         TabIndex        =   6
         Top             =   1515
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������(&X)"
         Height          =   180
         Index           =   1
         Left            =   30
         TabIndex        =   12
         Top             =   3360
         Width           =   990
      End
      Begin VB.Label lstFiles 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&T)"
         Height          =   180
         Index           =   4
         Left            =   390
         TabIndex        =   0
         Top             =   270
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "˵��(&Z)"
         Height          =   180
         Index           =   9
         Left            =   390
         TabIndex        =   10
         Top             =   2415
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   2
         Left            =   390
         TabIndex        =   3
         Top             =   645
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&C)"
         Height          =   180
         Index           =   0
         Left            =   390
         TabIndex        =   5
         Top             =   1020
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ļ�(&J)"
         Height          =   180
         Index           =   6
         Left            =   3090
         TabIndex        =   16
         Top             =   1500
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmAuditItemEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mEditMode               As �༭ģʽ

Private mblnDataChange          As Boolean  'ȷ�� or ȡ��

Private zlCheck                 As New clsCheck
Private mlngItemID              As Long  'ID
Private mlngItemTypeID          As Long  '����ID
Private mstrItemCode            As String   '����
Private mstrItemName            As String   '����
Private mstrItemMnemonicCode    As String   '����
Private mstrItemDescription     As String   '˵��
Private mstrItemAudit           As String   '�������
Private mintItemUsed            As Integer  '���ö���
Private mintItemFileID          As String   '���ö����ļ�
Private mstrItemLink            As Integer  '���û���
Private mintUsedListIndex       As Integer  '��¼���ö�������
Private mstrNumValue            As String   '��ֵ
Private mstrPalValue            As String   '����
Private mintSource              As Integer  '����Դ
Private mblnProgUsed            As Boolean  '�����Ƿ�ʹ��

Private Const conFieldFiles = "Select /*+ rule */ a.id as �ļ�ID,a.��� as �ļ�����,a.���� as �ļ�����,a.˵�� as �ļ�˵��" & vbCrLf & _
                         "from �����ļ��б� A, Table (Cast(f_Str2List([1])  As zlTools.t_StrList)) B " & vbCrLf & _
                         "where /*+ rule */a.id = b.COLUMN_VALUE And a.���� = [2]"
Private Const conEmrField = "Select /*+ Rule*/ Rawtohex(b.Id) As �ļ�id, b.Code As �ļ�����, b.Title As �ļ�����, b.Note As �ļ�˵��" & vbNewLine & _
                        "From (Select Hextoraw(Column_Value) As ID From Table(Zlcommunal.f_Str2list(:p0, ','))) A, Antetype_List B" & vbNewLine & _
                        "Where Hextoraw(a.Id) = b.Id And b.Kind = :p1" & vbNewLine & _
                        "Order By �ļ�����"

Public Property Get blnProgUsed() As Boolean
    blnProgUsed = mblnProgUsed
End Property

Public Property Let blnProgUsed(ByVal vNewValue As Boolean)
    mblnProgUsed = vNewValue
End Property

Public Property Get DataChange() As Boolean
    DataChange = mblnDataChange And mEditMode <> ���
End Property

Public Property Let DataChange(ByVal vNewValue As Boolean)
    mblnDataChange = vNewValue
End Property

Public Property Get EditMode() As �༭ģʽ
    EditMode = mEditMode
End Property

Public Property Let EditMode(ByVal vNewValue As �༭ģʽ)
    mEditMode = vNewValue
End Property

Public Property Get lngItemID() As Long
    lngItemID = mlngItemID
End Property

Public Property Let lngItemID(ByVal vNewValue As Long)
    mlngItemID = vNewValue
End Property

Public Property Get lngItemTypeID() As Long
    lngItemTypeID = mlngItemTypeID
End Property

Public Property Let lngItemTypeID(ByVal vNewValue As Long)
    mlngItemTypeID = vNewValue
End Property

Public Property Get strItemCode() As String
    strItemCode = mstrItemCode
End Property

Public Property Let strItemCode(ByVal vNewValue As String)
    mstrItemCode = vNewValue
End Property

Public Property Get strItemName() As String
    strItemName = mstrItemName
End Property

Public Property Let strItemName(ByVal vNewValue As String)
    mstrItemName = vNewValue
End Property

Public Property Get strItemMnemonicCode() As String
    strItemMnemonicCode = mstrItemMnemonicCode
End Property

Public Property Let strItemMnemonicCode(ByVal vNewValue As String)
    mstrItemMnemonicCode = vNewValue
End Property

Public Property Get strItemDescription() As String
    strItemDescription = mstrItemDescription
End Property

Public Property Let strItemDescription(ByVal vNewValue As String)
    mstrItemDescription = vNewValue
End Property

Public Property Get strItemAudit() As String
    strItemAudit = mstrItemAudit
End Property

Public Property Let strItemAudit(ByVal vNewValue As String)
    mstrItemAudit = vNewValue
End Property

Public Property Get intItemUsed() As Integer
    intItemUsed = mintItemUsed
End Property

Public Property Let intItemUsed(ByVal vNewValue As Integer)
    mintItemUsed = vNewValue
End Property

Public Property Get intItemFileID() As String
    intItemFileID = mintItemFileID
End Property

Public Property Let intItemFileID(ByVal vNewValue As String)
    mintItemFileID = vNewValue
End Property

Public Property Get strItemLink() As String
    strItemLink = mstrItemLink
End Property

Public Property Let strItemLink(ByVal vNewValue As String)
    mstrItemLink = vNewValue
End Property

Public Property Get strNumValue() As String
    strNumValue = mstrNumValue
End Property

Public Property Let strNumValue(ByVal vNewValue As String)
    mstrNumValue = vNewValue
End Property

Public Property Get strPalValue() As String
    strPalValue = mstrPalValue
End Property

Public Property Let strPalValue(ByVal vNewValue As String)
    mstrPalValue = vNewValue
End Property
Public Property Get intSource() As Integer
    intSource = mintSource
End Property

Public Property Let intSource(ByVal vNewValue As Integer)
    mintSource = vNewValue
End Property
Private Sub cboUsed_Click()
    Call cboUsed_LostFocus
End Sub

'==============================================================================
'=���ܣ� ���
'==============================================================================
Private Sub cmdCheck_Click()
    On Error GoTo ErrH
    If txtAudit_NotCheck.Text <> "" Then Call CheckAuditSql_IN(Trim(txtAudit_NotCheck.Text), True, IIf(optSource(0).Value, 0, 1))
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ѡ�����
'==============================================================================
Private Sub cmdSelect_Click()
    Dim intTypeID   As Integer
    Dim intLenght   As Integer
    Dim rsTemp      As ADODB.Recordset
    Dim rsTempType      As ADODB.Recordset
    On Error GoTo ErrH
    
    With frmAuditItemTypeSelect
        .lngLeft = frmAuditItem.Left + frmAuditItem.tvwAuditType.Width + txtTypeID.Left + 140
        .lngTop = frmAuditItem.Top + frmAuditItem.vsfAuditItem.Top + frmAuditItem.vsfAuditItem.Height + txtTypeID.Top - 3285
        .intTypeID = IIf(Val(txtTypeID.Tag) = 0, mlngItemTypeID, Val(txtTypeID.Tag))
        .Show vbModal
        If .blnCancel Then Set frmAuditItemTypeSelect = Nothing: Exit Sub
        intTypeID = .intTypeID
    End With
    gstrSQL = "select /*+ rule */id,�ϼ�ID,����,���� from ���������� a Where a.id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(intTypeID))
    If Not zlCheck.Connection_ChkRsState(rsTemp) Then
        txtTypeID.Tag = CStr(intTypeID)
        txtTypeID.Text = "[" + rsTemp!���� + "]" & rsTemp!����
    Else
        txtTypeID.Tag = "-1"
        txtTypeID.Text = "[ȫ��]����"
    End If
    
    '��ȡ����
    gstrSQL = "select /*+ rule */id,�ϼ�ID,����,���� from ���������� a Where a.id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txtTypeID.Tag)
    If Not zlCheck.Connection_ChkRsState(rsTemp) Then
        gstrSQL = "select nvl(Max(����),0) from �������Ŀ¼ where ����id=[1] and ���� like [2] || '%'"
        Set rsTempType = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txtTypeID.Tag, rsTemp!����)
        txtCode.Text = IncStr(rsTempType.Fields(0))
    Else
        txtCode.Text = "0001"
    End If
    If txtCode.Text = "1" Then
        txtCode.Text = rsTemp!���� & Format(txtCode.Text, "0000")
    End If
    txtCode.Text = InsertNewCode(txtCode.Text)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
 
'==============================================================================
'=���ܣ� ѡ���ļ�
'==============================================================================
Private Sub cmdSelectFile_Click()
    Call SelectFile(True)
End Sub

Private Sub SelectFile(Optional blnClick As Boolean = False)
 Dim rsTemp      As ADODB.Recordset, strType As String, strReturn As String
    
    On Error GoTo ErrH
    If cboUsed.Text = "" Or cboUsed.ListIndex = -1 Then
        zlCheck.Msg_OK "����ѡ�����ö���"
        Exit Sub
    End If
    strType = AuditFileTran(zlCheck.Cmb_ID(cboUsed), IIf(optSource(0).Value, 0, 1))
    With frmAuditItemEditFile
        .intItemFileID = txtFileID.Text
        .strType = strType
        .intSource = IIf(optSource(0).Value, 0, 1)
        .Show vbModal
        If .intItemFileID = txtFileID.Text Then
            Exit Sub
        End If
         txtFileID.Text = .intItemFileID
    End With
    Set frmAuditItemEditFile = Nothing
    'ˢ�±�������
    If optSource(0).Value Then
        gstrSQL = conFieldFiles
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txtFileID.Text, strType)
    Else
        gstrSQL = conEmrField
        strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, txtFileID.Text & "^" & DbType.T_String & "^p0|" & strType & "^" & DbType.T_String & "^p1", rsTemp)
        If strReturn <> "" Then
            zlCheck.Msg_OK strReturn
            Exit Sub
        End If
    End If
    Set vsfFiles.DataSource = rsTemp
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub optSource_Click(Index As Integer)
Dim rsUsed As New ADODB.Recordset
    '�����ѡ�ļ�
    gstrSQL = "Select /*+ rule */" & vbNewLine & _
            " a.Id As �ļ�id, a.��� As �ļ�����, a.���� As �ļ�����, a.˵�� As �ļ�˵��" & vbNewLine & _
            "From �����ļ��б� A" & vbNewLine & _
            "Where a.Id = [1]"
    Set rsUsed = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(0))
    Set vsfFiles.DataSource = rsUsed
    
    If Index = 0 Then
        gstrSQL = "select 1 as ID ,'סԺҽ��' as Name from dual union all" & vbCrLf & _
                    "select 2 as ID ,'סԺ����' as Name from dual union all" & vbCrLf & _
                    "select 3 as ID ,'������' as Name from dual union all" & vbCrLf & _
                    "select 4 as ID ,'�����¼' as Name from dual union all" & vbCrLf & _
                    "select 5 as ID ,'��ҳ��¼' as Name from dual union all" & vbCrLf & _
                    "select 6 as ID ,'ҽ������' as Name from dual union all" & vbCrLf & _
                    "select 7 as ID ,'����֤��' as Name from dual union all" & vbCrLf & _
                    "select 8 as ID ,'֪���ļ�' as Name from dual union all" & vbCrLf & _
                    "select 9 as ID ,'�ٴ�·��' as Name from dual"
        Set rsUsed = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        zlCheck.Cmb_List cboUsed, rsUsed
        lab����ID.Caption = "[����ID]��[��ҳID]Ϊϵͳ�������ֱ����ϵͳ�еĲ���ID����ҳID��"
        txtAudit_NotCheck.Tag = 0 '���ڷŴ�༭����
    Else
        gstrSQL = "select 2 as ID ,'סԺ����' as Name from dual union all" & vbCrLf & _
                    "select 3 as ID ,'������' as Name from dual union all" & vbCrLf & _
                    "select 7 as ID ,'����֤��' as Name from dual union all" & vbCrLf & _
                    "select 8 as ID ,'֪���ļ�' as Name from dual"
        Set rsUsed = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        zlCheck.Cmb_List cboUsed, rsUsed
        lab����ID.Caption = "[MID]��[ALIDIN]Ϊϵͳ�������ֱ����EMR�еĲ���ID����ԺID"
        lab����ID.ToolTipText = "ʹ��EMR�е�ID,��[MID]��[ALIDIN]�ⶼ��Ҫʹ��HextoRawת�����Ա��õ�������"
        txtAudit_NotCheck.Tag = 1 '���ڷŴ�༭����
    End If
End Sub
Private Sub txtFileID_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrH
    If KeyAscii = 13 Then
        Call SelectFile
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ҳ���ʼ��
'==============================================================================
Private Sub Form_Load()
    Dim rsUsed      As New ADODB.Recordset
    On Error GoTo ErrH
    mblnDataChange = False
    gstrSQL = "select 1 as ID ,'סԺҽ��' as Name from dual union all" & vbCrLf & _
                "select 2 as ID ,'סԺ����' as Name from dual union all" & vbCrLf & _
                "select 3 as ID ,'������' as Name from dual union all" & vbCrLf & _
                "select 4 as ID ,'�����¼' as Name from dual union all" & vbCrLf & _
                "select 5 as ID ,'��ҳ��¼' as Name from dual union all" & vbCrLf & _
                "select 6 as ID ,'ҽ������' as Name from dual union all" & vbCrLf & _
                "select 7 as ID ,'����֤��' as Name from dual union all" & vbCrLf & _
                "select 8 as ID ,'֪���ļ�' as Name from dual union all" & vbCrLf & _
                "select 9 as ID ,'�ٴ�·��' as Name from dual"
    Set rsUsed = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    zlCheck.Cmb_List cboUsed, rsUsed
    cboUsed.ListIndex = 0
    cboLink.Clear
    cboLink.AddItem "0" & strSplitCmb & "ȫ��"
    cboLink.AddItem "1" & strSplitCmb & "���"
    cboLink.AddItem "2" & strSplitCmb & "���"
    
    CboPalValue.Clear
    CboPalValue.AddItem "�۷���": CboPalValue.ItemData(CboPalValue.NewIndex) = 0
    CboPalValue.AddItem "�����": CboPalValue.ItemData(CboPalValue.NewIndex) = 1
    CboPalValue.ListIndex = 0
    
    '�ֶο��
    Set rsUsed = zlCheck.GetRsFieldWidth("�������Ŀ¼")
    rsUsed.Filter = "����='" & txtName.DataField & "'"
    If Not zlCheck.Connection_ChkRsState(rsUsed) Then txtName.MaxLength = "" & rsUsed.Fields("����")
    rsUsed.Filter = "����='" & txtCode.DataField & "'"
    If Not zlCheck.Connection_ChkRsState(rsUsed) Then txtCode.MaxLength = "" & rsUsed.Fields("����")
    rsUsed.Filter = "����='" & txtMnemonicCode.DataField & "'"
    If Not zlCheck.Connection_ChkRsState(rsUsed) Then txtMnemonicCode.MaxLength = "" & rsUsed.Fields("����")
    rsUsed.Filter = "����='" & txtDescription.DataField & "'"
    If Not zlCheck.Connection_ChkRsState(rsUsed) Then txtDescription.MaxLength = "" & rsUsed.Fields("����")
    rsUsed.Filter = "����='" & txtAudit_NotCheck.DataField & "'"
    If Not zlCheck.Connection_ChkRsState(rsUsed) Then txtAudit_NotCheck.MaxLength = "" & rsUsed.Fields("����")
    If gobjEmr Is Nothing Then
        fraSource.Visible = False
        optSource(0).Value = True
        optSource(1).Value = False
    Else
        fraSource.Visible = True
        optSource(0).Value = False
        optSource(1).Value = True
    End If
    
    zlCheck.Sys_System Me
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ����λ�õ���
'==============================================================================
Private Sub Form_Resize()
    On Error Resume Next
    fraAuditItem.Move 0, -75, Me.ScaleWidth, Me.ScaleHeight + 75
    txtTypeID.Move txtTypeID.Left, txtTypeID.Top, fraAuditItem.Width - txtTypeID.Left - 500 - IIf(gobjEmr Is Nothing, 0, fraSource.Width), txtTypeID.Height
    cmdSelect.Move txtTypeID.Left + txtTypeID.Width + 60, cmdSelect.Top, cmdSelect.Width, cmdSelect.Height
    fraSource.Move cmdSelect.Left + cmdSelect.Width
    
    txtMnemonicCode.Move txtTypeID.Left + txtTypeID.Width + IIf(gobjEmr Is Nothing, 0, fraSource.Width + 100) - txtMnemonicCode.Width - txtNumValue.Width - lbl(4).Width - lbl(8).Width - CboPalValue.Width - 200 '- txtNumValue.Width - 1000
    lbl(3).Move txtMnemonicCode.Left - lbl(3).Width - 50
    txtName.Move txtName.Left, txtName.Top, lbl(3).Left - txtTypeID.Left - 75, txtTypeID.Height
    
    
    lbl(8).Move txtMnemonicCode.Left + txtMnemonicCode.Width + 50
    CboPalValue.Move lbl(8).Left + lbl(8).Width + 50
    
    lbl(4).Move CboPalValue.Left + CboPalValue.Width + 50


    txtDescription.Move txtDescription.Left, txtDescription.Top, txtTypeID.Width + IIf(gobjEmr Is Nothing, 0, fraSource.Width + 100) + 50, txtDescription.Height
    txtAudit_NotCheck.Move txtAudit_NotCheck.Left, txtAudit_NotCheck.Top, txtTypeID.Width + IIf(gobjEmr Is Nothing, 0, fraSource.Width + 100) + 50, fraAuditItem.Height - txtAudit_NotCheck.Top - 300
    labNote.Top = txtAudit_NotCheck.Top + txtAudit_NotCheck.Height + 50
    lab����ID.Top = labNote.Top
    cmdCheck.Move txtAudit_NotCheck.Left + txtAudit_NotCheck.Width + 50
    
    vsfFiles.Move vsfFiles.Left, txtCode.Top, txtTypeID.Width + IIf(gobjEmr Is Nothing, 0, fraSource.Width + 100) - vsfFiles.Left + txtAudit_NotCheck.Left + 30
    cmdSelectFile.Move vsfFiles.Left + vsfFiles.Width + 60, vsfFiles.Top, cmdSelectFile.Width, cmdSelectFile.Height
    
     txtNumValue.Move lbl(4).Left + lbl(4).Width + 50, txtMnemonicCode.Top, 800 ', txtTypeID.Left + txtTypeID.Width - (lbl(4).Left + lbl(4).Width + 50), txtTypeID.Height
End Sub
 
'==============================================================================
'=���ܣ� �����ı���
'==============================================================================
Public Sub WinLock()
    Dim ctrAll           As Control
    
    On Error GoTo ErrH
    For Each ctrAll In Me.Controls
        If TypeName(ctrAll) = "TextBox" Then
            ctrAll.Locked = True
            ctrAll.BackColor = &H80000000
        ElseIf TypeName(ctrAll) = "CommandButton" Then
            ctrAll.Enabled = False
        ElseIf TypeName(ctrAll) = "ComboBox" Then
            ctrAll.Locked = True
            ctrAll.BackColor = &H80000000
        End If
    Next
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��ʼ�˵�������
'==============================================================================
Public Sub unWinLock(Optional blnCopy As Boolean)
    Dim ctrAll           As Control
    
    On Error GoTo ErrH
    
    For Each ctrAll In Me.Controls
        If TypeName(ctrAll) = "TextBox" Then
            If Not blnCopy Then ctrAll.Text = ""
            ctrAll.Locked = False
            ctrAll.BackColor = vbWhite
        ElseIf TypeName(ctrAll) = "CommandButton" Then
            ctrAll.Enabled = True
        ElseIf TypeName(ctrAll) = "ComboBox" Then
            ctrAll.Locked = False
            ctrAll.BackColor = vbWhite
        End If
    Next
    txtTypeID.Locked = True
    txtTypeID.BackColor = &H80000000
    
    txtFileID.Locked = True
    txtFileID.BackColor = &H80000000
    
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub AuditItemItemInsert()
    On Error GoTo ErrH
    mlngItemID = zlDatabase.GetNextId("�������Ŀ¼")
    gstrSQL = "Zl_�������Ŀ¼_Insert (" + CStr(mlngItemID) + "," + IIf(mlngItemTypeID = 0, "NULL", CStr(mlngItemTypeID)) + "," + "'" + mstrItemCode + "'" + "," + "'" + mstrItemName + "','" & mstrItemMnemonicCode & "','" & mstrItemDescription & "','" & mstrItemAudit & "'," & mintItemUsed & ",'" & mintItemFileID & "','" & mstrItemLink & "'," & Val(mstrPalValue) & " ," & Val(mstrNumValue) & "," & mintSource & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub AuditItemItemUpdate()
    On Error GoTo ErrH
    
    gstrSQL = "Zl_�������Ŀ¼_Update (" + CStr(mlngItemID) + "," + IIf(mlngItemTypeID = 0, "NULL", CStr(mlngItemTypeID)) + "," + "'" + mstrItemCode + "'" + "," + "'" + mstrItemName + "','" & mstrItemMnemonicCode & "','" & mstrItemDescription & "','" & mstrItemAudit & "'," & mintItemUsed & ",'" & mintItemFileID & "','" & mstrItemLink & "'," & Val(mstrPalValue) & ", " & Val(mstrNumValue) & "," & mintSource & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub AuditItemItemDelete()
    On Error GoTo ErrH
    Dim rsTmp As New ADODB.Recordset
    
    If blnProgUsed Then '����÷�����ʹ��,��Ҫ�����Ŀ�Ƿ��Ѿ�����,�����ʹ�ò��ܽ���ɾ��.
        Set rsTmp = gclsPackage.GetItemUse(mlngItemID)
        If rsTmp.RecordCount = 1 Then
            If rsTmp!���� > 0 Then
                Call MsgBox("�÷���Ŀ����ʹ�ù�,��ʱ���ܱ�ɾ��!", vbInformation, gstrSysName)
                Exit Sub
            End If
        End If
    End If
    
    gstrSQL = "Zl_�������Ŀ¼_Delete (" & mlngItemID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function InsertNewCode(strInCode) As String
    Dim rsTemp          As ADODB.Recordset
    On Error GoTo ErrH
    
    gstrSQL = "select 1 from �������Ŀ¼ where ���� = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strInCode)
    If zlCheck.Connection_ChkRsState(rsTemp) Then InsertNewCode = strInCode: Exit Function
    strInCode = IncStr(strInCode)
    InsertNewCode = InsertNewCode(strInCode)
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=���ܣ� ���������Ŀ���� ItemCopy
'==============================================================================
Public Sub ItemCopy()
    Dim strEditMode     As String
    Dim rsTemp          As ADODB.Recordset
    Dim rsTempType      As ADODB.Recordset
    On Error GoTo ErrH
    mEditMode = ��������
    unWinLock True
    '��ȡ����
    gstrSQL = "select /*+ rule */id,�ϼ�ID,����,���� from ���������� a Where a.id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txtTypeID.Tag)
    If Not zlCheck.Connection_ChkRsState(rsTemp) Then
        gstrSQL = "select nvl(Max(����),0) from �������Ŀ¼ where ����id=[1] and ���� like [2] || '%'"
        Set rsTempType = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txtTypeID.Tag, rsTemp!����)
        txtCode.Text = IncStr(rsTempType.Fields(0))
    Else
        txtCode.Text = "0001"
    End If
    If txtCode.Text = "1" Then
        txtCode.Text = rsTemp!���� & Format(txtCode.Text, "0000")
    End If
    txtCode.Text = InsertNewCode(txtCode.Text)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �����Ŀ���� ItemInsert
'==============================================================================
Public Sub ItemInsert()
    Dim strEditMode     As String
    Dim strID           As String
    Dim rsTemp          As ADODB.Recordset
    On Error GoTo ErrH
    strEditMode = ""
    
    mEditMode = ����
    unWinLock
    
    gstrSQL = "select /*+ rule */id,�ϼ�ID,����,���� from ���������� a Where a.id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngItemTypeID)
    If Not zlCheck.Connection_ChkRsState(rsTemp) Then
        txtTypeID.Tag = strID
        txtTypeID.Text = "[" + rsTemp!���� + "]" & rsTemp!����
    Else
        txtTypeID.Tag = "-1"
        txtTypeID.Text = "[ȫ��]����"
    End If
    txtName.SetFocus
    '��ȡ����
    
    CboPalValue.ListIndex = 0
    
    gstrSQL = "select nvl(Max(����),0) from �������Ŀ¼ where ����id=[1] and ���� like [2] || '%'"
    txtCode.Text = IncStr(zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngItemTypeID, rsTemp!����).Fields(0))
    If txtCode.Text = "1" Then
        txtCode.Text = rsTemp!���� & Format(txtCode.Text, "0000")
    End If
    txtCode.Text = InsertNewCode(txtCode.Text)
    If gobjEmr Is Nothing Then
        optSource(0).Value = True
        optSource(1).Value = False
    Else
        optSource(0).Value = False
        optSource(1).Value = True
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �޸���Ŀ���� ItemUpdate
'==============================================================================
Public Sub ItemUpdate()
    Dim strEditMode     As String
    On Error GoTo ErrH

    mEditMode = �޸�
    unWinLock True
    txtName.SetFocus
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ȡ���������� ItemDelete
'==============================================================================
Public Sub ItemDelete()
    Dim varPos  As Variant
    Dim strEditMode As String
    On Error GoTo ErrH
    
    If zlCheck.Msg_OKC("ȷ��ɾ�����������Ŀ��" & vbCrLf & "���룺��" & mstrItemCode & "��" & vbCrLf & "���ƣ���" & mstrItemName & "��") Then Exit Sub
    AuditItemItemDelete
    mEditMode = ���
    mblnDataChange = False
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ����������� ItemSave
'==============================================================================
Public Function ItemSave() As Boolean
    Dim strMsg      As String
    Dim varPos      As Variant
    Dim ctrSetF     As Control
    Dim lngRow      As Long
    Dim intCol      As Integer
    Dim bytMatch    As Byte
    Dim strSQL      As String
    Dim rsTemp      As ADODB.Recordset
    On Error GoTo ErrH
    ItemSave = True
    
    If txtTypeID.Tag <> "" Then mlngItemTypeID = Val(txtTypeID.Tag)
    mstrItemCode = txtCode.Text
    mstrItemName = txtName.Text
    mstrItemMnemonicCode = txtMnemonicCode.Text
    mstrItemDescription = txtDescription.Text
    mstrItemAudit = txtAudit_NotCheck.Text
    mintItemUsed = Val(zlCheck.Cmb_ID(cboUsed))
    mstrItemLink = Val(zlCheck.Cmb_ID(cboLink))
    mintItemFileID = txtFileID.Text
    mstrNumValue = txtNumValue.Text
    mstrPalValue = CboPalValue.ItemData(CboPalValue.ListIndex)
    mintSource = IIf(optSource(0).Value, 0, 1)
   
    If Not CheckAuditSql_IN(txtAudit_NotCheck.Text, False, IIf(optSource(0).Value, 0, 1)) Then Exit Function
    strMsg = ""
    strMsg = zlCheck.Chk_CheckTxtNull("����", txtCode, ctrSetF, strMsg)
    '�������ظ�
    strSQL = "select count(*) from �������Ŀ¼ where ���� = [1] and ID != [2]"
    If zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrItemCode, mlngItemID).Fields(0) <> "0" Then
        If ctrSetF Is Nothing Then Set ctrSetF = txtCode
        strMsg = strMsg & "���롾" & txtCode.Text & "���Ѵ��ڣ�������¼����޸ģ�" & vbCrLf
    End If
    strMsg = zlCheck.Chk_CheckTxtNull("����", txtName, ctrSetF, strMsg)
    '�������ظ�
    If mEditMode = ���� Or mEditMode = �������� Then
        strSQL = "select count(*) from �������Ŀ¼ where ���� = [1] and ����ID = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrItemName, mlngItemTypeID, mlngItemID)
    ElseIf mEditMode = �޸� Then
        strSQL = "select count(*) from �������Ŀ¼ where ���� = [1] and ����ID = [2] and ID != [3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrItemName, mlngItemTypeID, mlngItemID)
    End If
    If rsTemp.Fields(0) <> "0" Then
        If ctrSetF Is Nothing Then Set ctrSetF = txtName
        strMsg = strMsg & "���ơ�" & txtName.Text & "���ڷ��ࡾ" & txtTypeID.Text & "���Ѵ��ڣ�������¼����޸ģ�" & vbCrLf
    End If
    
    strMsg = zlCheck.Chk_CheckTxtNull("���ö���", cboUsed, ctrSetF, strMsg)
    DoEvents
    If zlCheck.Chk_CheckMsg(strMsg, ctrSetF) Then Exit Function
    If zlCheck.Cmb_ID(cboUsed) = "" Then
        zlCheck.Msg_OK "��ѡ�����ö���"
        Exit Function
    End If

    Select Case mEditMode
        Case �༭ģʽ.����, �༭ģʽ.��������
            AuditItemItemInsert
        Case �༭ģʽ.�޸�
            AuditItemItemUpdate
    End Select
    ItemSave = False
    mEditMode = ���
    mblnDataChange = False
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=���ܣ� ȡ���������� ItemCancel
'==============================================================================
Public Function ItemCancel() As Boolean
   
    On Error GoTo ErrH
    ItemCancel = True
    If mblnDataChange Then
        If zlCheck.Msg_OKC("ȷ��ȡ��������������") Then Exit Function
    End If
    mEditMode = ���
    WinLock
    mblnDataChange = False
    ItemCancel = False
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cboUsed_Validate(Cancel As Boolean)
    If mEditMode <> ��� Then mblnDataChange = True
End Sub

Private Sub txtAudit_Change()
    If mEditMode <> ��� Then mblnDataChange = True
End Sub

Private Sub txtCode_Change()
    If mEditMode <> ��� Then mblnDataChange = True
End Sub

Private Sub txtDescription_Change()
    If mEditMode <> ��� Then mblnDataChange = True
End Sub

Private Sub txtMnemonicCode_Change()
    If mEditMode <> ��� Then mblnDataChange = True
End Sub

Private Sub txtTypeID_Change()
    If mEditMode <> ��� Then mblnDataChange = True
End Sub

Private Sub cboUsed_Change()
    If mEditMode <> ��� Then mblnDataChange = True
End Sub

Private Sub cboUsed_GotFocus()
    On Error GoTo ErrH
    mintUsedListIndex = cboUsed.ListIndex
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboUsed_LostFocus()
    On Error GoTo ErrH
    If mintUsedListIndex <> cboUsed.ListIndex Then
        txtFileID.Text = ""
        txtFileID.Tag = ""
        vsfFiles.Rows = 1
    End If
'    gstrSQL = "Select Count(*) From �����ļ��б� where ���� = [1] "
    cmdSelectFile.Enabled = Not (Val(cboUsed.Text) = 1 Or Val(cboUsed.Text) = 5 Or Val(cboUsed.Text) = 9 Or cboUsed.Text = "")
    txtFileID.Enabled = cmdSelectFile.Enabled
    cmdSelectFile.Enabled = cmdSelectFile.Enabled
    lbl(6).Enabled = cmdSelectFile.Enabled
    vsfFiles.Enabled = cmdSelectFile.Enabled
    If vsfFiles.Enabled = True Then
        vsfFiles.BackColorBkg = &H80000005
    Else
        vsfFiles.BackColorBkg = &H8000000F
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtName_Change()
    On Error GoTo ErrH
    txtMnemonicCode.Text = zlCommFun.SpellCode(txtName.Text)
    If mEditMode <> ��� Then mblnDataChange = True
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
