VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.12#0"; "zlIDKind.ocx"
Begin VB.Form frmBillUses 
   Caption         =   "Ʊ����ϸ"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13260
   Icon            =   "frmBillUses.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   13260
   StartUpPosition =   1  '����������
   Begin VB.Frame fraCMD 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   11640
      TabIndex        =   26
      Top             =   540
      Width           =   1455
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   120
         TabIndex        =   16
         Top             =   510
         Width           =   1200
      End
      Begin VB.CommandButton cmdDistant 
         Caption         =   "��λ�Ϻ�(&H)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   120
         TabIndex        =   18
         Top             =   1710
         Width           =   1200
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   150
         TabIndex        =   20
         Top             =   2490
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "��λƱ��(&F)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   120
         TabIndex        =   21
         Top             =   2850
         Width           =   1200
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   120
         TabIndex        =   15
         Top             =   30
         Width           =   1200
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "����(&S)"
         Height          =   350
         Left            =   120
         TabIndex        =   17
         Top             =   990
         Width           =   1200
      End
      Begin VB.CommandButton cmdAllDO 
         Caption         =   "ȫ���˶�(&A)"
         Height          =   350
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   3510
         Width           =   1200
      End
      Begin VB.CommandButton cmdAllDO 
         Caption         =   "ȫ��ȡ��(&R)"
         Height          =   350
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   3870
         Width           =   1200
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&N)"
         Height          =   180
         Left            =   150
         TabIndex        =   19
         Top             =   2220
         Width           =   630
      End
      Begin VB.Line linBlack 
         BorderColor     =   &H00000000&
         X1              =   120
         X2              =   1300
         Y1              =   1590
         Y2              =   1590
      End
   End
   Begin VB.PictureBox picFilter 
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   90
      ScaleHeight     =   870
      ScaleWidth      =   11475
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11475
      Begin VB.CommandButton cmd������ 
         Caption         =   "��"
         Height          =   255
         Left            =   5100
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   120
         Width           =   285
      End
      Begin VB.ComboBox cboʹ����� 
         Height          =   300
         Left            =   9270
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   510
         Width           =   1350
      End
      Begin zlIDKind.IDKindNew cboFindType 
         Height          =   300
         Left            =   2310
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   90
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "����"
         IDKind          =   -1
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   3150
         MaxLength       =   200
         TabIndex        =   4
         Top             =   97
         Width           =   2265
      End
      Begin VB.PictureBox picTimeRange 
         BorderStyle     =   0  'None
         Height          =   390
         Left            =   2235
         ScaleHeight     =   390
         ScaleWidth      =   6105
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   6105
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "ˢ��(&R)"
            Height          =   350
            Left            =   4740
            TabIndex        =   10
            Top             =   30
            Width           =   1200
         End
         Begin MSComCtl2.DTPicker dtpStartDate 
            Height          =   300
            Left            =   0
            TabIndex        =   7
            Top             =   60
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   153288707
            CurrentDate     =   41520
         End
         Begin MSComCtl2.DTPicker dtpEndDate 
            Height          =   300
            Left            =   2550
            TabIndex        =   9
            Top             =   60
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   153288707
            CurrentDate     =   41520
         End
         Begin VB.Label lblEndDate 
            AutoSize        =   -1  'True
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2265
            TabIndex        =   8
            Top             =   120
            Width           =   210
         End
      End
      Begin VB.ComboBox cboʹ������ 
         Height          =   300
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   525
         Width           =   1350
      End
      Begin VB.ComboBox cboƱ�� 
         Height          =   300
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   90
         Width           =   1350
      End
      Begin VB.Label lblʹ����� 
         AutoSize        =   -1  'True
         Caption         =   "ʹ�����"
         Height          =   180
         Left            =   8490
         TabIndex        =   12
         Top             =   570
         Width           =   720
      End
      Begin VB.Line lineSplit 
         BorderColor     =   &H8000000A&
         X1              =   0
         X2              =   10365
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Label lblƱ�� 
         AutoSize        =   -1  'True
         Caption         =   "Ʊ������"
         Height          =   180
         Left            =   90
         TabIndex        =   1
         Top             =   150
         Width           =   720
      End
      Begin VB.Label lblʹ������ 
         AutoSize        =   -1  'True
         Caption         =   "ʹ��ʱ��"
         Height          =   180
         Left            =   90
         TabIndex        =   5
         Top             =   585
         Width           =   720
      End
   End
   Begin VB.TextBox txtInput 
      Height          =   270
      Left            =   7520
      MaxLength       =   200
      TabIndex        =   25
      Top             =   1815
      Visible         =   0   'False
      Width           =   1450
   End
   Begin VB.ComboBox cboResult 
      Height          =   300
      Left            =   6300
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   6735
      Left            =   60
      TabIndex        =   14
      Top             =   930
      Width           =   11370
      _ExtentX        =   20055
      _ExtentY        =   11880
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      BackColorSel    =   12320767
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorUnpopulated=   -2147483644
      GridColor       =   8421504
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "  ����  |   ���  |    ʹ��ʱ��    |ʹ����|   ʹ�����   |     �˶�ʱ��     |�˶���|   �˶Խ��  |      ��ע     |ID"
      MouseIcon       =   "frmBillUses.frx":0442
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin VB.Label lbl��ʾ 
      AutoSize        =   -1  'True
      Caption         =   "�����˵�����ʹ����ϸ�嵥"
      Height          =   180
      Left            =   7800
      TabIndex        =   27
      Top             =   1050
      Width           =   2160
   End
End
Attribute VB_Name = "frmBillUses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mbytInFun As Byte    '0-�鿴Ʊ��ʹ����ϸ,1-�˶�Ʊ����ϸ
Private mblnViewCheck As Boolean '��mbytInFun=0ʱ,�Ƿ���ʾ�˶�����ֶ�
Private mlngƱ�� As gBillType
Private mblnIsBIll As Boolean '��ǰƱ���Ƿ�ΪƱ��
Private mlng����ID As Long
Private mdblGiveCount As Double   '������Ʊ��������
Private mstrǰ׺�ı� As String
Private mblnUnClick As Boolean
Private mblnFirst As Boolean

Private Enum Col
    C0���� = 0
    C_��� = 1
    C1ʹ��ʱ�� = 2
    C2ʹ���� = 3
    C3ʹ����� = 4
    C4�˶�ʱ�� = 5
    C5�˶��� = 6
    C6�˶Խ�� = 7
    C7��ע = 8
    C8ID = 9
End Enum

Private Sub SetUnChecked(ByVal lngRow As Long)
    With mshDetail
        .TextMatrix(lngRow, Col.C4�˶�ʱ��) = ""
        .TextMatrix(lngRow, Col.C5�˶���) = ""
        .TextMatrix(lngRow, Col.C6�˶Խ��) = ""
        .TextMatrix(lngRow, Col.C7��ע) = ""
        
        .RowData(lngRow) = 1  '���ڱ���ʱ�ж�
        If Not cmdSave.Enabled Then cmdSave.Enabled = True
    End With
End Sub

Private Sub SetChecked(ByVal lngRow As Long, ByVal lngCol As Long, ByVal strContent As String, Optional ByVal strDate As String)
    With mshDetail
        If lngCol = Col.C6�˶Խ�� Then
            .TextMatrix(lngRow, Col.C4�˶�ʱ��) = strDate
            .TextMatrix(lngRow, Col.C5�˶���) = UserInfo.����
            .TextMatrix(lngRow, lngCol) = strContent
        ElseIf lngCol = Col.C7��ע Then
            .TextMatrix(lngRow, lngCol) = strContent
        End If
        
        .RowData(lngRow) = 1 '���ڱ���ʱ�ж�
        If Not cmdSave.Enabled Then cmdSave.Enabled = True
    End With
End Sub

Private Sub cboFindType_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If mblnUnClick Then Exit Sub
    cmd������.Visible = (cboFindType.IDKind = 1)
    SaveRegInFor g˽��ģ��, Me.Name, "�������ι��˷�ʽ", cboFindType.IDKind
    
    zlControl.ControlSetFocus txtFind: zlControl.TxtSelAll txtFind
End Sub

Private Sub cboResult_LostFocus()
    If cboResult.Visible Then cboResult.Visible = False
End Sub

Private Sub cboƱ��_Click()
    Dim lngƱ�� As gBillType
    Dim strKeyValue As String
    
    If cboƱ��.ListIndex <> -1 Then lngƱ�� = cboƱ��.ItemData(cboƱ��.ListIndex)
    
    mblnUnClick = True
    If lngƱ�� = gBillType.���￨ Or lngƱ�� = gBillType.���ѿ� Then
        cboFindType.IDKindStr = "��|������|0;��|����|0"
    Else
        cboFindType.IDKindStr = "��|������|0;��|��Ʊ��|0"
    End If
    mblnUnClick = False
    
    GetRegInFor g˽��ģ��, Me.Name, "�������ι��˷�ʽ", strKeyValue
    If Val(strKeyValue) < 1 Or Val(strKeyValue) > 2 Then
        cboFindType.IDKind = 1
    Else
        cboFindType.IDKind = Val(strKeyValue)
    End If
End Sub

Private Sub cboƱ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub SetRowHiddenAndTipText()
    '��������ʾ״̬����ʾ��Ϣ
    Dim i As Long, lngCount As Long, dblMoney As Double
    Dim dtDate As Date
    
    On Error GoTo ErrHandler
    With mshDetail
        .Redraw = False
        For i = 1 To .Rows - 1
            If cboʹ������.Text = "����" Then
                .RowHeight(i) = .RowHeight(0)
            ElseIf cboʹ������.Text = "ʱ�䷶Χ" Then
                If IsDate(.TextMatrix(i, Col.C1ʹ��ʱ��)) Then
                    dtDate = CDate(.TextMatrix(i, Col.C1ʹ��ʱ��))
                    If dtDate >= dtpStartDate.Value And dtDate <= dtpEndDate.Value Then
                        .RowHeight(i) = .RowHeight(0)
                    Else
                        .RowHeight(i) = 0
                    End If
                Else
                    .RowHeight(i) = 0
                End If
            ElseIf InStr(1, .TextMatrix(i, Col.C1ʹ��ʱ��), cboʹ������.Text) > 0 Then
                .RowHeight(i) = .RowHeight(0)
            Else
                .RowHeight(i) = 0
            End If
            
            If .RowHeight(i) <> 0 Then
                If Trim(cboʹ�����.Text) <> "" And cboʹ�����.Text <> .TextMatrix(i, Col.C3ʹ�����) Then
                    .RowHeight(i) = 0
                End If
            End If
            
            If .RowHeight(i) <> 0 Then
                lngCount = lngCount + 1
                dblMoney = dblMoney + Val(.TextMatrix(i, Col.C_���))
            End If
        Next
        .Redraw = True
    End With
    
    lbl��ʾ.Caption = lbl��ʾ.Tag
    If cboʹ������.Text <> "����" Or Trim(cboʹ�����.Text) <> "" Then
        lbl��ʾ.Caption = lbl��ʾ.Caption & IIf(lbl��ʾ.Caption = "", "", ".") & _
            "���е�ǰѡ�� " & lngCount & " ��" & IIf(mblnIsBIll, "Ʊ��", "��Ƭ")
    End If
    
    If mblnIsBIll Then
        lbl��ʾ.Caption = lbl��ʾ.Caption & IIf(lbl��ʾ.Caption = "", "", ",") & _
            "�ܽ�" & FormatEx(dblMoney, 2, , , 6)
    End If
    Exit Sub
ErrHandler:
    mshDetail.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cboʹ�����_Click()
    If mblnUnClick = True Then Exit Sub
    Call SetRowHiddenAndTipText
End Sub

Private Sub cboʹ�����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboʹ������_Click()
    On Error GoTo errHandle
    If mblnUnClick = True Then Exit Sub
    '����:29885
    picTimeRange.Visible = False
    If cboʹ������.Text = "ʱ�䷶Χ" Then
        picTimeRange.Visible = True
        If dtpStartDate.Visible And dtpStartDate.Enabled Then dtpStartDate.SetFocus
        Call Form_Resize
        Exit Sub
    End If
    Call SetRowHiddenAndTipText
    Call Form_Resize
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cboʹ������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdAllDO_Click(Index As Integer)
    Dim i As Long, strDate As String
    Dim blnSel As Boolean '�Ƿ���ڶ���ѡ��
    Dim lngRows As Long
    Dim lngStart As Long
    
    With mshDetail
        If .TextMatrix(1, Col.C0����) = "" Then Exit Sub
        blnSel = .Row <> .RowSel And .RowSel > .Row
        
        If Index = 0 Then
            strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        End If
        lngStart = IIf(blnSel, .Row, 1)
        lngRows = IIf(blnSel, .RowSel, .Rows - 1)
        .Redraw = False
        For i = lngStart To lngRows
            
            If .RowHeight(i) <> 0 Then
                If Index = 0 Then
                   '��ʹ�Ѻ˶Ե�Ҳ���º˶�,��д�µĺ˶��˺ͺ˶�ʱ��,���ע,��ǰ���˵�Ҳ�������
                   Call SetChecked(i, Col.C6�˶Խ��, .TextMatrix(i, Col.C3ʹ�����), strDate)
                Else
                    'û�к˶Թ���,����ȡ���˶�
                    If Trim(.TextMatrix(i, Col.C6�˶Խ��)) <> "" Then Call SetUnChecked(i)
                End If
            End If
        Next
        .Redraw = True
    End With
End Sub

Private Sub cboResult_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = vbKeyReturn Then
        With mshDetail
            If cboResult.ListIndex <= 0 Then
                Call SetUnChecked(.Row)
            Else
                Call SetChecked(.Row, Col.C6�˶Խ��, Trim(cboResult.Text), Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"))
            End If
            .SetFocus    '����lostfocus
            .Col = .Col + 1
        End With
    ElseIf KeyAscii >= 32 Then
        If Chr(KeyAscii) > 5 Or Chr(KeyAscii) < 0 Then Exit Sub
        lngIdx = zlControl.CboMatchIndex(cboResult.hWnd, KeyAscii, 0.008)
        If lngIdx = -1 And cboResult.ListCount > 0 And cboResult.ListIndex = -1 Then lngIdx = 0
        cboResult.ListIndex = lngIdx
    End If
End Sub

Private Function SaveData() As Boolean
    Dim i As Long, arrSQL As Variant, blnTrans As Boolean, bytAllChecked As Byte, bytAllCheckOK As Byte
    Dim strDate As String, lngGiveCount As Long
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    With mshDetail
        If .TextMatrix(1, Col.C0����) = "" Then Exit Function
        arrSQL = Array()
        For i = 1 To .Rows - 1
            If .RowData(i) = 1 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                strDate = Trim(.TextMatrix(i, Col.C4�˶�ʱ��))
                If strDate = "" Then
                    strDate = "Null"
                Else
                    strDate = "To_Date('" & .TextMatrix(i, Col.C4�˶�ʱ��) & "','YYYY-MM-DD HH24:MI:SS')"
                End If
                If mlngƱ�� = gBillType.���ѿ� Then
                    'Zl_���ѿ�ʹ�ü�¼_Check
                    strSQL = "Zl_���ѿ�ʹ�ü�¼_Check("
                    '  Id_In       In ���ѿ�ʹ�ü�¼.Id%Type,
                    strSQL = strSQL & "" & .TextMatrix(i, Col.C8ID) & ","
                    '  �˶Խ��_In In ���ѿ�ʹ�ü�¼.�˶Խ��%Type,
                    strSQL = strSQL & "" & ZVal(Val(.TextMatrix(i, Col.C6�˶Խ��))) & ","
                    '  �˶���_In   In ���ѿ�ʹ�ü�¼.�˶���%Type,
                    strSQL = strSQL & "'" & .TextMatrix(i, Col.C5�˶���) & "',"
                    '  ��ע_In     In ���ѿ�ʹ�ü�¼.��ע%Type,
                    strSQL = strSQL & "'" & .TextMatrix(i, Col.C7��ע) & "',"
                    '  �˶�ʱ��_In In ���ѿ�ʹ�ü�¼.�˶�ʱ��%Type
                    strSQL = strSQL & "" & strDate & ")"
                Else
                    'Zl_Ʊ��ʹ����ϸ_Check
                    strSQL = "Zl_Ʊ��ʹ����ϸ_Check("
                    '  Id_In       In Ʊ��ʹ����ϸ.Id%Type,
                    strSQL = strSQL & "" & .TextMatrix(i, Col.C8ID) & ","
                    '  �˶Խ��_In In Ʊ��ʹ����ϸ.�˶Խ��%Type,
                    strSQL = strSQL & "" & ZVal(Val(.TextMatrix(i, Col.C6�˶Խ��))) & ","
                    '  �˶���_In   In Ʊ��ʹ����ϸ.�˶���%Type,
                    strSQL = strSQL & "'" & .TextMatrix(i, Col.C5�˶���) & "',"
                    '  ��ע_In     In Ʊ��ʹ����ϸ.��ע%Type,
                    strSQL = strSQL & "'" & .TextMatrix(i, Col.C7��ע) & "',"
                    '  �˶�ʱ��_In In Ʊ��ʹ����ϸ.�˶�ʱ��%Type
                    strSQL = strSQL & "" & strDate & ")"
                End If
                arrSQL(UBound(arrSQL)) = strSQL
            End If
        Next
    End With
    
    On Error GoTo errH
    If UBound(arrSQL) >= 0 Then
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
        
        '����Ƿ���Ҫ��д����˶Լ�¼
        If mlngƱ�� = gBillType.���ѿ� Then
            strSQL = _
                "Select Nvl(Sum(Decode(�˶Խ��, Null, 1, 0)), 0) As δ�˶���, Count(Distinct ����) As ��ʹ����" & vbNewLine & _
                "From ���ѿ�ʹ�ü�¼" & vbNewLine & _
                "Where ����id = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
            If rsTmp!δ�˶��� = 0 And rsTmp!��ʹ���� = mdblGiveCount Then
                bytAllChecked = 1
                strSQL = "Select Count(ID) ������� From ���ѿ�ʹ�ü�¼ Where ����id = [1] And �˶Խ�� <> ԭ��"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
                If rsTmp!������� = 0 Then bytAllCheckOK = 1
            End If
                    
            If bytAllChecked = 1 Then
                strSQL = "Zl_���ѿ����ü�¼_Check(" & mlng����ID & "," & bytAllCheckOK & ",'" & UserInfo.���� & "',Null,1)"
            Else
                'ȡ������˶�
                strSQL = "Zl_���ѿ����ü�¼_Check(" & mlng����ID & ",Null,Null,Null,Null)"
            End If
        Else
            strSQL = _
                "Select Nvl(Sum(Decode(�˶Խ��, Null, 1, 0)), 0) As δ�˶���, Count(Distinct ����) As ��ʹ����" & vbNewLine & _
                "From Ʊ��ʹ����ϸ" & vbNewLine & _
                "Where ����id = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
            If rsTmp!δ�˶��� = 0 And rsTmp!��ʹ���� = mdblGiveCount Then
                bytAllChecked = 1
                strSQL = "Select Count(ID) ������� From Ʊ��ʹ����ϸ Where ����id = [1] And �˶Խ�� <> ԭ��"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
                If rsTmp!������� = 0 Then bytAllCheckOK = 1
            End If
                    
            If bytAllChecked = 1 Then
                strSQL = "zl_Ʊ�����ü�¼_check(" & mlng����ID & "," & bytAllCheckOK & ",'" & UserInfo.���� & "',Null,1)"
            Else
                'ȡ������˶�
                strSQL = "zl_Ʊ�����ü�¼_check(" & mlng����ID & ",Null,Null,Null,Null)"
            End If
        End If
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
       gcnOracle.CommitTrans: blnTrans = False
    End If
    SaveData = True
    
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
      Call SetRowHiddenAndTipText
End Sub

Private Sub cmdSave_Click()
    Dim i As Long
    
    If SaveData Then
        With mshDetail
            For i = 1 To .Rows - 1
                If .RowData(i) = 1 Then .RowData(i) = 0
            Next
        End With
        cmdSave.Enabled = False
    End If
End Sub

Private Sub cmd������_Click()
    Dim rsResult As ADODB.Recordset, strSQL As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim bytKind As Byte, varPara As Variant
    
    On Error GoTo ErrHandler
    bytKind = Val(zlStr.NeedCode(cboƱ��.Text))
    Call GetOperatorSql(bytKind, strSQL, varPara)
    '���궨λ
    vRect = zlControl.GetControlRect(txtFind.hWnd)
    Set rsResult = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "������ѡ��", False, "", "", _
        False, False, True, vRect.Left - 15, vRect.Top, txtFind.Height, blnCancel, False, False, varPara)
    If blnCancel Then zlControl.ControlSetFocus txtFind: zlControl.TxtSelAll txtFind: Exit Sub
    If rsResult Is Nothing Then zlControl.ControlSetFocus txtFind: zlControl.TxtSelAll txtFind: Exit Sub
    If rsResult.RecordCount = 0 Then zlControl.ControlSetFocus txtFind: zlControl.TxtSelAll txtFind: Exit Sub
    
    txtFind.Text = NVL(rsResult!����)
    Call txtFind_KeyPress(vbKeyReturn)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetOperatorSql(ByVal bytKind As gBillType, _
    ByRef strSQL As String, ByRef varPara As Variant)
    
    On Error GoTo ErrHandler
    If zlStr.IsHavePrivs(mstrPrivs, "���в���Ա") = False Then
        strSQL = _
            "Select A.ID, A.���, A.����" & vbNewLine & _
            "From ��Ա�� A" & vbNewLine & _
            "Where a.ID=[1]"
        varPara = UserInfo.ID
    Else
        strSQL = _
            "Select Distinct A.ID, A.���, A.����" & vbNewLine & _
            "From ��Ա�� A, ��Ա����˵�� B" & vbNewLine & _
            "Where A.ID = B.��Աid And B.��Ա���� = [1]" & vbNewLine & _
            "      And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            "      And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) Order By ����"
        If bytKind > 0 And bytKind <= 7 Then
            '�������Ժ�Ǽ�Ա������Ҫͬʱ���ö�Ӧ�ķ�����Ԥ����Ա�����������ʾ��������Ϣ����ͬ��Ҳ�����������
            varPara = Choose(bytKind, "�����շ�Ա", "Ԥ���տ�Ա", "סԺ����Ա", "����Һ�Ա", _
                "�����Ǽ���", "�����Ǽ���", "�����Ǽ���")
        End If
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub dtpEndDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub dtpStartDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub dtpEndDate_Change()
    If dtpEndDate.Value < dtpStartDate.Value Then dtpStartDate.Value = dtpEndDate.Value
End Sub
Private Sub dtpStartDate_Change()
    If dtpStartDate.Value > dtpEndDate.Value Then dtpEndDate.Value = dtpStartDate.Value
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    If mshDetail.Rows > 1 Then Call SetRow(1)
    If mlng����ID > 0 Then
        zlControl.ControlSetFocus cboʹ�����
    Else
        zlControl.ControlSetFocus cboʹ������
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub mshDetail_Click()
    If mbytInFun = 0 Or mshDetail.Row = 0 Then Exit Sub
    
    With mshDetail
        If .TextMatrix(1, Col.C0����) = "" Then Exit Sub
        Select Case .Col
            Case Col.C6�˶Խ��
                If .TextMatrix(.Row, .Col) <> "" Then
                    Call zlControl.CboLocate(cboResult, zlCommFun.GetNeedName(.TextMatrix(.Row, .Col)))
                Else
                    Call zlControl.CboLocate(cboResult, zlCommFun.GetNeedName(.TextMatrix(.Row, Col.C3ʹ�����)))
                End If
                Call SetCboResult
            Case Else
        End Select
    End With
End Sub

Private Sub mshDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    If mbytInFun = 0 Or mshDetail.Row = 0 Then Exit Sub
    
    If KeyCode = vbKeyDelete Then
        With mshDetail
            If .TextMatrix(1, Col.C0����) = "" Then Exit Sub
            Select Case .Col
                Case Col.C6�˶Խ��
                    Call SetUnChecked(.Row)
                Case Col.C7��ע
                    Call SetChecked(.Row, Col.C7��ע, "")
                Case Else
                
            End Select
        End With
    End If
End Sub

Private Sub mshDetail_KeyPress(KeyAscii As Integer)
    If mbytInFun = 0 Or mshDetail.Row = 0 Then Exit Sub
    
    With mshDetail
        If .TextMatrix(1, Col.C0����) = "" Then Exit Sub
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            If .Row = .Rows - 1 And (.Col = Col.C7��ע Or .Col = Col.C6�˶Խ�� And .TextMatrix(.Row, .Col) = "") Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If .Col = Col.C7��ע Then
                    .Row = .Row + 1
                    .Col = Col.C6�˶Խ��
                Else
                    .Col = .Col + 1
                End If
            End If
        Else
            Select Case .Col
                Case Col.C6�˶Խ��
                    Call SetCboResult
                    Call cboResult_KeyPress(KeyAscii)
                Case Col.C7��ע
                    If .TextMatrix(.Row, Col.C6�˶Խ��) <> "" Then
                        txtInput.Text = Chr(KeyAscii)
                        txtInput.SelStart = 2
                        Call SetTxtInput
                    End If
                Case Else
                
            End Select
        End If
    End With
End Sub

Private Sub picFilter_Resize()
    On Error Resume Next
    lineSplit.X1 = 0
    lineSplit.X2 = picFilter.ScaleWidth
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim strKey As String
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    strKey = Trim(txtFind.Text)
    If strKey = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    Call Select��������(txtFind, strKey, Val(zlStr.NeedCode(cboƱ��.Text)), cboFindType.IDKind)
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If zlCommFun.ActualLen(txtInput.Text) > txtInput.MaxLength Then
            MsgBox "���ֻ��������" & txtInput.MaxLength & "���ַ���", vbInformation, gstrSysName
            Exit Sub
        End If
        If InStr(1, txtInput.Text, "'") > 0 Then
            'MsgBox "ע��:��������ϵͳ��ֹ����������ַ�!", vbInformation, gstrSysName
            Beep
            Beep
            Exit Sub
        End If
        
        With mshDetail
            Call SetChecked(.Row, Col.C7��ע, Trim(txtInput.Text))
            txtInput.Visible = False
            .SetFocus  '����lostfocus
            If .Row = .Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                .Row = .Row + 1
                .Col = Col.C6�˶Խ��
            End If
        End With
    End If
End Sub

Private Sub cmdDistant_Click()
    Dim lngRow As Long, bln���� As Boolean
    Dim lngǰ׺ As Long
    
    MousePointer = vbHourglass
    lngǰ׺ = Len(mstrǰ׺�ı�) + 1
    With mshDetail
        lngRow = .Row + 1
        
        While True
            If lngRow > .Rows - 1 Then
                '���һ��
                If bln���� = False Then
                    If .Row = 1 Then
                        MsgBox "����δ���ֶϺ������", vbInformation, gstrSysName
                        MousePointer = vbDefault
                        Exit Sub
                    Else
                        If MsgBox("����δ���ֶϺŵ�������Ƿ��ͷ��ʼ��", vbQuestion Or vbYesNo, gstrSysName) = vbNo Then
                            MousePointer = vbDefault
                            Exit Sub
                        End If
                    End If
                    bln���� = True
                    lngRow = 1
                Else
                    MsgBox "����δ���ֶϺ������", vbInformation, gstrSysName
                    MousePointer = vbDefault
                    Exit Sub
                End If
            End If
            
            If lngRow > 1 Then
                If Val(Mid(.TextMatrix(lngRow - 1, 0), lngǰ׺)) < Val(Mid(.TextMatrix(lngRow, 0), lngǰ׺)) - 1 Then
                    '���ֶϺ�
                    If .RowHeight(lngRow) = 0 Then
                        If MsgBox("ע��:" & vbCrLf & "   �Ѿ����ҵ��˶Ϻţ������ڵ�ǰʱ�䷶Χ�ڣ��Ƿ���ж�λ��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                             If cboʹ������.Visible Then cboʹ������.ListIndex = 0:
                        Else
                            MousePointer = vbDefault
                            Exit Sub
                        End If
                    End If
                    Call SetRow(lngRow)
                    MousePointer = vbDefault
                    Exit Sub
                End If
            End If
            lngRow = lngRow + 1
        Wend
     End With
End Sub

Private Sub cmdFind_Click()
'����ָ������
    Dim strFind As String
    Dim lngRow As Long
    
    If txt����.Text = "" Then Exit Sub
    If Len(txt����.Text) > Len(mshDetail.TextMatrix(1, 0)) Then Exit Sub
    
    '�ѳ��Ȳ���
    strFind = UCase(Mid(mshDetail.TextMatrix(1, 0), 1, Len(mshDetail.TextMatrix(1, 0)) - Len(txt����.Text)) & txt����.Text)
    With mshDetail
        For lngRow = 1 To mshDetail.Rows - 1
            If mshDetail.TextMatrix(lngRow, 0) = strFind Then
                If .RowHeight(lngRow) = 0 Then
                    If MsgBox("ע��:" & vbCrLf & "   �������ҵ�" & IIf(mblnIsBIll, "����", "����") & "���ڵ�ǰʱ�䷶Χ�ڣ��Ƿ�Ҫ���ж�λ��", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                        If cboʹ������.Visible Then cboʹ������.ListIndex = 0:
                    Else
                        Exit Sub
                    End If
                End If
                Call SetRow(lngRow)
                Exit Sub
            End If
        Next
    End With
    MsgBox "δ�ҵ�" & IIf(mblnIsBIll, "����", "����") & "Ϊ " & strFind & " ��ʹ�ü�¼��", vbInformation, gstrSysName
End Sub

Private Sub cmdOK_Click()
    If mbytInFun = 1 Then
        Call SaveData
    End If
    Unload Me
End Sub

Private Sub SetHeader()
    Dim strHead As String, arrTmp As Variant, i As Long
    
    With mshDetail
        .Redraw = False
        If mbytInFun = 0 Then
            .SelectionMode = flexSelectionByRow
        Else
            .SelectionMode = flexSelectionFree
            .BackColorSel = &HE7CFBA
        End If
                
        If mbytInFun = 0 And Not mblnViewCheck Then
            strHead = "����,1,1000|���,7,900|ʹ��ʱ��,1,1800|ʹ����,4,800|ʹ�����,1,1000"
        Else
            strHead = "����,1,1000|���,7,900|ʹ��ʱ��,1,1800|ʹ����,4,800|ʹ�����,1,1000|�˶�ʱ��,1,1800|�˶���,4,800|�˶Խ��,1,1000|��ע,1,2000|ID,1,0"
        End If
        If mblnIsBIll = False Then strHead = Replace(strHead, "����", "����")
        arrTmp = Split(strHead, "|")
        
        .Cols = UBound(arrTmp) + 1
        For i = 0 To UBound(arrTmp)
            .TextMatrix(0, i) = Split(arrTmp(i), ",")(0)
            .ColAlignment(i) = Split(arrTmp(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(arrTmp(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        If mblnIsBIll = False Then .ColWidth(Col.C_���) = 0
        .Redraw = True
    End With
End Sub

Private Sub Form_Load()
    mblnFirst = True
    If mbytInFun = 0 And Not mblnViewCheck Then Me.Width = 8000
    
    If mlngƱ�� = gBillType.���￨ Then
        Me.Caption = IIf(mbytInFun = 0, "ҽ�ƿ���ϸ�嵥", "�˶�ҽ�ƿ���ϸ")
    ElseIf mlngƱ�� = gBillType.���ѿ� Then
        Me.Caption = IIf(mbytInFun = 0, "���ѿ���ϸ�嵥", "�˶����ѿ���ϸ")
    Else
        Me.Caption = IIf(mbytInFun = 0, "Ʊ����ϸ�嵥", "�˶�Ʊ����ϸ")
    End If
    Call InitContext
    If mlngƱ�� > 0 And mlngƱ�� - 1 < cboƱ��.ListCount Then cboƱ��.ListIndex = mlngƱ�� - 1
    
    Call RestoreFlexState(mshDetail, Me.Caption)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call SaveFlexState(mshDetail, Me.Caption)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If mbytInFun = 0 Then
        Set lbl��ʾ.Container = Me
        lbl��ʾ.Top = 100: lbl��ʾ.Left = 100
        mshDetail.Top = lbl��ʾ.Height + lbl��ʾ.Top + 20
    Else
        Set lbl��ʾ.Container = picFilter
        If picTimeRange.Visible Then
            lblʹ�����.Left = lblʹ������.Left
            cboʹ�����.Top = cboʹ������.Top + cboʹ������.Height + 50
        Else
            lblʹ�����.Left = cboʹ������.Left + cboʹ������.Width + 220
            cboʹ�����.Top = cboʹ������.Top
        End If
        lblʹ�����.Top = cboʹ�����.Top + (cboʹ�����.Height - lbl��ʾ.Height) / 2
        cboʹ�����.Left = lblʹ�����.Left + lblʹ�����.Width + 50
        lbl��ʾ.Left = cboʹ�����.Left + cboʹ�����.Width + 220
        lbl��ʾ.Top = cboʹ�����.Top + (cboʹ�����.Height - lbl��ʾ.Height) / 2
        picFilter.Height = cboʹ�����.Top + cboʹ�����.Height
        mshDetail.Top = picFilter.Height + picFilter.Top + 20
    End If
    
    mshDetail.Height = Me.ScaleHeight - mshDetail.Top - 120
    If Me.ScaleWidth > 3000 Then
        fraCMD.Left = Me.ScaleWidth - fraCMD.Width - 120
        mshDetail.Width = fraCMD.Left - mshDetail.Left - 120
    End If
    
    picFilter.Left = 0: picFilter.Width = Me.ScaleWidth
End Sub

Public Sub ShowMe(ByVal frmOwner As Form, ByVal strPrivs As String, _
    ByVal bytInFun As Byte, ByVal blnViewCheck As Boolean, ByVal blnNOMoved As Boolean, _
    ByVal lngƱ�� As gBillType, ByVal lng����ID As Long, ByVal strǰ׺ As String, _
    Optional strCondition As String, Optional lngԭ�� As Long, Optional lng���� As Long, _
    Optional strʹ���� As String, Optional str��ʾ As String)
    '����:bytInFun:0-�鿴Ʊ����ϸ,1-�˶�Ʊ����ϸ
    '   blnViewCheck:��bytInFun=0ʱ,�Ƿ���ʾ�˶�����ֶ�
    Dim strResult As String, arrTmp As Variant
    Dim i As Integer
    
    On Error GoTo errH
    mstrPrivs = strPrivs
    mbytInFun = bytInFun
    mblnViewCheck = blnViewCheck
    mlngƱ�� = lngƱ��
    mlng����ID = lng����ID
    mstrǰ׺�ı� = strǰ׺
    mblnIsBIll = CurrentIsBill(mlngƱ��)
    lbl����.Caption = IIf(mblnIsBIll, "����(&N)", "����(&N)")
    cmdFind.Caption = IIf(mblnIsBIll, "��λƱ��(&F)", "��λ����(&F)")
    
    If RefrashData(lng����ID, blnNOMoved, strCondition, lngԭ��, lng����, strʹ����, str��ʾ) = False Then Exit Sub
    
    cboResult.Visible = False
    txtInput.Visible = False
    If mbytInFun = 0 Then
        cmdOK.Caption = "�˳�(&X)"
        cmdOK.Cancel = True
        cmdCancel.Visible = False
        cmdSave.Visible = False
        cmdAllDO(0).Visible = False
        cmdAllDO(1).Visible = False
        picFilter.Visible = False
        lbl��ʾ.Left = picFilter.Left
    Else
        If mlngƱ�� = gBillType.���ѿ� Then
            strResult = " ,1-����ʹ��,2-�����ջ�,3-��������,4-�����ջ�,5-����"
            Call zlControl.CboSetWidth(cboResult.hWnd, 1500)
        Else
            strResult = " ,1-����ʹ��,2-�����ջ�,3-�ش򷢳�,4-�ش��ջ�,5-����,6-��Ʊ����"
            Call zlControl.CboSetWidth(cboResult.hWnd, 800)
        End If
        arrTmp = Split(strResult, ",")
        For i = 0 To UBound(arrTmp)
            cboResult.AddItem arrTmp(i)
            cboʹ�����.AddItem arrTmp(i)
        Next
        picFilter.Visible = True
    End If
    frmBillUses.Show vbModal, frmOwner
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 
Private Function RefrashData(ByVal lng����ID As Long, Optional ByVal blnNOMoved As Boolean, _
    Optional ByVal strCondition As String, Optional ByVal lngԭ�� As Long, _
    Optional ByVal lng���� As Long, Optional ByVal strʹ���� As String, Optional ByVal str��ʾ As String) As Boolean
    '��������
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long, j As Long
    Dim strTemp As String, strʹ������ As String
    Dim strMinDate As String, strMaxDate As String
    Dim varData As Variant, strNOs As String
    
    On Error GoTo errHandle
    If mlngƱ�� = gBillType.���ѿ� Then
        strSQL = _
            "Select ���� As ����, '' As Ʊ�ݽ��, To_Char(ʹ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ʹ��ʱ��, ʹ����," & vbNewLine & _
            "       Decode(ԭ��, 1, '1-����ʹ��', 2, '2-�����ջ�', 3, '3-��������', 4, '4-�����ջ�', '5-����') As ʹ�����," & vbNewLine & _
            "       To_Char(�˶�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �˶�ʱ��, �˶���," & vbNewLine & _
            "       Decode(�˶Խ��, 1, '1-����ʹ��', 2, '2-�����ջ�', 3, '3-�ش򷢳�', 4, '4-�ش��ջ�', 5,'5-����','') as �˶Խ��, ��ע, ID" & vbNewLine & _
            "From " & IIf(blnNOMoved, "H", "") & "���ѿ�ʹ�ü�¼" & vbNewLine & _
            "Where ����id = [1] " & strCondition & vbNewLine & _
            "Order By ����,ʹ��ʱ��"
        If mbytInFun = 0 And Not mblnViewCheck Then
            strSQL = _
                "Select ���� As ����, '' As Ʊ�ݽ��, To_Char(ʹ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ʹ��ʱ��, ʹ����," & vbNewLine & _
                "       Decode(ԭ��, 1, '1-����ʹ��', 2, '2-�����ջ�', 3, '3-��������', 4, '4-�����ջ�', '5-����') As ʹ�����" & vbNewLine & _
                "From " & IIf(blnNOMoved, "H", "") & "���ѿ�ʹ�ü�¼" & vbNewLine & _
                "Where ����id = [1] " & strCondition & vbNewLine & _
                "Order By ����,ʹ��ʱ��"
        End If
    Else
        strSQL = _
            "Select ����, Trim(To_Char(Ʊ�ݽ��, '99999999990.00')) As Ʊ�ݽ��, To_Char(ʹ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ʹ��ʱ��, ʹ����," & vbNewLine & _
            "       Decode(ԭ��, 1, '1-����ʹ��', 2, '2-�����ջ�', 3, '3-�ش򷢳�', 4, '4-�ش��ջ�', 6, '6-��Ʊ����', '5-����') As ʹ�����," & vbNewLine & _
            "       To_Char(�˶�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �˶�ʱ��, �˶���," & vbNewLine & _
            "       Decode(�˶Խ��, 1, '1-����ʹ��', 2, '2-�����ջ�', 3, '3-�ش򷢳�', 4, '4-�ش��ջ�', 5,'5-����','') as �˶Խ��, ��ע, ID" & vbNewLine & _
            "From " & IIf(blnNOMoved, "H", "") & "Ʊ��ʹ����ϸ" & vbNewLine & _
            "Where ����id = [1] " & strCondition & vbNewLine & _
            "Order By ����,ʹ��ʱ��"
        If mbytInFun = 0 And Not mblnViewCheck Then
            strSQL = _
                "Select ����, Trim(To_Char(Ʊ�ݽ��, '99999999990.00')) As Ʊ�ݽ��, To_Char(ʹ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ʹ��ʱ��, ʹ����," & vbNewLine & _
                "       Decode(ԭ��, 1, '1-����ʹ��', 2, '2-�����ջ�', 3, '3-�ش򷢳�', 4, '4-�ش��ջ�', 6, '6-��Ʊ����', '5-����') As ʹ�����" & vbNewLine & _
                "From " & IIf(blnNOMoved, "H", "") & "Ʊ��ʹ����ϸ" & vbNewLine & _
                "Where ����id = [1] " & strCondition & vbNewLine & _
                "Order By ����,ʹ��ʱ��"
        End If
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����", lng����ID, lngԭ��, lng����, strʹ����)
    If rsTmp.RecordCount = 0 Then
        mshDetail.Clear
        mshDetail.Rows = 2
    Else
        Set mshDetail.DataSource = rsTmp
    End If
    Call SetHeader
    
    lbl��ʾ.Tag = str��ʾ & IIf(str��ʾ = "", "", ",") & "���� " & rsTmp.RecordCount & " ��" & IIf(mblnIsBIll, "Ʊ��", "��Ƭ")
    strʹ������ = ""
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
            strTemp = "|" & Format(rsTmp!ʹ��ʱ��, "yyyy-MM-DD")
            If InStr(1, strʹ������ & "|", strTemp & "|") = 0 Then strʹ������ = strʹ������ & strTemp
            rsTmp.MoveNext
        Next
    End If
    If strʹ������ <> "" Then strʹ������ = Mid(strʹ������, 2)
    
    mblnUnClick = True
    If cboʹ�����.ListCount > 0 Then cboʹ�����.ListIndex = 0
    '��������С��������
    cboʹ������.Clear
    cboʹ������.AddItem "����": cboʹ������.ListIndex = cboʹ������.NewIndex
    cboʹ������.AddItem "ʱ�䷶Χ"
    mblnUnClick = False
    
    varData = Split(strʹ������, "|")
    For i = 0 To UBound(varData)
        If varData(i) <> "" Then
            For j = i + 1 To UBound(varData)
                If varData(j) < varData(i) Then
                    strTemp = varData(i)
                    varData(i) = varData(j)
                    varData(j) = strTemp
                End If
            Next
            If varData(i) < strMinDate Or strMinDate = "" Then strMinDate = varData(i)
            If varData(i) > strMaxDate Then strMaxDate = varData(i)
            cboʹ������.AddItem varData(i)
        End If
    Next
    
    dtpStartDate.MaxDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59")
    dtpEndDate.MaxDate = dtpStartDate.MaxDate
    If strMinDate <> "" And IsDate(strMinDate) Then
        dtpStartDate.MinDate = Format(CDate(strMinDate), "yyyy-mm-dd 00:00:00")
        dtpStartDate.Value = dtpStartDate.MinDate
        dtpEndDate.MinDate = dtpStartDate.MinDate
        If IsDate(strMaxDate) Then
            dtpEndDate.Value = Format(CDate(strMaxDate), "yyyy-mm-dd 23:59:59")
        Else
            dtpEndDate.Value = Format(dtpStartDate.MinDate, "yyyy-mm-dd 23:59:59")
        End If
    End If
    
    If mbytInFun <> 0 Then
        mdblGiveCount = 0
        If mlngƱ�� = gBillType.���ѿ� Then
            strSQL = _
                "Select To_Number(Replace(��ֹ����, ǰ׺�ı�)) - To_Number(Replace(��ʼ����, ǰ׺�ı�))+1 As ����," & vbNewLine & _
                "       ǰ׺�ı�" & vbNewLine & _
                "From ���ѿ����ü�¼" & vbNewLine & _
                "Where ID = [1]"
        Else
            strSQL = _
                "Select To_Number(Replace(��ֹ����, ǰ׺�ı�)) - To_Number(Replace(��ʼ����, ǰ׺�ı�))+1 As ����," & vbNewLine & _
                "       ǰ׺�ı�" & vbNewLine & _
                "From Ʊ�����ü�¼" & vbNewLine & _
                "Where ID = [1]"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
        If rsTmp.RecordCount > 0 Then
            mdblGiveCount = rsTmp!����
            mstrǰ׺�ı� = NVL(rsTmp!ǰ׺�ı�)
        End If
    End If
    picTimeRange.Visible = False: Call Form_Resize
    Call SetRowHiddenAndTipText
    
    RefrashData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mshDetail_DblClick()
    Dim strReportNO As String, strInvoiceNO As String
    
    If mlngƱ�� = gBillType.���ѿ� Then Exit Sub
    With mshDetail
        If .TextMatrix(1, Col.C0����) = "" Then Exit Sub
        Select Case .Col
            Case Col.C7��ע
                If mbytInFun = 0 Or mshDetail.Row = 0 Then Exit Sub
                If .TextMatrix(.Row, Col.C6�˶Խ��) = "" Then Exit Sub
                
                Call SetTxtInput
                txtInput.Text = .TextMatrix(.Row, .Col)
                Call zlControl.TxtSelAll(txtInput)
            Case Else
                strReportNO = "ZL" & glngSys \ 100 & "_INSIDE_1501"
                strInvoiceNO = .TextMatrix(.Row, Col.C0����)
                Call ReportOpen(gcnOracle, glngSys, strReportNO, Me, "Ʊ�ݺ�=" & strInvoiceNO & "", "Ʊ��=" & mlngƱ��, "ReportFormat=" & mlngƱ��, 1)
        End Select
    End With
End Sub

Private Sub SetCboResult()
    With mshDetail
        cboResult.Left = .Left + .CellLeft - 15
        cboResult.Top = .Top + .CellTop - 15
        cboResult.Width = .CellWidth + 15
        cboResult.Visible = True
        cboResult.SetFocus
    End With
End Sub

Private Sub SetTxtInput()
    With mshDetail
        txtInput.Left = .Left + .CellLeft - 15
        txtInput.Top = .Top + .CellTop - 15
        txtInput.Width = .CellWidth + 15
        txtInput.Height = .CellHeight
        txtInput.Visible = True
        txtInput.SetFocus
    End With
End Sub

Private Sub mshDetail_LeaveCell()
    If mbytInFun = 0 Or mshDetail.Row = 0 Then Exit Sub
    
    With mshDetail
        If .TextMatrix(1, Col.C0����) = "" Then Exit Sub
        If cboResult.Visible Then
            If .TextMatrix(.Row, .Col) <> Trim(cboResult.Text) Then
                If cboResult.ListIndex <= 0 Then
                    Call SetUnChecked(.Row)
                Else
                    Call SetChecked(.Row, Col.C6�˶Խ��, Trim(cboResult.Text), Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"))
                End If
            End If
        ElseIf txtInput.Visible Then
            If .TextMatrix(.Row, .Col) <> Trim(txtInput.Text) Then
                If zlCommFun.ActualLen(txtInput.Text) > txtInput.MaxLength Then
                    MsgBox "���ֻ��������" & txtInput.MaxLength & "���ַ�!", vbInformation, gstrSysName
                    Exit Sub
                End If
                If InStr(1, txtInput.Text, "'") > 0 Then
                    'MsgBox "ע��:��������ϵͳ��ֹ����������ַ�!", vbInformation, gstrSysName
                    Beep
                    Beep
                    Exit Sub
                End If
                Call SetChecked(.Row, Col.C7��ע, Trim(txtInput.Text))
            End If
        End If
    End With
End Sub

Private Sub mshDetail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mshDetail.MouseRow = 0 Then
        mshDetail.MousePointer = 99
    Else
        mshDetail.MousePointer = 0
    End If
End Sub

Private Sub mshDetail_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngCol As Long
    
    With mshDetail
        If .TextMatrix(1, Col.C0����) = "" Then Exit Sub
        If Button = 1 And .MousePointer = 99 Then
            lngCol = .MouseCol
            If .TextMatrix(0, lngCol) = "" Then Exit Sub
            
            .ColData(lngCol) = (.ColData(lngCol) + 1) Mod 2
            
            .Redraw = False
            .Col = lngCol: .ColSel = lngCol   '��������
            .Sort = IIf(.ColData(lngCol) = 1, 6, 5)
            .Col = 0
            .ColSel = .Cols - 1
            .Redraw = True
        End If
    End With
End Sub


Private Sub txtInput_LostFocus()
    If txtInput.Visible Then txtInput.Visible = False
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdFind_Click
        zlControl.TxtSelAll txt����
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub SetRow(ByVal lngRow As Long)
    Dim lngTop As Long
    With mshDetail
        .Row = lngRow
        lngTop = lngRow - 1
        If lngTop < 1 Then lngTop = 1
        If .RowIsVisible(lngTop) = False Then
            .TopRow = lngTop
        End If
        If mbytInFun = 0 Then
            .Col = 0
            .ColSel = .Cols - 1
        Else
            .Col = Col.C6�˶Խ��
        End If
    End With
End Sub

Private Sub InitContext()
    Dim blnҩ�� As Boolean
    
    On Error GoTo errHandle
    blnҩ�� = (glngSys \ 100 = 8)
    
    cboƱ��.Clear
    If blnҩ�� Then
        cboƱ��.AddItem "1-�շ��վ�":        cboƱ��.ItemData(cboƱ��.NewIndex) = 1
        cboƱ��.AddItem "5-��Ա��":          cboƱ��.ItemData(cboƱ��.NewIndex) = 5
    Else
        If InStr(1, mstrPrivs, ";�շ��վ�;") > 0 Then
            cboƱ��.AddItem "1-�շ��վ�":        cboƱ��.ItemData(cboƱ��.NewIndex) = 1
        End If
        If InStr(1, mstrPrivs, ";Ԥ���վ�;") > 0 Or _
          (InStr(1, mstrPrivs, ";Ԥ������Ʊ��;") > 0 _
          Or InStr(1, mstrPrivs, ";Ԥ��סԺƱ��;") > 0) Then
            cboƱ��.AddItem "2-Ԥ���վ�":        cboƱ��.ItemData(cboƱ��.NewIndex) = 2
        End If
        If InStr(1, mstrPrivs, ";�����վ�;") > 0 Then
          cboƱ��.AddItem "3-�����վ�":        cboƱ��.ItemData(cboƱ��.NewIndex) = 3
        End If
        If InStr(1, mstrPrivs, ";�Һ��վ�;") > 0 Then
          cboƱ��.AddItem "4-�Һ��վ�":        cboƱ��.ItemData(cboƱ��.NewIndex) = 4
        End If
        If InStr(1, mstrPrivs, ";ҽ�ƿ�;") > 0 Then
           cboƱ��.AddItem "5-ҽ�ƿ�":          cboƱ��.ItemData(cboƱ��.NewIndex) = 5
        End If
        If zlStr.IsHavePrivs(mstrPrivs, "���ѿ�") Then
           cboƱ��.AddItem "6-���ѿ�": cboƱ��.ItemData(cboƱ��.NewIndex) = 6
        End If
'        cboƱ��.AddItem "1-�շ��վ�":        cboƱ��.ItemData(cboƱ��.NewIndex) = 1
'        cboƱ��.AddItem "2-Ԥ���վ�":        cboƱ��.ItemData(cboƱ��.NewIndex) = 2
'        cboƱ��.AddItem "3-�����վ�":        cboƱ��.ItemData(cboƱ��.NewIndex) = 3
'        cboƱ��.AddItem "4-�Һ��վ�":        cboƱ��.ItemData(cboƱ��.NewIndex) = 4
'        cboƱ��.AddItem "5-ҽ�ƿ�":          cboƱ��.ItemData(cboƱ��.NewIndex) = 5
'        cboƱ��.AddItem "6-���ѿ�":          cboƱ��.ItemData(cboƱ��.NewIndex) = 6
        cboƱ��.ListIndex = 0
    End If
    
    cboFindType.NotAutoAppendKind = True
    cboFindType.ShowSortName = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Select��������(ByVal objCtl As Object, _
    ByVal strKey As String, ByVal intƱ�� As gBillType, ByVal bytMode As Byte) As Boolean
    '����:ѡ��ָ������������
    '���:
    '     strKey-����Ľ�ֵ
    '     intƱ��-��ǰѡ���Ʊ��
    '     bytMode-����ģʽ��1-�������˲��ң�2-����Ʊ�Ų���
    '����:
    '����:���ҳɹ�,����true,���򷵻�False
    Dim rsTemp As ADODB.Recordset, strWhere As String, strSQL As String
    Dim blnCancel As Boolean, vRect As RECT, blnFind As Boolean
    Dim strʹ����� As String
    
    Err = 0: On Error GoTo ErrHand:
    If strKey = "" Then Exit Function
    
    If bytMode = 1 Then '�������˲���
        If IsNumeric(strKey) Then '1.����ȫ������ʱֻƥ�����
            strWhere = " And ��� Like [2]"
        ElseIf zlCommFun.IsCharAlpha(strKey) Then '2.����ȫ����ĸʱֻƥ�����
            strWhere = " And ���� Like [2]"
        ElseIf zlStr.IsCharChinese(strKey) Then '2.���뺬�к���ʱֻƥ������
            strWhere = " And ���� Like [2]"
        Else
            strWhere = " And (���� Like [2] Or ���� Like [2] Or ��� Like [2])"
        End If
        strKey = GetMatchingSting(strKey, False)
        
        strWhere = " And a.������ In (Select ���� From  ��Ա�� Where 1=1 " & strWhere & ")"
    Else '����Ʊ�Ų���
        '˵���� And ���� Like '%%' ��һ���Ŀ����ʹ�ü�¼ֻ��һ��ʱ������ѡ����
        If intƱ�� = gBillType.���ѿ� Then
            strWhere = " And a.Id In(Select ����ID From ���ѿ�ʹ�ü�¼ Where ���� = [2] And ���� Like '%%')"
        Else
            strWhere = " And a.Id In(Select ����ID From Ʊ��ʹ����ϸ Where Ʊ�� = [1] And ���� = [2] And ���� Like '%%')"
        End If
    End If
    
    If intƱ�� = gBillType.���￨ Then
        strSQL = _
            "Select a.Id, Nvl(b.����, '���￨') As ʹ�����, a.��ʼ����, a.��ֹ����, a.������," & vbNewLine & _
            "       Decode(a.ʹ�÷�ʽ, 1, '����', '����') As ʹ�÷�ʽ, a.��ע, a.�Ǽ���, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd') As ����ʱ��," & vbNewLine & _
            "       Row_Number() Over(Partition By a.������ Order By a.�Ǽ�ʱ�� Desc) As �����к�" & vbNewLine & _
            "From Ʊ�����ü�¼ A, ҽ�ƿ���� B" & vbNewLine & _
            "Where To_Number(Nvl(a.ʹ�����, '0')) = b.Id(+) And a.Ʊ�� = [1]" & strWhere
    ElseIf intƱ�� = gBillType.���ѿ� Then
        strSQL = _
            "Select a.Id, Nvl(b.����, '���ѿ�') As ʹ�����, a.��ʼ���� As ��ʼ����, a.��ֹ���� As ��ֹ����, a.������," & vbNewLine & _
            "       Decode(a.ʹ�÷�ʽ, 1, '����', '����') As ʹ�÷�ʽ, a.��ע, a.�Ǽ���, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd') As ����ʱ��," & vbNewLine & _
            "       Row_Number() Over(Partition By a.������ Order By a.�Ǽ�ʱ�� Desc) As �����к�" & vbNewLine & _
            "From ���ѿ����ü�¼ A, ���ѿ����Ŀ¼ B" & vbNewLine & _
            "Where To_Number(Nvl(a.�ӿڱ��, '0')) = b.���(+)" & strWhere
    Else
        If intƱ�� = gBillType.�շ��վ� Or intƱ�� = gBillType.�����վ� Then
            strʹ����� = "a.ʹ�����,"
        ElseIf intƱ�� = gBillType.Ԥ���վ� Then
            strʹ����� = "Decode(Nvl(a.ʹ�����,'0'),'0','','1','����','סԺ') As ʹ�����,"
        End If
        
        strSQL = _
            "Select a.Id, " & strʹ����� & " a.��ʼ����, a.��ֹ����, a.������, Decode(a.ʹ�÷�ʽ, 1, '����', '����') As ʹ�÷�ʽ," & vbNewLine & _
            "       a.��ע, a.�Ǽ���, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd') As ����ʱ��," & vbNewLine & _
            "       Row_Number() Over(Partition By a.������ Order By a.�Ǽ�ʱ�� Desc) As �����к�" & vbNewLine & _
            "From Ʊ�����ü�¼ A" & vbNewLine & _
            "Where a.Ʊ�� = [1] " & strWhere
    End If
    
    strSQL = _
        "Select Id," & IIf(intƱ�� = gBillType.�Һ��վ�, "", " ʹ�����,") & _
        "       ��ʼ����, ��ֹ����, ������, ʹ�÷�ʽ, ��ע, �Ǽ���, ����ʱ��" & vbNewLine & _
        "From (" & strSQL & ")" & vbNewLine
    If bytMode = 1 Then '�������˲���ʱ����ʾ���10�����ü�¼
        strSQL = strSQL & _
            "Where �����к� < 11"
    End If
    
    '���궨λ
    vRect = zlControl.GetControlRect(objCtl.hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���ü�¼ѡ��", False, "", "", False, False, True, _
        vRect.Left - 15, vRect.Top, objCtl.Height, blnCancel, False, False, intƱ��, UCase(strKey))
   
   If blnCancel Then
        zlControl.ControlSetFocus objCtl: zlControl.TxtSelAll objCtl
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgbox "δ�ҵ��������������ü�¼�����飡"
        zlControl.ControlSetFocus objCtl: zlControl.TxtSelAll objCtl
        Exit Function
    End If
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "δ�ҵ��������������ü�¼�����飡"
        zlControl.ControlSetFocus objCtl: zlControl.TxtSelAll objCtl
        Exit Function
    End If
    
    mlngƱ�� = intƱ��
    mblnIsBIll = CurrentIsBill(mlngƱ��)
    mlng����ID = Val(NVL(rsTemp!ID))
    lbl����.Caption = IIf(mblnIsBIll, "����(&N)", "����(&N)")
    cmdFind.Caption = IIf(mblnIsBIll, "��λƱ��(&F)", "��λ����(&F)")
    If bytMode = 1 Then txtFind.Text = NVL(rsTemp!������)
    
    Call RefrashData(mlng����ID)
    zlCommFun.PressKey vbKeyTab
    
    Select�������� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

