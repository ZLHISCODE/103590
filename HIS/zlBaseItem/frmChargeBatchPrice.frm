VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmChargeBatchPrice 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�շ���Ŀ��������"
   ClientHeight    =   3525
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5685
   ClipControls    =   0   'False
   Icon            =   "frmChargeBatchPrice.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.ComboBox cbo�۸�ȼ� 
      Height          =   300
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   150
      Width           =   2535
   End
   Begin VB.TextBox txtChargeType 
      BackColor       =   &H8000000F&
      Height          =   270
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   885
      Width           =   2535
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "��"
      Height          =   260
      Left            =   3840
      TabIndex        =   2
      Top             =   540
      Width           =   255
   End
   Begin VB.TextBox txtType 
      Height          =   270
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   525
      Width           =   2535
   End
   Begin VB.Frame fra���۷�ʽ 
      Caption         =   "����"
      Height          =   2250
      Left            =   330
      TabIndex        =   15
      Top             =   1230
      Width           =   3795
      Begin VB.CheckBox chkByBase 
         Caption         =   "�۸�ȼ�δ���ü۸�ʱ����ȱʡ�۸��ԭ�ۻ����ϵ���(&P)"
         Height          =   375
         Left            =   270
         TabIndex        =   20
         Top             =   1830
         Width           =   3360
      End
      Begin VB.CheckBox chk�Ӽ� 
         Caption         =   "�����÷����������ӷ������Ŀ(&G)"
         Height          =   255
         Left            =   270
         TabIndex        =   9
         Top             =   1545
         Width           =   3195
      End
      Begin VB.TextBox txtEdit 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   2310
         TabIndex        =   7
         Top             =   750
         Width           =   885
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   2310
         TabIndex        =   5
         Top             =   330
         Width           =   885
      End
      Begin VB.OptionButton optAdjust 
         Caption         =   "��ԭ�ۻ����ϵ���(&B)"
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   6
         Top             =   750
         Width           =   2025
      End
      Begin VB.OptionButton optAdjust 
         Caption         =   "��ԭ�ۻ����ϵ���(&P)"
         Height          =   315
         Index           =   0
         Left            =   270
         TabIndex        =   4
         Top             =   330
         Value           =   -1  'True
         Width           =   2025
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   285
         Left            =   1350
         TabIndex        =   8
         Top             =   1140
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   172163075
         CurrentDate     =   36444
         MaxDate         =   401768
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ִ������(&E)"
         Height          =   180
         Index           =   15
         Left            =   300
         TabIndex        =   18
         Top             =   1200
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         TabIndex        =   16
         Top             =   330
         Width           =   150
      End
      Begin VB.Label Label5 
         Caption         =   "Ԫ"
         Height          =   180
         Left            =   3240
         TabIndex        =   17
         Top             =   810
         Width           =   180
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4350
      TabIndex        =   11
      Tag             =   "����"
      Top             =   690
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4350
      TabIndex        =   10
      Tag             =   "����"
      Top             =   240
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   4350
      TabIndex        =   12
      Tag             =   "����"
      Top             =   2820
      Width           =   1100
   End
   Begin VB.Label lbl�۸�ȼ� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�շѼ۸�ȼ���"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   330
      TabIndex        =   19
      Top             =   173
      Width           =   1320
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�շ���Ŀ���ࣺ"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   330
      TabIndex        =   14
      Top             =   900
      Width           =   1305
   End
   Begin VB.Label lbl��� 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�շ���Ŀ���"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   330
      TabIndex        =   13
      Top             =   555
      Width           =   1320
   End
End
Attribute VB_Name = "frmChargeBatchPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrPrivs As String
Public datSingle As Date '���������µ��������
Public datAll As Date    '������������Ŀ���������
Public dblSingle As Double   '���������µ���С���
Public dblAll As Double      '������������Ŀ����С���
Public mblnCanUpdateAll As Boolean '�Ƿ��������������Ŀ��δ���ü۸�ȼ��������˼۸�ȼ��С�����Ժ����Ȩ��

Private Sub cbo�۸�ȼ�_Click()
    chkByBase.Enabled = Not (cbo�۸�ȼ�.Text = "ȱʡ")
    chkByBase.value = vbUnchecked
End Sub

Private Sub chk�Ӽ�_Click()
    If chk�Ӽ�.value = 1 Then
        dtpBegin.MinDate = datAll
    Else
        dtpBegin.MinDate = datSingle
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function IsValid() As Boolean
'�жϺϷ�ֵ
    If cbo�۸�ȼ�.ListIndex = -1 Then
        MsgBox "��ѡ��۸�ȼ���", vbExclamation, gstrSysName
        If cbo�۸�ȼ�.Visible And cbo�۸�ȼ�.Enabled Then cbo�۸�ȼ�.SetFocus
        Exit Function
    End If
    If optAdjust(0).value = True Then
        If IsNumeric(txtEdit(0).Text) = False Then
            MsgBox "������һ����ֵ��", vbExclamation, gstrSysName
            zlControl.TxtSelAll txtEdit(0)
            txtEdit(0).SetFocus
            Exit Function
        End If
        If Val(txtEdit(0).Text) = 0 Then
            MsgBox "����ֵ����Ϊ�㡣", vbExclamation, gstrSysName
            zlControl.TxtSelAll txtEdit(0)
            txtEdit(0).SetFocus
            Exit Function
        End If
        If Val(txtEdit(0).Text) <= -100 Then
            MsgBox "����ֵ���ܵ���-100%��", vbExclamation, gstrSysName
            zlControl.TxtSelAll txtEdit(0)
            txtEdit(0).SetFocus
            Exit Function
        End If
        If Val(txtEdit(0).Text) > 9999 Then
            MsgBox "����ֵ̫���ˡ�", vbExclamation, gstrSysName
            zlControl.TxtSelAll txtEdit(0)
            txtEdit(0).SetFocus
            Exit Function
        End If
    Else
        If IsNumeric(txtEdit(1).Text) = False Then
            MsgBox "������һ����ֵ��", vbExclamation, gstrSysName
            zlControl.TxtSelAll txtEdit(1)
            txtEdit(1).SetFocus
            Exit Function
        End If
        If Val(txtEdit(1).Text) = 0 Then
            MsgBox "����ֵ����Ϊ�㡣", vbExclamation, gstrSysName
            zlControl.TxtSelAll txtEdit(1)
            txtEdit(1).SetFocus
            Exit Function
        End If
        If Val(txtEdit(1).Text) + IIF(chk�Ӽ�.value = 0, dblSingle, dblAll) <= 0 Then
            MsgBox "����ֵ����Ҫ����-" & IIF(chk�Ӽ�.value = 0, dblSingle, dblAll) & "��", vbExclamation, gstrSysName
            zlControl.TxtSelAll txtEdit(1)
            txtEdit(1).SetFocus
            Exit Function
        End If
        If Val(txtEdit(1).Text) > 9999999 Then
            MsgBox "����ֵ̫���ˡ�", vbExclamation, gstrSysName
            zlControl.TxtSelAll txtEdit(1)
            txtEdit(1).SetFocus
            Exit Function
        End If
    End If
    IsValid = True
End Function

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim int�������� As Integer 'ȡֵΪ1��������������Χ��2����������ȫ��Χ��3����ֵ������Χ��4����ֵ��ȫ��Χ��
    Dim dbl����ֵ   As Double
    Dim str�۸�ȼ� As String
    
    If IsValid = False Then Exit Sub
    If MsgBox("�������ۻ�Ӱ������Ŀ�ļ۸�" & vbCrLf & "��ȷ������ȷ���ã�", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    On Error GoTo errMass
    
    If chk�Ӽ�.value = 0 Then
        int�������� = IIF(optAdjust(0).value = True, 1, 3)
    Else
        int�������� = IIF(optAdjust(0).value = True, 2, 4)
    End If
    
    If cbo�۸�ȼ�.Text = "���м۸�ȼ�" Then
        str�۸�ȼ� = "'����'"
    ElseIf cbo�۸�ȼ�.Text = "ȱʡ" Then
        str�۸�ȼ� = "NULL"
    Else
        str�۸�ȼ� = "'" & cbo�۸�ȼ�.Text & "'"
    End If
    
    'Zl_�շ�ϸĿ_Raisemass(
    gstrSQL = "zl_�շ�ϸĿ_RaiseMass("
    '  ��������_In In Number,
    gstrSQL = gstrSQL & "" & int�������� & ","
    '  ����ֵ_In   In Number,
    gstrSQL = gstrSQL & "" & IIF(optAdjust(0).value = True, Val(txtEdit(0).Text) / 100, Val(txtEdit(1).Text)) & ","
    '  ִ������_In In Date,
    gstrSQL = gstrSQL & "" & "to_date('" & Format(dtpBegin.value, "yyyy-MM-dd") & "','YYYY-MM-DD')" & ","
    '  ��ֹ����_In In Date,
    gstrSQL = gstrSQL & "" & "to_date('" & Format(dtpBegin.value - 1, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')" & ","
    '  ������_In   In Varchar2,
    gstrSQL = gstrSQL & "'" & gstrUserName & "',"
    '  ����id_In   In �շ���ĿĿ¼.����id%Type := Null,
    gstrSQL = gstrSQL & "" & IIF(lbl����.Tag = "" Or lbl����.Tag = "0", "null", lbl����.Tag) & ","
    '  ���_In     In �շ���ĿĿ¼.���%Type := Null,
    gstrSQL = gstrSQL & "'" & lbl���.Tag & "',"
    '  �۸�ȼ�_In In �շѼ�Ŀ.�۸�ȼ�%Type := '����'
    gstrSQL = gstrSQL & "" & str�۸�ȼ� & ","
    '  վ��_In     In �շ���ĿĿ¼.վ��%Type := Null
    gstrSQL = gstrSQL & "" & IIF(mblnCanUpdateAll, "NULL", "'" & gstrNodeNo & "'") & ","
    '  ��ȱʡ�۸����_In Number := 0
    gstrSQL = gstrSQL & "" & IIF(chkByBase.value = vbChecked, 1, 0) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    If Not frmChargeManage.lvwMain_S.SelectedItem Is Nothing Then
        frmChargeManage.FillItem frmChargeManage.lvwMain_S.SelectedItem.Key
    End If
    Unload Me
    Exit Sub
errMass:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSel_Click()
On Error GoTo ErrHandle
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strReturn As String
    
    With frmSelCur
        strSQL = "Select Null,'�������' From Dual Union All Select ����,���� From �շ���Ŀ��� where not ���� in ('5','6','7') "
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            strReturn = .ShowCurrSel(Me, rsTmp, "����,800,0,2;���,1500,0,2", "���ѡ����", False, Me.lbl���.Tag, 0)
            If Trim(strReturn) <> "" Then
                txtType.Text = Split(strReturn, ",")(1)
                Me.lbl���.Tag = Split(strReturn, ",")(0)
            End If
        Else
            MsgBox "���κο��õ��������ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
            txtType.Text = "��"
            Me.lbl���.Tag = ""
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If txtEdit(0).Enabled = True Then txtEdit(0).SetFocus
End Sub

Private Sub Form_Load()
    Dim blnEnabled As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    cbo�۸�ȼ�.Clear
    cbo�۸�ȼ�.AddItem "���м۸�ȼ�"
    cbo�۸�ȼ�.AddItem "ȱʡ"
    If mblnCanUpdateAll Then
        strSQL = "Select Distinct a.���� As �۸�ȼ�" & vbNewLine & _
                " From �շѼ۸�ȼ� A" & vbNewLine & _
                " Where Nvl(a.�Ƿ�������ͨ��Ŀ, 0) = 1" & vbNewLine & _
                "       And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))"
    Else
        strSQL = "Select Distinct a.���� As �۸�ȼ�" & vbNewLine & _
                " From �շѼ۸�ȼ� A, �շѼ۸�ȼ�Ӧ�� B" & vbNewLine & _
                " Where a.���� = b.�۸�ȼ� And b.վ�� = [1] And Nvl(a.�Ƿ�������ͨ��Ŀ, 0) = 1" & vbNewLine & _
                "       And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))"

    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�۸�ȼ�", gstrNodeNo)
    Do While Not rsTemp.EOF
        cbo�۸�ȼ�.AddItem Nvl(rsTemp!�۸�ȼ�)
        rsTemp.MoveNext
    Loop
    If cbo�۸�ȼ�.ListCount > 0 Then cbo�۸�ȼ�.ListIndex = 0
        
    blnEnabled = IsPriceGradeEnabled()
    cbo�۸�ȼ�.Enabled = blnEnabled
    chkByBase.Enabled = blnEnabled
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub optAdjust_Click(Index As Integer)
    Dim lngOther As Long
    
    lngOther = IIF(Index = 0, 1, 0)
    txtEdit(Index).Enabled = True
    txtEdit(Index).BackColor = &H80000005
    txtEdit(Index).SetFocus
    txtEdit(lngOther).Enabled = False
    txtEdit(lngOther).BackColor = &H8000000F
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789.-", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then KeyAscii = 0
End Sub
