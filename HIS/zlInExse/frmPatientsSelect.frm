VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatientsSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��Լ��λ����ѡ��"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10320
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   10320
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ListView lvwPati 
      Height          =   4215
      Left            =   3240
      TabIndex        =   7
      Top             =   600
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   7435
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "�Ա�"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "�ѱ�"
         Object.Width           =   1677
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "����ʱ��"
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "δ����"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "��ѡ(&U)"
      Height          =   350
      Index           =   2
      Left            =   9120
      TabIndex        =   6
      Top             =   120
      Width           =   1100
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "ȫ��(&R)"
      Height          =   350
      Index           =   1
      Left            =   8040
      TabIndex        =   5
      Top             =   120
      Width           =   1100
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "ȫѡ(&A)"
      Height          =   350
      Index           =   0
      Left            =   6960
      TabIndex        =   4
      Top             =   120
      Width           =   1100
   End
   Begin VB.CommandButton cmdLookFor 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   1100
   End
   Begin VB.TextBox txtUnitName 
      Height          =   350
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7740
      TabIndex        =   8
      Top             =   4995
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9120
      TabIndex        =   9
      Top             =   4995
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshUnit 
      Height          =   4185
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   7382
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "ID|^  ����  |^      ����       "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmPatientsSelect.frx":0000
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��λ(&D)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   80
      TabIndex        =   0
      Top             =   175
      Width           =   840
   End
End
Attribute VB_Name = "frmPatientsSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Public mstrUnitName As String
Public mlngUnitID As Long
Public mrsPatients As ADODB.Recordset

Private Enum COLUNIT
    C0ID = 0
    C1���� = 1
    C2���� = 2
    C3����ʱ�� = 3
End Enum
Private Enum COLPATIENT
    C0���� = 0
    C1�Ա� = 1
    C2���� = 2
    C3�ѱ� = 3
    C4����ʱ�� = 4
    C5δ���� = 5
End Enum


Private Sub cmdCancel_Click()
    gblnOK = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, strPatiIDs As String
    
    strPatiIDs = ""
    For i = 1 To lvwPati.ListItems.Count
        If lvwPati.ListItems(i).Checked Then
            strPatiIDs = strPatiIDs & " Or ����ID=" & lvwPati.ListItems(i).Tag
        End If
    Next
    strPatiIDs = Mid(strPatiIDs, 4)
    If strPatiIDs = "" Then
        MsgBox "������ѡ��һλ���ˣ�", vbInformation, gstrSysName
        Exit Sub
    End If
        
    mlngUnitID = Val(mshUnit.TextMatrix(mshUnit.Row, COLUNIT.C0ID))
    mstrUnitName = mshUnit.TextMatrix(mshUnit.Row, COLUNIT.C2����)
    
    mrsPatients.Filter = strPatiIDs
    
    gblnOK = True
    Me.Hide
End Sub

Private Sub Form_Activate()
    lvwPati.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    
    Call LoadData
End Sub


Private Sub cmdLookFor_Click()
    mstrUnitName = txtUnitName.Text
    Call LoadData
    Call txtUnitName.SetFocus
End Sub

Private Sub LoadData()
    Dim rsTmp As ADODB.Recordset
    
    mshUnit.Redraw = False
    Set rsTmp = GetUnits(mstrUnitName)
    If rsTmp.RecordCount > 1 Then
        Set mshUnit.DataSource = rsTmp        'û������ʱ,ʹ�ô˷�ʽ,�´�������ʱ������ж�λ��λ
    Else
        Call grid.BandRec(mshUnit, rsTmp)
    End If
    
    mshUnit.ColWidth(COLUNIT.C0ID) = 0
    mshUnit.ColWidth(COLUNIT.C1����) = 550
    mshUnit.ColWidth(COLUNIT.C2����) = 2100
    mshUnit.ColWidth(COLUNIT.C3����ʱ��) = 0
    If mshUnit.Rows = 1 Then mshUnit.Rows = 2
    mshUnit.Redraw = True
    
    If mstrUnitName <> "" Then
        mshUnit.Row = 1: mshUnit.RowSel = mshUnit.Row
        mshUnit.Col = 0: mshUnit.ColSel = mshUnit.Cols - 1
        Call mshUnit_EnterCell
        Call mshUnit_SelChange
        Call mshUnit_LostFocus
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwPati.SortOrder = IIf(lvwPati.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    lvwPati.SortKey = ColumnHeader.Index - 1
    lvwPati.Sorted = True
End Sub

Private Sub lvwPati_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub mshUnit_Click()
    '��mstrUnitNameΪ��ʱ,���봰��ʱ,û����ʾ��ǰ��λ�Ĳ���,��Ҫ�ٵ�һ�²���ʾ
    'mshUnit_SelChange������mshUnit_Click֮ǰ,���Բ����ٴε���
    If Val(mshUnit.Tag) = 0 Then Call mshUnit_SelChange
End Sub

Private Sub mshUnit_EnterCell()
    mshUnit.ForeColorSel = mshUnit.CellForeColor
End Sub

Private Sub mshUnit_GotFocus()
    mshUnit.BackColorSel = &HC0C0C0
        
End Sub

Private Sub mshUnit_LostFocus()
    mshUnit.BackColorSel = &HE0E0E0
End Sub

Private Sub mshUnit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshUnit.MouseRow = 0 Then
        mshUnit.MousePointer = 99
    Else
        mshUnit.MousePointer = 0
    End If
End Sub

Private Sub mshUnit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long, strTime As String, blnDel As Boolean
    
    lngCol = mshUnit.MouseCol
    
    If Button = 1 And mshUnit.MousePointer = 99 Then
        If mshUnit.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshUnit.TextMatrix(1, 0) = "" Then Exit Sub
        
        mshUnit.ColData(lngCol) = (mshUnit.ColData(lngCol) + 1) Mod 2
        mshUnit.Redraw = False
        mshUnit.Col = lngCol: mshUnit.ColSel = lngCol   '��������
        mshUnit.Sort = IIf(mshUnit.ColData(lngCol) = 1, 6, 5)
        mshUnit.Col = 0
        mshUnit.ColSel = mshUnit.Cols - 1
        mshUnit.Redraw = True
        
    End If
End Sub

Private Sub mshUnit_SelChange()
    Dim lngUnit As Long, i As Long, blnHistory As Boolean
    Dim objItem As ListItem
    
    If mshUnit.Row = 0 Then Exit Sub
    
    lngUnit = Val(mshUnit.TextMatrix(mshUnit.Row, COLUNIT.C0ID))
    mshUnit.Tag = lngUnit
    If lngUnit = 0 Then
        Set mrsPatients = Nothing
    Else
        blnHistory = zlDatabase.DateMoved(Format(mshUnit.TextMatrix(mshUnit.Row, COLUNIT.C3����ʱ��), "yyyy-MM-dd 00:00:00"), 1, glngSys, Me.Caption)
        Set mrsPatients = GetPatients(lngUnit, blnHistory)
    End If
    
    lvwPati.ListItems.Clear
    If Not mrsPatients Is Nothing Then
        For i = 1 To mrsPatients.RecordCount
            Set objItem = lvwPati.ListItems.Add(, "_" & mrsPatients!����ID, mrsPatients!����)
            objItem.Tag = mrsPatients!����ID
            objItem.SubItems(COLPATIENT.C1�Ա�) = "" & mrsPatients!�Ա�
            objItem.SubItems(COLPATIENT.C2����) = "" & mrsPatients!����
            objItem.SubItems(COLPATIENT.C3�ѱ�) = "" & mrsPatients!�ѱ�
            objItem.SubItems(COLPATIENT.C4����ʱ��) = "" & mrsPatients!����ʱ��
            objItem.SubItems(COLPATIENT.C5δ����) = Format(mrsPatients!���ʽ��, gstrDec)
            
            mrsPatients.MoveNext
        Next
    End If
End Sub

Private Sub txtUnitName_GotFocus()
    Call zlcontrol.TxtSelAll(txtUnitName)
    Call OpenIme(gstrIme)
End Sub

Private Sub txtUnitName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtUnitName_Validate(Cancel As Boolean)
    txtUnitName.Text = Trim(txtUnitName.Text)
    Call OpenIme
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    Dim i As Long
    For i = 1 To lvwPati.ListItems.Count
        lvwPati.ListItems(i).Checked = Choose(Index + 1, True, False, Not lvwPati.ListItems(i).Checked)
    Next
End Sub



Private Function GetUnits(Optional strUnitName As String) As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    On Error GoTo errH
    
    strSql = "Select ID,����,����,To_Char(����ʱ��, 'YYYY-MM-DD HH24:MI') ����ʱ�� From ��Լ��λ"
    If strUnitName <> "" Then
        If zlCommFun.IsCharChinese(strUnitName) Then
            strSql = strSql & " Where ���� like [1]"
        ElseIf zlCommFun.IsCharAlpha(strUnitName) Then
            strSql = strSql & " Where ���� like [1]"
        ElseIf zlCommFun.IsNumOrChar(strUnitName) Then
            strSql = strSql & " Where ���� like [1]"
        Else
            strSql = strSql & " Where ���� like [1] or ���� like [1] or ���� like [1]"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UCase(strUnitName & "%"))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    End If
    
    If rsTmp.RecordCount > 0 Then rsTmp.Sort = "����"
    Set GetUnits = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPatients(lngUnitID As Long, blnHistory As Boolean) As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strNotZeroSQL As String, strHistorySQL As String
    
    On Error GoTo errH
    
    strNotZeroSQL = "" & _
     IIf(gblnZero, "", " And Not Exists ( Select 1 " & vbNewLine & _
                        "                  From ������ü�¼ B" & vbNewLine & _
                        "                  Where B.NO = A.NO And B.��¼���� = A.��¼���� And B.��� = A.���" & vbNewLine & _
                        "                  Group By B.NO, B.��¼����, B.���" & vbNewLine & _
                        "                  Having Nvl(Sum(B.ʵ�ս��), 0) = 0)" & vbNewLine)
    strHistorySQL = ""
    If blnHistory Then
        strHistorySQL = "" & _
        "              Union All" & vbNewLine & _
        "              Select B.����id, B.����, B.�Ա�, B.����, B.�ѱ�, Nvl(B.����ʱ��, B.�Ǽ�ʱ��) ����ʱ��, A.NO, A.���," & vbNewLine & _
        "                     A.��¼����, A.��¼״̬, A.ִ��״̬, A.ʵ�ս��, A.���ʽ��" & vbNewLine & _
        "              From H������ü�¼ A, ������Ϣ B" & vbNewLine & _
        "              Where A.����id = B.����id And A.����id Is Not Null And A.���ʷ��� = 1 And A.�����־ IN(1,4) And" & vbNewLine & _
        "                    Nvl(A.ʵ�ս��, 0) <> Nvl(A.���ʽ��, 0) And B.��ͬ��λid = [1] And B.��ǰ����id Is Null"
    End If
                        
    

    '1.���ʺͽ������ϵķ��ü�¼���ֱܷ������߱�ͺ󱸱���
    '2.�������ֻ�м��ʺͼ������ʵķ���,������ý��ʲ���Ϊ��ʱ,����ʾ�ò���
    strSql = "" & _
    " Select ����id, ����, �Ա�, ����, �ѱ�, ����ʱ��, Sum(δ����) ���ʽ��" & vbNewLine & _
    " From (Select B.����id, B.����, B.�Ա�, B.����, B.�ѱ�," & vbNewLine & _
    "              To_Char(Nvl(B.����ʱ��, B.�Ǽ�ʱ��), 'YYYY-MM-DD HH24:MI') ����ʱ��, Nvl(A.ʵ�ս��, 0) As δ����" & vbNewLine & _
    "       From ������ü�¼ A, ������Ϣ B" & vbNewLine & _
    "       Where A.����id = B.����id And A.����id Is Null And A.��¼״̬ <> 0 And A.���ʷ��� = 1 And A.�����־ IN(1,4) And" & vbNewLine & _
    "             B.��ͬ��λid = [1] And B.��ǰ����id Is Null " & vbNewLine & strNotZeroSQL & _
    "       Union All" & vbNewLine & _
    "       Select ����id, ����, �Ա�, ����, �ѱ�, To_Char(����ʱ��, 'YYYY-MM-DD HH24:MI') ����ʱ��," & vbNewLine & _
    "              Nvl(Sum(ʵ�ս��), 0) - Nvl(Sum(���ʽ��), 0) As δ����" & vbNewLine & _
    "       From (Select B.����id, B.����, B.�Ա�, B.����, B.�ѱ�, Nvl(B.����ʱ��, B.�Ǽ�ʱ��) ����ʱ��, NO, A.���, A.��¼����," & vbNewLine & _
    "                     A.��¼״̬, A.ִ��״̬, A.ʵ�ս��, A.���ʽ��" & vbNewLine & _
    "              From ������ü�¼ A, ������Ϣ B" & vbNewLine & _
    "              Where A.����id = B.����id And A.����id Is Not Null And A.���ʷ��� = 1 And A.�����־ IN(1,4) And" & vbNewLine & _
    "                    Nvl(A.ʵ�ս��, 0) <> Nvl(A.���ʽ��, 0) And B.��ͬ��λid = [1] And B.��ǰ����id Is Null" & vbNewLine & _
                   strHistorySQL & ")" & vbNewLine & _
    "       Group By ����id, ����, �Ա�, ����, �ѱ�, ����ʱ��, NO, ���, Mod(��¼����, 10), ��¼״̬, ִ��״̬" & vbNewLine & _
    "       Having Nvl(Sum(ʵ�ս��), 0) - Nvl(Sum(���ʽ��), 0) <> 0)" & vbNewLine & _
    " Group By ����id, ����, �Ա�, ����, �ѱ�, ����ʱ��"


    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngUnitID)
        
    If rsTmp.RecordCount > 0 Then rsTmp.Sort = "����"
    Set GetPatients = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

