VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmVItemEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������༭"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   Icon            =   "frmVItemEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraBase 
      Height          =   2790
      Left            =   105
      TabIndex        =   36
      Top             =   360
      Width           =   5460
      Begin VB.TextBox txt�ٴ����� 
         Height          =   600
         Left            =   1215
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   44
         ToolTipText     =   "Ҫ�ر༭ʱ����ʾ����"
         Top             =   2085
         Width           =   4065
      End
      Begin VB.ComboBox cbo�Ա��� 
         Height          =   300
         ItemData        =   "frmVItemEdit.frx":000C
         Left            =   3780
         List            =   "frmVItemEdit.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   1698
         Width           =   1500
      End
      Begin VB.TextBox txt��λ 
         Height          =   300
         Left            =   1215
         MaxLength       =   10
         TabIndex        =   16
         Top             =   1698
         Width           =   1170
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         ItemData        =   "frmVItemEdit.frx":0010
         Left            =   1215
         List            =   "frmVItemEdit.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1326
         Width           =   1185
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   1215
         MaxLength       =   8
         TabIndex        =   4
         Top             =   210
         Width           =   960
      End
      Begin VB.TextBox txt������ 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   1215
         MaxLength       =   40
         TabIndex        =   6
         Top             =   582
         Width           =   4065
      End
      Begin VB.TextBox txtӢ���� 
         Height          =   300
         Left            =   1215
         MaxLength       =   40
         TabIndex        =   8
         Top             =   954
         Width           =   4065
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   3375
         MaxLength       =   3
         TabIndex        =   12
         Top             =   1326
         Width           =   570
      End
      Begin VB.TextBox txtС�� 
         Height          =   300
         Left            =   4785
         MaxLength       =   1
         TabIndex        =   14
         Top             =   1326
         Width           =   495
      End
      Begin VB.Label lbl�ٴ����� 
         AutoSize        =   -1  'True
         Caption         =   "�ٴ�����(&M)"
         Height          =   180
         Left            =   180
         TabIndex        =   18
         Top             =   2130
         Width           =   990
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ŀ����(&R)"
         Height          =   180
         Left            =   180
         TabIndex        =   3
         Top             =   270
         Width           =   990
      End
      Begin VB.Label lbl������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&N)"
         Height          =   180
         Left            =   180
         TabIndex        =   5
         Top             =   642
         Width           =   990
      End
      Begin VB.Label lblӢ���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ӣ������(&E)"
         Height          =   180
         Left            =   180
         TabIndex        =   7
         Top             =   1014
         Width           =   990
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&T)"
         Height          =   180
         Left            =   180
         TabIndex        =   9
         Top             =   1386
         Width           =   990
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&L)"
         Height          =   180
         Left            =   2730
         TabIndex        =   11
         Top             =   1380
         Width           =   630
      End
      Begin VB.Label lblС�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "С��(&D)"
         Height          =   180
         Left            =   4110
         TabIndex        =   13
         Top             =   1380
         Width           =   630
      End
      Begin VB.Label lbl��λ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ֵ��λ(&U)"
         Height          =   180
         Left            =   180
         TabIndex        =   15
         Top             =   1758
         Width           =   990
      End
      Begin VB.Label lbl�Ա��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�����(&X)"
         Height          =   180
         Left            =   2730
         TabIndex        =   17
         Top             =   1755
         Width           =   990
      End
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "&P"
      Height          =   285
      Left            =   5235
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   75
      Width           =   315
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1140
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   1
      Top             =   60
      Width           =   4080
   End
   Begin VB.Frame fraScope 
      Height          =   2745
      Left            =   105
      TabIndex        =   35
      Top             =   3030
      Width           =   5460
      Begin VB.CheckBox chkDyn 
         Caption         =   "�Զ���"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2940
         TabIndex        =   42
         ToolTipText     =   "�������ַ�����Ϊ��ѡ/��ѡʱ�Ƿ�����ѡ���Զ���"
         Top             =   2325
         Width           =   915
      End
      Begin VB.CheckBox chkMust 
         Caption         =   "����"
         Height          =   300
         Left            =   4350
         TabIndex        =   41
         ToolTipText     =   "������д���ʱΪ�Ƿ������"
         Top             =   2325
         Width           =   660
      End
      Begin VB.CommandButton cmdMove 
         Height          =   345
         Index           =   1
         Left            =   5055
         Picture         =   "frmVItemEdit.frx":0014
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "����ƶ�"
         Top             =   1395
         Width           =   345
      End
      Begin VB.CommandButton cmdMove 
         Height          =   345
         Index           =   0
         Left            =   5055
         Picture         =   "frmVItemEdit.frx":0161
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "��ǰ�ƶ�"
         Top             =   1005
         Width           =   345
      End
      Begin VB.ComboBox cbo��ʾ�� 
         Height          =   300
         ItemData        =   "frmVItemEdit.frx":02AE
         Left            =   1215
         List            =   "frmVItemEdit.frx":02B0
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   195
         Width           =   2970
      End
      Begin VB.TextBox txt��ʼֵ 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1215
         MaxLength       =   250
         TabIndex        =   27
         Top             =   2325
         Visible         =   0   'False
         Width           =   1230
      End
      Begin ZL9BillEdit.BillEdit msh��ֵ�� 
         Height          =   1275
         Left            =   1215
         TabIndex        =   25
         Top             =   975
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   2249
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
         GridColor       =   -2147483633
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VB.TextBox txt��ֵ�� 
         Height          =   270
         Index           =   1
         Left            =   2940
         MaxLength       =   250
         TabIndex        =   24
         Top             =   600
         Width           =   1245
      End
      Begin VB.TextBox txt��ֵ�� 
         Height          =   300
         Index           =   0
         Left            =   1215
         MaxLength       =   250
         TabIndex        =   22
         Top             =   600
         Width           =   1230
      End
      Begin VB.Label lbl��ʾ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ַ���(&F)"
         Height          =   180
         Left            =   165
         TabIndex        =   19
         Top             =   255
         Width           =   990
      End
      Begin VB.Label lbl��ʼֵ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ��ֵ(&I)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   165
         TabIndex        =   26
         Top             =   2385
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Left            =   2610
         TabIndex        =   23
         Top             =   690
         Width           =   180
      End
      Begin VB.Label lbl��ֵ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȡֵ��Χ(&V)"
         Height          =   180
         Left            =   165
         TabIndex        =   21
         Top             =   660
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   105
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   5835
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3315
      TabIndex        =   32
      Top             =   5835
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4455
      TabIndex        =   33
      Top             =   5835
      Width           =   1100
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   3570
      Left            =   5730
      TabIndex        =   38
      TabStop         =   0   'False
      Tag             =   "1000"
      Top             =   420
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   6297
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
      Left            =   5760
      Top             =   135
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
            Picture         =   "frmVItemEdit.frx":02B2
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemEdit.frx":084C
            Key             =   "expend"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraWord 
      Height          =   1095
      Left            =   60
      TabIndex        =   37
      Top             =   6135
      Visible         =   0   'False
      Width           =   5460
      Begin VB.ComboBox cbo���ֱ��� 
         Height          =   300
         ItemData        =   "frmVItemEdit.frx":0DE6
         Left            =   2310
         List            =   "frmVItemEdit.frx":0DE8
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   210
         Width           =   2865
      End
      Begin VB.TextBox txt��ֵ���� 
         Height          =   300
         Left            =   2310
         MaxLength       =   100
         TabIndex        =   31
         Top             =   615
         Width           =   2865
      End
      Begin VB.Label lbl���ֱ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ת��Ϊ�ı��ı�������(&Y)"
         Height          =   180
         Left            =   180
         TabIndex        =   28
         Top             =   270
         Width           =   2070
      End
      Begin VB.Label lbl��ֵ���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ŀ��ֵΪ��ʱ��ʾΪ(&W)"
         Height          =   180
         Left            =   180
         TabIndex        =   30
         Top             =   675
         Width           =   2070
      End
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��Ŀ����(&K)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   105
      TabIndex        =   0
      Top             =   135
      Width           =   990
   End
End
Attribute VB_Name = "frmVItemEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'˵����
'   1���ϼ�����ͨ��������ShowMe�������������塢Ȩ�ޡ��༭��Ŀ�ķ���ID��ID,�༭״̬����Ϣ���ݽ��뱾����
'   2���༭״̬����Me.tag��ţ��ֱ�Ϊ"����"��"�޸�"��"����"�����ϼ�����ͨ��ShowMe����
'---------------------------------------------------
Private lngClassId As Long       '���༭�ķ���ID���ϼ�����ͨ��ShowMe���ݽ���
Private lngItemID As Long        '���༭����ĿID���޸ġ�����ʱ���ϼ�����ͨ��ShowMe���ݽ���,����ʱΪ0��

Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim strTemp As String, aryTemp() As String
Dim intCount As Integer

Public Sub ShowMe(ByVal frmParent As Object, ByVal byt״̬ As Byte, ByVal lng����id As Long, Optional ByVal lng��Ŀid As Long)
    '---------------------------------------------------
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '---------------------------------------------------
    Me.Tag = Switch(byt״̬ = 0, "����", byt״̬ = 1, "�޸�", byt״̬ = 2, "����")
    lngClassId = lng����id: lngItemID = lng��Ŀid
    
    '��д��Ҫѡ�������
    aryTemp = Split("0-��ֵ;1-����;2-����;3-�߼�", ";")
    Me.cbo����.Clear
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo����.AddItem aryTemp(intCount)
    Next
    Me.cbo����.ListIndex = 0
    
    aryTemp = Split("0-������;1-��;2-Ů", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo�Ա���.AddItem aryTemp(intCount)
    Next
    Me.cbo�Ա���.ListIndex = 0
    
    aryTemp = Split("1-��Ŀ��+��Ŀֵ+��λ;2-��Ŀֵ+��λ+��Ŀ��;3-��Ŀֵ+��λ", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo���ֱ���.AddItem aryTemp(intCount)
    Next
    Me.cbo���ֱ���.ListIndex = 0
    
    Err = 0: On Error GoTo errHand
    
    gstrSql = "select ID,�ϼ�ID,����,����,����" & _
            " From ������������" & _
            " Where ���� =(select ���� from ������������ where ID=[1])" & _
            " start with �ϼ�ID is null" & _
            " connect by prior ID=�ϼ�ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngClassId)
    
    With rsTemp
        If .BOF Or .EOF Then MsgBox "�����Ƚ������Ʒ�����Ŀ֮��������Ŀ", vbExclamation, gstrSysName: Unload Me: Exit Sub
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !���� & "]" & !����, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !�ϼ�ID, tvwChild, "_" & !ID, "[" & !���� & "]" & !����, "close")
            End If
            objNode.Tag = IIf(IsNull(!����), "", !����)
            objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
        Me.tvwClass.Nodes("_" & lng����id).Selected = True
        Me.txt����.Text = Me.tvwClass.SelectedItem.Text
        Me.txt����.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    End With
    
    '��ʾ����
    Me.Show 1, frmParent
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo��ʾ��_Click()
    Call zlSetGround
End Sub

Private Sub cbo��ʾ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo����_Click()
    '0-��ֵ��1-���֣�2-���ڣ�3-�߼�
    Me.txt����.Enabled = True: Me.txtС��.Enabled = True
    Select Case Left(Me.cbo����.Text, 1)
    Case 0
        aryTemp = Split("0-�ı�;1-����;2-����", ";")
    Case 1
        Me.txtС��.Text = 0: Me.txtС��.Enabled = False
        aryTemp = Split("0-�ı�;2-����;3-��ѡ;4-��ѡ", ";")
    Case 2
        Me.txt����.Text = 0: Me.txt����.Enabled = False
        Me.txtС��.Text = 0: Me.txtС��.Enabled = False
        aryTemp = Split("0-�ı�;2-����", ";")
    Case 3
        Me.txt����.Text = 0: Me.txt����.Enabled = False
        Me.txtС��.Text = 0: Me.txtС��.Enabled = False
        aryTemp = Split("3-��ѡ", ";")
    End Select
    Me.cbo��ʾ��.Clear
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo��ʾ��.AddItem aryTemp(intCount)
    Next
    Me.cbo��ʾ��.ListIndex = 0
    Call zlSetGround
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo���ֱ���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo�Ա���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdMove_Click(Index As Integer)
    If Index = 0 Then
        If msh��ֵ��.Row > 1 Then
            
            Call MoveItem(msh��ֵ��.Row, -1)
            msh��ֵ��.Row = msh��ֵ��.Row - 1

            
        End If
    ElseIf msh��ֵ��.Row < msh��ֵ��.Rows - 1 Then
        
        Call MoveItem(msh��ֵ��.Row, 1)
        msh��ֵ��.Row = msh��ֵ��.Row + 1

    End If
'    MSHFlexGrid1.TopRow = msh��ֵ��.Row
    If msh��ֵ��.MsfObj.RowIsVisible(msh��ֵ��.Row) = False Then
        msh��ֵ��.MsfObj.TopRow = msh��ֵ��.Row
    End If
    msh��ֵ��.SetFocus
End Sub

Private Function MoveItem(ByVal intCurRow As Integer, Optional ByVal intMove As Integer = 1) As Boolean
    '-----------------------------------------------------------------------------------------
    '����:
    '����:
    '-----------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim intCol As Integer
    
    On Error GoTo errHand
    
    strTmp = CStr(msh��ֵ��.RowData(intCurRow))
            
    msh��ֵ��.RowData(intCurRow) = msh��ֵ��.RowData(intCurRow + intMove)
    msh��ֵ��.RowData(intCurRow + intMove) = Val(strTmp)
    
    For intCol = 1 To msh��ֵ��.Cols - 1
        
        strTmp = msh��ֵ��.TextMatrix(intCurRow, intCol)
        
        msh��ֵ��.TextMatrix(msh��ֵ��.Row, intCol) = msh��ֵ��.TextMatrix(intCurRow + intMove, intCol)
        
        msh��ֵ��.TextMatrix(intCurRow + intMove, intCol) = strTmp
        
    Next
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdOK_Click()
    'һ�����Լ��
    If Trim(Me.txt����.Text) = "" Then MsgBox "��������Ŀ���룡", vbInformation, gstrSysName: Me.txt����.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > 8 Then MsgBox "���볬�������8���ַ�����", vbInformation, gstrSysName: Me.txt����.SetFocus: Exit Sub
    If Trim(Me.txt������.Text) = "" Then MsgBox "��������������", vbInformation, gstrSysName: Me.txt������.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt������.Text), vbFromUnicode)) > 40 Then MsgBox "���������������40���ַ���20�����֣���", vbInformation, gstrSysName: Me.txt������.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txtӢ����.Text), vbFromUnicode)) > 40 Then MsgBox "Ӣ�������������40���ַ�����", vbInformation, gstrSysName: Me.txtӢ����.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt��λ.Text), vbFromUnicode)) > 10 Then MsgBox "��λ���������10���ַ���5�����֣���", vbInformation, gstrSysName: Me.txt��λ.SetFocus: Exit Sub
'    If Me.cbo����.Text = "0-��ֵ" And IsNumeric(Me.txt��ʼֵ) = False Then MsgBox "����Ϊ��ֵʱ��ʼֵֻ���������֣�", vbInformation, gstrSysName: Me.txt��ʼֵ.SetFocus: Exit Sub
'    If Me.cbo����.Text = "2-����" And IsDate(Me.txt��ʼֵ) = False Then MsgBox "����Ϊ����ʱ��ʼֵֻ���������ڸ�ʽ��", vbInformation, gstrSysName: Me.txt��ʼֵ.SetFocus: Exit Sub
    
    gstrSql = Val(Me.txt����.Tag) & "," & _
            "'" & Trim(Me.txt����.Text) & "'," & _
            "'" & Trim(Me.txt������.Text) & "'," & _
            "'" & Trim(Me.txtӢ����.Text) & "'," & _
            Me.cbo����.ListIndex & "," & _
            IIf(Me.txt����.Enabled, Val(Me.txt����.Text), 0) & "," & _
            IIf(Me.txtС��.Enabled, Val(Me.txtС��.Text), 0) & "," & _
            "'" & Trim(Me.txt��λ.Text) & "'," & _
            "'" & Trim(Me.txt�ٴ�����.Text) & "'," & _
            Left(Me.cbo��ʾ��.Text, 1) & "," & _
            Me.cbo�Ա���.ListIndex & ","
    strTemp = ""
    If Me.txt��ֵ��(0).Enabled Then
        strTemp = Trim(Me.txt��ֵ��(0).Text) & ";" & Me.txt��ֵ��(1).Text
    End If
    If Me.msh��ֵ��.Active Then
        strTemp = ""
        For intCount = 1 To Me.msh��ֵ��.Rows - 1
            If Trim(Me.msh��ֵ��.TextMatrix(intCount, 1)) <> "" Then
                strTemp = strTemp & ";" & Trim(Me.msh��ֵ��.TextMatrix(intCount, 1))
            End If
        Next
        If strTemp <> "" Then strTemp = Mid(strTemp, 2)
'        If InStr(1, strTemp, Trim(Me.txt��ʼֵ.Text)) = 0 Then
'            MsgBox "��ʼֵû�а����ڿ�ѡ��ֵ�У�", vbInformation, gstrSysName
'            Me.msh��ֵ��.SetFocus: Exit Sub
'        End If
    End If
    gstrSql = gstrSql & "'" & strTemp & "',"
    gstrSql = gstrSql & "'" & Trim(Me.txt��ʼֵ.Text) & "'," & _
        Me.cbo���ֱ���.ListIndex + 1 & "," & _
        "'" & Trim(Me.txt��ֵ����.Text) & "'," & chkMust.Value & "," & chkDyn.Value
    '���ݱ���
    If Me.Tag = "����" Then
        lngItemID = zlDatabase.GetNextId("����������Ŀ")
        gstrSql = "ZL_������Ŀ_INSERT(" & lngItemID & "," & gstrSql & ")"
    Else
        gstrSql = "ZL_������Ŀ_UPDATE(" & lngItemID & "," & gstrSql & ")"
    End If
    
    Err = 0: On Error GoTo errHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    Unload Me
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd����_Click()
    With Me.tvwClass
        .Left = Me.txt����.Left
        .Top = Me.txt����.Top + Me.txt����.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
End Sub

Private Sub Form_Activate()
    
    '��ȡִ����Ŀ����Ϣ
    Err = 0: On Error GoTo errHand
    
    gstrSql = "select ID,����,������,Ӣ����,nvl(����,0) as ����,����,С��,С��,��λ," & _
            "        �ٴ�����,nvl(��ʾ��,0) as ��ʾ��,nvl(�Ա���,0) as �Ա���,��ֵ��,��ʼֵ,nvl(���ֱ���,1) as ���ֱ���,��ֵ����,����,��̬��" & _
            " from ����������Ŀ I" & _
            " where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    
    With rsTemp
        If .RecordCount > 0 Then
            Me.txt����.Text = !����
            Me.txt������.Text = IIf(IsNull(!������), "", !������)
            Me.txtӢ����.Text = IIf(IsNull(!Ӣ����), "", !Ӣ����)
            For intCount = 0 To Me.cbo����.ListCount - 1
                If Val(Left(Me.cbo����.List(intCount), 1)) = !���� Then
                    Me.cbo����.ListIndex = intCount: Exit For
                End If
            Next
            Me.txt����.Text = IIf(IsNull(!����), 0, !����)
            Me.txtС��.Text = IIf(IsNull(!С��), 0, !С��)
            Me.txt��λ.Text = IIf(IsNull(!��λ), "", !��λ)
            Me.txt�ٴ�����.Text = IIf(IsNull(!�ٴ�����), "", !�ٴ�����)
            
            For intCount = 0 To Me.cbo��ʾ��.ListCount - 1
                If Val(Left(Me.cbo��ʾ��.List(intCount), 1)) = !��ʾ�� Then
                    Me.cbo��ʾ��.ListIndex = intCount: Exit For
                End If
            Next
            Call zlSetGround
            For intCount = 0 To Me.cbo�Ա���.ListCount - 1
                If Val(Left(Me.cbo�Ա���.List(intCount), 1)) = !�Ա��� Then
                    Me.cbo�Ա���.ListIndex = intCount: Exit For
                End If
            Next
            aryTemp = Split(IIf(IsNull(!��ֵ��), "", !��ֵ��), ";")
            If Me.txt��ֵ��(0).Enabled And UBound(aryTemp) >= 0 Then
                Me.txt��ֵ��(0).Text = Val(aryTemp(0)): Me.txt��ֵ��(1).Text = 0
                If UBound(aryTemp) > 0 Then Me.txt��ֵ��(1).Text = Val(aryTemp(1))
                If Me.txt��ֵ��(0).Text = 0 Then Me.txt��ֵ��(0).Text = ""
                If Me.txt��ֵ��(1).Text = 0 Then Me.txt��ֵ��(1).Text = ""
            End If
            If Me.msh��ֵ��.Active Then
                With Me.msh��ֵ��
                    .ClearBill
                    .Rows = UBound(aryTemp) + 2
                    For intCount = 0 To UBound(aryTemp)
                        .TextMatrix(intCount + 1, 0) = intCount + 1
                        .TextMatrix(intCount + 1, 1) = aryTemp(intCount)
                    Next
                End With
            End If
            Me.txt��ʼֵ.Text = IIf(IsNull(!��ʼֵ), "", !��ʼֵ)
            For intCount = 0 To Me.cbo���ֱ���.ListCount - 1
                If Val(Left(Me.cbo���ֱ���.List(intCount), 1)) = !���ֱ��� Then
                    Me.cbo���ֱ���.ListIndex = intCount: Exit For
                End If
            Next
            Me.txt��ֵ����.Text = IIf(IsNull(!��ֵ����), "", !��ֵ����)
            chkMust.Value = !����
            chkDyn.Value = Nvl(!��̬��, 0)
        End If
        
        If Me.Tag = "����" Then
            lngItemID = 0

            gstrSql = "select nvl(max(I.����),'00000000') as ����" & _
                    " From ����������Ŀ I,������������ C" & _
                    " Where I.����ID=C.ID and C.����=(select ���� from ������������ where ID=[1])"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngClassId)
            
            If rsTemp.BOF = False Then
                Me.txt����.Text = Right(String(8, "0") & Val(rsTemp!����) + 1, Len(rsTemp!����))
            End If
            '���������Ϣ
            Me.txt������.Text = "": Me.txtӢ����.Text = "": Me.txt�ٴ�����.Text = ""
        End If
        If Me.Tag = "����" Then
            Me.fraBase.Enabled = False: Me.fraScope.Enabled = False: Me.fraWord.Enabled = False
            Me.cmd����.Enabled = False: Me.cmdOK.Visible = False
            Me.cmdCancel.Caption = "�ر�(&C)"
        End If
    End With
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.tvwClass.Visible Then
        Me.tvwClass.Visible = False: Me.txt����.SetFocus: Exit Sub
    End If
    Call cmdCancel_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    With Me.msh��ֵ��
        .Active = True
        .MsfObj.FixedCols = 1: .Rows = 2: .Cols = 2
        .MsfObj.AllowUserResizing = flexResizeNone
        .MsfObj.ScrollBars = flexScrollBarVertical
        .MsfObj.MergeCells = flexMergeFree
        .TextMatrix(0, 0) = "��ѡ��ֵ" & Space(30)
        .TextMatrix(0, 1) = "��ѡ��ֵ" & Space(30)
        .MsfObj.MergeRow(0) = True
        
        .ColData(0) = 5: .ColAlignment(0) = 1
        .ColData(1) = 4: .ColAlignment(1) = 1
        .ColWidth(0) = 250: .ColWidth(1) = .Width - .ColWidth(0) - 30
        .PrimaryCol = 1: .LocateCol = 1
        .Row = 1: .Col = 1
    End With
End Sub

Private Sub msh��ֵ��_AfterAddRow(Row As Long)
    With Me.msh��ֵ��
        For intCount = Row To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msh��ֵ��_AfterDeleteRow()
    With Me.msh��ֵ��
        For intCount = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msh��ֵ��_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If Me.txt��ʼֵ.Text = Me.msh��ֵ��.TextMatrix(Row, 1) Then
        Me.txt��ʼֵ.Text = ""
    End If
End Sub

Private Sub msh��ֵ��_DblClick(Cancel As Boolean)
    If Me.msh��ֵ��.Active Then
        Me.txt��ʼֵ.Text = Me.msh��ֵ��.TextMatrix(Me.msh��ֵ��.Row, 1): Cancel = True
    End If
End Sub

Private Sub msh��ֵ��_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.msh��ֵ��.Active = False Then Exit Sub
    With Me.msh��ֵ��
        If .Col <> 1 Then Exit Sub
        If .TxtVisible = False Then
            If .TextMatrix(.Row, 1) = "" Then
                If .Row = 1 Then Call zlCommFun.PressKey(vbKeyTab)
                Exit Sub
            End If
            strTemp = UCase(Trim(.TextMatrix(.Row, .Col)))
        Else
            If Trim(.Text) = "" Then
                If .Row = 1 Then .SetFocus: Call zlCommFun.PressKey(vbKeyTab)
                Exit Sub
            End If
            strTemp = UCase(Trim(.Text))
        End If
        Select Case Left(Me.cbo����.Text, 1)
        Case 0  '��ֵ
            If strTemp <> "0" And Val(strTemp) = 0 Then
                MsgBox "�������ݲ�����ֵ�ͣ�", vbExclamation, gstrSysName: Cancel = True: Exit Sub
            End If
        Case 1  '����
            strTemp = Replace(strTemp, "%", "")
            strTemp = Replace(strTemp, "&", "")
            strTemp = Replace(strTemp, ";", "")
            strTemp = Replace(strTemp, "'", "")
        Case 2  '����
            Err = 0: On Error Resume Next
            strTemp = CDate(strTemp)
            If Err <> 0 Then
                Err = 0
                MsgBox "�������ݲ������ڸ�ʽ��", vbExclamation, gstrSysName: Cancel = True: Exit Sub
            End If
        End Select
        .TextMatrix(.Row, 1) = strTemp
    End With
End Sub

Private Sub tvwClass_DblClick()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Me.txt����.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    Me.txt����.Text = Me.tvwClass.SelectedItem.Text
    Me.txt����.SetFocus
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
    If Me.cmd���� Is ActiveControl Then Exit Sub
    Me.tvwClass.Visible = False
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt��ʼֵ_GotFocus()
    Me.txt��ʼֵ.SelStart = 0: Me.txt��ʼֵ.SelLength = 100
End Sub

Private Sub txt��ʼֵ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt��λ_GotFocus()
    Me.txt��λ.SelStart = 0: Me.txt��λ.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��λ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt��λ_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt��ֵ����_GotFocus()
    Me.txt��ֵ����.SelStart = 0: Me.txt��ֵ����.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��ֵ����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt��ֵ����_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt�ٴ�����_GotFocus()
    Me.txt�ٴ�����.SelStart = 0: Me.txt�ٴ�����.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt�ٴ�����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("%&_|'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt�ٴ�����_LostFocus()
    Me.txt�ٴ�����.Text = Replace(Me.txt�ٴ�����, Chr(vbKeyReturn), "")
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt��ֵ��_GotFocus(Index As Integer)
    Me.txt��ֵ��(Index).SelStart = 0: Me.txt��ֵ��(Index).SelLength = 100
End Sub

Private Sub txt��ֵ��_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtС��_GotFocus()
    Me.txtС��.SelStart = 0: Me.txtС��.SelLength = 100
End Sub

Private Sub txtС��_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtӢ����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("&'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt������_GotFocus()
    Me.txt������.SelStart = 0: Me.txt������.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt������_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub zlSetGround()
    '----------------------------------------
    '���ܣ����ݱ�ʾ������Ŀ����ȷ����ֵ��Χ���뷽ʽ
    '----------------------------------------
    '0-�ı�;1-����;2-����;3-��ѡ;4-��ѡ
    
    cmdMove(0).Enabled = False
    cmdMove(1).Enabled = False
    Select Case Left(Me.cbo��ʾ��.Text, 1)
    Case 0
        '����Ϊ��ֵ���ı�������
        Select Case Left(Me.cbo����.Text, 1)
        Case 0  '��ֵ
            Me.txt��ֵ��(0).Enabled = True: Me.txt��ֵ��(1).Enabled = True
        Case 1  '����
            Me.txt��ֵ��(0).Enabled = False: Me.txt��ֵ��(1).Enabled = False
        Case 2  '����
            Me.txt��ֵ��(0).Enabled = True: Me.txt��ֵ��(1).Enabled = True
        End Select
        Me.msh��ֵ��.Active = False
        chkDyn.Enabled = False: chkDyn.Value = vbUnchecked
    Case 1
        'ֻ��������ֵ����
        Me.txt��ֵ��(0).Enabled = True: Me.txt��ֵ��(1).Enabled = True
        Me.msh��ֵ��.Active = False
        chkDyn.Enabled = False: chkDyn.Value = vbUnchecked
    Case 2
        '����Ϊ��ֵ���ı������ڣ�������ʾ����
        Me.txt��ֵ��(0).Enabled = False: Me.txt��ֵ��(1).Enabled = False
        Me.msh��ֵ��.Active = True
        cmdMove(0).Enabled = True
        cmdMove(1).Enabled = True
        chkDyn.Enabled = False: chkDyn.Value = vbUnchecked
    Case 3
        Me.txt��ֵ��(0).Enabled = False: Me.txt��ֵ��(1).Enabled = False
        '����Ϊ�ı����߼�
        Select Case Left(Me.cbo����.Text, 1)
        Case 1  '����
            Me.msh��ֵ��.Active = True
        Case 2  '�߼�
            Me.msh��ֵ��.Active = False
        End Select
        cmdMove(0).Enabled = True
        cmdMove(1).Enabled = True
        chkDyn.Enabled = True
    Case 4
        '����Ϊ��ֵ���ı���������ʾ����
        Me.txt��ֵ��(0).Enabled = False: Me.txt��ֵ��(1).Enabled = False
        Me.msh��ֵ��.Active = True
        cmdMove(0).Enabled = True
        cmdMove(1).Enabled = True
        chkDyn.Enabled = True
    End Select
    
    Me.txt��ʼֵ.Text = ""
    If Me.txt��ֵ��(0).Enabled = True Then
        Me.txt��ֵ��(0).BackColor = &H80000005
        Me.txt��ֵ��(1).BackColor = &H80000005
    Else
        Me.txt��ֵ��(0).BackColor = &H8000000F
        Me.txt��ֵ��(1).BackColor = &H8000000F
    End If
    
    If Me.msh��ֵ��.Active Then
'        Me.msh��ֵ��.ToolTipText = "˫�����ó�ʼ��ֵ"
        Call Me.msh��ֵ��.SetColColor(1, &H80000005)
        Me.msh��ֵ��.BackColorBkg = &H80000005
        Me.txt��ʼֵ.Enabled = False
        Me.txt��ʼֵ.BackColor = &H8000000F
    Else
        Me.msh��ֵ��.ToolTipText = ""
        Call Me.msh��ֵ��.SetColColor(1, &H8000000F)
        Me.msh��ֵ��.BackColorBkg = &H8000000F
        Me.txt��ʼֵ.Enabled = False
        Me.txt��ʼֵ.BackColor = &H80000005
    End If
    
End Sub

