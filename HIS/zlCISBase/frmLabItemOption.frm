VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmLabItemOption 
   BorderStyle     =   0  'None
   Caption         =   "������Ŀִ��ѡ��"
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraLine 
      Height          =   15
      Left            =   -30
      TabIndex        =   34
      Top             =   1785
      Width           =   7485
   End
   Begin VB.PictureBox picEdit 
      BorderStyle     =   0  'None
      Height          =   1995
      Left            =   3825
      ScaleHeight     =   1995
      ScaleWidth      =   4125
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1860
      Width           =   4125
      Begin VB.TextBox txt�����ʱ 
         Height          =   300
         Left            =   1155
         MaxLength       =   2
         TabIndex        =   45
         Top             =   630
         Width           =   615
      End
      Begin VB.TextBox txt�ͼ�ʱ�� 
         Height          =   300
         Left            =   1140
         MaxLength       =   2
         TabIndex        =   36
         Top             =   975
         Width           =   615
      End
      Begin VB.TextBox txt�������� 
         Height          =   300
         Left            =   1155
         MaxLength       =   3
         TabIndex        =   20
         Top             =   0
         Width           =   615
      End
      Begin VB.TextBox txt��ʱ��׼ 
         Height          =   300
         Left            =   1155
         MaxLength       =   2
         TabIndex        =   22
         Top             =   315
         Width           =   615
      End
      Begin VB.ComboBox cbo��ʱ��λ 
         Height          =   300
         Left            =   1845
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   315
         Width           =   825
      End
      Begin VB.TextBox txt����ص� 
         Height          =   300
         Left            =   645
         MaxLength       =   50
         TabIndex        =   25
         Top             =   1305
         Width           =   3000
      End
      Begin VB.TextBox txt����˵�� 
         Height          =   300
         Left            =   645
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   1635
         Width           =   3000
      End
      Begin VB.Label lbl�����ʱ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����걾        ���ֺ��ȡ����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   375
         TabIndex        =   46
         Top             =   690
         Width           =   2700
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ͼ�걾����        ���Ӻ�ܾ�����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   30
         TabIndex        =   35
         Top             =   1035
         Width           =   3060
      End
      Begin VB.Label lbl�������� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ĭ�ϸ�������        ,���˶Ա���ʷ���"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   30
         TabIndex        =   19
         Top             =   45
         Width           =   3330
      End
      Begin VB.Label lblִ��ʱ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ŀִ��ʱ��                  ���ȡ����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   30
         TabIndex        =   21
         Top             =   360
         Width           =   3600
      End
      Begin VB.Label lbl����ص� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ص�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   24
         Top             =   1365
         Width           =   360
      End
      Begin VB.Label lbl����˵�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "˵��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   26
         Top             =   1665
         Width           =   360
      End
   End
   Begin VB.Frame fraAppTo 
      Height          =   510
      Left            =   150
      TabIndex        =   28
      Top             =   3855
      Width           =   7710
      Begin VB.OptionButton optApplyTo 
         Caption         =   "���м�����Ŀ"
         Height          =   180
         Index           =   2
         Left            =   6045
         TabIndex        =   32
         Top             =   195
         Width           =   1395
      End
      Begin VB.OptionButton optApplyTo 
         Caption         =   "����""�ټ�""����Ŀ"
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   31
         Top             =   195
         Width           =   2670
      End
      Begin VB.OptionButton optApplyTo 
         Caption         =   "������Ŀ"
         Height          =   180
         Index           =   0
         Left            =   2250
         TabIndex        =   30
         Top             =   195
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.Label lblApplyTo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ������ͬʱӦ���ڣ�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   75
         TabIndex        =   29
         Top             =   195
         Width           =   2160
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfg�ɼ���ʽ 
      Height          =   1530
      Left            =   120
      TabIndex        =   17
      Top             =   2085
      Width           =   3615
      _cx             =   6376
      _cy             =   2699
      Appearance      =   0
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
      BackColorFixed  =   15790320
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
      Rows            =   3
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
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
      WordWrap        =   0   'False
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
   Begin VB.PictureBox picDept 
      BorderStyle     =   0  'None
      Height          =   1650
      Left            =   105
      ScaleHeight     =   1650
      ScaleWidth      =   7875
      TabIndex        =   33
      Top             =   45
      Width           =   7875
      Begin VB.CheckBox chk������� 
         Caption         =   "��������첡��(&I)"
         Height          =   225
         Index           =   2
         Left            =   5280
         TabIndex        =   41
         Top             =   0
         Width           =   1890
      End
      Begin VB.ComboBox cboִ�п��� 
         Height          =   300
         Index           =   2
         ItemData        =   "frmLabItemOption.frx":0000
         Left            =   6045
         List            =   "frmLabItemOption.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   660
         Width           =   1800
      End
      Begin VB.ComboBox cbo���Ƶ��� 
         Height          =   300
         Index           =   2
         ItemData        =   "frmLabItemOption.frx":0004
         Left            =   6045
         List            =   "frmLabItemOption.frx":0006
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   285
         Width           =   1800
      End
      Begin VB.ComboBox cboĬ������ 
         Height          =   300
         Index           =   2
         ItemData        =   "frmLabItemOption.frx":0008
         Left            =   6045
         List            =   "frmLabItemOption.frx":000A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1035
         Width           =   1800
      End
      Begin VB.CheckBox chk�����ֽ� 
         Caption         =   "�����ָ��Ĭ����������"
         Height          =   195
         Index           =   2
         Left            =   5535
         TabIndex        =   37
         Top             =   1425
         Width           =   2475
      End
      Begin VB.CheckBox chk�����ֽ� 
         Caption         =   "�����ָ��Ĭ����������"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   15
         Top             =   1425
         Width           =   2295
      End
      Begin VB.CheckBox chk�����ֽ� 
         Caption         =   "�����ָ��Ĭ����������"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   7
         Top             =   1425
         Width           =   2550
      End
      Begin VB.ComboBox cboĬ������ 
         Height          =   300
         Index           =   1
         ItemData        =   "frmLabItemOption.frx":000C
         Left            =   3390
         List            =   "frmLabItemOption.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1035
         Width           =   1800
      End
      Begin VB.ComboBox cboĬ������ 
         Height          =   300
         Index           =   0
         ItemData        =   "frmLabItemOption.frx":0010
         Left            =   780
         List            =   "frmLabItemOption.frx":0012
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1035
         Width           =   1800
      End
      Begin VB.ComboBox cbo���Ƶ��� 
         Height          =   300
         Index           =   1
         ItemData        =   "frmLabItemOption.frx":0014
         Left            =   3390
         List            =   "frmLabItemOption.frx":0016
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   285
         Width           =   1800
      End
      Begin VB.ComboBox cboִ�п��� 
         Height          =   300
         Index           =   1
         ItemData        =   "frmLabItemOption.frx":0018
         Left            =   3390
         List            =   "frmLabItemOption.frx":001F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   660
         Width           =   1800
      End
      Begin VB.ComboBox cbo���Ƶ��� 
         Height          =   300
         Index           =   0
         ItemData        =   "frmLabItemOption.frx":0030
         Left            =   795
         List            =   "frmLabItemOption.frx":0037
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   285
         Width           =   1800
      End
      Begin VB.ComboBox cboִ�п��� 
         Height          =   300
         Index           =   0
         ItemData        =   "frmLabItemOption.frx":0048
         Left            =   780
         List            =   "frmLabItemOption.frx":004F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   660
         Width           =   1800
      End
      Begin VB.CheckBox chk������� 
         Caption         =   "���������ﲡ��(&O)"
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Value           =   1  'Checked
         Width           =   1950
      End
      Begin VB.CheckBox chk������� 
         Caption         =   "������סԺ����(&I)"
         Height          =   225
         Index           =   1
         Left            =   2625
         TabIndex        =   8
         Top             =   0
         Value           =   1  'Checked
         Width           =   1890
      End
      Begin VB.Label lbl���Ƶ��� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���Ƶ���"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   5295
         TabIndex        =   43
         Top             =   345
         Width           =   705
      End
      Begin VB.Label lblִ�п��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ִ�п���"
         Height          =   180
         Index           =   2
         Left            =   5295
         TabIndex        =   44
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lblĬ������ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ĭ������"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   5295
         TabIndex        =   42
         Top             =   1095
         Width           =   705
      End
      Begin VB.Label lblĬ������ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ĭ������"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   2610
         TabIndex        =   13
         Top             =   1095
         Width           =   735
      End
      Begin VB.Label lblĬ������ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ĭ������"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   30
         TabIndex        =   5
         Top             =   1095
         Width           =   720
      End
      Begin VB.Label lbl���Ƶ��� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���Ƶ���"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   2610
         TabIndex        =   9
         Top             =   345
         Width           =   750
      End
      Begin VB.Label lbl���Ƶ��� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���Ƶ���"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   30
         TabIndex        =   1
         Top             =   345
         Width           =   720
      End
      Begin VB.Label lblִ�п��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ִ�п���"
         Height          =   180
         Index           =   0
         Left            =   30
         TabIndex        =   3
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lblִ�п��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ִ�п���"
         Height          =   180
         Index           =   1
         Left            =   2610
         TabIndex        =   11
         Top             =   720
         Width           =   780
      End
   End
   Begin VB.Label lbl�ɼ���ʽ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�걾�ɼ���ʽ(&G)"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   105
      TabIndex        =   16
      Top             =   1845
      Width           =   1350
   End
End
Attribute VB_Name = "frmLabItemOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngItemID As Long          '��ǰ��ʾ����Ŀid
Private mint��� As Integer

Private Enum mCol
    ID = 0: ��־: ����: ����: ����ʱ��
End Enum

'��ʱ����
Dim lngCount As Long
Dim strTemp As String, aryTemp() As String

'--------------------------------------------
'����Ϊ���幫������
'--------------------------------------------
Public Function zlRefresh(lngItemId As Long) As Boolean
    '���ܣ�������Ŀidˢ�µ�ǰ��ʾ����
    '��������ǰ��Ŀid
    Dim rsTemp As New ADODB.Recordset
    Dim j As Integer
    mlngItemID = lngItemId: mint��� = 0
    
    '�����ǰ��Ŀ����ʾ
    Me.chk�������(0).Value = 0: Me.cbo���Ƶ���(0).ListIndex = -1: Me.cboִ�п���(0).ListIndex = -1
    Me.cboĬ������(0).ListIndex = -1: Me.chk�����ֽ�(0).Value = vbUnchecked: Me.chk�����ֽ�(0).Enabled = False
    Me.chk�������(1).Value = 0: Me.cbo���Ƶ���(1).ListIndex = -1: Me.cboִ�п���(1).ListIndex = -1
    Me.cboĬ������(1).ListIndex = -1: Me.chk�����ֽ�(1).Value = vbUnchecked: Me.chk�����ֽ�(1).Enabled = False
    Me.chk�������(2).Value = 0: Me.cbo���Ƶ���(2).ListIndex = -1: Me.cboִ�п���(2).ListIndex = -1
    Me.cboĬ������(2).ListIndex = -1: Me.chk�����ֽ�(2).Value = vbUnchecked: Me.chk�����ֽ�(2).Enabled = False
    Me.txt��������.Text = "": Me.txt��ʱ��׼.Text = "": Me.cbo��ʱ��λ.ListIndex = 0
    Me.txt����ص�.Text = "": Me.txt����˵��.Text = "": Me.txt�ͼ�ʱ��.Text = ""
    Me.txt�����ʱ.Text = ""
    
    With Me.vfg�ɼ���ʽ
        For lngCount = .FixedRows To .Rows - 1
            .Row = lngCount: .Col = mCol.����
            .CellChecked = flexUnchecked
        Next
    End With
    If lngItemId = 0 Then zlRefresh = True: Exit Function
    
    '��ȡָ����Ŀ����Ϣ
    Err = 0: On Error GoTo ErrHand
    'װ����ͬ�������͵�����
    gstrSql = "Select M.ID, M.����, M.���� From �������� M, ������ĿĿ¼ I Where M.�������� = I.�������� And I.ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
    With rsTemp
        Me.cboĬ������(0).Clear: Me.cboĬ������(1).Clear: Me.cboĬ������(2).Clear
        Do While Not .EOF
            Me.cboĬ������(0).AddItem !���� & "-" & !����: Me.cboĬ������(0).ItemData(Me.cboĬ������(0).NewIndex) = !ID
            Me.cboĬ������(1).AddItem !���� & "-" & !����: Me.cboĬ������(1).ItemData(Me.cboĬ������(1).NewIndex) = !ID
            Me.cboĬ������(2).AddItem !���� & "-" & !����: Me.cboĬ������(2).ItemData(Me.cboĬ������(2).NewIndex) = !ID
            .MoveNext
        Loop
    End With
    
    gstrSql = "Select I.�������, I.��������, I.�����Ŀ, O.��������id, O.���������ֽ�, O.סԺ����id, O.סԺ�����ֽ�, O.��������," & vbNewLine & _
            "       O.��ʱ��׼, O.��ʱ��λ, O.ȡ����ص�, O.����˵��,O.�ͼ�ʱ��,O.�������id,O.��������ֽ�,O.�����ʱ " & vbNewLine & _
            "From ������ĿĿ¼ I, ������Ŀѡ�� O" & vbNewLine & _
            "Where I.ID = O.������Ŀid(+) And I.ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
    With rsTemp
        If .RecordCount > 0 Then
            If Val("" & !�����Ŀ) = 1 Then mint��� = 1: Me.chk�����ֽ�(0).Enabled = True: Me.chk�����ֽ�(1).Enabled = True
            
            If Val("" & !�������) = 4 Then
                Me.chk�������(2).Value = vbChecked: Me.chk�������(0).Value = vbUnchecked: Me.chk�������(1).Value = vbUnchecked
            ElseIf Val("" & !�������) = 3 Then
                Me.chk�������(0).Value = vbChecked: Me.chk�������(1).Value = vbChecked
            ElseIf Val("" & !�������) = 1 Then
                Me.chk�������(0).Value = vbChecked
            ElseIf Val("" & !�������) = 2 Then
                Me.chk�������(1).Value = vbChecked
            End If
            Me.optApplyTo(1).Caption = "����""" & !�������� & """����Ŀ"
            
            For lngCount = 0 To Me.cboĬ������(0).ListCount - 1
                If Me.cboĬ������(0).ItemData(lngCount) = Val("" & !��������id) Then Me.cboĬ������(0).ListIndex = lngCount: Exit For
            Next
            Me.chk�����ֽ�(0).Value = IIf(Val("" & !���������ֽ�) = 0, vbUnchecked, vbChecked)
            
            For lngCount = 0 To Me.cboĬ������(1).ListCount - 1
                If Me.cboĬ������(1).ItemData(lngCount) = Val("" & !סԺ����id) Then Me.cboĬ������(1).ListIndex = lngCount: Exit For
            Next
            Me.chk�����ֽ�(1).Value = IIf(Val("" & !סԺ�����ֽ�) = 0, vbUnchecked, vbChecked)
            
            For lngCount = 0 To Me.cboĬ������(2).ListCount - 1
                If Me.cboĬ������(2).ItemData(lngCount) = Val("" & !�������id) Then Me.cboĬ������(2).ListIndex = lngCount: Exit For
            Next
            Me.chk�����ֽ�(2).Value = IIf(Val("" & !��������ֽ�) = 0, vbUnchecked, vbChecked)
            
            Me.txt��������.Text = "" & !��������
            Me.txt��ʱ��׼.Text = "" & !��ʱ��׼
            For lngCount = 0 To Me.cbo��ʱ��λ.ListCount - 1
                If Me.cbo��ʱ��λ.List(lngCount) = !��ʱ��λ Then Me.cbo��ʱ��λ.ListIndex = lngCount: Exit For
            Next
            Me.txt�����ʱ.Text = "" & !�����ʱ
            
            Me.txt����ص�.Text = "" & !ȡ����ص�
            Me.txt����˵��.Text = "" & !����˵��
            Me.txt�ͼ�ʱ��.Text = IIf(Val(Nvl(!�ͼ�ʱ��)) = 0, "", Nvl(!�ͼ�ʱ��))
        End If
    End With
    
    gstrSql = "Select �÷�id From �����÷����� Where ��Ŀid = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
    Do While Not rsTemp.EOF
        With Me.vfg�ɼ���ʽ
            For lngCount = .FixedRows To .Rows - 1
                .Row = lngCount: .Col = mCol.����
                If Val(.TextMatrix(lngCount, mCol.����ʱ��)) = 0 Then
                    If Val(.TextMatrix(lngCount, mCol.ID)) = Val("" & rsTemp!�÷�ID) Then .CellChecked = flexChecked
                ElseIf Format(.TextMatrix(lngCount, mCol.����ʱ��), "yyyy-mm-dd") = "3000-01-01" Then
                    If Val(.TextMatrix(lngCount, mCol.ID)) = Val("" & rsTemp!�÷�ID) Then .CellChecked = flexChecked
                Else
                    If Val(.TextMatrix(lngCount, mCol.ID)) = Val("" & rsTemp!�÷�ID) Then .CellChecked = flexTSGrayed
                End If
            Next
        End With
        rsTemp.MoveNext
    Loop
    
    gstrSql = "Select ������Դ, ִ�п���id From ����ִ�п��� Where ������Ŀid = [1] And ��������id Is Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
    Do While Not rsTemp.EOF
        If Val("" & rsTemp!������Դ) = 4 Then
            For lngCount = 0 To Me.cboִ�п���(2).ListCount - 1
                If Me.cboִ�п���(2).ItemData(lngCount) = rsTemp!ִ�п���ID Then Me.cboִ�п���(2).ListIndex = lngCount: Exit For
            Next
        ElseIf Val("" & rsTemp!������Դ) = 1 Then
            For lngCount = 0 To Me.cboִ�п���(0).ListCount - 1
                If Me.cboִ�п���(0).ItemData(lngCount) = rsTemp!ִ�п���ID Then Me.cboִ�п���(0).ListIndex = lngCount: Exit For
            Next
        ElseIf Val("" & rsTemp!������Դ) = 2 Then
            For lngCount = 0 To Me.cboִ�п���(1).ListCount - 1
                If Me.cboִ�п���(1).ItemData(lngCount) = rsTemp!ִ�п���ID Then Me.cboִ�п���(1).ListIndex = lngCount: Exit For
            Next
        End If
        rsTemp.MoveNext
    Loop
    
    gstrSql = "Select Ӧ�ó���, �����ļ�id From ��������Ӧ�� Where ������Ŀid = [1] And Ӧ�ó��� In (1, 2, 4)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
    Do While Not rsTemp.EOF
        If Val("" & rsTemp!Ӧ�ó���) = 1 Then
            For lngCount = 0 To Me.cbo���Ƶ���(0).ListCount - 1
                If Me.cbo���Ƶ���(0).ItemData(lngCount) = rsTemp!�����ļ�id Then Me.cbo���Ƶ���(0).ListIndex = lngCount: Exit For
            Next
        ElseIf Val("" & rsTemp!Ӧ�ó���) = 2 Then
            For lngCount = 0 To Me.cbo���Ƶ���(1).ListCount - 1
                If Me.cbo���Ƶ���(1).ItemData(lngCount) = rsTemp!�����ļ�id Then Me.cbo���Ƶ���(1).ListIndex = lngCount: Exit For
            Next
        ElseIf Val("" & rsTemp!Ӧ�ó���) = 4 Then
            For lngCount = 0 To Me.cbo���Ƶ���(2).ListCount - 1
                If Me.cbo���Ƶ���(2).ItemData(lngCount) = rsTemp!�����ļ�id Then Me.cbo���Ƶ���(2).ListIndex = lngCount: Exit For
            Next
        End If
        rsTemp.MoveNext
    Loop
    
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart() As Boolean
    '���ܣ���ʼ��Ŀ�༭
    '������ lngItemId-ָ���༭����Ŀ
    Me.Tag = "�༭": Call Form_Resize
    zlEditStart = True: Exit Function
End Function

Public Sub zlEditCancel()
    '���ܣ��������ڽ��еı༭
    Me.Tag = "": Call Form_Resize
    Call Me.zlRefresh(mlngItemID)
End Sub

Public Function zlEditSave() As Long
    '���ܣ��������ڽ��еı༭,���������ڱ༭��Ŀid,����ʧ�ܷ���0
    Dim strList As String
    Dim strDept0 As String, strDept1 As String, strDept2 As String
    Dim strBill0 As String, strBill1 As String, strBill2 As String
    Dim strApt0 As String, strApt1 As String, strApt2 As String
    Dim strAllot0 As String, strAllot1 As String
    
    If Me.cbo��ʱ��λ.ListIndex = -1 Then Me.cbo��ʱ��λ.ListIndex = 0
    '���ݱ��������֯
    strList = ""
    With Me.vfg�ɼ���ʽ
        For lngCount = .FixedRows To .Rows - 1
            .Row = lngCount: .Col = mCol.����
            If .CellChecked = flexChecked Then strList = strList & "," & .TextMatrix(lngCount, mCol.ID)
            If .CellChecked = flexTSGrayed Then strList = strList & "," & .TextMatrix(lngCount, mCol.ID)
        Next
    End With
    If strList <> "" Then strList = Mid(strList, 2)

    If Me.cboִ�п���(0).ListIndex = -1 Then
        strDept0 = "Null"
    Else
        strDept0 = Me.cboִ�п���(0).ItemData(Me.cboִ�п���(0).ListIndex)
    End If
    If Me.cboִ�п���(1).ListIndex = -1 Then
        strDept1 = "Null"
    Else
        strDept1 = Me.cboִ�п���(1).ItemData(Me.cboִ�п���(1).ListIndex)
    End If
    If Me.cbo���Ƶ���(0).ListIndex = -1 Then
        strBill0 = "Null"
    Else
        strBill0 = Me.cbo���Ƶ���(0).ItemData(Me.cbo���Ƶ���(0).ListIndex)
    End If
    If Me.cbo���Ƶ���(1).ListIndex = -1 Then
        strBill1 = "Null"
    Else
        strBill1 = Me.cbo���Ƶ���(1).ItemData(Me.cbo���Ƶ���(1).ListIndex)
    End If
    If Me.cboĬ������(0).ListIndex = -1 Then
        strApt0 = "Null"
    Else
        strApt0 = Me.cboĬ������(0).ItemData(Me.cboĬ������(0).ListIndex)
    End If
    If Me.cboĬ������(1).ListIndex = -1 Then
        strApt1 = "Null"
    Else
        strApt1 = Me.cboĬ������(1).ItemData(Me.cboĬ������(1).ListIndex)
    End If
    strAllot0 = IIf(Me.chk�����ֽ�(0).Value = vbChecked, 1, 0)
    strAllot1 = IIf(Me.chk�����ֽ�(1).Value = vbChecked, 1, 0)
    
    If Me.chk�������(2) = vbChecked Then
        gstrSql = mlngItemID & ",'" & strList & "',4," & strDept0 & "," & strDept1 & "," & strBill0 & "," & strBill1
        gstrSql = gstrSql & "," & strApt0 & "," & strAllot0 & "," & strApt1 & "," & strAllot1
    ElseIf Me.chk�������(0).Value = vbChecked And Me.chk�������(1).Value = vbChecked Then
        gstrSql = mlngItemID & ",'" & strList & "',3," & strDept0 & "," & strDept1 & "," & strBill0 & "," & strBill1
        gstrSql = gstrSql & "," & strApt0 & "," & strAllot0 & "," & strApt1 & "," & strAllot1
    ElseIf Me.chk�������(0).Value <> vbChecked And Me.chk�������(1).Value = vbChecked Then
        gstrSql = mlngItemID & ",'" & strList & "',2,Null," & strDept1 & ",Null," & strBill1
        gstrSql = gstrSql & ",Null,0," & strApt1 & "," & strAllot1
    ElseIf Me.chk�������(0).Value = vbChecked And Me.chk�������(1).Value <> vbChecked Then
        gstrSql = mlngItemID & ",'" & strList & "',1," & strDept0 & ",Null," & strBill0 & ",Null"
        gstrSql = gstrSql & "," & strApt0 & "," & strAllot0 & ",Null,0"
    Else
        gstrSql = mlngItemID & ",'" & strList & "',0,Null,Null,Null,Null,Null,Null,Null,Null"
    End If
    gstrSql = gstrSql & "," & Val(Me.txt��������.Text)
    gstrSql = gstrSql & "," & Val(Me.txt��ʱ��׼.Text)
    gstrSql = gstrSql & ",'" & Me.cbo��ʱ��λ.Text & "'"
    gstrSql = gstrSql & ",'" & Me.txt����ص�.Text & "'"
    gstrSql = gstrSql & ",'" & Me.txt����˵��.Text & "'"
    
    If optApplyTo.Item(0).Value = True Then
        gstrSql = gstrSql & "," & 0
    ElseIf optApplyTo.Item(1).Value = True Then
        gstrSql = gstrSql & "," & 1
    ElseIf optApplyTo.Item(2).Value = True Then
        gstrSql = gstrSql & "," & 2
    End If
    gstrSql = gstrSql & "," & IIf(Trim(Me.txt�ͼ�ʱ��) = "", "NULL", Val(Me.txt�ͼ�ʱ��))
    
    strDept2 = "Null"
    strBill2 = "Null"
    strApt2 = "Null"
    If Me.cboִ�п���(2).ListIndex <> -1 Then strDept2 = Me.cboִ�п���(2).ItemData(Me.cboִ�п���(2).ListIndex)
    If Me.cbo���Ƶ���(2).ListIndex <> -1 Then strBill2 = Me.cbo���Ƶ���(2).ItemData(Me.cbo���Ƶ���(2).ListIndex)
    If Me.cboĬ������(2).ListIndex <> -1 Then strApt2 = Me.cboĬ������(2).ItemData(Me.cboĬ������(2).ListIndex)
    gstrSql = gstrSql & "," & strDept2 & "," & strBill2 & "," & strApt2 & "," & IIf(Me.chk�����ֽ�(2).Value = vbChecked, 1, 0)
    
    gstrSql = gstrSql & "," & IIf(Trim(Me.txt�����ʱ) = "", "Null", Val(Me.txt�����ʱ))
    
    gstrSql = "Zl_������Ŀѡ��_Edit(" & gstrSql & ")"
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    Me.Tag = "": Call Form_Resize
    zlEditSave = mlngItemID: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function



'--------------------------------------------
'����Ϊ����ؼ���Ӧ�¼�
'--------------------------------------------
Private Sub cbo��ʱ��λ_Click()
    If Me.cbo��ʱ��λ.ListCount > 0 Then
        Me.lbl�����ʱ.Caption = "����걾        " & Me.cbo��ʱ��λ.List(Me.cbo��ʱ��λ.ListIndex) & "���ȡ����"
    End If
End Sub

Private Sub cbo��ʱ��λ_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo��ʱ��λ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboĬ������_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cboĬ������_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo���Ƶ���_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo���Ƶ���_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboִ�п���_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cboִ�п���_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk�������_Click(Index As Integer)


    
    If Me.chk�������(Index).Value Then
        Me.cboִ�п���(Index).Enabled = True: Me.cbo���Ƶ���(Index).Enabled = True
        Me.cboĬ������(Index).Enabled = True: Me.chk�����ֽ�(Index).Enabled = (mint��� = 1)
        
        If Index = 2 Then
            If Me.chk�������(0).Value <> 0 Then
                Me.chk�������(0).Value = 0
                Me.cboִ�п���(0).Enabled = False: Me.cbo���Ƶ���(0).Enabled = False
                Me.cboĬ������(0).Enabled = False: Me.chk�����ֽ�(0).Enabled = False
            End If
            If Me.chk�������(1).Value <> 0 Then
                Me.chk�������(1).Value = 0
                Me.cboִ�п���(1).Enabled = False: Me.cbo���Ƶ���(1).Enabled = False
                Me.cboĬ������(1).Enabled = False: Me.chk�����ֽ�(1).Enabled = False
            End If
        Else
            If Me.chk�������(2).Value <> 0 Then
                Me.chk�������(2).Value = 0
                Me.cboִ�п���(2).Enabled = False: Me.cbo���Ƶ���(2).Enabled = False
                Me.cboĬ������(2).Enabled = False: Me.chk�����ֽ�(2).Enabled = False
            End If
        End If
        
    Else
        Me.cboִ�п���(Index).Enabled = False: Me.cbo���Ƶ���(Index).Enabled = False
        Me.cboĬ������(Index).Enabled = False: Me.chk�����ֽ�(Index).Enabled = False
    End If
    Me.cboִ�п���(2).Enabled = True: Me.cbo���Ƶ���(2).Enabled = True
    
End Sub

Private Sub chk�������_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chk�������_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk�����ֽ�_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chk�����ֽ�_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    '��������װ��
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    Dim j As Integer
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select I.ID, 0 As ��־, I.����, I.����,i.����ʱ�� From ������ĿĿ¼ I Where I.��� = 'E' And I.�������� = '6' " & vbNewLine & _
              "  Order By I.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With Me.vfg�ɼ���ʽ
        .Redraw = flexRDNone
         Set .DataSource = rsTemp
        .TextMatrix(0, mCol.ID) = "ID": .TextMatrix(0, mCol.��־) = ""
        .TextMatrix(0, mCol.����) = "����": .TextMatrix(0, mCol.����) = "����"
        .TextMatrix(0, mCol.����ʱ��) = "����ʱ��"
        .ColWidth(mCol.ID) = 0: .ColWidth(mCol.��־) = 0: .ColWidth(mCol.����ʱ��) = 0
        .ColWidth(mCol.����) = 800: .ColWidth(mCol.����) = 2000
        For j = .FixedRows To .Rows - 1
            If Val(.TextMatrix(j, mCol.����ʱ��)) = 0 Or Format(.TextMatrix(j, mCol.����ʱ��), "yyyy-mm-dd") = "3000-01-01" Then
                
            Else
                For i = mCol.���� To .Cols - 1
                    .Cell(flexcpForeColor, j, i, j, i) = vbRed
                
                    .TextMatrix(j, i) = .TextMatrix(j, i) & "(��ͣ��)"
                Next
            End If
        Next
        For lngCount = .FixedCols To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        .Redraw = flexRDDirect
    End With
    
    aryTemp = Split("����;Сʱ;��", ";")
    Me.cbo��ʱ��λ.Clear
    For lngCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo��ʱ��λ.AddItem aryTemp(lngCount)
    Next
    If Me.cbo��ʱ��λ.ListCount > 0 Then Me.cbo��ʱ��λ.ListIndex = 0
    
    cboִ�п���(0).Clear: cboִ�п���(1).Clear: cboִ�п���(2).Clear
    cbo���Ƶ���(0).Clear: cbo���Ƶ���(1).Clear: cbo���Ƶ���(2).Clear
    Call zlControl.CboSetWidth(cbo���Ƶ���(0).hWnd, 2650)
    Call zlControl.CboSetWidth(cbo���Ƶ���(1).hWnd, 2650)
    Call zlControl.CboSetWidth(cbo���Ƶ���(2).hWnd, 2650)
    
    '--- ���� ���
    gstrSql = "Select ID, ����, ����" & vbNewLine & _
            "From ���ű� D, ��������˵�� P" & vbNewLine & _
            "Where D.ID = P.����id And P.�������� = '����' And P.������� In (1, 3) And" & vbNewLine & _
            "      (To_Char(D.����ʱ��, 'YYYY-MM-DD') = '3000-01-01' Or D.����ʱ�� Is Null)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Do While Not .EOF
            Me.cboִ�п���(0).AddItem !���� & "-" & !����
            Me.cboִ�п���(0).ItemData(Me.cboִ�п���(0).NewIndex) = !ID
            
            Me.cboִ�п���(2).AddItem !���� & "-" & !����
            Me.cboִ�п���(2).ItemData(Me.cboִ�п���(2).NewIndex) = !ID
            
            .MoveNext
        Loop
    End With
    gstrSql = "Select ID, ���, ���� From �����ļ��б� Where ���� = 7"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Do While Not .EOF
            Me.cbo���Ƶ���(0).AddItem !��� & "-" & !����
            Me.cbo���Ƶ���(0).ItemData(Me.cbo���Ƶ���(0).NewIndex) = !ID
            
            Me.cbo���Ƶ���(2).AddItem !��� & "-" & !����
            Me.cbo���Ƶ���(2).ItemData(Me.cbo���Ƶ���(2).NewIndex) = !ID
            .MoveNext
        Loop
    End With
    '--- סԺ
    gstrSql = "Select ID, ����, ����" & vbNewLine & _
            "From ���ű� D, ��������˵�� P" & vbNewLine & _
            "Where D.ID = P.����id And P.�������� = '����' And P.������� In (2, 3) And" & vbNewLine & _
            "      (To_Char(D.����ʱ��, 'YYYY-MM-DD') = '3000-01-01' Or D.����ʱ�� Is Null)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Do While Not .EOF
            Me.cboִ�п���(1).AddItem !���� & "-" & !����
            Me.cboִ�п���(1).ItemData(Me.cboִ�п���(1).NewIndex) = !ID
            .MoveNext
        Loop
    End With
    gstrSql = "Select ID, ���, ���� From �����ļ��б� Where ���� = 7"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Do While Not .EOF
            Me.cbo���Ƶ���(1).AddItem !��� & "-" & !����
            Me.cbo���Ƶ���(1).ItemData(Me.cbo���Ƶ���(1).NewIndex) = !ID
            .MoveNext
        Loop
    End With
        
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    Me.fraAppTo.Top = Me.ScaleHeight - Me.fraAppTo.Height - 180
    If Me.Tag = "�༭" Then
        Me.vfg�ɼ���ʽ.Height = Me.fraAppTo.Top - Me.vfg�ɼ���ʽ.Top
        Me.picEdit.Height = Me.fraAppTo.Top - Me.picEdit.Top
        Me.picEdit.Enabled = True: Me.picDept.Enabled = True
        Me.fraAppTo.Enabled = True: Me.fraAppTo.Visible = True
    Else
        Me.vfg�ɼ���ʽ.Height = Me.ScaleHeight - Me.vfg�ɼ���ʽ.Top - 180
        Me.picEdit.Height = Me.ScaleHeight - Me.picEdit.Top - 180
        Me.picEdit.Enabled = False: Me.picDept.Enabled = False
        Me.fraAppTo.Enabled = False: Me.fraAppTo.Visible = False
    End If
End Sub

Private Sub optApplyTo_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub optApplyTo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub picEdit_Resize()
    Err = 0: On Error Resume Next
    Me.txt����˵��.Height = Me.picEdit.ScaleHeight - Me.txt����˵��.Top
End Sub

Private Sub txt����ص�_GotFocus()
    Me.txt����ص�.SelStart = 0: Me.txt����ص�.SelLength = 1000
End Sub

Private Sub txt����ص�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt����˵��_GotFocus()
    Me.txt����˵��.SelStart = 0: Me.txt����˵��.SelLength = 1000
End Sub

Private Sub txt����˵��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt��������_GotFocus()
    Me.txt��������.SelStart = 0: Me.txt��������.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt��������_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt��ʱ��׼_GotFocus()
    Me.txt��ʱ��׼.SelStart = 0: Me.txt��ʱ��׼.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt��ʱ��׼_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt�ͼ�ʱ��_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt�����ʱ_GotFocus()
    Me.txt�����ʱ.SelStart = 0: Me.txt�����ʱ.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt�����ʱ_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub vfg�ɼ���ʽ_DblClick()
    If Me.vfg�ɼ���ʽ.MouseRow < Me.vfg�ɼ���ʽ.FixedRows Then Exit Sub
    If Me.Tag <> "�༭" Then Exit Sub
    With Me.vfg�ɼ���ʽ
        If .Row < .FixedRows And .Row > .Rows - 1 Then Exit Sub
        .Col = mCol.����
        If Val(.TextMatrix(.Row, mCol.����ʱ��)) = 0 Or Format(.TextMatrix(.Row, mCol.����ʱ��), "yyyy-mm-dd") = "3000-01-01" Then
                
        Else
            MsgBox "��ͣ�õĲɼ���ʽ�����ܹ�ѡ��", vbInformation, Me.Caption
            Exit Sub
        End If
        If .CellChecked = flexChecked Then
            .CellChecked = flexUnchecked
        Else
            .CellChecked = flexChecked
        End If
    End With
End Sub

Private Sub vfg�ɼ���ʽ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If KeyAscii <> vbKeySpace Then Exit Sub
    Call vfg�ɼ���ʽ_DblClick
End Sub
