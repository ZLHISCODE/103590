VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Begin VB.Form frmLabItemRef 
   BorderStyle     =   0  'None
   Caption         =   "������Ŀ�ο�ֵ"
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txt���쾯ʾ�� 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Left            =   7230
      MaxLength       =   12
      TabIndex        =   38
      Top             =   465
      Width           =   645
   End
   Begin VB.TextBox txt�ȶ�ʧ���� 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Left            =   3750
      MaxLength       =   12
      TabIndex        =   5
      Top             =   90
      Width           =   1020
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   1005
      Left            =   120
      TabIndex        =   7
      Top             =   825
      Width           =   7755
      _cx             =   13679
      _cy             =   1773
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
      Cols            =   6
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
   Begin VB.TextBox txt�ȶԾ�ʾ�� 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Left            =   1950
      MaxLength       =   12
      TabIndex        =   3
      Top             =   90
      Width           =   1020
   End
   Begin VB.TextBox txt���챨���� 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Left            =   5685
      MaxLength       =   12
      TabIndex        =   1
      Top             =   465
      Width           =   645
   End
   Begin VB.PictureBox picEdit 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2055
      Left            =   120
      ScaleHeight     =   2055
      ScaleWidth      =   7740
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1860
      Width           =   7740
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   900
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   975
         Width           =   2100
      End
      Begin VB.TextBox txt�������� 
         Height          =   300
         Left            =   4185
         TabIndex        =   32
         Top             =   975
         Width           =   1020
      End
      Begin VB.TextBox txt�������� 
         Height          =   300
         Left            =   5640
         TabIndex        =   33
         Top             =   975
         Width           =   1020
      End
      Begin VB.TextBox txt�������� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   4170
         MaxLength       =   12
         TabIndex        =   28
         Top             =   660
         Width           =   1020
      End
      Begin VB.TextBox txt�������� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   5625
         MaxLength       =   12
         TabIndex        =   29
         Top             =   660
         Width           =   1020
      End
      Begin VB.CheckBox chkĬ�� 
         Alignment       =   1  'Right Justify
         Caption         =   "Ĭ��"
         Height          =   180
         Left            =   6825
         TabIndex        =   30
         Top             =   720
         Width           =   840
      End
      Begin VB.ComboBox cbo�ο�ֵ 
         Height          =   300
         ItemData        =   "frmLabItemRef.frx":0000
         Left            =   3765
         List            =   "frmLabItemRef.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   345
         Width           =   1200
      End
      Begin VB.TextBox txt��ƫ���� 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   7005
         MaxLength       =   13
         TabIndex        =   24
         Top             =   345
         Width           =   750
      End
      Begin VB.TextBox txt��ע 
         Height          =   300
         Left            =   900
         MaxLength       =   50
         TabIndex        =   34
         Top             =   1290
         Width           =   6690
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   900
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   660
         Width           =   2100
      End
      Begin VB.ComboBox cbo�ٴ����� 
         Height          =   300
         Left            =   900
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   345
         Width           =   2100
      End
      Begin VB.TextBox txt�ο�ֵ 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   4905
         MaxLength       =   13
         TabIndex        =   21
         Top             =   345
         Width           =   900
      End
      Begin VB.TextBox txt�ο�ֵ 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   3780
         MaxLength       =   13
         TabIndex        =   20
         Top             =   345
         Width           =   900
      End
      Begin VB.ComboBox cbo��λ 
         Height          =   300
         Left            =   7020
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   30
         Width           =   750
      End
      Begin VB.TextBox txt���� 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   6495
         MaxLength       =   3
         TabIndex        =   15
         Top             =   30
         Width           =   495
      End
      Begin VB.ComboBox cbo�Ա� 
         Height          =   300
         Left            =   3765
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   30
         Width           =   1200
      End
      Begin VB.TextBox txt���� 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   5775
         MaxLength       =   3
         TabIndex        =   14
         Top             =   30
         Width           =   495
      End
      Begin VB.ComboBox cbo�걾 
         Height          =   300
         Left            =   900
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   30
         Width           =   2100
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ӧ����"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   75
         TabIndex        =   40
         Top             =   1035
         Width           =   720
      End
      Begin VB.Label lbl�������� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʾ�ο�              ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3345
         TabIndex        =   27
         Top             =   720
         Width           =   2160
      End
      Begin VB.Label lbl�������� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ο�              ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3345
         TabIndex        =   39
         Top             =   1035
         Width           =   2160
      End
      Begin VB.Label lbl��ƫ���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ƫ����"
         Height          =   180
         Left            =   6240
         TabIndex        =   23
         Top             =   405
         Width           =   720
      End
      Begin VB.Label lbl��ע 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ע"
         Height          =   180
         Left            =   420
         TabIndex        =   35
         Top             =   1350
         Width           =   360
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ӧ����"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   75
         TabIndex        =   25
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lbl�ٴ����� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ٴ�����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   75
         TabIndex        =   17
         Top             =   405
         Width           =   720
      End
      Begin VB.Label lbl�ο� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ο�           ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3330
         TabIndex        =   19
         Top             =   405
         Width           =   1530
      End
      Begin VB.Label lbl�Ա� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3345
         TabIndex        =   11
         Top             =   90
         Width           =   360
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����      ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5400
         TabIndex        =   13
         Top             =   90
         Width           =   1080
      End
      Begin VB.Label lbl�걾 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�걾����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   75
         TabIndex        =   9
         Top             =   90
         Width           =   720
      End
      Begin XtremeCommandBars.CommandBars cbsThis 
         Left            =   60
         Top             =   1665
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         VisualTheme     =   2
      End
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���쾯ʾ"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6450
      TabIndex        =   37
      Top             =   525
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(��ͬ�����������ϼ�����ıȶ�)"
      Height          =   180
      Left            =   4890
      TabIndex        =   36
      Top             =   135
      Width           =   2880
   End
   Begin VB.Label lblʧ���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ʧ����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3150
      TabIndex        =   4
      Top             =   150
      Width           =   540
   End
   Begin VB.Label lblList 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�ο�ֵ��ϸ�б�:"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   105
      TabIndex        =   6
      Top             =   525
      Width           =   1350
   End
   Begin VB.Label lbl�ȶԾ�ʾ�� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����ȶԼ�飺��ʾ��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   105
      TabIndex        =   2
      Top             =   150
      Width           =   1800
   End
   Begin VB.Label lbl���챨���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���챨��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4920
      TabIndex        =   0
      Top             =   525
      Width           =   720
   End
End
Attribute VB_Name = "frmLabItemRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngItemID As Long          '��ǰ��ʾ����Ŀid
Private mInt���� As Integer         '��ǰ��Ŀ�Ľ������

Private Enum mCol
    Ĭ�� = 0: �걾: �Ա���: �Ա�: ��������: ��������: ���䵥λ: ������ʾ: �ٴ�����: �ο���ֵ: �ο���ֵ: �ο���ʾ: ��ʾ��ֵ: ��ʾ��ֵ: ��ʾ��ʾ: �����ֵ: �����ֵ: ������ʾ: ��ƫ����: ����Id: ������: �������Id: �������: ��ע
End Enum

'��ʱ����
Dim cbrControl As CommandBarControl

Dim lngCount As Long
Dim strTemp As String, aryTemp() As String

'--------------------------------------------
'����Ϊ���幫������
'--------------------------------------------
Private Sub RowRefresh(lngRow As Long)
    '���ܣ�����ִ�������е�����Ͳο�ֵ��ʾ
    '������ lngRow-������
    Dim lngList As Long
    
    With Me.vfgList
        
        Select Case Val(.TextMatrix(lngRow, mCol.�Ա���))
        Case 1: .TextMatrix(lngRow, mCol.�Ա�) = "��"
        Case 2: .TextMatrix(lngRow, mCol.�Ա�) = "Ů"
        Case Else: .TextMatrix(lngRow, mCol.�Ա�) = ""
        End Select
        
        .TextMatrix(lngRow, mCol.��������) = Replace(Replace(" " & Trim(.TextMatrix(lngRow, mCol.��������)), " .", "0."), " ", "")
        .TextMatrix(lngRow, mCol.��������) = Replace(Replace(" " & Trim(.TextMatrix(lngRow, mCol.��������)), " .", "0."), " ", "")
        .TextMatrix(lngRow, mCol.�ο���ֵ) = Replace(Replace(" " & Trim(.TextMatrix(lngRow, mCol.�ο���ֵ)), " .", "0."), " ", "")
        .TextMatrix(lngRow, mCol.�ο���ֵ) = Replace(Replace(" " & Trim(.TextMatrix(lngRow, mCol.�ο���ֵ)), " .", "0."), " ", "")
        .TextMatrix(lngRow, mCol.��ʾ��ֵ) = Replace(Replace(" " & Trim(.TextMatrix(lngRow, mCol.��ʾ��ֵ)), " .", "0."), " ", "")
        .TextMatrix(lngRow, mCol.��ʾ��ֵ) = Replace(Replace(" " & Trim(.TextMatrix(lngRow, mCol.��ʾ��ֵ)), " .", "0."), " ", "")
        .TextMatrix(lngRow, mCol.�����ֵ) = Replace(Replace(" " & Trim(.TextMatrix(lngRow, mCol.�����ֵ)), " .", "0."), " ", "")
        .TextMatrix(lngRow, mCol.�����ֵ) = Replace(Replace(" " & Trim(.TextMatrix(lngRow, mCol.�����ֵ)), " .", "0."), " ", "")
        
        
        If Val(.TextMatrix(lngRow, mCol.��������)) <> 0 And Val(.TextMatrix(lngRow, mCol.��������)) <> 0 Then
            .TextMatrix(lngRow, mCol.������ʾ) = .TextMatrix(lngRow, mCol.��������) & "��" & .TextMatrix(lngRow, mCol.��������) & .TextMatrix(lngRow, mCol.���䵥λ)
        ElseIf Val(.TextMatrix(lngRow, mCol.��������)) <> 0 Then
            .TextMatrix(lngRow, mCol.������ʾ) = .TextMatrix(lngRow, mCol.��������) & .TextMatrix(lngRow, mCol.��������) & .TextMatrix(lngRow, mCol.���䵥λ) & "��"
        ElseIf Val(.TextMatrix(lngRow, mCol.��������)) <> 0 Then
            .TextMatrix(lngRow, mCol.������ʾ) = "��" & .TextMatrix(lngRow, mCol.��������) & .TextMatrix(lngRow, mCol.���䵥λ)
        Else
            .TextMatrix(lngRow, mCol.������ʾ) = ""
        End If
        
        .TextMatrix(lngRow, mCol.�ο���ʾ) = ""
        .TextMatrix(lngRow, mCol.��ʾ��ʾ) = ""
        .TextMatrix(lngRow, mCol.������ʾ) = ""
        Select Case mInt����
        Case 1, 3  '�����Ͷ���(����Ҳ��Ҫ�вο�ֵ)
'            .TextMatrix(lngRow, mcol.�ο���ֵ) = FormatReference(mlngItemID, .TextMatrix(lngRow, mcol.�ο���ֵ))
'            .TextMatrix(lngRow, mcol.�ο���ֵ) = FormatReference(mlngItemID, .TextMatrix(lngRow, mcol.�ο���ֵ))
            If IsNumeric(.TextMatrix(lngRow, mCol.�ο���ֵ)) = True And IsNumeric(.TextMatrix(lngRow, mCol.�ο���ֵ)) = True Then
                .TextMatrix(lngRow, mCol.�ο���ʾ) = .TextMatrix(lngRow, mCol.�ο���ֵ) & "��" & .TextMatrix(lngRow, mCol.�ο���ֵ)
            ElseIf IsNumeric(.TextMatrix(lngRow, mCol.�ο���ֵ)) = True Then
                .TextMatrix(lngRow, mCol.�ο���ʾ) = .TextMatrix(lngRow, mCol.�ο���ֵ) & "��"
            ElseIf IsNumeric(.TextMatrix(lngRow, mCol.�ο���ֵ)) = True Then
                .TextMatrix(lngRow, mCol.�ο���ʾ) = "��" & .TextMatrix(lngRow, mCol.�ο���ֵ)
            End If
        
            If IsNumeric(.TextMatrix(lngRow, mCol.��ʾ��ֵ)) = True And IsNumeric(.TextMatrix(lngRow, mCol.��ʾ��ֵ)) = True Then
                .TextMatrix(lngRow, mCol.��ʾ��ʾ) = .TextMatrix(lngRow, mCol.��ʾ��ֵ) & "��" & .TextMatrix(lngRow, mCol.��ʾ��ֵ)
            ElseIf IsNumeric(.TextMatrix(lngRow, mCol.��ʾ��ֵ)) = True Then
                .TextMatrix(lngRow, mCol.��ʾ��ʾ) = .TextMatrix(lngRow, mCol.��ʾ��ֵ) & "��"
            ElseIf IsNumeric(.TextMatrix(lngRow, mCol.��ʾ��ֵ)) = True Then
                .TextMatrix(lngRow, mCol.��ʾ��ʾ) = "��" & .TextMatrix(lngRow, mCol.��ʾ��ֵ)
            End If
        
            If IsNumeric(.TextMatrix(lngRow, mCol.�����ֵ)) = True And IsNumeric(.TextMatrix(lngRow, mCol.�����ֵ)) = True Then
                .TextMatrix(lngRow, mCol.������ʾ) = .TextMatrix(lngRow, mCol.�����ֵ) & "��" & .TextMatrix(lngRow, mCol.�����ֵ)
            ElseIf IsNumeric(.TextMatrix(lngRow, mCol.�����ֵ)) = True Then
                .TextMatrix(lngRow, mCol.������ʾ) = .TextMatrix(lngRow, mCol.�����ֵ) & "��"
            ElseIf IsNumeric(.TextMatrix(lngRow, mCol.�����ֵ)) = True Then
                .TextMatrix(lngRow, mCol.������ʾ) = "��" & .TextMatrix(lngRow, mCol.�����ֵ)
            End If
        Case 2  '�붨��
            For lngList = 0 To Me.cbo�ο�ֵ.ListCount - 1
                If lngList = Val(.TextMatrix(lngRow, mCol.�ο���ֵ)) And IsNumeric(.TextMatrix(lngRow, mCol.�ο���ֵ)) = True Then
                    .TextMatrix(lngRow, mCol.�ο���ʾ) = Me.cbo�ο�ֵ.List(lngList + 1): Exit For
                End If
            Next
        End Select
    End With
End Sub

Private Sub setListFormat(Optional blnKeepData As Boolean)
    '���ܣ���ʼ�����òο�ֵ�б�
    '������ blnKeepData-�Ƿ������ݣ���ֻ���������ø�ʽ
    With Me.vfgList
        .Redraw = flexRDNone
        If blnKeepData = False Then
            .Clear
            .Rows = 1: .FixedRows = 1: .Cols = 24: .FixedCols = 0
        End If
        .TextMatrix(0, mCol.Ĭ��) = "Ĭ��":
        .TextMatrix(0, mCol.�걾) = "�걾": .TextMatrix(0, mCol.�Ա���) = "�Ա���": .TextMatrix(0, mCol.�Ա�) = "�Ա�"
        .TextMatrix(0, mCol.��������) = "��������": .TextMatrix(0, mCol.��������) = "��������"
        .TextMatrix(0, mCol.���䵥λ) = "���䵥λ": .TextMatrix(0, mCol.������ʾ) = "����"
        .TextMatrix(0, mCol.�ٴ�����) = "�ٴ�����"
        .TextMatrix(0, mCol.�ο���ֵ) = "�ο���ֵ": .TextMatrix(0, mCol.�ο���ֵ) = "�ο���ֵ": .TextMatrix(0, mCol.�ο���ʾ) = "�ο�ֵ"
        .TextMatrix(0, mCol.��ʾ��ֵ) = "��ʾ��ֵ": .TextMatrix(0, mCol.��ʾ��ֵ) = "��ʾ��ֵ": .TextMatrix(0, mCol.��ʾ��ʾ) = "��ʾֵ"
        .TextMatrix(0, mCol.�����ֵ) = "�����ֵ": .TextMatrix(0, mCol.�����ֵ) = "�����ֵ": .TextMatrix(0, mCol.������ʾ) = "����ֵ"
        .TextMatrix(0, mCol.��ƫ����) = "��ƫ����"
        .TextMatrix(0, mCol.����Id) = "����id": .TextMatrix(0, mCol.������) = "������"
        .TextMatrix(0, mCol.�������Id) = "����id": .TextMatrix(0, mCol.�������) = "�������"
        .TextMatrix(0, mCol.��ע) = "��ע"
        
        .ColWidth(mCol.Ĭ��) = 500
        .ColWidth(mCol.�걾) = 1000: .ColWidth(mCol.�Ա���) = 0: .ColWidth(mCol.�Ա�) = 700
        .ColWidth(mCol.��������) = 0: .ColWidth(mCol.��������) = 0
        .ColWidth(mCol.���䵥λ) = 0: .ColWidth(mCol.������ʾ) = 1000
        .ColWidth(mCol.�ٴ�����) = 1200
        .ColWidth(mCol.�ο���ֵ) = 0: .ColWidth(mCol.�ο���ֵ) = 0: .ColWidth(mCol.�ο���ʾ) = 1300
        .ColWidth(mCol.��ʾ��ֵ) = 0: .ColWidth(mCol.��ʾ��ֵ) = 0: .ColWidth(mCol.��ʾ��ʾ) = 1300
        .ColWidth(mCol.�����ֵ) = 0: .ColWidth(mCol.�����ֵ) = 0: .ColWidth(mCol.������ʾ) = 1300
        .ColWidth(mCol.��ƫ����) = 900
        .ColWidth(mCol.����Id) = 0: .ColWidth(mCol.������) = 1500
        .ColWidth(mCol.�������Id) = 0: .ColWidth(mCol.�������) = 1500
        .ColWidth(mCol.��ע) = 1500
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        .ColDataType(mCol.Ĭ��) = flexDTBoolean

        '��������Ͳο�ֵ����ʾ
        For lngCount = .FixedRows To .Rows - 1
            Call RowRefresh(lngCount)
        Next
        .Redraw = flexRDDirect
    End With
    
End Sub

Public Function zlRefresh(lngItemID As Long) As Boolean
    '���ܣ�������Ŀidˢ�µ�ǰ��ʾ����
    '��������ǰ��Ŀid
    Dim rsTemp As New ADODB.Recordset
    
    mlngItemID = lngItemID
    
    '�����ǰ��Ŀ����ʾ
    Me.txt��������.Text = "": Me.txt��������.Text = ""
    Me.txt���챨����.Text = "": Me.txt�ȶԾ�ʾ��.Text = "": Me.txt�ȶ�ʧ����.Text = "": Me.txt���쾯ʾ��.Text = ""
        
    If lngItemID = 0 Then Call setListFormat: zlRefresh = True: Exit Function
    
    '��ȡָ����Ŀ����Ϣ
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select I.�������, I.��������, I.��������, I.���챨����, I.�ȶԾ�ʾ��, I.�ȶ�ʧ����, I.ȡֵ����,I.���쾯ʾ�� " & vbNewLine & _
            "From ������Ŀ I, ���鱨����Ŀ R" & vbNewLine & _
            "Where I.������Ŀid = R.������Ŀid And R.������Ŀid = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    With rsTemp
        mInt���� = 1
        If .RecordCount > 0 Then
'            Me.txt��������.Text = "" & !��������: Me.txt��������.Text = "" & !��������
            Me.txt���챨����.Text = "" & !���챨����: Me.txt���쾯ʾ��.Text = "" & !���쾯ʾ��
            Me.txt�ȶԾ�ʾ��.Text = "" & !�ȶԾ�ʾ��: Me.txt�ȶ�ʧ����.Text = "" & !�ȶ�ʧ����
            Me.cbo�ο�ֵ.Tag = "" & !ȡֵ����

            mInt���� = Val("" & !�������)
            
            'Me.txt��������.Text = Replace(Replace(" " & Trim(Me.txt��������.Text), " .", "0."), " ", "")
            'Me.txt��������.Text = Replace(Replace(" " & Trim(Me.txt��������.Text), " .", "0."), " ", "")
            Me.txt���챨����.Text = Replace(Replace(" " & Trim(Me.txt���챨����.Text), " .", "0."), " ", "")
            Me.txt���쾯ʾ��.Text = Replace(Replace(" " & Trim(Me.txt���쾯ʾ��.Text), " .", "0."), " ", "")
            Me.txt�ȶԾ�ʾ��.Text = Replace(Replace(" " & Trim(Me.txt�ȶԾ�ʾ��.Text), " .", "0."), " ", "")
            Me.txt�ȶ�ʧ����.Text = Replace(Replace(" " & Trim(Me.txt�ȶ�ʧ����.Text), " .", "0."), " ", "")
        End If
    End With
    
    Me.lbl�ο�.Visible = False
    Me.txt�ο�ֵ(0).Visible = False: Me.txt�ο�ֵ(1).Visible = False: Me.txt��ƫ����.Visible = False: lbl��ƫ����.Visible = False
    Me.txt��������.Visible = False: Me.txt��������.Visible = False
    Me.txt��������.Visible = False: Me.txt��������.Visible = False
    Me.lbl��������.Visible = False: Me.lbl��������.Visible = False
    Me.cbo�ο�ֵ.Visible = False
    Select Case mInt����
    Case 3 '�붨��
        Me.lbl�ο�.Visible = True
        Me.txt�ο�ֵ(0).Visible = True: Me.txt�ο�ֵ(1).Visible = True
        Me.lbl��������.Visible = True: Me.lbl��������.Visible = True
        Me.txt��������.Visible = True: Me.txt��������.Visible = True
        Me.txt��������.Visible = True: Me.txt��������.Visible = True
    Case 2 '������
        Me.lbl�ο�.Visible = True
        With Me.cbo�ο�ֵ
            .Clear
            aryTemp() = Split(.Tag, ";")
            .AddItem ""
            For lngCount = LBound(aryTemp) To UBound(aryTemp)
                .AddItem aryTemp(lngCount)
            Next
            .Visible = True
        End With
        
    Case Else '������
        Me.lbl�ο�.Visible = True
        Me.txt�ο�ֵ(0).Visible = True: Me.txt�ο�ֵ(1).Visible = True: Me.txt��ƫ����.Visible = True: lbl��ƫ����.Visible = True
        Me.lbl��������.Visible = True: Me.lbl��������.Visible = True
        Me.txt��������.Visible = True: Me.txt��������.Visible = True
        Me.txt��������.Visible = True: Me.txt��������.Visible = True
    End Select
        
    gstrSql = "Select nvl(L.Ĭ��,0) As Ĭ��,L.�걾���� As �걾, L.�Ա���, '' As �Ա�, L.��������, L.��������, L.���䵥λ, '' As ����, L.�ٴ�����," & vbNewLine & _
            "       L.�ο���ֵ, L.�ο���ֵ, '' As �ο�,L.��ʾ����,L.��ʾ����,'' as ��ʾ,L.��������,L.��������,'' as ����, L.��ƫ����, L.����id, M.���� As ������, L.�������ID, N.���� as �������, L.��ע " & vbNewLine & _
            "From ������Ŀ�ο� L, ���鱨����Ŀ R, �������� M, ���ű� N" & vbNewLine & _
            "Where L.��Ŀid = R.������Ŀid And L.����id = M.ID(+) And L.�������ID=N.ID(+) And R.������Ŀid = [1] Order by L.�걾����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    Set Me.vfgList.DataSource = rsTemp:
    Call setListFormat(True)
    If Me.vfgList.Rows <= Me.vfgList.FixedRows Then
        Me.vfgList.Rows = Me.vfgList.FixedRows + 1
        Me.vfgList.Row = Me.vfgList.FixedRows
    End If
    Call vfgList_RowColChange
    
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart() As Boolean
    '���ܣ���ʼ��Ŀ�༭
    '������ lngItemId-ָ���༭����Ŀ
        
    If Me.cbo�걾.ListCount = 0 Then
        MsgBox "�������ֵ��г�ʼ��������걾����", vbInformation, gstrSysName
        zlEditStart = False: Exit Function
    End If
    
    Me.Tag = "�༭": Call Form_Resize
    If Me.Visible Then Me.txt�ȶԾ�ʾ��.SetFocus
    zlEditStart = True: Exit Function

End Function

Public Sub zlEditCancel()
    '���ܣ��������ڽ��еı༭
    Me.Tag = "": Call Form_Resize
    Call Me.zlRefresh(mlngItemID)
End Sub

Public Function zlEditSave() As Long
    '���ܣ��������ڽ��еı༭,���������ڱ༭��Ŀid,����ʧ�ܷ���0
    Dim strValue As String
    'һ�����Լ��
    If Val(Me.txt��������.Text) <> 0 And Val(Me.txt��������.Text) <> 0 Then
        If Val(Me.txt��������.Text) >= Val(Me.txt��������.Text) Then
            MsgBox "�������ޱ�����ھ�������", vbInformation, gstrSysName
            Me.txt��������.SetFocus: zlEditSave = 0: Exit Function
        End If
    End If
    If Val(Me.txt��������.Text) <> 0 Then
        If Val(Me.txt��������.Text) > 999999999 Or Val(Val(Me.txt��������.Text) * 100000) - Int(Val(Val(Me.txt��������.Text) * 100000)) > 0 Then
            MsgBox "����������ֵ̫��򾫶�̫�ߣ�", vbInformation, gstrSysName
            Me.txt��������.SetFocus: zlEditSave = 0: Exit Function
        End If
    End If
    If Val(Me.txt��������.Text) <> 0 Then
        If Val(Me.txt��������.Text) > 999999999 Or Val(Val(Me.txt��������.Text) * 100000) - Int(Val(Val(Me.txt��������.Text) * 100000)) > 0 Then
            MsgBox "����������ֵ̫��򾫶�̫�ߣ�", vbInformation, gstrSysName
            Me.txt��������.SetFocus: zlEditSave = 0: Exit Function
        End If
    End If
    If Val(Me.txt���챨����.Text) <> 0 Then
        If Val(Me.txt���챨����.Text) > 999999999 Or Val(Val(Me.txt���챨����.Text) * 100000) - Int(Val(Val(Me.txt���챨����.Text) * 100000)) > 0 Then
            MsgBox "���챨����̫��򾫶�̫�ߣ�", vbInformation, gstrSysName
            Me.txt���챨����.SetFocus: zlEditSave = 0: Exit Function
        End If
    End If
    
    If Val(Me.txt���쾯ʾ��.Text) <> 0 Then
        If Val(Me.txt���쾯ʾ��.Text) > 999999999 Or Val(Val(Me.txt���쾯ʾ��.Text) * 100000) - Int(Val(Val(Me.txt���쾯ʾ��.Text) * 100000)) > 0 Then
            MsgBox "���챨����̫��򾫶�̫�ߣ�", vbInformation, gstrSysName
            Me.txt���쾯ʾ��.SetFocus: zlEditSave = 0: Exit Function
        End If
    End If
    
    
    If Val(Me.txt�ȶԾ�ʾ��.Text) <> 0 Then
        If Val(Me.txt�ȶԾ�ʾ��.Text) > 999999999 Or Val(Val(Me.txt�ȶԾ�ʾ��.Text) * 100000) - Int(Val(Val(Me.txt�ȶԾ�ʾ��.Text) * 100000)) > 0 Then
            MsgBox "�ȶԾ�ʾ��̫��򾫶�̫�ߣ�", vbInformation, gstrSysName
            Me.txt�ȶԾ�ʾ��.SetFocus: zlEditSave = 0: Exit Function
        End If
    End If
    If Val(Me.txt�ȶ�ʧ����.Text) <> 0 Then
        If Val(Me.txt�ȶ�ʧ����.Text) > 999999999 Or Val(Val(Me.txt�ȶ�ʧ����.Text) * 100000) - Int(Val(Val(Me.txt�ȶ�ʧ����.Text) * 100000)) > 0 Then
            MsgBox "�ȶ�ʧ����̫��򾫶�̫�ߣ�", vbInformation, gstrSysName
            Me.txt�ȶ�ʧ����.SetFocus: zlEditSave = 0: Exit Function
        End If
    End If
    
    '�ο�ֵ�б���
    Dim strLists As String, strItems As String
    With Me.vfgList
        strLists = ""
        For lngCount = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(lngCount, mCol.�걾)) = "" Then
                MsgBox "��" & lngCount & "�ο�ֵδ��д�걾������д��ɾ����", vbInformation, gstrSysName
                Me.cbo�걾.SetFocus: zlEditSave = 0: Exit Function
            End If
            
            strItems = Trim(.TextMatrix(lngCount, mCol.�걾)) & ";"
            If Val(.TextMatrix(lngCount, mCol.�Ա���)) <> 0 Then strItems = strItems & Val(.TextMatrix(lngCount, mCol.�Ա���))
            strItems = strItems & ";"
            If IsNumeric(.TextMatrix(lngCount, mCol.��������)) = True Then strItems = strItems & Val(.TextMatrix(lngCount, mCol.��������))
            strItems = strItems & ";"
            If IsNumeric(.TextMatrix(lngCount, mCol.��������)) = True Then strItems = strItems & Val(.TextMatrix(lngCount, mCol.��������))
            strItems = strItems & ";" & Trim(.TextMatrix(lngCount, mCol.���䵥λ))
            strItems = strItems & ";" & Trim(.TextMatrix(lngCount, mCol.�ٴ�����))
            strItems = strItems & ";"
            If IsNumeric(.TextMatrix(lngCount, mCol.�ο���ֵ)) = True Then
                strItems = strItems & .TextMatrix(lngCount, mCol.�ο���ֵ)
                strValue = .TextMatrix(lngCount, mCol.�ο���ֵ)
                If Val(strValue) > 999999999 Or Val(Val(strValue) * 10000) - Int(Val(Val(strValue) * 10000)) > 0 Then
                    MsgBox "��" & lngCount & "�ο�ֵ̫��򾫶�̫�ߣ�", vbInformation, gstrSysName
                    Me.cbo�걾.SetFocus: zlEditSave = 0: Exit Function
                End If
            End If
            strItems = strItems & ";"
            If IsNumeric(.TextMatrix(lngCount, mCol.�ο���ֵ)) = True Then
                strItems = strItems & .TextMatrix(lngCount, mCol.�ο���ֵ)
                strValue = .TextMatrix(lngCount, mCol.�ο���ֵ)
                If Val(strValue) > 999999999 Or Val(Val(strValue) * 10000) - Int(Val(Val(strValue) * 10000)) > 0 Then
                    MsgBox "��" & lngCount & "�ο�ֵ̫��򾫶�̫�ߣ�", vbInformation, gstrSysName
                    Me.cbo�걾.SetFocus: zlEditSave = 0: Exit Function
                End If
            End If
            strItems = strItems & ";"
            If IsNumeric(.TextMatrix(lngCount, mCol.��ƫ����)) = True Then
                strItems = strItems & Val(.TextMatrix(lngCount, mCol.��ƫ����))
                strValue = .TextMatrix(lngCount, mCol.��ƫ����)
                If Val(strValue) > 999999999 Or Val(Val(strValue) * 10000) - Int(Val(Val(strValue) * 10000)) > 0 Then
                    MsgBox "��" & lngCount & "��ƫ����̫��򾫶�̫�ߣ�", vbInformation, gstrSysName
                    Me.cbo�걾.SetFocus: zlEditSave = 0: Exit Function
                End If
            End If
            strItems = strItems & ";"
            If Val(.TextMatrix(lngCount, mCol.����Id)) <> 0 Then strItems = strItems & Val(.TextMatrix(lngCount, mCol.����Id))
            strItems = strItems & ";" & Trim(.TextMatrix(lngCount, mCol.��ע))
            strItems = strItems & ";" & Trim(.TextMatrix(lngCount, mCol.Ĭ��))
            
            strItems = strItems & ";"
            If IsNumeric(.TextMatrix(lngCount, mCol.��ʾ��ֵ)) = True Then
                strItems = strItems & .TextMatrix(lngCount, mCol.��ʾ��ֵ)
                strValue = .TextMatrix(lngCount, mCol.��ʾ��ֵ)
                If Val(strValue) > 999999999 Or Val(Val(strValue) * 10000) - Int(Val(Val(strValue) * 10000)) > 0 Then
                    MsgBox "��" & lngCount & "��ʾֵ̫��򾫶�̫�ߣ�", vbInformation, gstrSysName
                    Me.cbo�걾.SetFocus: zlEditSave = 0: Exit Function
                End If
            End If
            strItems = strItems & ";"
            If IsNumeric(.TextMatrix(lngCount, mCol.��ʾ��ֵ)) = True Then
                strItems = strItems & .TextMatrix(lngCount, mCol.��ʾ��ֵ)
                strValue = .TextMatrix(lngCount, mCol.��ʾ��ֵ)
                If Val(strValue) > 999999999 Or Val(Val(strValue) * 10000) - Int(Val(Val(strValue) * 10000)) > 0 Then
                    MsgBox "��" & lngCount & "��ʾֵ̫��򾫶�̫�ߣ�", vbInformation, gstrSysName
                    Me.cbo�걾.SetFocus: zlEditSave = 0: Exit Function
                End If
            End If
            
            strItems = strItems & ";"
            If IsNumeric(.TextMatrix(lngCount, mCol.�����ֵ)) = True Then
                strItems = strItems & .TextMatrix(lngCount, mCol.�����ֵ)
                strValue = .TextMatrix(lngCount, mCol.�����ֵ)
                If Val(strValue) > 999999999 Or Val(Val(strValue) * 10000) - Int(Val(Val(strValue) * 10000)) > 0 Then
                    MsgBox "��" & lngCount & "����ֵ̫��򾫶�̫�ߣ�", vbInformation, gstrSysName
                    Me.cbo�걾.SetFocus: zlEditSave = 0: Exit Function
                End If
            End If
            strItems = strItems & ";"
            If IsNumeric(.TextMatrix(lngCount, mCol.�����ֵ)) = True Then
                strItems = strItems & .TextMatrix(lngCount, mCol.�����ֵ)
                strValue = .TextMatrix(lngCount, mCol.�����ֵ)
                If Val(strValue) > 999999999 Or Val(Val(strValue) * 10000) - Int(Val(Val(strValue) * 10000)) > 0 Then
                    MsgBox "��" & lngCount & "����ֵ̫��򾫶�̫�ߣ�", vbInformation, gstrSysName
                    Me.cbo�걾.SetFocus: zlEditSave = 0: Exit Function
                End If
            End If
            strItems = strItems & ";"
            If Val(.TextMatrix(lngCount, mCol.�������Id)) <> 0 Then strItems = strItems & Val(.TextMatrix(lngCount, mCol.�������Id))
            
            strLists = strLists & "|" & strItems
        Next
    End With
    If strLists <> "" Then strLists = Mid(strLists, 2)

    '���ݱ��������֯
    gstrSql = "Zl_������Ŀ�ο�_Edit(" & mlngItemID & "," & IIf(Trim(Me.txt��������.Text) = "", "''", Val(Me.txt��������.Text)) & "," & _
              IIf(Val(Me.txt��������.Text) = 0, "''", Val(Me.txt��������.Text)) & "," & IIf(Val(Me.txt���챨����.Text) = 0, "''", Val(Me.txt���챨����.Text)) & _
              "," & Val(Me.txt���쾯ʾ��.Text) & _
              "," & IIf(Val(Me.txt�ȶԾ�ʾ��.Text) = 0, "''", Val(Me.txt�ȶԾ�ʾ��.Text)) & _
              "," & IIf(Val(Me.txt�ȶ�ʧ����.Text) = 0, "''", Val(Me.txt�ȶ�ʧ����.Text)) & ",'" & strLists & "')"
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    Me.Tag = "": Call Form_Resize
    zlEditSave = mlngItemID: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

Private Sub cbo�걾_Click()
'   ���ݱ걾�����Ա�
    If Me.cbo�걾.ListIndex >= 0 Then
        If Me.cbo�걾.ItemData(Me.cbo�걾.ListIndex) = 1 Then
            Me.cbo�Ա�.ListIndex = 1
            Me.cbo�Ա�.Enabled = False
        ElseIf cbo�걾.ItemData(Me.cbo�걾.ListIndex) = 2 Then
            Me.cbo�Ա�.ListIndex = 2
            Me.cbo�Ա�.Enabled = False
        Else
            Me.cbo�Ա�.Enabled = True
        End If
    End If
End Sub

'--------------------------------------------
'����Ϊ����ؼ���Ӧ�¼�
'--------------------------------------------

Private Sub cbo�걾_GotFocus()
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub cbo�걾_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo�ο�ֵ_GotFocus()
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub cbo�ο�ֵ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo��λ_GotFocus()
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub cbo��λ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo�ٴ�����_GotFocus()
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub cbo�ٴ�����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo�Ա�_GotFocus()
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub cbo�Ա�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo����_GotFocus()
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngCurRow As Long, lngRow As Long, lngCol As Long
    Dim Str�걾 As String, intRow As Integer
    With Me.vfgList
        Select Case Control.ID
        Case conMenu_Edit_NewItem
            .Rows = .Rows + 1: .Row = .Rows - 1
        Case conMenu_Edit_Delete
            If .Row = .Rows - 1 Then
                .Rows = .Rows - 1: .Row = .Rows - 1
            Else
                lngCurRow = .Row
                For lngRow = lngCurRow To .Rows - 2
                    For lngCol = 0 To .Cols - 1
                        .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow + 1, lngCol)
                    Next
                Next
                .Rows = .Rows - 1
            End If
        Case conMenu_Edit_Adjust
            If Me.cbo�걾.ListIndex = -1 Then
                .TextMatrix(.Row, mCol.�걾) = ""
            Else
                .TextMatrix(.Row, mCol.�걾) = Mid(Me.cbo�걾.Text, 4)
            End If
            .TextMatrix(.Row, mCol.�Ա���) = Val(Left(Me.cbo�Ա�.Text, 1))
            .TextMatrix(.Row, mCol.��������) = Me.txt����(0).Text
            .TextMatrix(.Row, mCol.��������) = Me.txt����(1).Text
            
            If Me.cbo��λ.ListIndex = -1 Then
                .TextMatrix(.Row, mCol.���䵥λ) = "��"
            Else
                .TextMatrix(.Row, mCol.���䵥λ) = Mid(Me.cbo��λ.Text, 3)
            End If
            If Me.cbo�ٴ�����.ListIndex = -1 Then
                .TextMatrix(.Row, mCol.�ٴ�����) = ""
            Else
                .TextMatrix(.Row, mCol.�ٴ�����) = Mid(Me.cbo�ٴ�����.Text, 4)
            End If
            Select Case mInt����
            Case 1, 3
                .TextMatrix(.Row, mCol.�ο���ֵ) = FormatReference(mlngItemID, IIf(IsNumeric(txt�ο�ֵ(0)), Me.txt�ο�ֵ(0), ""))
                .TextMatrix(.Row, mCol.�ο���ֵ) = FormatReference(mlngItemID, IIf(IsNumeric(txt�ο�ֵ(1)), Me.txt�ο�ֵ(1), ""))
                .TextMatrix(.Row, mCol.��ʾ��ֵ) = FormatReference(mlngItemID, IIf(IsNumeric(txt��������), Me.txt��������, ""))
                .TextMatrix(.Row, mCol.��ʾ��ֵ) = FormatReference(mlngItemID, IIf(IsNumeric(txt��������), Me.txt��������, ""))
                .TextMatrix(.Row, mCol.�����ֵ) = FormatReference(mlngItemID, IIf(IsNumeric(txt��������), Me.txt��������, ""))
                .TextMatrix(.Row, mCol.�����ֵ) = FormatReference(mlngItemID, IIf(IsNumeric(txt��������), Me.txt��������, ""))
                .TextMatrix(.Row, mCol.��ƫ����) = IIf(Val(Me.txt��ƫ����.Text) = 0, "", Val(Me.txt��ƫ����.Text))
                
                '--- 2012-1-19 ��Ϊ���������������������Ϲ�ҽ���޸ģ���Ϊ���û������趨ÿ���ο��ľ�ʾ���ޣ����ޡ�
'                For intRow = .FixedRows To .Rows - 1
'                    If .TextMatrix(intRow, mCol.�걾) = .TextMatrix(.Row, mCol.�걾) Then
'                        .TextMatrix(intRow, mCol.��ʾ��ֵ) = FormatReference(mlngItemID, IIf(IsNumeric(txt��������), Me.txt��������, ""))
'                        .TextMatrix(intRow, mCol.��ʾ��ֵ) = FormatReference(mlngItemID, IIf(IsNumeric(txt��������), Me.txt��������, ""))
'                        .TextMatrix(intRow, mCol.�����ֵ) = FormatReference(mlngItemID, IIf(IsNumeric(txt��������), Me.txt��������, ""))
'                        .TextMatrix(intRow, mCol.�����ֵ) = FormatReference(mlngItemID, IIf(IsNumeric(txt��������), Me.txt��������, ""))
'                        Call RowRefresh(CLng(intRow))
'                    End If
'                Next
                
            Case 2
                .TextMatrix(.Row, mCol.�ο���ֵ) = IIf(Me.cbo�ο�ֵ.ListIndex = 0, "", Me.cbo�ο�ֵ.ListIndex - 1)
                .TextMatrix(.Row, mCol.�ο���ֵ) = IIf(Me.cbo�ο�ֵ.ListIndex = 0, "", Me.cbo�ο�ֵ.ListIndex - 1)
                .TextMatrix(.Row, mCol.��ƫ����) = ""
            End Select
            If Me.cbo����.ListIndex = -1 Then
                .TextMatrix(.Row, mCol.����Id) = 0: .TextMatrix(.Row, mCol.������) = ""
            Else
                .TextMatrix(.Row, mCol.����Id) = Me.cbo����.ItemData(Me.cbo����.ListIndex)
                .TextMatrix(.Row, mCol.������) = Mid(Me.cbo����.Text, 5)
            End If
            .TextMatrix(.Row, mCol.��ע) = Trim(Me.txt��ע.Text)
            .TextMatrix(.Row, mCol.Ĭ��) = Me.chkĬ��.Value
            If Me.chkĬ��.Value = 1 Then
                For lngRow = .FixedRows To .Rows - 1
                    If lngRow <> .Row And .TextMatrix(lngRow, mCol.�걾) = .TextMatrix(.Row, mCol.�걾) Then
                        If .TextMatrix(lngRow, mCol.Ĭ��) = 1 Then .TextMatrix(lngRow, mCol.Ĭ��) = 0
                    End If
                Next
            End If
            If Me.cbo����.ListIndex = -1 Then
                .TextMatrix(.Row, mCol.�������Id) = 0: .TextMatrix(.Row, mCol.�������Id) = ""
            Else
                .TextMatrix(.Row, mCol.�������Id) = Me.cbo����.ItemData(Me.cbo����.ListIndex)
                .TextMatrix(.Row, mCol.�������) = Mid(Me.cbo����.Text, 5)
            End If
            Call RowRefresh(.Row)
        End Select
    End With
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_Edit_NewItem: Control.Enabled = (Me.Tag = "�༭")
    Case conMenu_Edit_Delete, conMenu_Edit_Adjust: Control.Enabled = (Me.Tag = "�༭" And Me.vfgList.Row >= Me.vfgList.FixedRows)
    End Select
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    '�ڲ��˵�����������
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set cbsThis.Icons = zlcommfun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
    End With
    Me.cbsThis.EnableCustomization False
    
    Me.cbsThis.ActiveMenuBar.Title = "�˵�"
    Me.cbsThis.ActiveMenuBar.Position = xtpBarBottom
    Me.cbsThis.ActiveMenuBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    With Me.cbsThis.ActiveMenuBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "��������"): cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ������"): cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "���µ��ο�ֵ�б���"): cbrControl.Flags = xtpFlagRightAlign: cbrControl.Style = xtpButtonIconAndCaption
    End With

    '��������װ��
    aryTemp = Split("0-������;1-����;2-Ů��", ";")
    For lngCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo�Ա�.AddItem aryTemp(lngCount)
    Next
    Me.cbo�Ա�.ListIndex = 0
    
    aryTemp = Split("1-��;2-��;3-��;4-Сʱ", ";")
    For lngCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo��λ.AddItem aryTemp(lngCount)
    Next
    Me.cbo��λ.ListIndex = 0
    
    Err = 0: On Error GoTo ErrHand
 
    gstrSql = "Select ����,����,�����Ա� From ���Ƽ���걾"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Do While Not rsTemp.EOF
        Me.cbo�걾.AddItem rsTemp!���� & "-" & rsTemp!����
        If InStr(Trim("" & rsTemp!�����Ա�), "��") > 0 Then
            Me.cbo�걾.ItemData(Me.cbo�걾.NewIndex) = 1
        ElseIf InStr(Trim("" & rsTemp!�����Ա�), "Ů") > 0 Then
            Me.cbo�걾.ItemData(Me.cbo�걾.NewIndex) = 2
        Else
            Me.cbo�걾.ItemData(Me.cbo�걾.NewIndex) = 0
        End If
        rsTemp.MoveNext
    Loop
    If Me.cbo�걾.ListCount > 0 Then Me.cbo�걾.ListIndex = 0
        
    gstrSql = "Select ����,���� From �ٴ�����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Me.cbo�ٴ�����.AddItem "-": Me.cbo�ٴ�����.ListIndex = 0
    Do While Not rsTemp.EOF
        Me.cbo�ٴ�����.AddItem rsTemp!���� & "-" & rsTemp!����
        rsTemp.MoveNext
    Loop

    
    gstrSql = "Select ID, ����, ���� From ��������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Me.cbo����.AddItem "-": Me.cbo����.ItemData(Me.cbo����.NewIndex) = 0: Me.cbo����.ListIndex = 0
    Do While Not rsTemp.EOF
        Me.cbo����.AddItem rsTemp!���� & "-" & rsTemp!����
        Me.cbo����.ItemData(Me.cbo����.NewIndex) = Val("" & rsTemp!ID)
        rsTemp.MoveNext
    Loop
    
    
    gstrSql = "Select Distinct a.Id, a.����, a.����, b.�������" & vbNewLine & _
            "From ���ű� A, ��������˵�� B" & vbNewLine & _
            "Where a.Id = b.����id And b.�������� = '�ٴ�' Order by a.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Me.cbo����.AddItem "-": Me.cbo����.ItemData(Me.cbo����.NewIndex) = 0: Me.cbo����.ListIndex = 0
    Do While Not rsTemp.EOF
        Me.cbo����.AddItem rsTemp!���� & "-" & rsTemp!����
        Me.cbo����.ItemData(Me.cbo����.NewIndex) = Val("" & rsTemp!ID)
        rsTemp.MoveNext
    Loop

    
    '�б�����
    Call setListFormat

    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    Me.picEdit.Top = Me.ScaleHeight - Me.picEdit.Height - 180
    If Me.Tag = "�༭" Then
        Me.vfgList.Height = Me.picEdit.Top - Me.vfgList.Top
        Me.txt��������.Enabled = True: Me.txt��������.Enabled = True
        Me.txt���챨����.Enabled = True: Me.txt�ȶԾ�ʾ��.Enabled = True: Me.txt�ȶ�ʧ����.Enabled = True
        Me.picEdit.Enabled = True: Me.picEdit.Visible = True: Me.txt���쾯ʾ��.Enabled = True
    Else
        Me.vfgList.Height = Me.picEdit.Top + Me.picEdit.Height - Me.vfgList.Top
        Me.txt��������.Enabled = False: Me.txt��������.Enabled = False
        Me.txt���챨����.Enabled = False: Me.txt�ȶԾ�ʾ��.Enabled = False: Me.txt�ȶ�ʧ����.Enabled = False
        Me.picEdit.Enabled = False: Me.picEdit.Visible = False: Me.txt���쾯ʾ��.Enabled = False
    End If
End Sub

Private Sub txt��ע_GotFocus()
    Me.txt��ע.SelStart = 0: Me.txt��ע.SelLength = 1000
    Call zlcommfun.OpenIme(True)
End Sub

Private Sub txt��ע_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlcommfun.PressKey(vbKeyTab)
        '�Զ����µ��б���
        Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Adjust)
        If cbrControl Is Nothing Then Exit Sub
        If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
        Call cbsThis_Execute(cbrControl)
        Exit Sub
    End If
    If InStr(GCST_INVALIDCHAR & ";|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt�ȶԾ�ʾ��_GotFocus()
    Me.txt�ȶԾ�ʾ��.SelStart = 0: Me.txt�ȶԾ�ʾ��.SelLength = 1000
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txt�ȶԾ�ʾ��_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or InStr(".", Chr(KeyAscii)) > 0 Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt�ȶ�ʧ����_GotFocus()
    Me.txt�ȶ�ʧ����.SelStart = 0: Me.txt�ȶ�ʧ����.SelLength = 1000
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txt�ȶ�ʧ����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or InStr(".", Chr(KeyAscii)) > 0 Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt���챨����_GotFocus()
    Me.txt���챨����.SelStart = 0: Me.txt���챨����.SelLength = 1000
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txt���챨����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or InStr(".", Chr(KeyAscii)) > 0 Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt���쾯ʾ��_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or InStr(".", Chr(KeyAscii)) > 0 Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt�ο�ֵ_GotFocus(Index As Integer)
    Me.txt�ο�ֵ(Index).SelStart = 0: Me.txt�ο�ֵ(Index).SelLength = 1000
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txt�ο�ֵ_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or InStr("-.", Chr(KeyAscii)) > 0 Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt��������_GotFocus()
    Me.txt��������.SelStart = 0: Me.txt��������.SelLength = 1000
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txt��������_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or InStr("-.", Chr(KeyAscii)) > 0 Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt��������_GotFocus()
    Me.txt��������.SelStart = 0: Me.txt��������.SelLength = 1000
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txt��������_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or InStr("-.", Chr(KeyAscii)) > 0 Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt��ƫ����_GotFocus()
    Me.txt��ƫ����.SelStart = 0: Me.txt��ƫ����.SelLength = 1000
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txt��ƫ����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or InStr(".", Chr(KeyAscii)) > 0 Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����_GotFocus(Index As Integer)
    Me.txt����(Index).SelStart = 0: Me.txt����(Index).SelLength = 1000
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txt����_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub vfgList_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    With Me.vfgList
        If .Row = .Rows - 1 Then
            Call zlcommfun.PressKey(vbKeyTab)
        Else
            .Row = .Row + 1
        End If
    End With
End Sub

Private Sub vfgList_RowColChange()
    Dim lngList As Long
    With Me.vfgList
        Me.cbo�걾.ListIndex = -1: Me.cbo�Ա�.ListIndex = 0
        Me.txt����(0).Text = "": Me.txt����(1).Text = "": Me.cbo��λ.ListIndex = -1
        Me.cbo�ٴ�����.ListIndex = 0
        Me.txt�ο�ֵ(0).Text = "": Me.txt�ο�ֵ(1).Text = "": Me.txt��ƫ����.Text = ""
        Me.cbo�ο�ֵ.ListIndex = -1
        Me.cbo����.ListIndex = 0: Me.txt��ע.Text = ""
        Me.cbo����.ListIndex = 0
        If .Row = 0 Then Exit Sub
        
        For lngCount = 0 To Me.cbo�걾.ListCount - 1
            If Mid(Me.cbo�걾.List(lngCount), 4) = .TextMatrix(.Row, mCol.�걾) Then Me.cbo�걾.ListIndex = lngCount: Exit For
        Next
        
        For lngCount = 0 To Me.cbo�Ա�.ListCount - 1
            If Val(Left(Me.cbo�Ա�.List(lngCount), 1)) = Val(.TextMatrix(.Row, mCol.�Ա���)) Then Me.cbo�Ա�.ListIndex = lngCount: Exit For
        Next
        
        Me.txt����(0).Text = .TextMatrix(.Row, mCol.��������)
        Me.txt����(1).Text = .TextMatrix(.Row, mCol.��������)
        For lngCount = 0 To Me.cbo��λ.ListCount - 1
            If Mid(Me.cbo��λ.List(lngCount), 3) = .TextMatrix(.Row, mCol.���䵥λ) Then Me.cbo��λ.ListIndex = lngCount: Exit For
        Next
        
        For lngCount = 1 To Me.cbo�ٴ�����.ListCount - 1
            If Mid(Me.cbo�ٴ�����.List(lngCount), 4) = .TextMatrix(.Row, mCol.�ٴ�����) Then Me.cbo�ٴ�����.ListIndex = lngCount: Exit For
        Next
        
        Select Case mInt����
        Case 1, 2, 3
            Me.txt�ο�ֵ(0).Text = FormatReference(mlngItemID, .TextMatrix(.Row, mCol.�ο���ֵ))
            Me.txt�ο�ֵ(1).Text = FormatReference(mlngItemID, .TextMatrix(.Row, mCol.�ο���ֵ))
            
            Me.txt��������.Text = FormatReference(mlngItemID, .TextMatrix(.Row, mCol.��ʾ��ֵ))
            Me.txt��������.Text = FormatReference(mlngItemID, .TextMatrix(.Row, mCol.��ʾ��ֵ))
            Me.txt��������.Text = FormatReference(mlngItemID, .TextMatrix(.Row, mCol.�����ֵ))
            Me.txt��������.Text = FormatReference(mlngItemID, .TextMatrix(.Row, mCol.�����ֵ))
            
            Me.txt��ƫ����.Text = .TextMatrix(.Row, mCol.��ƫ����)
        Case 3
'            For lngCount = 0 To Me.cbo�ο�ֵ.ListCount - 1
'                If lngCount = Val(.TextMatrix(.Row, mcol.�ο���ֵ)) Then Me.cbo�ο�ֵ.ListIndex = lngCount: Exit For
'            Next
'            Me.txt��ƫ����.Text = ""
        End Select
        
        For lngList = 1 To Me.cbo����.ListCount - 1
            If Me.cbo����.ItemData(lngList) = Val(.TextMatrix(.Row, mCol.����Id)) Then Me.cbo����.ListIndex = lngList: Exit For
        Next
        Me.txt��ע.Text = .TextMatrix(.Row, mCol.��ע)
        Me.chkĬ��.Value = Val(.TextMatrix(.Row, mCol.Ĭ��))
        
        For lngList = 1 To Me.cbo����.ListCount - 1
            If Me.cbo����.ItemData(lngList) = Val(.TextMatrix(.Row, mCol.�������Id)) Then Me.cbo����.ListIndex = lngList: Exit For
        Next
    End With
End Sub

Private Function FormatReference(lngID As Long, strReference As String) As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    strSQL = "select max( D.С��λ��)  as С��λ�� from   ���鱨����Ŀ a ,  ������ĿĿ¼ b , ����������Ŀ c ,����������Ŀ d " & _
                     " Where a.������Ŀid = b.ID And a.������Ŀid = c.ID And c.ID = d.��Ŀid And d.С��λ�� Is Not Null " & _
                     " and b.id = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, mlngItemID)
    
    If IsNull(rsTmp(0)) = False And rsTmp(0) > 0 Then
        strReference = Format(strReference, "0." & Replace(Space(rsTmp(0)), " ", "0"))
    End If
    FormatReference = strReference
End Function
