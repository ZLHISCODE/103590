VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmMediPriceDiffCard 
   Caption         =   "����ҩƷ����"
   ClientHeight    =   7995
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14010
   Icon            =   "frmMediPriceDiffCard.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   14010
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picCondition 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   575
      Left            =   120
      ScaleHeight     =   570
      ScaleWidth      =   13575
      TabIndex        =   6
      Top             =   600
      Width           =   13575
      Begin VB.Image imgNote 
         Height          =   480
         Left            =   0
         Picture         =   "frmMediPriceDiffCard.frx":058A
         Top             =   0
         Width           =   480
      End
      Begin VB.Label lblComment 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "˵����1.��ѡ����ȡ���۹���ҩƷ 2.����ҩƷȫԺֻ��һ���۸��ۼۺͳɱ�����ͬ�� 3.ʱ��ҩƷͬ�ⷿ���ε��ۼۺͳɱ���Ҫһ��"
         Height          =   180
         Left            =   600
         TabIndex        =   7
         Top             =   150
         Width           =   10800
      End
   End
   Begin VB.PictureBox picInfo 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   13335
      TabIndex        =   1
      Top             =   7200
      Width           =   13335
      Begin VB.PictureBox picAdjustTime 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   7800
         ScaleHeight     =   375
         ScaleWidth      =   5535
         TabIndex        =   8
         Top             =   120
         Width           =   5535
         Begin VB.OptionButton optʱ�� 
            BackColor       =   &H80000003&
            Caption         =   "ָ������"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   10
            Top             =   15
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optʱ�� 
            BackColor       =   &H80000003&
            Caption         =   "����ִ��"
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   9
            Top             =   15
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker dtpRunDate 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
            Height          =   300
            Left            =   3000
            TabIndex        =   11
            Top             =   0
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
            Format          =   145031171
            CurrentDate     =   36846.5833333333
         End
         Begin VB.Label lblִ��ʱ�� 
            BackColor       =   &H80000003&
            Caption         =   "ִ��ʱ��"
            Height          =   180
            Left            =   0
            TabIndex        =   12
            Top             =   45
            Width           =   855
         End
      End
      Begin VB.TextBox txtSummary 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   4800
         MaxLength       =   100
         TabIndex        =   2
         Top             =   120
         Width           =   2805
      End
      Begin VB.TextBox txtValuer 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   120
         Width           =   1125
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   600
         MaxLength       =   100
         TabIndex        =   15
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "����"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   180
         Width           =   360
      End
      Begin VB.Label lblValuer 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "������"
         Height          =   180
         Left            =   1920
         TabIndex        =   5
         Top             =   180
         Width           =   540
      End
      Begin VB.Label lblSummary 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "����˵��"
         Height          =   180
         Left            =   3960
         TabIndex        =   4
         Top             =   180
         Width           =   720
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfPrice 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   13575
      _cx             =   23945
      _cy             =   8281
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   15191994
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   22
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMediPriceDiffCard.frx":0E54
      ScrollTrack     =   0   'False
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
      Editable        =   2
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
   Begin VSFlex8Ctl.VSFlexGrid vsfPrint 
      Height          =   1695
      Left            =   1080
      TabIndex        =   13
      Top             =   8160
      Visible         =   0   'False
      Width           =   11175
      _cx             =   19711
      _cy             =   2990
      Appearance      =   1
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   0
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMediPriceDiffCard.frx":1199
      ScrollTrack     =   0   'False
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgList 
      Bindings        =   "frmMediPriceDiffCard.frx":1312
      Left            =   480
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMediPriceDiffCard.frx":1326
   End
End
Attribute VB_Name = "frmMediPriceDiffCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���ܰ�ť
Private Const mconMenu_Save = 100 'ȷ��(&A)
Private Const mconMenu_Quit = 101 'ȡ��(&Q)
Private Const mconMenu_PrintStore = 102 '��ӡ���䶯��(&P)
Private Const mconMenu_ClearAll = 103 '����б�(&C)
Private Const mconMenu_ClearAllPrice = 109 '����ּ۸�
Private Const mconMenu_ClearAllDate = 110 '��ս�������
Private Const mconMenu_Adjust = 104 '�Զ����۷�ʽ
Private Const mconMenu_AdjustByCost = 105 '���۷�ʽ���Գɱ���Ϊ׼�����ۼ�
Private Const mconMenu_AdjustByPrice = 106 '���۷�ʽ�����ۼ�Ϊ׼�����ɱ���
Private Const mconMenu_AllDrug = 107 'ѡ�����������۹����ҩƷ
Private Const mconMenu_AllDiff = 108 '��ȡ�����ۼۡ��ɱ��۲���ȵ�ҩƷ
Private Const mconMenu_BatchExtraction = 111 '������ȡ����ҩƷ
Private Const mconMenu_Location = 112 '���ٶ�λ��һ��δ�����۸�ļ�¼����
Private Const mconMenu_Find = 113 '����

Private mintCostDigit As Integer        '�ɱ���С��λ��
Private mintPriceDigit As Integer       '�ۼ�С��λ��
'��ɫ����
'Private Const mconlngColor As Long = &HFFFFFF        '�����޸�����ɫΪ��ɫ
Private Const mconlngCanColColor As Long = &HE7CFBA    '���޸�����ɫΪ����ɫ
Private Const mlngBorderColor As Long = &H0&    'ѡ���б߿���ɫ
Private Const mlngNoneBorderColor As Long = &HE0E0E0    ' ûѡ���б߿���ɫ

Private mstr��¼��ˮ�� As String '��¼����ģʽ
Private marrSql() As Variant     '��¼�洢���̵�����

Private mintUnit As Integer     '������¼���õ���ʲô��λ
Private mstrҩƷID As String
Private mintRow As Integer    'û����д�۸�ľ����к�
Private mrsFindName As ADODB.Recordset '��¼��ѯ���ݼ�
Private mlngFindCurrRow As Long             '��ѯ���ĵ�ǰ��
Private Const MStrCaption As String = "����ҩƷ����"

Private Sub GetPartPriceDiff(Optional bln��ʾ As Boolean = True)
    '��ȡ���������������۹����۸�һ�µ�ҩƷ
    Dim rsData As ADODB.Recordset
    Dim bln����δִ�м۸� As Boolean
    Dim int��� As Integer
    
    On Error GoTo errHandle
       
    Call setNOtExcetePrice
    
    gstrSQL = "Select ҩƷid, ͨ����, ���, 0 As �ⷿid, '' As �ⷿ, ������, '' As ����, ����, ��λ, ��װϵ��, �ۼ�, Sum(�ɱ��� * ʵ������) / Sum(ʵ������) As �ɱ���, �Ƿ�ʱ��," & vbNewLine & _
                    "       �п��, �۸�id, ������Ŀid, Null As �ϴι�Ӧ��id, Null As Ч��" & vbNewLine & _
                    "From (Select a.ҩƷid, '[' || c.���� || ']' || c.���� As ͨ����, c.���, c.���� As ������, 0 As ����," & vbNewLine & _
                    "              Decode([1], 0, a.ҩ�ⵥλ, 2, a.סԺ��λ, 1, a.���ﵥλ, c.���㵥λ) As ��λ," & vbNewLine & _
                    "              Decode([1], 0, a.ҩ���װ, 2, a.סԺ��װ, 1, a.�����װ, 1) As ��װϵ��, b.�ּ� As �ۼ�, decode(d.ƽ���ɱ���,null,a.�ɱ���,d.ƽ���ɱ���) As �ɱ���, 0 As �Ƿ�ʱ��, d.ʵ������," & vbNewLine & _
                    "              1 As �п��, b.Id As �۸�id, b.������Ŀid" & vbNewLine & _
                    "       From ҩƷ��� A, �շѼ�Ŀ B, �շ���ĿĿ¼ C, ҩƷ��� D" & vbNewLine & _
                    "       Where a.ҩƷid = b.�շ�ϸĿid And a.ҩƷid = c.Id And a.ҩƷid = d.ҩƷid And d.���� = 1 And (Sysdate Between b.ִ������ And b.��ֹ����) And" & vbNewLine & _
                    "             (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.�Ƿ��� = 0 And Nvl(a.�Ƿ����۹���, 0) = 1 And" & vbNewLine & _
                    "             b.�ּ� <> decode(d.ƽ���ɱ���,null,a.�ɱ���,d.ƽ���ɱ���) " & vbNewLine & _
                    "  And Not (zl_fun_getbatchpro(d.�ⷿid,d.ҩƷid)=1 And Nvl(d.����,0) = 0 And d.�������� < 0 And d.ʵ������ = 0 And d.ʵ�ʽ�� = 0 And d.ʵ�ʲ�� = 0)) " & vbNewLine & _
                    "Group By ҩƷid, ͨ����, ���, ������, ����, ��λ, ��װϵ��, �ۼ�, �۸�id, ������Ŀid, �Ƿ�ʱ��, �п�� " & vbNewLine & _
                    "Union All "

    gstrSQL = gstrSQL & " Select a.ҩƷid, '[' || c.���� || ']' || c.���� As ͨ����, c.���, d.�ⷿid, e.���� As �ⷿ, d.�ϴβ��� As ������, d.�ϴ����� As ����, d.����," & vbNewLine & _
                    "       Decode([1], 0, a.ҩ�ⵥλ, 2, a.סԺ��λ, 1, a.���ﵥλ, c.���㵥λ) As ��λ," & vbNewLine & _
                    "       Decode([1], 0, a.ҩ���װ, 2, a.סԺ��װ, 1, a.�����װ, 1) As ��װϵ��, d.���ۼ� As �ۼ�, decode(d.ƽ���ɱ���,null,a.�ɱ���,d.ƽ���ɱ���) As �ɱ���, 1 As �Ƿ�ʱ��, 1 As �п��," & vbNewLine & _
                    "       b.Id As �۸�id, b.������Ŀid, nvl(d.�ϴι�Ӧ��id,0) As �ϴι�Ӧ��id, d.Ч��" & vbNewLine & _
                    "From ҩƷ��� A, �շѼ�Ŀ B, �շ���ĿĿ¼ C, ҩƷ��� D, ���ű� E" & vbNewLine & _
                    "Where a.ҩƷid = b.�շ�ϸĿid And a.ҩƷid = c.Id And a.ҩƷid = d.ҩƷid And d.�ⷿid = e.Id And d.���� = 1 And" & vbNewLine & _
                    "      (Sysdate Between b.ִ������ And b.��ֹ����) And c.�Ƿ��� = 1 And" & vbNewLine & _
                    "      (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And Nvl(a.�Ƿ����۹���, 0) = 1 And d.���ۼ� <> decode(d.ƽ���ɱ���,null,a.�ɱ���,d.ƽ���ɱ���)  " & vbNewLine & _
                    "  And Not (zl_fun_getbatchpro(d.�ⷿid,d.ҩƷid)=1 And Nvl(d.����,0) = 0 And d.�������� < 0 And d.ʵ������ = 0 And d.ʵ�ʽ�� = 0 And d.ʵ�ʲ�� = 0) " & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select a.ҩƷid, '[' || c.���� || ']' || c.���� As ͨ����, c.���, 0 As �ⷿid, '' As �ⷿ, '' As ������, '' As ����, 0 As ����," & vbNewLine & _
                    "       Decode([1], 0, a.ҩ�ⵥλ, 2, a.סԺ��λ, 1, a.���ﵥλ, c.���㵥λ) As ��λ," & vbNewLine & _
                    "       Decode([1], 0, a.ҩ���װ, 2, a.סԺ��װ, 1, a.�����װ, 1) As ��װϵ��, b.�ּ� As �ۼ�, a.�ɱ���, c.�Ƿ��� As �Ƿ�ʱ��, 0 As �п��," & vbNewLine & _
                    "       b.Id As �۸�id, b.������Ŀid, Null As �ϴι�Ӧ��id, Null As Ч��" & vbNewLine & _
                    "From ҩƷ��� A, �շѼ�Ŀ B, �շ���ĿĿ¼ C" & vbNewLine & _
                    "Where a.ҩƷid = c.Id And a.ҩƷid = b.�շ�ϸĿid And (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
                    "      Nvl(a.�Ƿ����۹���, 0) = 1 And b.�ּ� <> a.�ɱ��� And (Sysdate Between b.ִ������ And b.��ֹ����) And Not Exists" & vbNewLine & _
                    " (Select 1 From ҩƷ��� D Where d.ҩƷid = a.ҩƷid And d.���� = 1)" & vbNewLine & _
                    "Order By ҩƷid, �ⷿid, ����,����"

    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "GetPartPriceDiff", mintUnit)
    
    With vsfPrice
        .Redraw = flexRDNone
        .MergeCells = flexMergeNever
        .rows = 1
        
        If rsData.RecordCount = 0 Then
            .rows = 2
            .Cell(flexcpText, 1, 1, 1, .Cols - 1) = "û���ҵ����������۹���ģʽ��ҩƷ......"
            .MergeCells = flexMergeRestrictRows
            .MergeRow(1) = True
        Else
            Do While Not rsData.EOF
                '����Ƿ����δִ�м۸�������ھͲ�ȡ����
                If CheckExistExecutePrice(Val(rsData!ҩƷid)) = False Then
                    .rows = .rows + 1
                    
                    .TextMatrix(.rows - 1, .ColIndex("���")) = int��� + 1
                    .TextMatrix(.rows - 1, .ColIndex("ҩƷid")) = rsData!ҩƷid
                    .TextMatrix(.rows - 1, .ColIndex("ҩ������")) = IIf(rsData!�Ƿ�ʱ�� = 1, "ʱ��", "����")
                    .TextMatrix(.rows - 1, .ColIndex("Ʒ��")) = rsData!ͨ����
                    .TextMatrix(.rows - 1, .ColIndex("���")) = rsData!���
                    .TextMatrix(.rows - 1, .ColIndex("������")) = Nvl(rsData!������, "")
                    .TextMatrix(.rows - 1, .ColIndex("�ⷿid")) = rsData!�ⷿid
                    .TextMatrix(.rows - 1, .ColIndex("�ⷿ")) = Nvl(rsData!�ⷿ, "")
                    .TextMatrix(.rows - 1, .ColIndex("����")) = Nvl(rsData!����, "")
                    .TextMatrix(.rows - 1, .ColIndex("��λ")) = rsData!��λ
                    .TextMatrix(.rows - 1, .ColIndex("��װϵ��")) = rsData!��װϵ��
                    .TextMatrix(.rows - 1, .ColIndex("ԭ�ۼ�")) = zlStr.FormatEx(rsData!�ۼ� * rsData!��װϵ��, mintPriceDigit, , True)
                    .TextMatrix(.rows - 1, .ColIndex("ԭ�ɱ���")) = zlStr.FormatEx(rsData!�ɱ��� * rsData!��װϵ��, mintCostDigit, , True)
                    .TextMatrix(.rows - 1, .ColIndex("ԭʼ�ۼ�")) = rsData!�ۼ�
                    .TextMatrix(.rows - 1, .ColIndex("ԭʼ�ɱ���")) = rsData!�ɱ���
                    .TextMatrix(.rows - 1, .ColIndex("�п��")) = rsData!�п��
                    .TextMatrix(.rows - 1, .ColIndex("�۸�id")) = rsData!�۸�id
                    .TextMatrix(.rows - 1, .ColIndex("������Ŀid")) = rsData!������ĿID
                    .TextMatrix(.rows - 1, .ColIndex("����")) = Nvl(rsData!����, 0)
                    .TextMatrix(.rows - 1, .ColIndex("�ϴι�Ӧ��ID")) = Nvl(rsData!�ϴι�Ӧ��ID)
                    .TextMatrix(.rows - 1, .ColIndex("Ч��")) = Nvl(rsData!Ч��)
                    
                    .Cell(flexcpForeColor, .rows - 1, .ColIndex("ҩ������"), .rows - 1, .ColIndex("ҩ������")) = IIf(rsData!�Ƿ�ʱ�� = 1, vbRed, vbBlack)
                    int��� = int��� + 1
                Else
                    bln����δִ�м۸� = True
                End If
                
                rsData.MoveNext
            Loop
            
            If .rows >= 2 Then
                .Cell(flexcpBackColor, 1, .ColIndex("�ּ۸�"), .rows - 1, .ColIndex("�ּ۸�")) = mconlngCanColColor
                .Cell(flexcpForeColor, 1, .ColIndex("�ּ۸�"), .rows - 1, .ColIndex("�ּ۸�")) = vbBlue
                .Cell(flexcpFontBold, 1, .ColIndex("�ּ۸�"), .rows - 1, .ColIndex("�ּ۸�")) = True
            End If
            
            .rows = .rows + 1
            .RowHidden(.rows - 1) = True
        End If
        
        .Redraw = flexRDDirect
    End With
    
    txtValuer.Text = UserInfo.�û�����
    txtSummary.Text = "���۵���"
    
    txtValuer.Tag = "��������ҩƷ"
    
    If bln����δִ�м۸� = True Then
        If bln��ʾ Then
            MsgBox "�������۹���ҩƷ������δִ�е�Ԥ���ۼ�¼�������������������б�����ʾ��ЩҩƷ����ע��鿴��", vbInformation, gstrSysName
        Else
            MsgBox "�������۹���ҩƷ�����������в���ҩƷδ���е��ۣ���ע��鿴��", vbInformation, gstrSysName
        End If
    End If
        
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetBorder()
    '������ѡ�б߿�
    Dim intRow As Integer
    
    With vsfPrice
        If .rows <> 1 Then
            For intRow = 1 To .rows - 2
                .CellBorderRange intRow, 0, intRow, .Cols - 1, mlngNoneBorderColor, 0, 0, 0, 0, 0, 0
            Next
            
            .CellBorderRange .Row, .ColIndex("���"), .Row, .ColIndex("�ּ۸�"), mlngBorderColor, 0, 2, 0, 2, 0, 2
        End If
    End With
End Sub


Private Function CheckExistExecutePrice(ByVal lngDrugID As Long) As Boolean
    '���� ������Ƿ����δִ�еļ۸�
    '���أ�true-����δִ�м۸�false-������δִ�м۸�
    Dim RecCheck As New ADODB.Recordset
    
    On Error GoTo errHandle

    '�ж��Ƿ���δִ�е���ʷ�۸�
    gstrSQL = " Select 1 Records From �շѼ�Ŀ Where �䶯ԭ��=0 And ִ������ > Sysdate And �շ�ϸĿID=[1]"
    Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, "CheckExistExecutePrice", lngDrugID)
    
    If Not RecCheck.EOF Then CheckExistExecutePrice = True: Exit Function
    
    '����Ƿ���δִ�еĳɱ��۵��ۼƻ�
    gstrSQL = "Select 1 From ҩƷ�۸��¼ Where ҩƷid = [1] And ��¼״̬ = 0 And Rownum < 2 "
    Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, "CheckExistExecutePrice", lngDrugID)
    
    If Not RecCheck.EOF Then CheckExistExecutePrice = True: Exit Function
    
    CheckExistExecutePrice = False
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub initCommandBars()
    Dim cbrToolBar As CommandBar
    Dim cbrControl As CommandBarControl
    Dim cbrControlPopu As CommandBarControl
    Dim lngCount As Integer
    
    With CommandBarsGlobalSettings
        .App = App
        .CompanyName = "����������Ϣ��ҵ�������ι�˾" '��˾����
        .ResourceFile = .OcxPath & "\XTPResourceZhCn.dll" '��������������Դ�ļ�
        .ColorManager.SystemTheme = xtpSystemThemeAuto  '�ؼ��������ɫ����
    End With

    With cbsMain.Options
        .ShowExpandButtonAlways = False '�����ڹ������Ҳ���ʾѡ�ť,��ʹ�������㹻��
        .ToolBarAccelTips = True '��ʾ��ť��ʾ
        .AlwaysShowFullMenus = False '�����õĲ˵���������
        .UseFadedIcons = True 'ͼ����ʾΪ��ɫЧ��
        .IconsWithShadow = True '���ָ�������ͼ����ʾ��ӰЧ��
        .UseDisabledIcons = True '��������ť����ʱͼ����ʾΪ������ʽ
        .LargeIcons = True '��������ʾΪ��ͼ��
        .SetIconSize True, 24, 24 '���ô�ͼ��ĳߴ�
        .SetIconSize False, 16, 16 '����Сͼ��ĳߴ�
    End With

    With cbsMain
        .VisualTheme = xtpThemeOffice2003 '���ÿؼ���ʾ���
        .EnableCustomization False '�Ƿ������Զ�������
        Set .Icons = imgList.Icons '���ù�����ͼ��ؼ�
        .ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap '����仯ʱ�������ʾ����˵�Ҳ������
        .ActiveMenuBar.Title = "�˵�"
    End With
    
    'ɾ�����ڵĹ������������˵���
    For lngCount = cbsMain.ActiveMenuBar.Controls.count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsMain.count To 1 Step -1
        cbsMain(lngCount).Delete
    Next
    
    '����������
    Set cbrToolBar = cbsMain.Add("������", xtpBarTop)
    cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    cbrToolBar.ContextMenuPresent = False

    With cbrToolBar
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_PrintStore, "��ӡ���䶯��")
        
        Set cbrControl = .Controls.Add(xtpControlPopup, mconMenu_ClearAll, "���")
        cbrControl.BeginGroup = True
        cbrControl.Id = mconMenu_ClearAll
        cbrControl.IconId = mconMenu_ClearAll
        Set cbrControlPopu = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_ClearAllPrice, "����ּ۸�(&A)", -1, False)
        Set cbrControlPopu = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_ClearAllDate, "��ս�������(&Q)", -1, False)

        Set cbrControl = .Controls.Add(xtpControlPopup, mconMenu_Adjust, "���۷�ʽ")
        cbrControl.BeginGroup = True
        cbrControl.Id = mconMenu_Adjust
        cbrControl.IconId = mconMenu_Adjust
        Set cbrControlPopu = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_AdjustByCost, "���ɱ��۵����ۼ�(&C)", -1, False)
        Set cbrControlPopu = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_AdjustByPrice, "���ۼ۵����ɱ���(&P)", -1, False)
        
        Set cbrControl = .Controls.Add(xtpControlPopup, mconMenu_BatchExtraction, "��ȡ����ҩƷ")
        cbrControl.BeginGroup = True
        cbrControl.Id = mconMenu_BatchExtraction
        cbrControl.IconId = mconMenu_BatchExtraction
        Set cbrControlPopu = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_AllDrug, "������ȡ����ҩƷ(&E)", -1, False)
        Set cbrControlPopu = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_AllDiff, "ֻ��ȡ�ۼۺͳɱ��۲�һ�µ�ҩƷ(&R)", -1, False)

        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_Location, "��λ��δ�����۸����")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_Save, "ȷ��")
        cbrControl.BeginGroup = True
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_Quit, "�˳�")
                
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_Find, "����")
        cbrControl.Visible = False
    End With

    For Each cbrControl In cbrToolBar.Controls  '�ù������а�ťͬʱ��ʾͼ�������
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    With Me.cbsMain.KeyBindings
        .Add 0, VK_F3, mconMenu_Find
    End With

End Sub

Private Sub SavePriceAdjust()
    '�����ִ�е���
    Dim int�޸ļ�¼ As Integer
    Dim i As Integer
    Dim Array��ˮ�� As Variant
    Dim blnTrans As Boolean
  
    On Error GoTo ErrHand
    
    marrSql = Array()
    Array��ˮ�� = Array()
    mstr��¼��ˮ�� = ""
    
    If vsfPrice.rows <= 1 Then Exit Sub
    
    With vsfPrice
        '���۸��Ƿ�ȫΪ��
        For i = 1 To .rows - 2
            If .TextMatrix(i, .ColIndex("�ּ۸�")) = "" Then
                int�޸ļ�¼ = int�޸ļ�¼ + 1
                If int�޸ļ�¼ = .rows - 2 Then
                    MsgBox "�����е�ҩƷ�ּ۸�Ϊ�գ�����ִ�е��ۣ�", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        Next
        
        '����ּ۸��Ƿ���ԭ�ۼۡ�ԭ�ɱ��۶����
        For i = 1 To .rows - 2
            If .TextMatrix(i, .ColIndex("�ּ۸�")) = .TextMatrix(i, .ColIndex("ԭ�ۼ�")) And .TextMatrix(i, .ColIndex("�ּ۸�")) = .TextMatrix(i, .ColIndex("ԭ�ɱ���")) Then
                MsgBox "�ڡ�" & i & "�����ּ۸���ԭ�ۼۺ�ԭ�ɱ�����ȣ�����ִ�е��ۣ�", vbInformation, gstrSysName
                .Select i, .ColIndex("�ּ۸�")
                Exit Sub
            End If
        Next
        
        '����ּ��Ƿ�̫��
        For i = 1 To .rows - 2
            If Val(.TextMatrix(i, .ColIndex("�ּ۸�"))) > 100000 Then
                MsgBox "�ڡ�" & i & "��������ļ۸�������������룡", vbInformation, gstrSysName
                .Select i, .ColIndex("�ּ۸�")
                Exit Sub
            End If
        Next

    End With
    
    Call ModifyCostPrice          '���ɱ���
    Call ModifyRetailPrice        '���ۼ�
    Call ModifyAllPrice             '�ɱ����ۼ�һ���
             
    Array��ˮ�� = Split(Mid(mstr��¼��ˮ��, 2), ";")
    
    For i = 0 To UBound(Array��ˮ��)
        '��ˮ��
        gstrSQL = "Zl_���ۻ��ܼ�¼_Insert(" & Split(Array��ˮ��(i), "|")(0) & ","
        '����
        gstrSQL = gstrSQL & Split(Array��ˮ��(i), "|")(1) & ","
        'ִ������
        If optʱ��(0).Value = True Then
            gstrSQL = gstrSQL & "sysdate" & ","
        Else
            gstrSQL = gstrSQL & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
        End If
        '˵��
        gstrSQL = gstrSQL & "'" & MoveSpecialChar(txtSummary.Text) & "',"
        '���ࡢ������
        gstrSQL = gstrSQL & "0,'" & MoveSpecialChar(txtValuer.Text) & "')"
        
        ReDim Preserve marrSql(UBound(marrSql) + 1)
        marrSql(UBound(marrSql)) = gstrSQL
    Next
                   
    gcnOracle.BeginTrans: blnTrans = True          '��������
    For i = 0 To UBound(marrSql)
        Call zlDatabase.ExecuteProcedure(CStr(marrSql(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False     '�ύ����
    
    If int�޸ļ�¼ = 0 Then
        Unload Me
    ElseIf int�޸ļ�¼ <> 0 Then
        If txtValuer.Tag = "ȫ������ҩƷ" Then
            Call GetAllPriceDiff(False)
        Else
            Call GetPartPriceDiff(False)
        End If
    End If
        
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub SetAdjust(ByVal intAdjustType As Integer)
    '�������õ���
    'intAdjustType��0-��ԭ�ɱ���Ϊ׼�����ۼۣ�1-��ԭ�ۼ�Ϊ׼�����ɱ���
    Dim i As Integer
    
    With vsfPrice
        If .rows <= 1 Then Exit Sub
        If Val(.TextMatrix(1, .ColIndex("ҩƷid"))) = 0 Then Exit Sub
        
        If intAdjustType = 0 Then
            If MsgBox("����������ԭ�ɱ�����Ϊ�ּ۸��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Else
            If MsgBox("����������ԭ�ۼ���Ϊ�ּ۸��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        .Redraw = flexRDNone
        
        For i = 1 To .rows - 2
            .TextMatrix(i, .ColIndex("�ּ۸�")) = IIf(intAdjustType = 0, .TextMatrix(i, .ColIndex("ԭ�ɱ���")), .TextMatrix(i, .ColIndex("ԭ�ۼ�")))
        Next
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case mconMenu_Save  'ִ�е���
            Call SavePriceAdjust
        Case mconMenu_PrintStore    '��ӡ���䶯��
            Call PrintPrice            '��ӡ
        Case mconMenu_AdjustByCost  '���۷�ʽ���Գɱ���Ϊ׼�����ۼ�
            Call SetAdjust(0)
        Case mconMenu_AdjustByPrice  '���۷�ʽ�����ۼ�Ϊ׼�����ɱ���
            Call SetAdjust(1)
        Case mconMenu_ClearAllPrice  '��������ּ۸�
            Call ClearAllPrice
        Case mconMenu_ClearAllDate  '��ս�������
            Call ClearAllDate
            mintRow = 1
        Case mconMenu_AllDrug  'ѡ������ҩƷ
            If txtValuer.Tag = "��������ҩƷ" And vsfPrice.rows > 1 Then
                 If MsgBox("�ò���������������ݲ�������ȡ��ѡ������۹���ҩƷ���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
            Call AllDrug
            mintRow = 1
        Case mconMenu_AllDiff  '��ȡ�۸񲻵�����ҩƷ
            If vsfPrice.rows > 1 Then
                If MsgBox("�ò�������ս������ݲ�ֻ��ȡ�ۼۺͳɱ��۲�һ�µ�ҩƷ���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
            Call GetPartPriceDiff
            mstrҩƷID = ""
            mintRow = 1
        Case mconMenu_Location '���ٶ�λ��һ��δ�����۸�ļ�¼����
            Call FindLocation
        Case mconMenu_Find '����
            txtCode.SetFocus
            If Trim(txtCode.Text) <> "" Then Call FindGridRow(txtCode.Text)
        Case mconMenu_Quit  'ȡ��
            If vsfPrice.rows > 1 Then
                If Val(vsfPrice.TextMatrix(1, vsfPrice.ColIndex("ҩƷid"))) > 0 Then
                    If MsgBox("δ���汾�ε��ۣ��Ƿ��˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    
                    Unload Me
                Else
                    Unload Me
                End If
            Else
                Unload Me
            End If
            
            mintRow = 1
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    Me.picCondition.Move lngLeft, lngTop, lngRight - lngLeft
    
    Me.picInfo.Move lngLeft, Me.ScaleHeight - Me.picInfo.Height, lngRight - lngLeft
    
    Me.vsfPrice.Move lngLeft, picCondition.Top + picCondition.Height + 50, lngRight - lngLeft, Me.picInfo.Top - Me.picCondition.Top - Me.picCondition.Height - 100
End Sub

Private Sub Form_Load()
    Dim intUnitTemp As Integer
    
    '��ȡ���õĵ�λ
    mintUnit = Val(zlDatabase.GetPara("ҩƷ��λ", glngSys, 1333, "1"))
    
    Select Case mintUnit
        Case 0 'ҩ��
            intUnitTemp = 4
        Case 1 '����
            intUnitTemp = 2
        Case 2 'סԺ
            intUnitTemp = 3
        Case 3 '�ۼ�
            intUnitTemp = 1
    End Select
    '��ȡ������λ����
    mintCostDigit = GetDigitTiaoJia(1, 1, intUnitTemp)
    mintPriceDigit = GetDigitTiaoJia(1, 2, intUnitTemp)
 
    '��ʼ��ʱ��Ϊ��ǰʱ��+1��
    dtpRunDate.Value = DateAdd("d", 1, CDate(Format(Sys.Currentdate(), "yyyy-MM-dd hh:mm:ss")))
    
    Call initCommandBars
    Call RestoreWinState(Me, App.ProductName, MStrCaption)
    
    vsfPrice.rows = 1
    mlngFindCurrRow = 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrҩƷID = ""
    mintRow = 1
End Sub

Private Sub optʱ��_Click(Index As Integer)
    If Index = 0 Then
        dtpRunDate.Enabled = False
    Else
        dtpRunDate.Enabled = True
    End If
End Sub

Private Sub picCondition_Resize()
    On Error Resume Next
    
    With lblComment
        .Left = 50
        .Height = picCondition.Width - 50
    End With
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    
    With picAdjustTime
        .Left = picInfo.Width - .Width - 100
    End With

    With txtSummary
        .Width = picAdjustTime.Left - .Left - 100
    End With
    
End Sub

Private Sub vsfPrice_EnterCell()
    With vsfPrice
        .Editable = flexEDNone
        If .Col = .ColIndex("�ּ۸�") Then
            .FocusRect = flexFocusSolid
            .Editable = flexEDKbdMouse
        Else
            .FocusRect = flexFocusLight
        End If
        
        Call SetBorder '������ѡ�б߿�
    End With
End Sub

Private Sub vsfPrice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strkey As String
    Dim intDigit As Integer

    With vsfPrice
        strkey = .EditText
        If Col = .ColIndex("�ּ۸�") Then
            If KeyAscii = vbKeyReturn Then
                .EditText = zlStr.FormatEx(.EditText, mintPriceDigit, , True)
                If Row <> .rows - 2 Then
                    .Row = Row + 1
                    .Col = Col
                End If
                Exit Sub
            End If
            
            If KeyAscii <> vbKeyBack Then
                If KeyAscii = vbKeyDelete Then
                    If InStr(1, strkey, ".") > 0 Then
                        KeyAscii = 0
                    End If
                ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                    If vsfPrice.EditSelLength = Len(strkey) Then Exit Sub
                    If InStr(strkey, ".") <> 0 And Chr(KeyAscii) = "." Then   'ֻ�ܴ���һ��С����
                        KeyAscii = 0
                        Exit Sub
                    End If
'                    If Len(Mid(strkey, InStr(1, strkey, ".") + 1)) > mintPriceDigit And strkey Like "*.*" Then
'                        KeyAscii = 0
'                        Exit Sub
'                    Else
'                        Exit Sub
'                    End If
                Else
                    KeyAscii = 0
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfPrice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim dbl�ּ۸� As Double
    Dim dblԭʼ�ۼ� As Double
    Dim dblԭʼ�ɱ��� As Double
    Dim intRow As Integer
    
    With vsfPrice
        If Col = .ColIndex("�ּ۸�") Then
            If Trim(.EditText) = "" Then Exit Sub
            
            .EditText = zlStr.FormatEx(Val(.EditText), mintPriceDigit, , True)
            .TextMatrix(Row, .ColIndex("�ּ۸�")) = .EditText
            
            dbl�ּ۸� = Val(zlStr.FormatEx(Val(.TextMatrix(Row, .ColIndex("�ּ۸�"))) / Val(.TextMatrix(Row, .ColIndex("��װϵ��"))), gtype_UserDrugDigits.Digit_���ۼ�, , True))
            dblԭʼ�ۼ� = Val(.TextMatrix(Row, .ColIndex("ԭʼ�ۼ�")))
            dblԭʼ�ɱ��� = Val(.TextMatrix(Row, .ColIndex("ԭʼ�ɱ���")))
            
            If dbl�ּ۸� = dblԭʼ�ۼ� And dbl�ּ۸� = dblԭʼ�ɱ��� Then
                MsgBox "ע�⣺���ۼۺ�ԭ��һ���ˣ�������¼�룡", vbInformation, gstrSysName
                Cancel = True
                Exit Sub
            End If
        End If
    End With
End Sub

Private Sub ClearAllPrice()
    '��������ּ۸�
    Dim i As Integer
    
    If vsfPrice.rows <= 1 Then Exit Sub
    If MsgBox("δ���汾�ε��ۣ��Ƿ�������м۸�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

    For i = 1 To vsfPrice.rows - 2
        vsfPrice.TextMatrix(i, vsfPrice.ColIndex("�ּ۸�")) = ""
    Next
End Sub

Private Sub ModifyCostPrice()
    '���ɱ���
    Dim i As Integer
    Dim bln�Ƿ�����ɱ��� As Boolean
    Dim strCost��ˮ�� As String
    Dim str�����ɱ���ID As String
    Dim Array���ɱ���ID As Variant
    Dim rsTemp As ADODB.Recordset
    Dim dbl��װ As Double
    Dim dtToday As Date
    
    Array���ɱ���ID = Array()
    On Error GoTo ErrHand
    
    With vsfPrice
        For i = 1 To .rows - 2
            If Val(.TextMatrix(i, .ColIndex("ԭ�ɱ���"))) <> Val(.TextMatrix(i, .ColIndex("�ּ۸�"))) And Val(.TextMatrix(i, .ColIndex("ԭ�ۼ�"))) = Val(.TextMatrix(i, .ColIndex("�ּ۸�"))) And .TextMatrix(i, .ColIndex("�ּ۸�")) <> "" Then
                bln�Ƿ�����ɱ��� = True
                If strCost��ˮ�� = "" Then
                    gstrSQL = "select nextno(135) as ��ˮ�� from dual"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "������ˮ��")
                    If rsTemp.RecordCount = 0 Then
                        MsgBox "������ˮ��δ�ܳ�ʼ���ɹ����������Ա��ϵ��", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    strCost��ˮ�� = rsTemp!��ˮ��

                    mstr��¼��ˮ�� = mstr��¼��ˮ�� & ";" & strCost��ˮ�� & "|" & 1
                    dtToday = Sys.Currentdate() - 1 / 24 / 60 / 60
                End If

                If InStr(str�����ɱ���ID & ";", ";" & Val(.TextMatrix(i, .ColIndex("ҩƷID"))) & ";") = 0 Then
                    str�����ɱ���ID = str�����ɱ���ID & ";" & Val(.TextMatrix(i, .ColIndex("ҩƷID")))
                End If

                dbl��װ = Val(vsfPrice.TextMatrix(i, vsfPrice.ColIndex("��װϵ��")))
                
                If .TextMatrix(i, .ColIndex("�п��")) = 1 And .TextMatrix(i, .ColIndex("ҩ������")) = "����" Then
                    gstrSQL = "Select s.�ⷿid, s.ҩƷid, d.���� As �ⷿ, '[' || m.���� || ']' || m.���� As ҩƷ, m.���,s.�ϴβ��� as ����," & vbNewLine & _
                                    "       Decode([2], 0, p.ҩ�ⵥλ, 2, p.סԺ��λ, 1, p.���ﵥλ, m.���㵥λ) As ��λ," & vbNewLine & _
                                    "       Decode([2], 0, p.ҩ���װ, 2, p.סԺ��װ, 1, p.�����װ, 1) As ��װϵ��," & vbNewLine & _
                                    "       s.�ϴ����� As ����, Nvl(s.ʵ������, 0) As ����, Nvl(s.����,0) as ����," & vbNewLine & _
                                    "       Nvl(m.�Ƿ���, 0) ���, m.Id, Decode(Nvl(m.�Ƿ���, 0), 0, e.�ּ�, Decode(Nvl(s.���ۼ�, 0),0,s.ʵ�ʽ��/s.ʵ������,s.���ۼ�)) As ʱ���ۼ�, p.�ӳ���, Decode(Nvl(s.ƽ���ɱ���, 0), 0, p.�ɱ���, s.ƽ���ɱ���) As �ɱ���, nvl(s.�ϴι�Ӧ��id,0) As �ϴι�Ӧ��id," & vbNewLine & _
                                    "       n.���� As ��Ӧ��, s.Ч��" & vbNewLine & _
                                    "From ҩƷ��� S, ���ű� D, �շ���ĿĿ¼ M, ҩƷ��� P, ��Ӧ�� N, �շѼ�Ŀ E" & vbNewLine & _
                                    "Where d.Id = s.�ⷿid And s.ҩƷid = m.Id And m.Id = p.ҩƷid And Nvl(s.�ϴι�Ӧ��id, 0) = n.Id(+) And m.Id = e.�շ�ϸĿid And" & vbNewLine & _
                                    "      s.���� = 1 And s.ҩƷid = [1] And Sysdate Between e.ִ������ And e.��ֹ����  And e.�۸�ȼ� Is Null" & vbNewLine & _
                                    "Order By s.ҩƷid,s.�ⷿid, s.�ϴ�����,s.���� "

                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ҩƷ��Ϣ", Val(.TextMatrix(i, .ColIndex("ҩƷID"))), mintUnit)
                    
                    With rsTemp
                        Do While Not .EOF
                            gstrSQL = "Zl_ҩƷ�۸��¼_Stop("
                            '�۸�����_In
                            gstrSQL = gstrSQL & 2
                            '�ⷿid_In
                            gstrSQL = gstrSQL & "," & !�ⷿid
                            'ҩƷid_In
                            gstrSQL = gstrSQL & "," & !ҩƷid
                            '����_In
                            gstrSQL = gstrSQL & "," & Nvl(!����, 0)
                            '��ֹ����_In
                            If optʱ��(0).Value = True Then
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            gstrSQL = gstrSQL & ")"
                            
                            ReDim Preserve marrSql(UBound(marrSql) + 1)
                            marrSql(UBound(marrSql)) = gstrSQL
                            
                            gstrSQL = "Zl_ҩƷ�۸��¼_Insert("
                            '��������_In
                            gstrSQL = gstrSQL & 1
                            '�۸�����_In
                            gstrSQL = gstrSQL & ",2"
                            '�ⷿid_In
                            gstrSQL = gstrSQL & "," & !�ⷿid
                            'ҩƷid_In
                            gstrSQL = gstrSQL & "," & !ҩƷid
                            '����_In
                            gstrSQL = gstrSQL & "," & Nvl(!����, 0)
                            'ԭ��_In
                            gstrSQL = gstrSQL & "," & zlStr.FormatEx(!�ɱ���, gtype_UserDrugDigits.Digit_���ۼ�, , True)
                            '�ּ�_In
                            gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(vsfPrice.TextMatrix(i, vsfPrice.ColIndex("�ּ۸�"))) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�, , True)
                            'ִ������_In
                            If optʱ��(0).Value = True Then
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            '����˵��_In
                            gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                            '������_In
                            gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                            '���ۻ��ܺ�_In
                            gstrSQL = gstrSQL & ",'" & strCost��ˮ�� & "'"
                            '��ҩ��λid_In
                            gstrSQL = gstrSQL & "," & IIf(!�ϴι�Ӧ��ID = 0, "Null", !�ϴι�Ӧ��ID)
                            '����_In
                            gstrSQL = gstrSQL & ",'" & Nvl(!����) & "'"
                            'Ч��_In
                            gstrSQL = gstrSQL & "," & "to_date('" & Format(IIf(IsNull(!Ч��), "", !Ч��), "YYYY-MM-DD") & "','yyyy-mm-dd')"
                            '����_In
                            gstrSQL = gstrSQL & ",'" & Nvl(!����) & "'"
                            gstrSQL = gstrSQL & ")"
                            
                            ReDim Preserve marrSql(UBound(marrSql) + 1)
                            marrSql(UBound(marrSql)) = gstrSQL
                            
                        .MoveNext
                        Loop
                    End With
                End If
                            
                If .TextMatrix(i, .ColIndex("�п��")) = 1 And .TextMatrix(i, .ColIndex("ҩ������")) = "ʱ��" Then
                    '�п��
                    gstrSQL = "Zl_ҩƷ�۸��¼_Stop("
                    '�۸�����_In
                    gstrSQL = gstrSQL & 2
                    '�ⷿid_In
                    gstrSQL = gstrSQL & "," & Val(.TextMatrix(i, .ColIndex("�ⷿid")))
                    'ҩƷid_In
                    gstrSQL = gstrSQL & "," & Val(.TextMatrix(i, .ColIndex("ҩƷid")))
                    '����_In
                    gstrSQL = gstrSQL & "," & Nvl(.TextMatrix(i, .ColIndex("����")), 0)
                    '��ֹ����_In
                    If optʱ��(0).Value = True Then
                        gstrSQL = gstrSQL & "," & "to_date('" & Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        gstrSQL = gstrSQL & "," & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    gstrSQL = gstrSQL & ")"
                    
                    ReDim Preserve marrSql(UBound(marrSql) + 1)
                    marrSql(UBound(marrSql)) = gstrSQL
                    
                    gstrSQL = "Zl_ҩƷ�۸��¼_Insert("
                    '��������_In
                    gstrSQL = gstrSQL & 1
                    '�۸�����_In
                    gstrSQL = gstrSQL & ",2"
                    '�ⷿid_In
                    gstrSQL = gstrSQL & "," & Val(.TextMatrix(i, .ColIndex("�ⷿid")))
                    'ҩƷid_In
                    gstrSQL = gstrSQL & "," & Val(.TextMatrix(i, .ColIndex("ҩƷid")))
                    '����_In
                    gstrSQL = gstrSQL & "," & Nvl(.TextMatrix(i, .ColIndex("����")), 0)
                    'ԭ��_In
                    gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(i, .ColIndex("ԭ�ɱ���"))) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�, , True)
                    '�ּ�_In
                    gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(i, .ColIndex("�ּ۸�"))) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�, , True)
                    'ִ������_In
                    If optʱ��(0).Value = True Then
                        gstrSQL = gstrSQL & "," & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        gstrSQL = gstrSQL & "," & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    '����˵��_In
                    gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                    '������_In
                    gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                    '���ۻ��ܺ�_In
                    gstrSQL = gstrSQL & ",'" & strCost��ˮ�� & "'"
                    '��ҩ��λid_In
                    gstrSQL = gstrSQL & "," & IIf(Val(.TextMatrix(i, .ColIndex("�ϴι�Ӧ��ID"))) = 0, "Null", Val(.TextMatrix(i, .ColIndex("�ϴι�Ӧ��ID"))))
                    '����_In
                    gstrSQL = gstrSQL & ",'" & Nvl(.TextMatrix(i, .ColIndex("����"))) & "'"
                    'Ч��_In
                    gstrSQL = gstrSQL & "," & "to_date('" & Format(IIf(IsNull(.TextMatrix(i, .ColIndex("Ч��"))), "", .TextMatrix(i, .ColIndex("Ч��"))), "YYYY-MM-DD") & "','yyyy-mm-dd')"
                    '����_In
                    gstrSQL = gstrSQL & ",'" & Nvl(.TextMatrix(i, .ColIndex("������"))) & "'"
                    gstrSQL = gstrSQL & ")"
                    
                    ReDim Preserve marrSql(UBound(marrSql) + 1)
                    marrSql(UBound(marrSql)) = gstrSQL
                End If
                    
                If .TextMatrix(i, .ColIndex("�п��")) = 0 Then
                    '�޿��
                    gstrSQL = "Zl_ҩƷ�۸��¼_Insert("
                    '��������_In
                    gstrSQL = gstrSQL & 1
                    '�۸�����_In
                    gstrSQL = gstrSQL & ",2"
                    '�ⷿid_In
                    gstrSQL = gstrSQL & ",Null"
                    'ҩƷid_In
                    gstrSQL = gstrSQL & "," & Val(.TextMatrix(i, .ColIndex("ҩƷid")))
                    '����_In
                    gstrSQL = gstrSQL & ",0"
                    'ԭ��_In
                    gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(i, .ColIndex("ԭ�ɱ���"))) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�, , True)
                    '�ּ�_In
                    gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(i, .ColIndex("�ּ۸�"))) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�, , True)
                    'ִ������_In
                    If optʱ��(0).Value = True Then
                        gstrSQL = gstrSQL & "," & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        gstrSQL = gstrSQL & "," & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    '����˵��_In
                    gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                    '������_In
                    gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                    '���ۻ��ܺ�_In
                    gstrSQL = gstrSQL & ",'" & strCost��ˮ�� & "'"
                    gstrSQL = gstrSQL & ")"
                    
                    ReDim Preserve marrSql(UBound(marrSql) + 1)
                    marrSql(UBound(marrSql)) = gstrSQL
                End If
  
            End If
        Next
    End With
    
    If optʱ��(0).Value = True Then
        Array���ɱ���ID = Split(Mid(str�����ɱ���ID, 2), ";")
        If bln�Ƿ�����ɱ��� Then
            For i = 0 To UBound(Array���ɱ���ID)
                gstrSQL = "zl_ҩƷ�շ���¼_Adjust(" & Array���ɱ���ID(i) & ",2)"
                ReDim Preserve marrSql(UBound(marrSql) + 1)
                marrSql(UBound(marrSql)) = gstrSQL
            Next
            bln�Ƿ�����ɱ��� = False
        End If
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ModifyRetailPrice()
    '���ۼ�
    Dim i As Integer
    Dim n As Integer
    Dim intͬҩƷID�� As Integer
    Dim int�շѼ�Ŀ��� As Integer
    Dim bln�Ƿ�����ۼ� As Boolean
    Dim dbl���ۼ� As Double
    Dim strRetail��ˮ�� As String
    Dim lngAdjId As Long
    Dim dtToday As Date
    Dim rsTemp As ADODB.Recordset
    Dim strNo As String
    Dim LngCurID As Long
    Dim dbl��װ As Double
    
    On Error GoTo ErrHand
    
     With vsfPrice
        For i = 1 To .rows - 2
            If Val(.TextMatrix(i, .ColIndex("ҩƷid"))) = Val(.TextMatrix(i + 1, .ColIndex("ҩƷid"))) Then
                intͬҩƷID�� = intͬҩƷID�� + 1
            Else
                For n = i - intͬҩƷID�� To i
                    dbl���ۼ� = dbl���ۼ� + Val(.TextMatrix(n, .ColIndex("�ּ۸�")))
                    If Val(.TextMatrix(n, .ColIndex("ԭ�ۼ�"))) <> Val(.TextMatrix(n, .ColIndex("�ּ۸�"))) And Val(.TextMatrix(n, .ColIndex("ԭ�ɱ���"))) = Val(.TextMatrix(n, .ColIndex("�ּ۸�"))) And .TextMatrix(i, .ColIndex("�ּ۸�")) <> "" Then
                        bln�Ƿ�����ۼ� = True
                    End If
                Next

                If bln�Ƿ�����ۼ� Then
                    If strRetail��ˮ�� = "" Then
                        strNo = Sys.GetNextNo(9)
                        gstrSQL = "select nextno(135) as ��ˮ�� from dual"
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "������ˮ��")
                        If rsTemp.RecordCount = 0 Then
                            MsgBox "������ˮ��δ�ܳ�ʼ���ɹ����������Ա��ϵ��", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        strRetail��ˮ�� = rsTemp!��ˮ��

                        gstrSQL = "select �շѼ�Ŀ_ID.nextval from dual"
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�շѼ�Ŀ���")
                        lngAdjId = rsTemp.Fields(0).Value

                        mstr��¼��ˮ�� = mstr��¼��ˮ�� & ";" & strRetail��ˮ�� & "|" & 0
                        dtToday = Sys.Currentdate() - 1 / 24 / 60 / 60
                    End If
                    
                    int�շѼ�Ŀ��� = int�շѼ�Ŀ��� + 1
                    dbl���ۼ� = Round(dbl���ۼ� / (intͬҩƷID�� + 1), 2)
                    LngCurID = Sys.NextId("�շѼ�Ŀ")
                    dbl��װ = Val(.TextMatrix(i, .ColIndex("��װϵ��")))
            
                    If CLng(.TextMatrix(i, .ColIndex("�۸�ID"))) <> 0 Then
                        '������һ�εļ۸��¼��ִֹ��
                        gstrSQL = "zl_�շѼ�Ŀ_stop(" & .TextMatrix(i, .ColIndex("ҩƷid")) & ","
                        If optʱ��(0).Value = True Then
                            gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                        Else
                            gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -2, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                        End If
                        gstrSQL = gstrSQL & ")"
                        
                        ReDim Preserve marrSql(UBound(marrSql) + 1)
                        marrSql(UBound(marrSql)) = gstrSQL
            
                        '�����۸��¼
                        'ID
                        gstrSQL = "zl_�շѼ�Ŀ_Insert(" & LngCurID & ","
                        'ԭ��ID
                        gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("�۸�ID"))) & ","
                        '�շ�ϸĿID
                        gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("ҩƷid"))) & ","
                        '������ĿID
                        gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("������ĿID"))) & ","
                        'ԭ��
                        gstrSQL = gstrSQL & zlStr.FormatEx(Val(.TextMatrix(i, .ColIndex("ԭ�ۼ�"))) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�, , True) & ","
                        '�ּ�
                        gstrSQL = gstrSQL & zlStr.FormatEx(dbl���ۼ� / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�, , True) & ","
                        '�����շ��ʡ��Ӱ�Ӽ��ʡ�����˵��
                        gstrSQL = gstrSQL & "NULL,NULL,'" & MoveSpecialChar(txtSummary.Text) & "',"
                        '����id��������
                        gstrSQL = gstrSQL & lngAdjId & ",'" & MoveSpecialChar(txtValuer.Text) & "',"
                        'ִ������
                        If optʱ��(0).Value = True Then
                            gstrSQL = gstrSQL & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                        Else
                            gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                        End If
                        '�䶯ԭ��
                        gstrSQL = gstrSQL & "0,"
                        'NO
                        gstrSQL = gstrSQL & "'" & strNo & "',"
                        '��š�ȱʡ�۸�
                        gstrSQL = gstrSQL & int�շѼ�Ŀ��� & ",Null,"
                        '���ۻ��ܺ�
                        gstrSQL = gstrSQL & strRetail��ˮ�� & ")"
                        
                        ReDim Preserve marrSql(UBound(marrSql) + 1)
                        marrSql(UBound(marrSql)) = gstrSQL
                    End If
                    
                    For n = i - intͬҩƷID�� To i
                        If Val(.TextMatrix(n, .ColIndex("ԭ�ۼ�"))) <> Val(.TextMatrix(n, .ColIndex("�ּ۸�"))) And Val(.TextMatrix(n, .ColIndex("ԭ�ɱ���"))) = Val(.TextMatrix(n, .ColIndex("�ּ۸�"))) And .TextMatrix(i, .ColIndex("�ּ۸�")) <> "" And .TextMatrix(n, .ColIndex("ҩ������")) = "ʱ��" And .TextMatrix(n, .ColIndex("�п��")) = 1 Then
                            'ʱ��ҩƷ�п�����
                            gstrSQL = "Zl_ҩƷ�۸��¼_Stop("
                            '�۸�����_In
                            gstrSQL = gstrSQL & 1
                            '�ⷿid_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("�ⷿID")))
                            'ҩƷid_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("ҩƷID")))
                            '����_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("����")))
                            '��ֹ����_In
                            If optʱ��(0).Value = True Then
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            gstrSQL = gstrSQL & ")"
                            
                            ReDim Preserve marrSql(UBound(marrSql) + 1)
                            marrSql(UBound(marrSql)) = gstrSQL
                            
                            gstrSQL = "Zl_ҩƷ�۸��¼_Insert("
                            '��������_In
                            gstrSQL = gstrSQL & 1
                            '�۸�����_In
                            gstrSQL = gstrSQL & ",1"
                            '�ⷿid_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("�ⷿID")))
                            'ҩƷid_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("ҩƷID")))
                            '����_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("����")))
                            'ԭ��_In
                            gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(n, .ColIndex("ԭ�ۼ�"))) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�, , True)
                            '�ּ�_In
                            gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(n, .ColIndex("�ּ۸�"))) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�, , True)
                            'ִ������_In
                            If optʱ��(0).Value = True Then
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            '����˵��_In
                            gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                            '������_In
                            gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                            '���ۻ��ܺ�_In
                            gstrSQL = gstrSQL & ",'" & strRetail��ˮ�� & "'"
                            '��ҩ��λid_In
                            gstrSQL = gstrSQL & "," & IIf(Val(.TextMatrix(n, .ColIndex("�ϴι�Ӧ��ID"))) = 0, "null", Val(.TextMatrix(n, .ColIndex("�ϴι�Ӧ��ID"))))
                            '����_In
                            gstrSQL = gstrSQL & ",'" & Nvl(.TextMatrix(n, .ColIndex("����"))) & "'"
                            'Ч��_In
                            gstrSQL = gstrSQL & "," & "to_date('" & Format(Nvl(.TextMatrix(n, .ColIndex("Ч��"))), "YYYY-MM-DD") & "','yyyy-mm-dd')"
                            '����_In
                            gstrSQL = gstrSQL & ",'" & Nvl(.TextMatrix(n, .ColIndex("������"))) & "'"
                            gstrSQL = gstrSQL & ")"
                            
                            ReDim Preserve marrSql(UBound(marrSql) + 1)
                            marrSql(UBound(marrSql)) = gstrSQL
                        End If
                        
                        If Val(.TextMatrix(n, .ColIndex("ԭ�ۼ�"))) <> Val(.TextMatrix(n, .ColIndex("�ּ۸�"))) And Val(.TextMatrix(n, .ColIndex("ԭ�ɱ���"))) = Val(.TextMatrix(n, .ColIndex("�ּ۸�"))) And .TextMatrix(i, .ColIndex("�ּ۸�")) <> "" And .TextMatrix(n, .ColIndex("ҩ������")) = "ʱ��" And .TextMatrix(n, .ColIndex("�п��")) = 0 Then
                            'ʱ��ҩƷ�޿�����
                            gstrSQL = "Zl_ҩƷ�۸��¼_Insert("
                            '��������_In
                            gstrSQL = gstrSQL & 1
                            '�۸�����_In
                            gstrSQL = gstrSQL & ",1"
                            '�ⷿid_In
                            gstrSQL = gstrSQL & ",Null"
                            'ҩƷid_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("ҩƷID")))
                            '����_In
                            gstrSQL = gstrSQL & ",0"
                            'ԭ��_In
                            gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(n, .ColIndex("ԭ�ۼ�"))) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�, , True)
                            '�ּ�_In
                            gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(n, .ColIndex("�ּ۸�"))) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�, , True)
                            'ִ������_In
                            If optʱ��(0).Value = True Then
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            '����˵��_In
                            gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                            '������_In
                            gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                            '���ۻ��ܺ�_In
                            gstrSQL = gstrSQL & ",'" & strRetail��ˮ�� & "'"
                            gstrSQL = gstrSQL & ")"

                            ReDim Preserve marrSql(UBound(marrSql) + 1)
                            marrSql(UBound(marrSql)) = gstrSQL
                        End If
                    Next
                    
                    If optʱ��(0).Value = True Then
                        gstrSQL = "zl_ҩƷ�շ���¼_Adjust(" & Val(.TextMatrix(i, .ColIndex("ҩƷid"))) & ",1)"
                        
                        ReDim Preserve marrSql(UBound(marrSql) + 1)
                        marrSql(UBound(marrSql)) = gstrSQL
                    End If
                End If
                
                bln�Ƿ�����ۼ� = False
                dbl���ۼ� = 0
                intͬҩƷID�� = 0
            End If
        Next
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ModifyAllPrice()
    '�ɱ��ۡ��ۼ�һ���
    Dim strAll��ˮ�� As String
    Dim rsTemp As ADODB.Recordset
    Dim dbl��װ As Double
    Dim i As Integer
    Dim n As Integer
    Dim intͬҩƷID�� As Integer
    Dim int�շѼ�Ŀ��� As Integer
    Dim bln�Ƿ�����ۼ� As Boolean
    Dim dbl���ۼ� As Double
    Dim lngAdjId As Long
    Dim dtToday As Date
    Dim strNo As String
    Dim LngCurID As Long

    On Error GoTo ErrHand
    
    With vsfPrice
        For i = 1 To .rows - 2
            If Val(.TextMatrix(i, .ColIndex("ԭ�ۼ�"))) <> Val(.TextMatrix(i, .ColIndex("�ּ۸�"))) And Val(.TextMatrix(i, .ColIndex("ԭ�ɱ���"))) <> Val(.TextMatrix(i, .ColIndex("�ּ۸�"))) And .TextMatrix(i, .ColIndex("�ּ۸�")) <> "" Then
                '�ȴ���ɱ���
                If strAll��ˮ�� = "" Then
                    gstrSQL = "select nextno(135) as ��ˮ�� from dual"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "������ˮ��")
                    If rsTemp.RecordCount = 0 Then
                        MsgBox "������ˮ��δ�ܳ�ʼ���ɹ����������Ա��ϵ��", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    strAll��ˮ�� = rsTemp!��ˮ��

                    mstr��¼��ˮ�� = mstr��¼��ˮ�� & ";" & strAll��ˮ�� & "|" & 2
                    dtToday = Sys.Currentdate()
                End If
                              
                dbl��װ = Val(vsfPrice.TextMatrix(i, vsfPrice.ColIndex("��װϵ��")))
                
                If .TextMatrix(i, .ColIndex("�п��")) = 1 And .TextMatrix(i, .ColIndex("ҩ������")) = "����" Then
                    gstrSQL = "Select s.�ⷿid, s.ҩƷid, d.���� As �ⷿ, '[' || m.���� || ']' || m.���� As ҩƷ, m.���,s.�ϴβ��� as ����," & vbNewLine & _
                                    "       Decode([2], 0, p.ҩ�ⵥλ, 2, p.סԺ��λ, 1, p.���ﵥλ, m.���㵥λ) As ��λ," & vbNewLine & _
                                    "       Decode([2], 0, p.ҩ���װ, 2, p.סԺ��װ, 1, p.�����װ, 1) As ��װϵ��," & vbNewLine & _
                                    "       s.�ϴ����� As ����, Nvl(s.ʵ������, 0) As ����, Nvl(s.����,0) as ����," & vbNewLine & _
                                    "       Nvl(m.�Ƿ���, 0) ���, m.Id, Decode(Nvl(m.�Ƿ���, 0), 0, e.�ּ�, Decode(Nvl(s.���ۼ�, 0),0,s.ʵ�ʽ��/s.ʵ������,s.���ۼ�)) As ʱ���ۼ�, p.�ӳ���, Decode(Nvl(s.ƽ���ɱ���, 0), 0, p.�ɱ���, s.ƽ���ɱ���) As �ɱ���, nvl(s.�ϴι�Ӧ��id,0) As �ϴι�Ӧ��id," & vbNewLine & _
                                    "       n.���� As ��Ӧ��, s.Ч��" & vbNewLine & _
                                    "From ҩƷ��� S, ���ű� D, �շ���ĿĿ¼ M, ҩƷ��� P, ��Ӧ�� N, �շѼ�Ŀ E" & vbNewLine & _
                                    "Where d.Id = s.�ⷿid And s.ҩƷid = m.Id And m.Id = p.ҩƷid And Nvl(s.�ϴι�Ӧ��id, 0) = n.Id(+) And m.Id = e.�շ�ϸĿid And" & vbNewLine & _
                                    "      s.���� = 1 And s.ҩƷid = [1] And Sysdate Between e.ִ������ And e.��ֹ����   And e.�۸�ȼ� Is Null" & vbNewLine & _
                                    "Order By s.ҩƷid,s.�ⷿid, s.�ϴ�����,s.���� "

                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ҩƷ��Ϣ", Val(.TextMatrix(i, .ColIndex("ҩƷID"))), mintUnit)
                    
                    With rsTemp
                        Do While Not .EOF
                            gstrSQL = "Zl_ҩƷ�۸��¼_Stop("
                            '�۸�����_In
                            gstrSQL = gstrSQL & 2
                            '�ⷿid_In
                            gstrSQL = gstrSQL & "," & !�ⷿid
                            'ҩƷid_In
                            gstrSQL = gstrSQL & "," & !ҩƷid
                            '����_In
                            gstrSQL = gstrSQL & "," & Nvl(!����, 0)
                            '��ֹ����_In
                            If optʱ��(0).Value = True Then
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            gstrSQL = gstrSQL & ")"
                            
                            ReDim Preserve marrSql(UBound(marrSql) + 1)
                            marrSql(UBound(marrSql)) = gstrSQL
                            
                            gstrSQL = "Zl_ҩƷ�۸��¼_Insert("
                            '��������_In
                            gstrSQL = gstrSQL & 1
                            '�۸�����_In
                            gstrSQL = gstrSQL & ",2"
                            '�ⷿid_In
                            gstrSQL = gstrSQL & "," & !�ⷿid
                            'ҩƷid_In
                            gstrSQL = gstrSQL & "," & !ҩƷid
                            '����_In
                            gstrSQL = gstrSQL & "," & Nvl(!����, 0)
                            'ԭ��_In
                            gstrSQL = gstrSQL & "," & zlStr.FormatEx(!�ɱ���, gtype_UserDrugDigits.Digit_���ۼ�, , True)
                            '�ּ�_In
                            gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(vsfPrice.TextMatrix(i, vsfPrice.ColIndex("�ּ۸�"))) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�, , True)
                            'ִ������_In
                            If optʱ��(0).Value = True Then
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            '����˵��_In
                            gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                            '������_In
                            gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                            '���ۻ��ܺ�_In
                            gstrSQL = gstrSQL & ",'" & strAll��ˮ�� & "'"
                            '��ҩ��λid_In
                            gstrSQL = gstrSQL & "," & IIf(!�ϴι�Ӧ��ID = 0, "null", !�ϴι�Ӧ��ID)
                            '����_In
                            gstrSQL = gstrSQL & ",'" & Nvl(!����) & "'"
                            'Ч��_In
                            gstrSQL = gstrSQL & "," & "to_date('" & Format(IIf(IsNull(!Ч��), "", !Ч��), "YYYY-MM-DD") & "','yyyy-mm-dd')"
                            '����_In
                            gstrSQL = gstrSQL & ",'" & Nvl(!����) & "'"
                            gstrSQL = gstrSQL & ")"
                            
                            ReDim Preserve marrSql(UBound(marrSql) + 1)
                            marrSql(UBound(marrSql)) = gstrSQL
                            
                        .MoveNext
                        Loop
                    End With
                End If
            
                If .TextMatrix(i, .ColIndex("�п��")) = 1 And .TextMatrix(i, .ColIndex("ҩ������")) = "ʱ��" Then
                    '�п��
                    gstrSQL = "Zl_ҩƷ�۸��¼_Stop("
                    '�۸�����_In
                    gstrSQL = gstrSQL & 2
                    '�ⷿid_In
                    gstrSQL = gstrSQL & "," & Val(.TextMatrix(i, .ColIndex("�ⷿid")))
                    'ҩƷid_In
                    gstrSQL = gstrSQL & "," & Val(.TextMatrix(i, .ColIndex("ҩƷid")))
                    '����_In
                    gstrSQL = gstrSQL & "," & Nvl(.TextMatrix(i, .ColIndex("����")), 0)
                    '��ֹ����_In
                    If optʱ��(0).Value = True Then
                        gstrSQL = gstrSQL & "," & "to_date('" & Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        gstrSQL = gstrSQL & "," & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    gstrSQL = gstrSQL & ")"
                    
                    ReDim Preserve marrSql(UBound(marrSql) + 1)
                    marrSql(UBound(marrSql)) = gstrSQL
                    
                    gstrSQL = "Zl_ҩƷ�۸��¼_Insert("
                    '��������_In
                    gstrSQL = gstrSQL & 1
                    '�۸�����_In
                    gstrSQL = gstrSQL & ",2"
                    '�ⷿid_In
                    gstrSQL = gstrSQL & "," & Val(.TextMatrix(i, .ColIndex("�ⷿid")))
                    'ҩƷid_In
                    gstrSQL = gstrSQL & "," & Val(.TextMatrix(i, .ColIndex("ҩƷid")))
                    '����_In
                    gstrSQL = gstrSQL & "," & Nvl(.TextMatrix(i, .ColIndex("����")), 0)
                    'ԭ��_In
                    gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(i, .ColIndex("ԭ�ɱ���"))) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�, , True)
                    '�ּ�_In
                    gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(i, .ColIndex("�ּ۸�"))) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�, , True)
                    'ִ������_In
                    If optʱ��(0).Value = True Then
                        gstrSQL = gstrSQL & "," & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        gstrSQL = gstrSQL & "," & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    '����˵��_In
                    gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                    '������_In
                    gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                    '���ۻ��ܺ�_In
                    gstrSQL = gstrSQL & ",'" & strAll��ˮ�� & "'"
                    '��ҩ��λid_In
                    gstrSQL = gstrSQL & "," & IIf(Val(.TextMatrix(i, .ColIndex("�ϴι�Ӧ��ID"))) = 0, "null", Val(.TextMatrix(i, .ColIndex("�ϴι�Ӧ��ID"))))
                    '����_In
                    gstrSQL = gstrSQL & ",'" & Nvl(.TextMatrix(i, .ColIndex("����"))) & "'"
                    'Ч��_In
                    gstrSQL = gstrSQL & "," & "to_date('" & Format(IIf(IsNull(.TextMatrix(i, .ColIndex("Ч��"))), "", .TextMatrix(i, .ColIndex("Ч��"))), "YYYY-MM-DD") & "','yyyy-mm-dd')"
                    '����_In
                    gstrSQL = gstrSQL & ",'" & Nvl(.TextMatrix(i, .ColIndex("������"))) & "'"
                    gstrSQL = gstrSQL & ")"
                    
                    ReDim Preserve marrSql(UBound(marrSql) + 1)
                    marrSql(UBound(marrSql)) = gstrSQL
                End If
                
                If .TextMatrix(i, .ColIndex("�п��")) = 0 Then
                    '�޿��
                    gstrSQL = "Zl_ҩƷ�۸��¼_Insert("
                    '��������_In
                    gstrSQL = gstrSQL & 1
                    '�۸�����_In
                    gstrSQL = gstrSQL & ",2"
                    '�ⷿid_In
                    gstrSQL = gstrSQL & ",Null"
                    'ҩƷid_In
                    gstrSQL = gstrSQL & "," & Val(.TextMatrix(i, .ColIndex("ҩƷid")))
                    '����_In
                    gstrSQL = gstrSQL & ",0"
                    'ԭ��_In
                    gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(i, .ColIndex("ԭ�ɱ���"))) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�, , True)
                    '�ּ�_In
                    gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(i, .ColIndex("�ּ۸�"))) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�, , True)
                    'ִ������_In
                    If optʱ��(0).Value = True Then
                        gstrSQL = gstrSQL & "," & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        gstrSQL = gstrSQL & "," & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    '����˵��_In
                    gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                    '������_In
                    gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                    '���ۻ��ܺ�_In
                    gstrSQL = gstrSQL & ",'" & strAll��ˮ�� & "'"
                    gstrSQL = gstrSQL & ")"

                    ReDim Preserve marrSql(UBound(marrSql) + 1)
                    marrSql(UBound(marrSql)) = gstrSQL
                End If
            End If
        Next

        For i = 1 To .rows - 2
            '�ٴ����ۼ�
            If Val(.TextMatrix(i, .ColIndex("ҩƷid"))) = Val(.TextMatrix(i + 1, .ColIndex("ҩƷid"))) Then
                intͬҩƷID�� = intͬҩƷID�� + 1
            Else
                For n = i - intͬҩƷID�� To i
                    dbl���ۼ� = dbl���ۼ� + Val(.TextMatrix(n, .ColIndex("�ּ۸�")))
                    If Val(.TextMatrix(n, .ColIndex("ԭ�ۼ�"))) <> Val(.TextMatrix(n, .ColIndex("�ּ۸�"))) And Val(.TextMatrix(n, .ColIndex("ԭ�ɱ���"))) <> Val(.TextMatrix(n, .ColIndex("�ּ۸�"))) And .TextMatrix(n, .ColIndex("�ּ۸�")) <> "" Then
                        bln�Ƿ�����ۼ� = True
                    End If
                Next

                If bln�Ƿ�����ۼ� Then
                    If lngAdjId = 0 Then
                        strNo = Sys.GetNextNo(9)

                        gstrSQL = "select �շѼ�Ŀ_ID.nextval from dual"
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�շѼ�Ŀ���")
                        lngAdjId = rsTemp.Fields(0).Value

                        dtToday = Sys.Currentdate()
                    End If

                    int�շѼ�Ŀ��� = int�շѼ�Ŀ��� + 1
                    dbl���ۼ� = Round(dbl���ۼ� / (intͬҩƷID�� + 1), 2)
                    LngCurID = Sys.NextId("�շѼ�Ŀ")
                    dbl��װ = Val(.TextMatrix(i, .ColIndex("��װϵ��")))
            
                    If CLng(.TextMatrix(i, .ColIndex("�۸�ID"))) <> 0 Then
                        '������һ�εļ۸��¼��ִֹ��
                        gstrSQL = "zl_�շѼ�Ŀ_stop(" & .TextMatrix(i, .ColIndex("ҩƷid")) & ","
                        If optʱ��(0).Value = True Then
                            gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                        Else
                            gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                        End If
                        gstrSQL = gstrSQL & ")"
                        
                        ReDim Preserve marrSql(UBound(marrSql) + 1)
                        marrSql(UBound(marrSql)) = gstrSQL
            
                        '�����۸��¼
                        'ID
                        gstrSQL = "zl_�շѼ�Ŀ_Insert(" & LngCurID & ","
                        'ԭ��ID
                        gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("�۸�ID"))) & ","
                        '�շ�ϸĿID
                        gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("ҩƷid"))) & ","
                        '������ĿID
                        gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("������ĿID"))) & ","
                        'ԭ��
                        gstrSQL = gstrSQL & zlStr.FormatEx(Val(.TextMatrix(i, .ColIndex("ԭ�ۼ�"))) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�, , True) & ","
                        '�ּ�
                        gstrSQL = gstrSQL & zlStr.FormatEx(dbl���ۼ� / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�, , True) & ","
                        '�����շ��ʡ��Ӱ�Ӽ��ʡ�����˵��
                        gstrSQL = gstrSQL & "NULL,NULL,'" & MoveSpecialChar(txtSummary.Text) & "',"
                        '����id��������
                        gstrSQL = gstrSQL & lngAdjId & ",'" & MoveSpecialChar(txtValuer.Text) & "',"
                        'ִ������
                        If optʱ��(0).Value = True Then
                            gstrSQL = gstrSQL & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                        Else
                            gstrSQL = gstrSQL & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                        End If
                        '�䶯ԭ��
                        gstrSQL = gstrSQL & "0,"
                        'NO
                        gstrSQL = gstrSQL & "'" & strNo & "',"
                        '��š�ȱʡ�۸�
                        gstrSQL = gstrSQL & int�շѼ�Ŀ��� & ",Null,"
                        '���ۻ��ܺ�
                        gstrSQL = gstrSQL & strAll��ˮ�� & ")"
                        
                        ReDim Preserve marrSql(UBound(marrSql) + 1)
                        marrSql(UBound(marrSql)) = gstrSQL
                    End If
                
                    For n = i - intͬҩƷID�� To i
                        If Val(.TextMatrix(n, .ColIndex("ԭ�ۼ�"))) <> Val(.TextMatrix(n, .ColIndex("�ּ۸�"))) And Val(.TextMatrix(n, .ColIndex("ԭ�ɱ���"))) <> Val(.TextMatrix(n, .ColIndex("�ּ۸�"))) And .TextMatrix(n, .ColIndex("�ּ۸�")) <> "" And .TextMatrix(n, .ColIndex("ҩ������")) = "ʱ��" And .TextMatrix(n, .ColIndex("�п��")) = 1 Then
                            gstrSQL = "Zl_ҩƷ�۸��¼_Stop("
                            '�۸�����_In
                            gstrSQL = gstrSQL & 1
                            '�ⷿid_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("�ⷿID")))
                            'ҩƷid_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("ҩƷID")))
                            '����_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("����")))
                            '��ֹ����_In
                            If optʱ��(0).Value = True Then
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            gstrSQL = gstrSQL & ")"
                            
                            ReDim Preserve marrSql(UBound(marrSql) + 1)
                            marrSql(UBound(marrSql)) = gstrSQL
                            
                            gstrSQL = "Zl_ҩƷ�۸��¼_Insert("
                            '��������_In
                            gstrSQL = gstrSQL & 1
                            '�۸�����_In
                            gstrSQL = gstrSQL & ",1"
                            '�ⷿid_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("�ⷿID")))
                            'ҩƷid_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("ҩƷID")))
                            '����_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("����")))
                            'ԭ��_In
                            gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(n, .ColIndex("ԭ�ۼ�"))) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�, , True)
                            '�ּ�_In
                            gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(n, .ColIndex("�ּ۸�"))) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�, , True)
                            'ִ������_In
                            If optʱ��(0).Value = True Then
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            '����˵��_In
                            gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                            '������_In
                            gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                            '���ۻ��ܺ�_In
                            gstrSQL = gstrSQL & ",'" & strAll��ˮ�� & "'"
                            '��ҩ��λid_In
                            gstrSQL = gstrSQL & "," & IIf(Val(.TextMatrix(n, .ColIndex("�ϴι�Ӧ��ID"))) = 0, "null", Val(.TextMatrix(n, .ColIndex("�ϴι�Ӧ��ID"))))
                            '����_In
                            gstrSQL = gstrSQL & ",'" & Nvl(.TextMatrix(n, .ColIndex("����"))) & "'"
                            'Ч��_In
                            gstrSQL = gstrSQL & "," & "to_date('" & Format(Nvl(.TextMatrix(n, .ColIndex("Ч��"))), "YYYY-MM-DD") & "','yyyy-mm-dd')"
                            '����_In
                            gstrSQL = gstrSQL & ",'" & Nvl(.TextMatrix(n, .ColIndex("������"))) & "'"
                            gstrSQL = gstrSQL & ")"
                            
                            ReDim Preserve marrSql(UBound(marrSql) + 1)
                            marrSql(UBound(marrSql)) = gstrSQL
                        End If
                        
                        If Val(.TextMatrix(n, .ColIndex("ԭ�ۼ�"))) <> Val(.TextMatrix(n, .ColIndex("�ּ۸�"))) And Val(.TextMatrix(n, .ColIndex("ԭ�ɱ���"))) <> Val(.TextMatrix(n, .ColIndex("�ּ۸�"))) And .TextMatrix(n, .ColIndex("�ּ۸�")) <> "" And .TextMatrix(n, .ColIndex("ҩ������")) = "ʱ��" And .TextMatrix(n, .ColIndex("�п��")) = 0 Then
                            'ʱ��ҩƷ�޿�����
                            gstrSQL = "Zl_ҩƷ�۸��¼_Insert("
                            '��������_In
                            gstrSQL = gstrSQL & 1
                            '�۸�����_In
                            gstrSQL = gstrSQL & ",1"
                            '�ⷿid_In
                            gstrSQL = gstrSQL & ",Null"
                            'ҩƷid_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("ҩƷID")))
                            '����_In
                            gstrSQL = gstrSQL & ",0"
                            'ԭ��_In
                            gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(n, .ColIndex("ԭ�ۼ�"))) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�, , True)
                            '�ּ�_In
                            gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(n, .ColIndex("�ּ۸�"))) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�, , True)
                            'ִ������_In
                            If optʱ��(0).Value = True Then
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            '����˵��_In
                            gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                            '������_In
                            gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                            '���ۻ��ܺ�_In
                            gstrSQL = gstrSQL & ",'" & strAll��ˮ�� & "'"
                            gstrSQL = gstrSQL & ")"

                            ReDim Preserve marrSql(UBound(marrSql) + 1)
                            marrSql(UBound(marrSql)) = gstrSQL
                        End If
                    Next
                    
                    If optʱ��(0).Value = True Then
                        gstrSQL = "zl_ҩƷ�շ���¼_Adjust(" & Val(.TextMatrix(i, .ColIndex("ҩƷid"))) & ")"
                        
                        ReDim Preserve marrSql(UBound(marrSql) + 1)
                        marrSql(UBound(marrSql)) = gstrSQL
                    End If
                End If
                bln�Ƿ�����ۼ� = False
                dbl���ۼ� = 0
                intͬҩƷID�� = 0
            End If
        Next
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub setNOtExcetePrice()
    '����ʱ�仹δִ�е���ҩƷִ�е���
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim blnTrans As Boolean
    Dim arrSql() As Variant
    
    arrSql = Array()
    On Error GoTo errHandle
    
    gstrSQL = "Select Distinct i.Id As ҩƷid " & _
               " From �շ���ĿĿ¼ I, �շѼ�Ŀ N, ҩƷ��� P" & _
               " Where i.Id = n.�շ�ϸĿid And i.Id = p.ҩƷid And (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) And" & _
                   " n.�䶯ԭ�� = 0 And Sysdate>n.ִ������" & GetPriceClassString("N") & _
               " Union " & _
               " Select Distinct a.ҩƷid From ҩƷ�۸��¼ A Where a.��¼״̬ = 0 And a.ִ������ <= Sysdate " & _
               " Order By ҩƷid "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ִ�е���")
    
    If rsTemp.RecordCount = 0 Then Exit Sub
    
    For i = 0 To rsTemp.RecordCount - 1
        gstrSQL = "Zl_ҩƷ�շ���¼_Adjust(" & rsTemp!ҩƷid & ")"
        
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
        rsTemp.MoveNext
    Next
                   
    gcnOracle.BeginTrans: blnTrans = True          '��������
    For i = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False     '�ύ����
    
    Exit Sub
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub PrintPrice()
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    Call Summary
    
    If vsfPrint.rows = 1 Then
        MsgBox "û�п��䶯��¼��", vbInformation, gstrSysName
        Exit Sub
    End If

    objPrint.Title.Text = "���ۿ��䶯��"

    Set objRow = New zlTabAppRow
    objRow.Add "����˵��:" & Me.txtSummary.Text
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "ִ��ʱ��:" & Format(IIf(optʱ��(0).Value = True, Sys.Currentdate, Me.dtpRunDate.Value), "yyyy��MM��DD�� HH:mm:ss")
    objRow.Add "������:" & Me.txtValuer.Text
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.�û�����
    objRow.Add "��ӡʱ��:" & Format(Sys.Currentdate, "yyyy��MM��DD�� HH:mm:ss")
    objPrint.BelowAppRows.Add objRow

    Set objPrint.Body = Me.vsfPrint.Object
    objPrint.PageFooter = 2

    Select Case zlPrintAsk(objPrint)
    Case 1
         zlPrintOrView1Grd objPrint, 1
    Case 2
        zlPrintOrView1Grd objPrint, 2
    Case 3
        zlPrintOrView1Grd objPrint, 3
    End Select
    Set objPrint = Nothing
End Sub

Private Sub Summary()
    '���ܿ��䶯��
    Dim i As Integer
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    vsfPrint.rows = 1
    
    For i = 1 To vsfPrice.rows - 2
        If vsfPrice.TextMatrix(i, vsfPrice.ColIndex("�п��")) = 1 And vsfPrice.TextMatrix(i, vsfPrice.ColIndex("ҩ������")) = "ʱ��" And vsfPrice.TextMatrix(i, vsfPrice.ColIndex("�ּ۸�")) <> "" Then
            vsfPrint.rows = vsfPrint.rows + 1
            vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("ҩ������")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("ҩ������"))
            vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("�ⷿ")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("�ⷿ"))
            vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("Ʒ��")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("Ʒ��"))
            vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("���")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("���"))
            vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("������")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("������"))
            vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("����")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("����"))
            vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("��λ")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("��λ"))
            vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("ԭ�ۼ�")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("ԭ�ۼ�"))
            vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("ԭ�ɱ���")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("ԭ�ɱ���"))
            vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("���ۼ�")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("�ּ۸�"))
            vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("�ֳɱ���")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("�ּ۸�"))
        End If

        If vsfPrice.TextMatrix(i, vsfPrice.ColIndex("�п��")) = 1 And vsfPrice.TextMatrix(i, vsfPrice.ColIndex("ҩ������")) = "����" And vsfPrice.TextMatrix(i, vsfPrice.ColIndex("�ּ۸�")) <> "" Then
            
            gstrSQL = "Select s.�ⷿid, s.ҩƷid, d.���� As �ⷿ, '[' || m.���� || ']' || m.���� As ҩƷ, m.���,s.�ϴβ��� as ����," & vbNewLine & _
                            "       Decode([2], 0, p.ҩ�ⵥλ, 2, p.סԺ��λ, 1, p.���ﵥλ, m.���㵥λ) As ��λ," & vbNewLine & _
                            "       Decode([2], 0, p.ҩ���װ, 2, p.סԺ��װ, 1, p.�����װ, 1) As ��װϵ��," & vbNewLine & _
                            "       s.�ϴ����� As ����, Nvl(s.ʵ������, 0) As ����, s.����," & vbNewLine & _
                            "       Nvl(m.�Ƿ���, 0) ���, m.Id, Decode(Nvl(m.�Ƿ���, 0), 0, e.�ּ�, Decode(Nvl(s.���ۼ�, 0),0,s.ʵ�ʽ��/s.ʵ������,s.���ۼ�)) As ʱ���ۼ�, p.�ӳ���, Decode(Nvl(s.ƽ���ɱ���, 0), 0, p.�ɱ���, s.ƽ���ɱ���) As �ɱ���, nvl(s.�ϴι�Ӧ��id,0) As �ϴι�Ӧ��id," & vbNewLine & _
                            "       n.���� As ��Ӧ��, s.Ч��" & vbNewLine & _
                            "From ҩƷ��� S, ���ű� D, �շ���ĿĿ¼ M, ҩƷ��� P, ��Ӧ�� N, �շѼ�Ŀ E" & vbNewLine & _
                            "Where d.Id = s.�ⷿid And s.ҩƷid = m.Id And m.Id = p.ҩƷid And Nvl(s.�ϴι�Ӧ��id, 0) = n.Id(+) And m.Id = e.�շ�ϸĿid And" & vbNewLine & _
                            "      s.���� = 1 And s.ҩƷid = [1] And Sysdate Between e.ִ������ And e.��ֹ���� And e.�۸�ȼ� Is Null" & vbNewLine & _
                            "Order By s.ҩƷid,s.�ⷿid, s.�ϴ�����,s.���� "

                            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ҩƷ��Ϣ", Val(vsfPrice.TextMatrix(i, vsfPrice.ColIndex("ҩƷID"))), mintUnit)
            
            With rsTemp
                Do While Not .EOF
                    vsfPrint.rows = vsfPrint.rows + 1
                    vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("ҩ������")) = "����"
                    vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("�ⷿ")) = !�ⷿ
                    vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("Ʒ��")) = !ҩƷ
                    vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("���")) = !���
                    vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("������")) = Nvl(!����)
                    vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("����")) = Nvl(!����)
                    vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("��λ")) = !��λ
                    vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("ԭ�ۼ�")) = zlStr.FormatEx(!ʱ���ۼ� * !��װϵ��, mintPriceDigit, , True)
                    vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("ԭ�ɱ���")) = zlStr.FormatEx(!�ɱ��� * !��װϵ��, mintCostDigit, , True)
                    vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("���ۼ�")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("�ּ۸�"))
                    vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("�ֳɱ���")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("�ּ۸�"))
                    .MoveNext
                Loop
            End With
        End If
    Next
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ClearAllDate()
    '��ս�������
    Dim i As Integer
    If vsfPrice.rows <= 1 Then Exit Sub
    If MsgBox("δ���汾�ε��ۣ��Ƿ�����������ݣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    mstrҩƷID = ""
    vsfPrice.rows = 1
End Sub

Private Sub AllDrug()
    Dim intRow As Integer
    Dim rsReturn As ADODB.Recordset
    Dim blnOK As Boolean
    
    frmBatchSelect.ShowME Me, rsReturn, blnOK, 1

    On Error GoTo errHandle
    If blnOK = False Then Exit Sub
    If rsReturn.RecordCount = 0 Then Exit Sub
    
'    If txtValuer.Tag = "��������ҩƷ" And vsfPrice.rows > 1 Then
'         If MsgBox("�ò���������������ݲ�������ȡ��ѡ������۹���ҩƷ���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'    End If
    
    If mstrҩƷID <> "" Then
        Select Case MsgBox("�Ƿ�����������ݣ�", vbQuestion + vbYesNoCancel + vbDefaultButton2, gstrSysName)
            Case vbYes
                mstrҩƷID = ""
            Case vbCancel
                Exit Sub
        End Select
    End If
        
    rsReturn.MoveFirst
    Do While Not rsReturn.EOF
        If InStr(mstrҩƷID & ",", "," & rsReturn!ҩƷid & ",") = 0 Then
            mstrҩƷID = mstrҩƷID & "," & rsReturn!ҩƷid
        End If
        rsReturn.MoveNext
    Loop

    Call GetAllPriceDiff
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetAllPriceDiff(Optional bln��ʾ As Boolean = True)
    '��ȡ���������������۹����ҩƷ�����۸�һ�µ�ҩƷ
    Dim rsData As ADODB.Recordset
    Dim bln����δִ�м۸� As Boolean
    Dim int��� As Integer
    
    On Error GoTo errHandle
       
    Call setNOtExcetePrice
    
    gstrSQL = "select * from (" & vbNewLine & _
                    "Select ҩƷid, ͨ����, ���, 0 As �ⷿid, '' As �ⷿ, ������, '' As ����, ����, ��λ, ��װϵ��, �ۼ�, Sum(�ɱ��� * ʵ������) / Sum(ʵ������) As �ɱ���, �Ƿ�ʱ��," & vbNewLine & _
                    "       �п��, �۸�id, ������Ŀid, Null As �ϴι�Ӧ��id, Null As Ч��" & vbNewLine & _
                    "From (Select a.ҩ��id,a.ҩƷid, '[' || c.���� || ']' || c.���� As ͨ����, c.���, c.���� As ������, 0 As ����," & vbNewLine & _
                    "              Decode([1], 0, a.ҩ�ⵥλ, 2, a.סԺ��λ, 1, a.���ﵥλ, c.���㵥λ) As ��λ," & vbNewLine & _
                    "              Decode([1], 0, a.ҩ���װ, 2, a.סԺ��װ, 1, a.�����װ, 1) As ��װϵ��, b.�ּ� As �ۼ�, decode(d.ƽ���ɱ���,null,a.�ɱ���,d.ƽ���ɱ���) As �ɱ���, 0 As �Ƿ�ʱ��, d.ʵ������," & vbNewLine & _
                    "              1 As �п��, b.Id As �۸�id, b.������Ŀid" & vbNewLine & _
                    "       From ҩƷ��� A, �շѼ�Ŀ B, �շ���ĿĿ¼ C, ҩƷ��� D" & vbNewLine & _
                    "       Where a.ҩƷid = b.�շ�ϸĿid And a.ҩƷid = c.Id And a.ҩƷid = d.ҩƷid And d.���� = 1 And (Sysdate Between b.ִ������ And b.��ֹ����) And" & vbNewLine & _
                    "             (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.�Ƿ��� = 0 And Nvl(a.�Ƿ����۹���, 0) = 1 " & vbNewLine & _
                    " And Not (zl_fun_getbatchpro(d.�ⷿid,d.ҩƷid)=1 And Nvl(d.����,0) = 0 And d.�������� < 0 And d.ʵ������ = 0 And d.ʵ�ʽ�� = 0 And d.ʵ�ʲ�� = 0)) " & vbNewLine & _
                    "Group By ҩƷid, ͨ����, ���, ������, ����, ��λ, ��װϵ��, �ۼ�, �۸�id, ������Ŀid, �Ƿ�ʱ��, �п�� " & vbNewLine & _
                    "Union All "

    gstrSQL = gstrSQL & "Select a.ҩƷid, '[' || c.���� || ']' || c.���� As ͨ����, c.���, d.�ⷿid, e.���� As �ⷿ, d.�ϴβ��� As ������, d.�ϴ����� As ����, d.����," & vbNewLine & _
                    "       Decode([1], 0, a.ҩ�ⵥλ, 2, a.סԺ��λ, 1, a.���ﵥλ, c.���㵥λ) As ��λ," & vbNewLine & _
                    "       Decode([1], 0, a.ҩ���װ, 2, a.סԺ��װ, 1, a.�����װ, 1) As ��װϵ��, d.���ۼ� As �ۼ�, decode(d.ƽ���ɱ���,null,a.�ɱ���,d.ƽ���ɱ���) As �ɱ���, 1 As �Ƿ�ʱ��, 1 As �п��," & vbNewLine & _
                    "       b.Id As �۸�id, b.������Ŀid, nvl(d.�ϴι�Ӧ��id,0) As �ϴι�Ӧ��id, d.Ч��" & vbNewLine & _
                    "From ҩƷ��� A, �շѼ�Ŀ B, �շ���ĿĿ¼ C, ҩƷ��� D, ���ű� E" & vbNewLine & _
                    "Where a.ҩƷid = b.�շ�ϸĿid And a.ҩƷid = c.Id And a.ҩƷid = d.ҩƷid And d.�ⷿid = e.Id And d.���� = 1 And" & vbNewLine & _
                    "      (Sysdate Between b.ִ������ And b.��ֹ����) And c.�Ƿ��� = 1 And" & vbNewLine & _
                    "      (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And Nvl(a.�Ƿ����۹���, 0) = 1  " & vbNewLine & _
                    " And Not (zl_fun_getbatchpro(d.�ⷿid,d.ҩƷid)=1 And Nvl(d.����,0) = 0 And d.�������� < 0 And d.ʵ������ = 0 And d.ʵ�ʽ�� = 0 And d.ʵ�ʲ�� = 0) " & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select a.ҩƷid, '[' || c.���� || ']' || c.���� As ͨ����, c.���, 0 As �ⷿid, '' As �ⷿ, '' As ������, '' As ����, 0 As ����," & vbNewLine & _
                    "       Decode([1], 0, a.ҩ�ⵥλ, 2, a.סԺ��λ, 1, a.���ﵥλ, c.���㵥λ) As ��λ," & vbNewLine & _
                    "       Decode([1], 0, a.ҩ���װ, 2, a.סԺ��װ, 1, a.�����װ, 1) As ��װϵ��, b.�ּ� As �ۼ�, a.�ɱ���, c.�Ƿ��� As �Ƿ�ʱ��, 0 As �п��, b.Id As �۸�id," & vbNewLine & _
                    "       b.������Ŀid, Null As �ϴι�Ӧ��id, Null As Ч��" & vbNewLine & _
                    "From ҩƷ��� A, �շѼ�Ŀ B, �շ���ĿĿ¼ C" & vbNewLine & _
                    "Where a.ҩƷid = c.Id And a.ҩƷid = b.�շ�ϸĿid And (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
                    "      Nvl(a.�Ƿ����۹���, 0) = 1 And (Sysdate Between b.ִ������ And b.��ֹ����) And Not Exists" & vbNewLine & _
                    " (Select 1 From ҩƷ��� D Where d.ҩƷid = a.ҩƷid And d.���� = 1)" & vbNewLine & _
                    "Order By ҩƷid, �ⷿid, ����,����) m" & vbNewLine & _
                    "where m.ҩƷid In (Select Column_Value From Table(f_num2list([2]))) Order By ҩƷid, �ⷿid, ����,���� "

    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "GetAllPriceDiff", mintUnit, Mid(mstrҩƷID, 2))
    
    With vsfPrice
        .Redraw = flexRDNone
        .MergeCells = flexMergeNever
        .rows = 1
    
        Do While Not rsData.EOF
            '����Ƿ����δִ�м۸�������ھͲ�ȡ����
            If CheckExistExecutePrice(Val(rsData!ҩƷid)) = False Then
                .rows = .rows + 1
                
                .TextMatrix(.rows - 1, .ColIndex("���")) = int��� + 1
                .TextMatrix(.rows - 1, .ColIndex("ҩƷid")) = rsData!ҩƷid
                .TextMatrix(.rows - 1, .ColIndex("ҩ������")) = IIf(rsData!�Ƿ�ʱ�� = 1, "ʱ��", "����")
                .TextMatrix(.rows - 1, .ColIndex("Ʒ��")) = rsData!ͨ����
                .TextMatrix(.rows - 1, .ColIndex("���")) = rsData!���
                .TextMatrix(.rows - 1, .ColIndex("������")) = Nvl(rsData!������, "")
                .TextMatrix(.rows - 1, .ColIndex("�ⷿid")) = rsData!�ⷿid
                .TextMatrix(.rows - 1, .ColIndex("�ⷿ")) = Nvl(rsData!�ⷿ, "")
                .TextMatrix(.rows - 1, .ColIndex("����")) = Nvl(rsData!����, "")
                .TextMatrix(.rows - 1, .ColIndex("��λ")) = rsData!��λ
                .TextMatrix(.rows - 1, .ColIndex("��װϵ��")) = rsData!��װϵ��
                .TextMatrix(.rows - 1, .ColIndex("ԭ�ۼ�")) = zlStr.FormatEx(rsData!�ۼ� * rsData!��װϵ��, mintPriceDigit, , True)
                .TextMatrix(.rows - 1, .ColIndex("ԭ�ɱ���")) = zlStr.FormatEx(rsData!�ɱ��� * rsData!��װϵ��, mintCostDigit, , True)
                .TextMatrix(.rows - 1, .ColIndex("ԭʼ�ۼ�")) = rsData!�ۼ�
                .TextMatrix(.rows - 1, .ColIndex("ԭʼ�ɱ���")) = rsData!�ɱ���
                .TextMatrix(.rows - 1, .ColIndex("�п��")) = rsData!�п��
                .TextMatrix(.rows - 1, .ColIndex("�۸�id")) = rsData!�۸�id
                .TextMatrix(.rows - 1, .ColIndex("������Ŀid")) = rsData!������ĿID
                .TextMatrix(.rows - 1, .ColIndex("����")) = Nvl(rsData!����, 0)
                .TextMatrix(.rows - 1, .ColIndex("�ϴι�Ӧ��ID")) = Nvl(rsData!�ϴι�Ӧ��ID)
                .TextMatrix(.rows - 1, .ColIndex("Ч��")) = Nvl(rsData!Ч��)
                
                .Cell(flexcpForeColor, .rows - 1, .ColIndex("ҩ������"), .rows - 1, .ColIndex("ҩ������")) = IIf(rsData!�Ƿ�ʱ�� = 1, vbRed, vbBlack)
                int��� = int��� + 1
            Else
                bln����δִ�м۸� = True
            End If
            
            rsData.MoveNext
        Loop
            
        If .rows >= 2 Then
            .Cell(flexcpBackColor, 1, .ColIndex("�ּ۸�"), .rows - 1, .ColIndex("�ּ۸�")) = mconlngCanColColor
            .Cell(flexcpForeColor, 1, .ColIndex("�ּ۸�"), .rows - 1, .ColIndex("�ּ۸�")) = vbBlue
            .Cell(flexcpFontBold, 1, .ColIndex("�ּ۸�"), .rows - 1, .ColIndex("�ּ۸�")) = True
        End If
        
        .rows = .rows + 1
        .RowHidden(.rows - 1) = True
        .Redraw = flexRDDirect
    End With
    
    txtValuer.Text = UserInfo.�û�����
    txtSummary.Text = "���۵���"
    
    txtValuer.Tag = "ȫ������ҩƷ"
    If bln����δִ�м۸� = True Then
        If bln��ʾ Then
            MsgBox "�������۹���ҩƷ������δִ�е�Ԥ���ۼ�¼�������������������б�����ʾ��ЩҩƷ����ע��鿴��", vbInformation, gstrSysName
        Else
            MsgBox "�������۹���ҩƷ�����������в���ҩƷδ���е��ۣ���ע��鿴��", vbInformation, gstrSysName
        End If
    End If
        
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FindLocation()
    Dim i As Integer
    
    With vsfPrice
        If .rows > 1 Then
            For i = mintRow To .rows - 2
                If .TextMatrix(mintRow, .ColIndex("�ּ۸�")) = "" Then
                    .TopRow = mintRow
                    .Row = mintRow
                    .Col = .ColIndex("�ּ۸�")
                    mintRow = mintRow + 1
                    Exit For
                Else
                    mintRow = mintRow + 1
                End If
            Next

            If mintRow = .rows - 1 And .TextMatrix(mintRow - 1, .ColIndex("�ּ۸�")) <> "" Then
                mintRow = 1
                For i = mintRow To .rows - 2
                    If .TextMatrix(mintRow, .ColIndex("�ּ۸�")) = "" Then
                        .TopRow = mintRow
                        .Row = mintRow
                        .Col = .ColIndex("�ּ۸�")
                        mintRow = mintRow + 1
                        Exit For
                    Else
                        mintRow = mintRow + 1
                    End If
                Next
            End If
            
            If mintRow = .rows - 1 Then
                mintRow = 1
            End If
        End If
    End With
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 And Trim(txtCode.Text) <> "" Then
        Call FindGridRow(txtCode.Text)
    End If
End Sub

Private Sub FindGridRow(ByVal strInput As String)
    Dim n As Integer
    Dim lngFindRow As Long
    Dim strҩ�� As String
    Dim lngRow As Long
    
    '����ҩƷ
    On Error GoTo errHandle
    If strInput <> txtCode.Tag Then
        '��ʾ�µĲ���
        txtCode.Tag = strInput
        
        gstrSQL = "Select Distinct A.Id,'[' || A.���� || ']' As ҩƷ����, A.���� As ͨ����, B.���� As ��Ʒ�� " & _
                  "From �շ���ĿĿ¼ A,�շ���Ŀ���� B " & _
                  "Where (A.վ�� = [3] Or A.վ�� is Null) And A.Id =B.�շ�ϸĿid And A.��� In ('5','6','7') " & _
                  "  And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2] ) " & _
                  "Order By ҩƷ���� "
        Set mrsFindName = zlDatabase.OpenSQLRecord(gstrSQL, "ȡƥ���ҩƷID", strInput & "%", "%" & strInput & "%", gstrNodeNo)
        
        If mrsFindName.RecordCount = 0 Then Exit Sub
        mrsFindName.MoveFirst
    End If
    
    '��ʼ����
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub
    
    For n = 1 To mrsFindName.RecordCount
        '��������ˣ��򷵻ص�1����¼
        If mrsFindName.EOF Then mrsFindName.MoveFirst
        If mrsFindName.RecordCount = 1 Then mlngFindCurrRow = 1
        
        If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
            strҩ�� = mrsFindName!ҩƷ���� & mrsFindName!ͨ����
        Else
            strҩ�� = mrsFindName!ҩƷ���� & IIf(IsNull(mrsFindName!��Ʒ��), mrsFindName!ͨ����, mrsFindName!��Ʒ��)
        End If
        lngFindRow = vsfPrice.FindRow(strҩ��, mlngFindCurrRow, CLng(vsfPrice.ColIndex("Ʒ��")), True, True)
        
        If lngFindRow > 0 Then '��ѯ�����ݺ���ƶ��µ���һ�У����������һ���Ƿ�����ͬ��ҩƷ
'            vsfPrice.Select lngFindRow, 1, lngFindRow, vsfPrice.Cols - 1
            vsfPrice.TopRow = lngFindRow
            vsfPrice.Row = lngFindRow
            vsfPrice.Col = vsfPrice.ColIndex("�ּ۸�")
                        
            If lngFindRow < vsfPrice.rows - 2 Then
                mlngFindCurrRow = lngFindRow + 1
            Else
                mlngFindCurrRow = 1
                mrsFindName.MoveNext 'δ��ѯ���������ƶ�����һ�����ݼ�������ѯ
            End If
            Exit For
        Else
            mrsFindName.MoveNext 'δ��ѯ���������ƶ�����һ�����ݼ�������ѯ
            mlngFindCurrRow = 1 '�����ӵ�һ�п�ʼ�Ƚ�����ҩƷ
        End If
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

