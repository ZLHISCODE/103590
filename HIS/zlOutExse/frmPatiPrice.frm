VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatiPrice 
   Caption         =   "���˻��۵�����"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10065
   Icon            =   "frmPatiPrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   10065
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboDiagnose 
      Height          =   300
      ItemData        =   "frmPatiPrice.frx":058A
      Left            =   960
      List            =   "frmPatiPrice.frx":0591
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   1530
      Width           =   3500
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   13
      Top             =   5955
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatiPrice.frx":059F
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12674
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraOk 
      Height          =   615
      Left            =   60
      TabIndex        =   23
      Top             =   5370
      Width           =   9855
      Begin VB.CommandButton cmdAllCls 
         Caption         =   "ȫ��(&R)"
         Height          =   350
         Left            =   1140
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   195
         Width           =   945
      End
      Begin VB.CommandButton cmdAllSel 
         Caption         =   "ȫѡ(&A)"
         Height          =   350
         Left            =   120
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   8640
         TabIndex        =   25
         Top             =   165
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   7500
         TabIndex        =   24
         ToolTipText     =   "�ȼ���F2"
         Top             =   165
         Width           =   1100
      End
   End
   Begin VB.Frame fraDays 
      Caption         =   "ѡ�񻮼۵�"
      Height          =   1455
      Left            =   7965
      TabIndex        =   28
      Top             =   60
      Width           =   1830
      Begin VB.CheckBox chkȱʡ 
         Caption         =   "���л��۵�"
         Height          =   360
         Index           =   2
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   990
         Width           =   1605
      End
      Begin VB.CheckBox chkȱʡ 
         Caption         =   "��Ч�����Ļ��۵�"
         Height          =   360
         Index           =   1
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   615
         Value           =   1  'Checked
         Width           =   1605
      End
      Begin VB.CheckBox chkȱʡ 
         Caption         =   "�����ڻ��۵�"
         Height          =   360
         Index           =   0
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   225
         Width           =   1605
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid mshDetail 
      Height          =   1410
      Left            =   45
      TabIndex        =   20
      Top             =   3960
      Width           =   9825
      _cx             =   17330
      _cy             =   2487
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
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483633
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin VB.Frame fraHsc 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   45
      MousePointer    =   7  'Size N S
      TabIndex        =   9
      Top             =   3825
      Width           =   9840
   End
   Begin VB.Frame fraPati 
      Caption         =   " ������Ϣ "
      Height          =   1455
      Left            =   45
      TabIndex        =   8
      Top             =   60
      Width           =   7860
      Begin VB.CommandButton cmdFind 
         Caption         =   "����(&F)"
         Height          =   350
         Left            =   6525
         TabIndex        =   29
         Top             =   1020
         Width           =   1155
      End
      Begin VB.ComboBox cbo�����ѱ� 
         Height          =   300
         Left            =   4500
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   690
         Width           =   1200
      End
      Begin VB.TextBox txt���ʽ 
         Height          =   300
         Left            =   6510
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   315
         Width           =   1200
      End
      Begin VB.TextBox txt����� 
         Height          =   300
         Left            =   750
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   675
         Width           =   1185
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   4500
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   315
         Width           =   1200
      End
      Begin VB.TextBox txt�Ա� 
         Height          =   300
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   315
         Width           =   1080
      End
      Begin VB.TextBox txtPatient 
         Height          =   300
         Left            =   750
         MaxLength       =   100
         TabIndex        =   0
         Top             =   315
         Width           =   1185
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   3615
         TabIndex        =   7
         Top             =   1035
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483636
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   177012739
         CurrentDate     =   38073
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   750
         TabIndex        =   6
         Top             =   1035
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483636
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   177012739
         CurrentDate     =   38073
      End
      Begin VB.TextBox txt�ѱ� 
         Height          =   300
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   690
         Width           =   1080
      End
      Begin VB.Label lbl�����ѱ� 
         AutoSize        =   -1  'True
         Caption         =   "�����ѱ�"
         Height          =   180
         Left            =   3750
         TabIndex        =   22
         Top             =   750
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ʽ"
         Height          =   180
         Left            =   5745
         TabIndex        =   18
         Top             =   375
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ԭ�ѱ�"
         Height          =   180
         Left            =   2070
         TabIndex        =   17
         Top             =   750
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   165
         TabIndex        =   16
         Top             =   735
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   4110
         TabIndex        =   15
         Top             =   375
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   2250
         TabIndex        =   14
         Top             =   375
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   345
         TabIndex        =   12
         Top             =   375
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   180
         Left            =   3195
         TabIndex        =   11
         Top             =   1095
         Width           =   90
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʱ��"
         Height          =   180
         Left            =   345
         TabIndex        =   10
         Top             =   1095
         Width           =   360
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid mshList 
      Height          =   2070
      Left            =   0
      TabIndex        =   19
      Top             =   1830
      Width           =   9825
      _cx             =   17330
      _cy             =   3651
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
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483633
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      ExplorerBar     =   2
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
   Begin VB.Label lblDiagnose 
      BackStyle       =   0  'Transparent
      Caption         =   "����Ϲ���"
      Height          =   255
      Left            =   30
      TabIndex        =   34
      Top             =   1590
      Width           =   915
   End
End
Attribute VB_Name = "frmPatiPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrBills As String
Private mstrPrivs As String
Private mlngModule As String
Private mlng����ID As Long
Private mrsList As ADODB.Recordset  '�����б�
Private mrsDetail As ADODB.Recordset
Private mbln������൥�� As Boolean
Private mlng�Һſ��� As Long        '��ͨ���Һŵ�����ʱ,���벡�˵�ǰ�Һŵ��ĹҺſ���
Private mblnCard As Boolean
Private mblnסԺ���������շ� As Boolean '34182
Private mblnNotClick As Boolean
Private mblnPreCard As Boolean
'-----------------------------------------------------------------------------------
'���㿨���
Private mstrPassWord As String
'-----------------------------------------------------------------------------------
Private Const VK_RETURN = &HD
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private mblnCacheKeyReturn As Boolean   '41025:�Ƿ񻺴��˻س���,���ܴ������շѽ���ˢ���б�������˻س�,�����Ҫ�ж�

Public Function FindBill(frmParent As Object, _
    ByVal strPrivs As String, Optional ByVal lng����ID As Long, _
    Optional ByVal bln������൥�� As Boolean, _
    Optional ByVal lng�Һſ��� As Long, _
    Optional ByVal blnסԺ���������շ� As Boolean = False, _
    Optional blnCard As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���һ��۵���
    '���:lng����ID=����ָ������(�ò���֮ǰ�϶���ȷ���л��۵���)
    '        lng�Һſ���,��ͨ���Һŵ�����ʱ,���벡�˵�ǰ�Һŵ��ĹҺſ���
    '        blnסԺ���������շ�-סԺ���˰���������շ�:34182
    '����:
    '����:
    '����:���˺�
    '����:2010-11-19 14:37:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnסԺ���������շ� = blnסԺ���������շ�
    mstrPrivs = strPrivs
    mlng����ID = lng����ID: mblnPreCard = blnCard
    mbln������൥�� = bln������൥��
    mlng�Һſ��� = lng�Һſ���
    Me.Show 1, frmParent
    FindBill = mstrBills
End Function

Public Function GetPriceBillString(ByVal lng����ID As Long, ByVal bln������൥�� As Boolean, ByVal lng�Һſ��� As Long, _
    Optional ByVal blnסԺ���������շ� As Boolean = False, Optional blnCard As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ָ������ʱ�䷶Χ�ڵĻ��۵�(����������ѡ��)
    '���:bln������൥��,���������൥��,ֻ������󻮼۵�һ�ŵ��ݺ�
    '        lng�Һſ���,��ͨ���Һŵ�����ʱ,���벡�˵�ǰ�Һŵ��ĹҺſ���
    '        blnסԺ���������շ�-סԺ���˰���������շ�:34182
    '       blnCard-�Ƿ����
    '����:
    '����:"G0001112,G0001113,G0001114..."
    '����:���˺�
    '����:2010-11-19 14:39:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, DatBegin As Date, DatEnd As Date
    Dim i As Long, strTmp As String
    mblnסԺ���������շ� = blnסԺ���������շ�
    DatEnd = zlDatabase.Currentdate
    DatBegin = DatEnd - gintSeekDays
    mblnPreCard = blnCard
    Set rsTmp = GetPriceBills(lng����ID, lng�Һſ���, DatBegin, DatEnd)
    For i = 1 To rsTmp.RecordCount
        strTmp = strTmp & IIf(strTmp = "", "", ",") & rsTmp!���ݺ�
        If bln������൥�� Then Exit For
        rsTmp.MoveNext
    Next
        
    If gblnCheckTest Then
        'ֻҪ����ҩƷƤ�Խ��Ϊ���ԵĶ��������շ�
        If Not CheckTest(strTmp, DatBegin, DatEnd) Then strTmp = ""
    End If
    
    GetPriceBillString = strTmp

End Function
Private Sub Local�ѱ�(ByVal str�ѱ� As String, Optional blnNotFindAdd As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݷѱ𣬲���ָ���ķѱ�
    '���:blnNotFindAdd-û�ҵ���ֱ������
    '����:
    '����:���˺�
    '����:2011-04-17 21:45:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To cbo�����ѱ�.ListCount - 1
        If str�ѱ� = cbo�����ѱ�.List(i) Then
            cbo�����ѱ�.ListIndex = i: Exit Sub
        End If
    Next
    If blnNotFindAdd = False Then Exit Sub
    cbo�����ѱ�.AddItem str�ѱ�
    cbo�����ѱ�.ListIndex = cbo�����ѱ�.NewIndex
End Sub

Private Sub cboDiagnose_Click()
    '74296,Ƚ����,2014-7-4,�����ݵ���Ϲ���,�ѵ����е�������һ�������б�ѡ��
    If mblnNotClick Then Exit Sub
    If mrsList Is Nothing Then Exit Sub
    mshList.Clear
    mshList.Rows = 2
    mshDetail.Clear
    mshDetail.Rows = 2
    stbThis.Panels(2).Text = ""
    mrsList.Filter = IIf(cboDiagnose.Text = "�������", "", "���='" & cboDiagnose.Text & "'")
    Set mshList.DataSource = mrsList
    Call SetHeader
    Call SetDetail
    Call mshList_EnterCell
    stbThis.Panels(2).Text = GetBillNote
End Sub

Private Sub cbo�����ѱ�_Click()
    Dim strSql As String, strComMand As String
    Dim i As Integer, strNos As String, blnChange As Boolean
    
    If mblnNotClick Then Exit Sub
    If cbo�����ѱ�.ListIndex < 0 Then Exit Sub
    '79870:���ϴ�,2015/4/10,�������ֵ��ݵķѱ�
    '��Ϊ���ܴ��ڲ��ֵ����벡����Ϣ�ķѱ�һ�µ���������Բ��ټ������ķѱ��Ƿ���ԭ�ѱ���ͬ
    If InStr(1, mstrPrivs, ";�������˷ѱ�;") = 0 Then Exit Sub
    
    strComMand = zlCommFun.ShowMsgbox("ע��", "���Ƿ�Ҫ���ѱ�" & Trim(txt�ѱ�.Text) & "������Ϊ��" & cbo�����ѱ�.Text & "����?" & vbCrLf & "  �����ѱ�ֱ��Ӱ����ص��շѻ��۵�!" & vbCrLf & vbCrLf & _
    "�����л��۵���:������δ�շѵĻ��۵�ȫ�����µķѱ����,ҽ���¿��Ĵ������·ѱ����" & vbCrLf & vbCrLf & _
    "��ѡ�л��۵���:ֻ������ѡ�еĻ��۵�,��Ӱ���¿��Ĵ�����δѡ�еĻ��۵�" & vbCrLf & vbCrLf & _
    "����������:�������ѱ�,��ԭ����" & vbCrLf, "���л��۵�,ѡ�л��۵�,������", Me, vbQuestion)
    Select Case strComMand
    Case "���л��۵�"
    Case "ѡ�л��۵�"
        With mshList
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("ѡ��")) = "��" Then strNos = strNos & "," & .TextMatrix(i, .ColIndex("���ݺ�"))
            Next
        End With
        If strNos = "" Then
            MsgBox "���ȹ�ѡ��Ҫ�����ĵ��ݣ���ѡ��ѱ����ͣ�", vbInformation, gstrSysName
            Exit Sub
        End If
        strNos = Mid(strNos, 2)
    Case Else
        mblnNotClick = True
        Local�ѱ� txt�ѱ�.Text, True
        mblnNotClick = False
        Exit Sub
    End Select
    
    fraOk.Enabled = False: cmdAllSel.Enabled = False: cmdAllCls.Enabled = False
    zlCommFun.ShowFlash "���ڽ�������ѱ������ʵ�ս����㣬���Ժ�..."
    Screen.MousePointer = 11
    On Error GoTo errHandle
    
    'Zl_���ﻮ��_Recalcmoney
    strSql = "Zl_���ﻮ��_Recalcmoney("
    '  ����id_In ������ü�¼.����id%Type,
    strSql = strSql & "" & mlng����ID & ","
    '  �ѱ�_In   ������ü�¼.�ѱ�%Type,
    strSql = strSql & "'" & Trim(cbo�����ѱ�.Text) & "',"
    '  Nos_In    ������ü�¼.NO%Type := Null
    strSql = strSql & IIf(strNos = "", "NULL,", "'" & strNos & "',")
    '  ��¼����_In    ������ü�¼.��¼���� %Type := 1
    strSql = strSql & "1,"
    '   �����ѱ�_In integer:=0
    strSql = strSql & IIf(strNos = "", "0)", "1)")
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    zlCommFun.StopFlash
    MsgBox "�����ѱ�ɹ�!", vbInformation + vbOKOnly, gstrSysName
    
    '�Ƴ����뵥��ʱ���ܼ������Ч�ѱ�97338
    For i = cbo�����ѱ�.ListCount - 1 To 0 Step -1
        If Val(cbo�����ѱ�.ItemData(i)) = 0 And cbo�����ѱ�.ListIndex <> i Then
            cbo�����ѱ�.RemoveItem i: Exit For
        End If
    Next
    
    fraOk.Enabled = True: cmdAllSel.Enabled = True: cmdAllCls.Enabled = True
    If strComMand = "���л��۵�" Then txt�ѱ�.Text = cbo�����ѱ�.Text
    Call cmdFind_Click
    Screen.MousePointer = 0
    Exit Sub
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    fraOk.Enabled = True: cmdAllSel.Enabled = True: cmdAllCls.Enabled = True
End Sub

Private Sub chkȱʡ_Click(Index As Integer)
    Dim i As Long, j As Long
    If mblnNotClick Then Exit Sub
    mblnNotClick = True
    If chkȱʡ(Index).Value = 1 Then
        For i = 0 To chkȱʡ.Count - 1
            If i <> Index Then
                chkȱʡ(i).Value = 0
            End If
        Next
    Else
        j = IIf(Index > 1, Index - 1, chkȱʡ.Count - 1)
        For i = 0 To chkȱʡ.Count - 1
            If i <> j Then
                chkȱʡ(i).Value = 0
            Else
                chkȱʡ(i).Value = 1
            End If
        Next
    End If
    mblnNotClick = False
    Call ShowBills
End Sub

Private Sub cmdAllCls_Click()
    Call SelBill(True)
End Sub

Private Sub cmdAllSel_Click()
 Call SelBill(False)
End Sub

Private Sub cmdCancel_Click()
    mstrBills = ""
    Unload Me
End Sub

Private Sub cmdFind_Click()
    
    If dtpBegin.Value >= dtpEnd.Value Then
        If Visible Then
            MsgBox "��ʼʱ��ӦС�ڽ���ʱ�䡣", vbInformation, gstrSysName
            dtpBegin.SetFocus
        End If
        Exit Sub
    End If
    
    Call ShowBills
    
    If Visible And mshList.Rows > 1 Then
        If mshList.TextMatrix(1, mshList.ColIndex("���ݺ�")) <> "" Then
            mshList.SetFocus
        Else
            txtPatient.SetFocus
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim strPati As String, i As Long
    Dim strDept As String
    Dim strNos As String, strNos1 As String
    Dim strNo As String
    Dim cllPro As New Collection
    Dim strSql As String
    
    If mblnCacheKeyReturn Then mblnCacheKeyReturn = False: Exit Sub
    mstrBills = ""
    If mshList.Rows < 2 Then
        MsgBox "�ò���û���κλ��۵��ݿ����շѡ�", vbInformation, gstrSysName
        txtPatient.SetFocus: Exit Sub
    End If
    If mshList.TextMatrix(1, mshList.ColIndex("���ݺ�")) = "" Then
        MsgBox "�ò���û���κλ��۵��ݿ����շѡ�", vbInformation, gstrSysName
        txtPatient.SetFocus: Exit Sub
    End If
    
    strNos = "": strNos1 = ""
    For i = 1 To mshList.Rows - 1
        strNo = Trim(mshList.TextMatrix(i, mshList.ColIndex("���ݺ�")))
        If mshList.TextMatrix(i, mshList.ColIndex("ѡ��")) <> "" Then
            If InStr(1, strNos & ",", "," & strNo & ",") = 0 Then
                strNos = strNos & "," & strNo
            End If
            '102748,���ݺŴ�С����
            mstrBills = mshList.TextMatrix(i, mshList.ColIndex("���ݺ�")) & "," & mstrBills
            If InStr(strPati & ",", "," & mshList.TextMatrix(i, mshList.ColIndex("����")) & ",") = 0 Then
                strPati = strPati & "," & mshList.TextMatrix(i, mshList.ColIndex("����"))
            End If
            If InStr(strDept & ",", "," & mshList.TextMatrix(i, mshList.ColIndex("��������")) & ",") = 0 Then
                strDept = strDept & "," & mshList.TextMatrix(i, mshList.ColIndex("��������"))
            End If
        Else
            If InStr(1, strNos1 & ",", "," & strNo & ",") = 0 Then
                strNos1 = strNos1 & "," & strNo
            End If
        End If
        If Len(strNos) >= 4000 Then
            strNos = Mid(strNos, 2)
            'Zl_�շѻ���_�ݲ�ִ��
            strSql = "Zl_�շѻ���_�ݲ�ִ��("
            '  Nos_In      Varchar2,
            strSql = strSql & "'" & strNos & "',"
            '  �ݲ�ִ��_In Integer:=-1
            strSql = strSql & "0)"
            zlAddArray cllPro, strSql
            strNos = ""
        End If
        If Len(strNos1) >= 4000 Then
            strNos1 = Mid(strNos1, 2)
            'Zl_�շѻ���_�ݲ�ִ��
            strSql = "Zl_�շѻ���_�ݲ�ִ��("
            '  Nos_In      Varchar2,
            strSql = strSql & "'" & strNos1 & "',"
            '  �ݲ�ִ��_In Integer:=-1
            strSql = strSql & "-1)"
            zlAddArray cllPro, strSql
            strNos1 = ""
        End If
    Next
    If strNos <> "" Then
         strNos = Mid(strNos, 2)
         'Zl_�շѻ���_�ݲ�ִ��
         strSql = "Zl_�շѻ���_�ݲ�ִ��("
         '  Nos_In      Varchar2,
         strSql = strSql & "'" & strNos & "',"
         '  �ݲ�ִ��_In Integer:=-1
         strSql = strSql & "0)"
         zlAddArray cllPro, strSql
         strNos = ""
     End If
     If strNos1 <> "" Then
        strNos1 = Mid(strNos1, 2)
        'Zl_�շѻ���_�ݲ�ִ��
        strSql = "Zl_�շѻ���_�ݲ�ִ��("
        '  Nos_In      Varchar2,
        strSql = strSql & "'" & strNos1 & "',"
        '  �ݲ�ִ��_In Integer:=-1
        strSql = strSql & "-1)"
        zlAddArray cllPro, strSql
        strNos1 = ""
    End If
    Err = 0: On Error GoTo ErrHand:
    '�ȴ����۵�:38281
    zlExecuteProcedureArrAy cllPro, Me.Caption
    
    Err = 0: On Error GoTo ErrHand1:
    
    If mstrBills <> "" Then '102748,���ݺŴ�С����,ȥ�����һ���ָ���
        mstrBills = Left(mstrBills, Len(mstrBills) - 1)
    End If
    If strPati <> "" Then strPati = Mid(strPati, 2)
    If strDept <> "" Then strDept = Mid(strDept, 2)
    
    If mbln������൥�� Then
        If UBound(Split(mstrBills, ",")) > 0 Then
            MsgBox "������ѡ����Ż��۵��շ�!", vbInformation, gstrSysName
            mstrBills = ""
            mshList.SetFocus: Exit Sub
        End If
    End If
    
    If mstrBills = "" Then
        MsgBox "������ѡ��һ����Ҫ�շѵĻ��۵��ݡ�", vbInformation, gstrSysName
        mstrBills = ""
        mshList.SetFocus: Exit Sub
    '    ElseIf UBound(Split(strDept, ",")) > 0 Then
    '        MsgBox "��ѡ��Ķ��ŵ������Բ�ͬ�Ŀ������ң���ֿ��շѡ�", vbInformation, gstrSysName
    '        mshList.SetFocus: Exit Sub
    ElseIf UBound(Split(mstrBills, ",")) + 1 >= 200 Then
        MsgBox "��������̫�࣬��ֳɶ���շѡ�", vbInformation, gstrSysName
        mstrBills = ""
        mshList.SetFocus: Exit Sub
    ElseIf UBound(Split(strPati, ",")) > 0 Then
        If MsgBox("ѡ��ĵ����а��������ͬ�Ĳ���������Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            mstrBills = ""
            mshList.SetFocus: Exit Sub
        End If
    End If
    
    If gblnCheckTest Then
        If Not CheckTest(mstrBills, dtpBegin.Value, dtpEnd.Value) Then
            mstrBills = ""
            mshList.SetFocus: Exit Sub
        End If
    End If
    Unload Me
    Exit Sub
ErrHand:
     gcnOracle.RollbackTrans
ErrHand1:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Activate()
 '   If mlng����ID <> 0 Then mshList.SetFocus
    If cmdOK.Enabled Then cmdOK.SetFocus
End Sub
Private Sub SelBill(Optional blnCls As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ�񵥾�
    '���:blnCls-�Ƿ����
    '����:���˺�
    '����:2011-03-16 11:21:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With mshList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("���ݺ�")) <> "" Then
                 .TextMatrix(i, .ColIndex("ѡ��")) = IIf(blnCls, "", "��")
            End If
        Next
    End With
      stbThis.Panels(2).Text = GetBillNote
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    If KeyCode = 13 Then
        If Me.ActiveControl Is mshList Then
            If Me.cmdOK.Enabled And Me.cmdOK.Visible Then
                cmdOK.SetFocus
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
         End If
    ElseIf KeyCode = vbKeyF2 Then
        If cmdOK.Visible And cmdOK.Enabled Then Call cmdOK_Click
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call SelBill(False)
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
       Call SelBill(True)
    '54538:������,2014-02-24,��ѡ�񻮼۵�������ȡʱ������ݼ�F3��֧��
    ElseIf KeyCode = vbKeyF3 Then
        mshList.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then KeyAscii = 0
End Sub
Private Function Init�ѱ�() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���طѱ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2011-04-17 21:28:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle

    strSql = _
        "Select a.����, a.����, a.����, Nvl(a.ȱʡ��־, 0) As ȱʡ, Nvl(a.���޳���, 0) As ����" & vbNewLine & _
        "From �ѱ� A" & vbNewLine & _
        "Where Nvl(a.�������, 3) In (1, 3) And a.���� = 1 And Trunc(Sysdate) Between Nvl(a.��Ч��ʼ, To_Date('1900-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
        "      And Nvl(a.��Ч����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
        "Order By a.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    mblnNotClick = True
    With cbo�����ѱ�
        .Clear
        Do While Not rsTemp.EOF
            .AddItem NVL(rsTemp!����)
            .ItemData(.NewIndex) = 1 '�����Ч�ѱ�
            If Val(NVL(rsTemp!ȱʡ)) = 1 Then .ListIndex = .NewIndex
            rsTemp.MoveNext
        Loop
    End With
    mblnNotClick = False
    Init�ѱ� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub Form_Load()
    Dim Curdate As Date, blnCancel As Boolean
    Dim i As Integer, j As Long
    
    'ѡ������������Ƿ����˻س�����
    mblnCacheKeyReturn = False
    If mblnPreCard Then
        mblnCacheKeyReturn = (GetAsyncKeyState(VK_RETURN) And &H1) <> 0
    End If
    mlngModule = 1121
    i = Val(zlDatabase.GetPara("ȱʡѡ�񻮼۵�", glngSys, mlngModule, "1", Array(chkȱʡ(0), chkȱʡ(1), chkȱʡ(2)), InStr(1, mstrPrivs, ";��������;") > 0))
    i = IIf(i > 2, 1, i): i = IIf(i < 0, 1, i)
    mblnNotClick = True
    chkȱʡ(0).Value = 0
    chkȱʡ(1).Value = 0
    chkȱʡ(2).Value = 0
    chkȱʡ(i).Value = 1
    mblnNotClick = False
    RestoreWinState Me, App.ProductName
    If InStr(1, mstrPrivs, ";�������˷ѱ�;") > 0 Then Call Init�ѱ�
    cbo�����ѱ�.Visible = InStr(1, mstrPrivs, ";�������˷ѱ�;") > 0
    lbl�����ѱ�.Visible = InStr(1, mstrPrivs, ";�������˷ѱ�;") > 0

    mstrBills = ""
    Curdate = zlDatabase.Currentdate
    dtpBegin.Value = Format(Curdate - gintSeekDays, "yyyy-MM-dd HH:mm:ss")
    dtpEnd.Value = Format(Curdate, "yyyy-MM-dd HH:mm:ss")
    dtpBegin.MaxDate = Curdate
    dtpEnd.MaxDate = Curdate
    '74296,Ƚ����,2014-7-4,�����ݵ���Ϲ���,�ѵ����е�������һ�������б�ѡ��
    cboDiagnose.Clear
    cboDiagnose.AddItem "�������"
    cboDiagnose.ListIndex = cboDiagnose.NewIndex
    
    Call SetHeader
    Call SetDetail
    Call SetActiveList
    If mlng����ID <> 0 Then
        txtPatient.Locked = True
        txtPatient.TabStop = False
        txtPatient.BackColor = &HE0E0E0
        txt�Ա�.BackColor = &HE0E0E0
        txt����.BackColor = &HE0E0E0
        txt�ѱ�.BackColor = &HE0E0E0
        txt�����.BackColor = &HE0E0E0
        txt���ʽ.BackColor = &HE0E0E0
        
        txtPatient.Text = "-" & mlng����ID
        Call txtPatient_Validate(blnCancel)
        If Not blnCancel Then
            Call cmdFind_Click
        End If
    End If
End Sub

Private Sub Form_Resize()
    Dim sngHeight As Single
    
     On Error Resume Next
    fraPati.Left = 20
    fraPati.Top = 50
    fraDays.Top = 50
    sngHeight = stbThis.Height + fraOk.Height
    
    '74296,Ƚ����,2014-7-4,�����ݵ���Ϲ���,�ѵ����е�������һ�������б�ѡ��
    lblDiagnose.Top = fraPati.Top + fraPati.Height + 50
    cboDiagnose.Top = lblDiagnose.Top - 40
    
    '59399
    With mshList
         .Left = 0
         .Top = cboDiagnose.Top + cboDiagnose.Height + 20
         .Width = ScaleWidth
         .Height = ScaleHeight - mshList.Top - IIf(mshDetail.Height < 0, 0, mshDetail.Height) - fraHsc.Height - sngHeight
    End With
    
    fraHsc.Left = 0
    fraHsc.Top = mshList.Top + mshList.Height
    fraHsc.Width = ScaleWidth
    
    mshDetail.Left = 0
    mshDetail.Top = fraHsc.Top + fraHsc.Height
    mshDetail.Width = ScaleWidth
    If Me.ScaleHeight - mshDetail.Top - sngHeight <= 1600 Then
        mshDetail.Height = 2000
    Else
        mshDetail.Height = Me.ScaleHeight - mshDetail.Top - sngHeight
    End If
    fraOk.Top = Me.ScaleHeight - sngHeight
    fraOk.Width = Me.ScaleWidth - fraOk.Left
    fraOk.ZOrder
    stbThis.ZOrder
    cmdCancel.Left = fraOk.Width + fraOk.Left - cmdCancel.Width - 50
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 20
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlng����ID = 0
    Set mrsList = Nothing
    Set mrsDetail = Nothing
    zl_vsGrid_Para_Save 0, mshList, Me.Caption, "��ͷ�б�", False, True
    SaveWinState Me, App.ProductName
    Call zlDatabase.SetPara("ȱʡѡ�񻮼۵�", IIf(chkȱʡ(0).Value = 1, 0, IIf(chkȱʡ(1).Value = 1, 1, 2)), glngSys, mlngModule, InStr(1, mstrPrivs, ";��������;") > 0)
End Sub

Private Sub fraHsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If mshList.Height + Y < 1000 Or mshDetail.Height - Y < 1000 Then Exit Sub
        fraHsc.Top = fraHsc.Top + Y
        mshList.Height = mshList.Height + Y
        mshDetail.Top = mshDetail.Top + Y
        mshDetail.Height = mshDetail.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub mshDetail_LostFocus()
    Call SetActiveList
End Sub

Private Sub mshList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save 0, mshList, Me.Caption, "��ͷ�б�", False, True
End Sub

Private Sub mshList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
  zl_vsGrid_Para_Save 0, mshList, Me.Caption, "��ͷ�б�", False, True
  With mshList
        If .ColIndex("���") >= 0 Then
            .AutoSize .ColIndex("���"), .ColIndex("���")
        End If
    End With
End Sub
Private Sub mshList_DblClick()
    Call mshList_KeyPress(32)
End Sub

Private Sub mshList_EnterCell()
    Dim strNo As String
    strNo = mshList.TextMatrix(mshList.Row, mshList.ColIndex("���ݺ�"))
    If mshList.Row = 0 Or strNo = "" Then Exit Sub
    Call ShowDetail(strNo)
End Sub

Private Sub mshList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        KeyAscii = 0
        If mshList.TextMatrix(mshList.Row, mshList.ColIndex("���ݺ�")) <> "" Then
            If mshList.TextMatrix(mshList.Row, mshList.ColIndex("ѡ��")) = "" Then
                mshList.TextMatrix(mshList.Row, mshList.ColIndex("ѡ��")) = "��"
            Else
                mshList.TextMatrix(mshList.Row, mshList.ColIndex("ѡ��")) = ""
            End If
        End If
        stbThis.Panels(2).Text = GetBillNote
    End If
End Sub

Private Sub mshList_LostFocus()
    Call SetActiveList
End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Integer
    strHead = "���,1,3500|ѡ��,4,500|���ݺ�,4,850|��������,1,1200|ҽ��,1,800|����,1,800|�Ա�,4,500|����,4,500|Ӧ�ս��,7,850|ʵ�ս��,7,850|������,1,800|����ʱ��,4,1850|Ƥ��,4,500"
    With mshList
        .Redraw = flexRDNone
        .COLS = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            .ColKey(i) = .TextMatrix(0, i)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .RowHeight(0) = 320
        zl_vsGrid_Para_Restore 0, mshList, Me.Caption, "��ͷ�б�", False, False
        'If Not Visible Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name)
        .Col = 0: .ColSel = .COLS - 1
        If .ColIndex("���") >= 0 Then
            .AutoSize .ColIndex("���"), .ColIndex("���")
        End If
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Integer
    
    strHead = "���,1,750|��Ŀ,1,2000" & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "|��Ʒ��,1,2000", "") & "|���,1,1000|��λ,4,500|����,7,850|�ѱ�,1,750|����,7,850|Ӧ�ս��,7,850|ʵ�ս��,7,850|ִ�п���,1,850|ժҪ,1,2000"
    
    With mshDetail
        .Redraw = False
        
        .COLS = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        If Not Visible Then Call RestoreFlexState(mshDetail, App.ProductName & "\" & Me.Name)
        For i = 0 To .COLS - 1
            If .TextMatrix(0, i) = "��Ʒ��" Then
                If gTy_System_Para.bytҩƷ������ʾ = 2 Then
                    If .ColWidth(i) <= 0 Then .ColWidth(i) = 2000
                Else
                    .ColWidth(i) = 0
                End If
            End If
        Next
        .RowHeight(0) = 320
        .Row = 1: .Col = 0: .ColSel = .COLS - 1
        .Redraw = True
    End With
End Sub

Private Sub SetActiveList(Optional obj As Object)
    If obj Is mshList Then
        mshList.BackColorSel = &H8000000D
        mshDetail.BackColorSel = &H8000000C
    ElseIf obj Is mshDetail Then
        mshList.BackColorSel = &H8000000C
        mshDetail.BackColorSel = &H8000000D
    Else
        mshList.BackColorSel = &H8000000C
        mshDetail.BackColorSel = &H8000000C
    End If
End Sub

Private Sub mshDetail_GotFocus()
    Call SetActiveList(mshDetail)
End Sub

Private Sub mshList_GotFocus()
    Call SetActiveList(mshList)
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Integer
    For i = 0 To mshList.COLS - 1
        If mshList.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
End Function

Private Sub ShowDetail(ByVal strNo As String)
    Dim i As Integer, strSql As String
    
    On Error GoTo errH
    
    strSql = _
    " Select C.���� as ���,'['||B.����||']'||Nvl(E.����,B.����) as ��Ŀ," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "E1.���� as ��Ʒ��,", "") & "B.���," & _
            IIf(gblnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & gstrҩ����λ & ")", "A.���㵥λ") & " as ��λ," & _
    "       Ltrim(To_Char(Avg(Nvl(A.����,1)*A.����)" & _
            IIf(gblnҩ����λ, "/Nvl(X." & gstrҩ����װ & ",1)", "") & ",'999990.00000')) as ����, " & _
    "       A.�ѱ�,Ltrim(To_Char(Sum(A.��׼����)" & _
            IIf(gblnҩ����λ, "*Nvl(X." & gstrҩ����װ & ",1)", "") & ",'99999" & gstrFeePrecisionFmt & "')) as ����, " & _
    "       Ltrim(To_Char(Sum(A.Ӧ�ս��),'99999" & gstrDec & "')) as Ӧ�ս��, " & _
    "       Ltrim(To_Char(Sum(A.ʵ�ս��),'99999" & gstrDec & "')) as ʵ�ս��, " & _
    "       D.���� as ִ�п���,A.ժҪ" & _
    " From ������ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,�շ���Ŀ���� E,ҩƷ��� X" & _
            IIf(gTy_System_Para.bytҩƷ������ʾ = 2, ",�շ���Ŀ���� E1", "") & _
    " Where A.�շ�ϸĿID=B.ID and A.�շ����=C.���� And A.ִ�в���ID=D.ID(+) And A.�շ�ϸĿID=X.ҩƷID(+)" & _
    "       And A.��¼����=1 and A.��¼״̬ IN(0,1,3) And A.NO=[1]" & _
    "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
            IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3", "") & _
    " Group by Nvl(A.�۸񸸺�,A.���),C.����,B.����,Nvl(E.����,B.����)," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "E1.����,", "") & " B.���,A.���㵥λ,A.�ѱ�," & _
    "       D.����,A.ժҪ,X.ҩƷID,X." & gstrҩ����λ & ",Nvl(X." & gstrҩ����װ & ",1)" & _
    " Order by Nvl(A.�۸񸸺�,A.���)"
    Set mrsDetail = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo)
    mshDetail.Clear
    mshDetail.Rows = 2
    If Not mrsDetail.EOF Then
        Set mshDetail.DataSource = mrsDetail
    End If
    Call SetDetail
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub ShowBills()
    Dim strSql As String, i As Long
    Dim bytType As Byte
    Dim strTemp As String, strTempCbo As String, sngCboWidth As Long
    
    On Error GoTo errHandle
    bytType = IIf(chkȱʡ(0).Value = 1, 0, IIf(chkȱʡ(1).Value = 1, 1, 2))
    strTempCbo = IIf(cboDiagnose.Text = "�������", "", cboDiagnose.Text)
    Screen.MousePointer = 11
    Set mrsList = GetPriceBills(mlng����ID, mlng�Һſ���, dtpBegin.Value, dtpEnd.Value, True, bytType)
    mshList.Clear
    mshList.Rows = 2
    mshDetail.Clear
    mshDetail.Rows = 2
    stbThis.Panels(2).Text = ""
    
    '74296,Ƚ����,2014-7-4,�����ݵ���Ϲ���,�ѵ����е�������һ�������б�ѡ��
    If Not mrsList.EOF Then
        mrsList.MoveFirst
        Do While Not mrsList.EOF
            strTemp = NVL(mrsList!���)
            mblnNotClick = True
            If zlControl.CboLocate(cboDiagnose, strTemp) = False Then
                cboDiagnose.AddItem strTemp '����������
            End If
            mblnNotClick = False
            If sngCboWidth < Me.TextWidth(strTemp) Then
                sngCboWidth = Me.TextWidth(strTemp) '�����ı������
            End If
            mrsList.MoveNext
        Loop
        '����������Ŀ��
        If sngCboWidth + 300 > cboDiagnose.Width Then zlControl.CboSetWidth cboDiagnose.hWnd, sngCboWidth + 300
        
        mrsList.Filter = IIf(strTempCbo = "", "", "���='" & strTempCbo & "'")
        Set mshList.DataSource = mrsList
        stbThis.Panels(2).Text = GetBillNote
    End If
    
    Call SetHeader
    Call SetDetail
    Call mshList_EnterCell
    mblnNotClick = True
    If zlControl.CboLocate(cboDiagnose, IIf(strTempCbo = "", "�������", strTempCbo)) = False Then cboDiagnose.ListIndex = 0 '���¶�λ
    mblnNotClick = False
    
    Me.Refresh
    Screen.MousePointer = 0
    Exit Sub
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
End Sub

Private Function GetBillNote() As String
    Dim curTotal As Currency, i As Long, k As Long
    
    k = 0
    For i = 1 To mshList.Rows - 1
        If mshList.TextMatrix(i, mshList.ColIndex("ѡ��")) <> "" Then
            k = k + 1
            curTotal = curTotal + Val(mshList.TextMatrix(i, mshList.ColIndex("ʵ�ս��")))
        End If
    Next
    If k > 0 Then
        GetBillNote = "��ǰѡ���� " & k & " �ŵ��ݣ��ϼ� " & Format(curTotal, gstrDec) & " Ԫ"
    End If
End Function

Private Sub txtPatient_GotFocus()
    mblnCard = False
    Call zlControl.TxtSelAll(txtPatient)
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    mblnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, gobjSquare.blnȱʡ��������)
    'ˢ���Զ�ȷ��
    If mblnCard And Len(txtPatient.Text) = gobjSquare.blnȱʡ�������� - 1 And KeyAscii <> 8 And KeyAscii <> 13 Then
        txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
        txtPatient.SelStart = Len(txtPatient.Text)
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
        
    If txtPatient.Text <> "" Then
        If txtPatient.Text <> txtPatient.Tag Then
            Set rsTmp = GetPatient(txtPatient.Text, mblnCard)
            If rsTmp Is Nothing Then
                If Visible Then MsgBox "δ�ҵ����˵���Ϣ��", vbInformation, gstrSysName
                txtPatient.Text = ""
            Else
                
                '���￨������
                If Mid(gstrCardPass, 3, 1) = "1" And mblnCard Then
                    If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, rsTmp!����, "" & rsTmp!�Ա�, "" & rsTmp!����) Then
                        Set rsTmp = Nothing: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
                    End If
                End If
            
                txtPatient.PasswordChar = ""
                '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
                txtPatient.IMEMode = 0
                txtPatient.Text = NVL(rsTmp!����)
                txtPatient.Tag = txtPatient.Text
                mlng����ID = NVL(rsTmp!����ID, 0)
                txt�Ա�.Text = NVL(rsTmp!�Ա�)
                txt����.Text = NVL(rsTmp!����)
                txt�ѱ�.Text = NVL(rsTmp!�ѱ�)
                txt�����.Text = NVL(rsTmp!�����)
                txt���ʽ.Text = NVL(rsTmp!ҽ�Ƹ��ʽ)
                If InStr(1, mstrPrivs, ";�������˷ѱ�;") > 0 Then
                    mblnNotClick = True
                    Local�ѱ� Trim(txt�ѱ�.Text), True
                    mblnNotClick = False
                End If
            End If
        End If
    End If
    If txtPatient.Text = "" Then
        mlng����ID = 0
        txtPatient.Tag = ""
        txt�Ա�.Text = ""
        txt����.Text = ""
        txt�ѱ�.Text = ""
        txt�����.Text = ""
        txt���ʽ.Text = ""
        Cancel = True: Exit Sub
    End If
End Sub

Private Function GetPatient(ByVal strInput As String, ByVal blnCard As Boolean) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:blnCard=�Ƿ���￨ˢ��
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-03 16:47:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strWhere As String
    Dim strPati As String, vRect As RECT
    
    strInput = UCase(strInput)
    
    '���������Ȩ��
    If gint������Դ = 1 Then
        'strWhere = " And Nvl(A.��ǰ����ID,0)=0"
        If Not mblnסԺ���������շ� Then    '34182
            strWhere = " And Not Exists(Select 1 From ������ҳ Where ����ID=A.����ID And ��ҳID<>0 And ��ҳID=A.��ҳID And Nvl(��������,0)=0 And ��Ժ���� is Null)"
        End If
    ElseIf gint������Դ = 2 Then
        If Not mblnסԺ���������շ� Then    '34182
            strWhere = " And Nvl(A.��ǰ����ID,0)<>0"
        End If
    End If
    
    '��ȡ������Ϣ
    strSql = "Select A.����ID,A.�����,A.����,A.�Ա�,A.����,A.���￨��,A.����֤��,A.�ѱ�,A.ҽ�Ƹ��ʽ,A.��������,A.���� From ������Ϣ A Where 1=1"
    If blnCard Then '���￨��
        If gint������Դ = 1 And Not gblnInputCard Then Exit Function
        '������:27364
        If Not gobjSquare.objDefaultCard Is Nothing And gobjSquare.bln��ȱʡ������ Then
            lng�����ID = gobjSquare.objDefaultCard.�ӿ����
        Else
            lng�����ID = "-1"
        End If
        
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg, lng�����ID) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strSql = strSql & strWhere & " And A.����ID=[1] "
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
        strSql = strSql & strWhere & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
        strSql = strSql & strWhere & " And A.�����=[1]"
        '75087,Ƚ����,2014-7-29,���ﲡ���շ�ʱ,����Ҫ���������������,ֻ��Ҫ��������ŵ����˳��ż����ҵ��������Ĳ�����Ϣ������
        strInput = "*" & zlCommFun.GetFullNO(Mid(strInput, 2), 3)
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��
        strSql = strSql & strWhere & " And A.סԺ��=[1]"
    ElseIf Left(strInput, 1) = "." Then '�Һŵ���
        If gint������Դ = 1 And Not gblnInputNO Then Exit Function
        '���ջ���˳���Ź���
        strInput = GetFullNO(Mid(strInput, 2), 12)
        strSql = "" & _
        " Select A.����ID,Nvl(A.�����,B.��ʶ��) as �����,A.��������,A.����," & _
        "       Nvl(A.����,B.����) as ����,Nvl(A.�Ա�,B.�Ա�) as �Ա�," & _
        "       Nvl(A.����,B.����) as ����,A.���￨��,A.����֤��,Nvl(A.�ѱ�,B.�ѱ�) as �ѱ�," & _
        "       Nvl(A.ҽ�Ƹ��ʽ,C.����) as ҽ�Ƹ��ʽ" & _
        " From ������Ϣ A,������ü�¼ B,ҽ�Ƹ��ʽ C" & _
        " Where B.��¼����=4 And B.��¼״̬=1 And B.NO=[2]" & _
        "       And B.����ID=A.����ID(+) And B.���ʽ=C.����(+)" & strWhere & _
             zlGetRegEventsCons("�Ӱ��־", "B")
    Else
        'ͨ������ģ�����Ҳ���
        If gblnSeekName Then
            strWhere = " A.���� Like '" & strInput & "%' " & strWhere
            strPati = _
                " Select A.����ID as ID,A.����ID,A.����,A.�Ա�,A.����,A.���￨��,A.����֤��,A.�����,A.��������,A.���֤��,A.��ͥ��ַ,A.������λ" & _
                " From ������Ϣ A Where " & strWhere & _
                IIf(gintNameDays = 0, "", " And (A.����ʱ��>Trunc(Sysdate-" & gintNameDays & ") Or A.�Ǽ�ʱ��>Trunc(Sysdate-" & gintNameDays & "))") & _
                " And Rownum<101" & _
                " Order by A.����"
            vRect = zlControl.GetControlRect(txtPatient.hWnd)
            Set rsTmp = zlDatabase.ShowSelect(Me, strPati, 0, "����Find", , , , , , True, vRect.Left, vRect.Top, txtPatient.Height, , , True)
            If rsTmp Is Nothing Then Exit Function
            If rsTmp.EOF Then Exit Function
            strInput = rsTmp!����ID
            strSql = strSql & strWhere & " And A.����ID=[2]"
        Else
            Exit Function
        End If
    End If
        
    On Error GoTo errH
    '75259:���ϴ�,2014-7-10������������ʾ��ɫ����
    txtPatient.ForeColor = Me.ForeColor
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Mid(strInput, 2)), strInput)
    If Not rsTmp.EOF Then
        Call SetPatiColor(txtPatient, NVL(rsTmp!��������), IIf(IsNull(rsTmp!����), txtPatient.ForeColor, vbRed))
        mstrPassWord = strPassWord
        If Not blnHavePassWord Then mstrPassWord = NVL(rsTmp!����֤��)
        Set GetPatient = rsTmp
    End If
    Exit Function
NotFoundPati:
    Set GetPatient = Nothing
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
