VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ŀѡ��"
   ClientHeight    =   7695
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   12270
   Icon            =   "frmClinicSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   12270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   12270
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7065
      Width           =   12270
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   9810
         TabIndex        =   10
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   11055
         TabIndex        =   9
         Top             =   135
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Height          =   350
         Left            =   870
         TabIndex        =   8
         Top             =   135
         Width           =   1845
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "��λ(&L)"
         Height          =   350
         Left            =   2700
         TabIndex        =   7
         Top             =   135
         Width           =   1100
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "��Ŀ����"
         Height          =   180
         Left            =   90
         TabIndex        =   12
         Top             =   210
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "�������������"
         ForeColor       =   &H00008000&
         Height          =   180
         Left            =   3975
         TabIndex        =   11
         Top             =   210
         Width           =   1260
      End
   End
   Begin VB.CommandButton cmdBound 
      Caption         =   "������"
      Height          =   350
      Left            =   4350
      TabIndex        =   5
      Top             =   3945
      Width           =   1100
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "�������"
      Height          =   350
      Left            =   11115
      TabIndex        =   4
      Top             =   3855
      Width           =   1100
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5685
      Left            =   3630
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5685
      ScaleWidth      =   45
      TabIndex        =   2
      Top             =   240
      Width           =   45
   End
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   12515
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgItem 
      Height          =   3750
      Left            =   3675
      TabIndex        =   1
      Top             =   30
      Width           =   8550
      _cx             =   15081
      _cy             =   6615
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      ExplorerBar     =   3
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
      Begin MSComctlLib.ImageList imgSort 
         Left            =   930
         Top             =   900
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   9
         ImageHeight     =   8
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClinicSelect.frx":6852
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClinicSelect.frx":68B0
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgOften 
      Left            =   0
      Top             =   645
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":690E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":7008
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":7702
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":7DFC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":84F6
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":8A90
            Key             =   "Expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":902A
            Key             =   "��ҩ"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":95C4
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":9B5E
            Key             =   "��ҩ"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":A0F8
            Key             =   "����"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgBound 
      Height          =   2790
      Left            =   3675
      TabIndex        =   3
      Top             =   4320
      Width           =   8550
      _cx             =   15081
      _cy             =   4921
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      ExplorerBar     =   3
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
End
Attribute VB_Name = "frmClinicSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mblnDown As Boolean
Private mstrIDs As String
Private mstrNAMEs As String
Private mstrPreNode As String
Private mstrLike As String
Private mrsItem As New ADODB.Recordset
Private mrsFind As New ADODB.Recordset

Private Function FillTree() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objNode As node
    
    On Error GoTo errH
    
    strSQL = _
        " Select 0 as ��,����,-���� as ID,-Null as �ϼ�ID,����||'' as ����," & _
        " ����||'.'||Decode(����,1,'����ҩ',2,'�г�ҩ',3,'�в�ҩ',4,'��ҩ�䷽',5,'������Ŀ',6,'��������','7','��������') as ����" & _
        " From ���Ʒ���Ŀ¼ Where ����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Group by ����"
    strSQL = strSQL & " Union ALL " & _
        " Select Level as ��,����,ID,Nvl(�ϼ�ID,-����) as �ϼ�ID,����,���� From ���Ʒ���Ŀ¼" & _
        " Where ����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD')" & _
        " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
        " Order by ��,����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name)
        
    For i = 1 To rsTmp.RecordCount
        If IsNull(rsTmp!�ϼ�ID) Then
            Set objNode = tvw_s.Nodes.Add(, , "_" & rsTmp!ID, rsTmp!����, "Close")
        Else
            Set objNode = tvw_s.Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!ID, "[" & rsTmp!���� & "]" & rsTmp!����, "Close")
        End If
        objNode.Tag = rsTmp!���� '��ŷ�������
        objNode.ExpandedImage = "Expend"
        rsTmp.MoveNext
    Next
    If tvw_s.Nodes.Count > 0 Then
        tvw_s.Nodes(1).Expanded = True
        If tvw_s.Nodes(1).Children > 0 Then
            tvw_s.Nodes(1).Child.Selected = True
        Else
            tvw_s.Nodes(1).Selected = True
        End If
        tvw_s.SelectedItem.EnsureVisible
        Call tvw_s_NodeClick(tvw_s.SelectedItem)
    End If
    
    FillTree = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function FillList() As Boolean
'���ܣ����ݵ�ǰ��������װ��������ĿĿ¼
    Dim objNode As node, objItem As ListItem
    Dim strSQL As String, strInside As String
    Dim arrClass As Variant, strclass As String
    Dim strSub As String, str�������� As String
    Dim str�Ա� As String, strStock As String
    Dim strInput As String, lngҩ��ID As Long
    Dim blnLoad As Boolean, objTab As MSComctlLib.Tab
    Dim str��Χ As String, strҩƷ As String
    Dim blnOften As Boolean, blnStock As Boolean
    Dim str������� As String, strPriv As String
    Dim i As Long, j As Long
    Dim strCommIF As String, strScope As String
    
    Dim lng����ID As Long, int���� As Integer, str��� As String

    Set objNode = tvw_s.SelectedItem '����ΪNothing
    
    '�����Ŀ�嵥�����࿨Ƭ
    '------------------------------------------------------------------------
    vfgItem.Rows = vfgItem.FixedRows
    vfgItem.Rows = vfgItem.FixedRows + 1
    Me.Refresh
    
    '��ȡ����
    int���� = Val(objNode.Tag)
    lng����ID = Val(Mid(objNode.Key, 2))
    If Val(Mid(objNode.Key, 2)) < 0 Then
        strSub = " And A.����ID IN(" & _
            " Select ID From ���Ʒ���Ŀ¼ Where ����=[1] And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " )"
    Else
        strSub = " And A.����ID IN(" & _
            " Select ID From ���Ʒ���Ŀ¼ Where ����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD')" & _
            " Start With ID=[2] Connect by Prior ID=�ϼ�ID)"
    End If
    
    '��Ʒ���´�ĳ���
    blnLoad = InStr(",1,2,3,", Val(objNode.Tag)) > 0
    If blnLoad Then
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select A.��� As ���ID,A.ID as ������ĿID,-Null as �շ�ϸĿID," & _
                " F.���� As ���,Null as ����,A.����,A.����,Null as ��Ʒ��," & _
                "A.���㵥λ,Null as ���,Null as ����, D.ҩƷ����," & _
                "Null as ��������,Null as ˵��,D.����ְ�� as ����ְ��ID" & _
            " From ҩƷ���� D,������Ŀ��� F,������ĿĿ¼ A" & _
            " Where A.ID=D.ҩ��ID And A.���=F.���� And A.��� IN ('5','6','7')" & strCommIF & strSub
    End If
        
    '2.��ҩƷ���ĵ�������Ŀ����:���಻��ҩƷ����ʱ���ض�ȡ
    '--------------------------------------------------------------------------------------
    blnLoad = InStr(",1,2,3,7,", Val(objNode.Tag)) = 0
    If blnLoad Then
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select " & _
                " A.��� As ���ID,A.ID as ������ĿID,-Null as �շ�ϸĿID,D.���� As ���,Null as ����," & _
                " A.����,A.����,Null as ��Ʒ��,A.���㵥λ,A.�걾��λ as ���,Null as ����," & _
                " Null as ҩƷ����,Null as ��������,Null as ˵��,Null as ����ְ��ID" & _
            " From ������Ŀ��� D,������ĿĿ¼ A" & _
            " Where A.���=D.���� And A.��� Not IN('4','5','6','7')" & strCommIF & strSub
    End If
    
    blnLoad = Val(objNode.Tag) = 7
    If blnLoad Then
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select A.��� AS ���ID,E.ID as ������ĿID,A.ID as �շ�ϸĿID," & _
                " F.���� AS ���,Null as ����,A.����,A.���� as ����,Null as ��Ʒ��,A.���㵥λ,A.���,A.����," & _
                " Null as ҩƷ����,Null as ��Ŀ����,A.��������,A.˵��,Null as ����ְ��ID" & _
            " From �շ���ĿĿ¼ A,�������� C,������ĿĿ¼ E,�շ���Ŀ��� F" & _
            " Where A.ID=C.����ID And C.����ID=E.ID And A.���=F.���� And E.���='4' And C.�������=0" & _
                " And A.���='4'" & strCommIF & strSub & _
                " And (E.������� IN(2,3)) " & _
                " And (E.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or E.����ʱ�� IS NULL)"
    End If
    
    strSQL = "Select Rownum as KeyID,A.* From (" & strSQL & ") A Order by Decode(���ID,'4','Z',���ID),���,����"
    
    On Error GoTo errH
    Screen.MousePointer = 11
    'Set mrsItem = New ADODB.Recordset
    Set mrsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Name, int����, lng����ID)
    
    '������
    '--------------------------------------------------------------------------
    vfgItem.Redraw = flexRDNone
    
    vfgItem.Rows = 2
    vfgItem.FixedRows = 1
    vfgItem.Cols = 2
    vfgItem.FixedCols = 1
    vfgItem.RowData(1) = 0
    
    vfgItem.ScrollBars = flexScrollBarNone
    Set vfgItem.DataSource = mrsItem
    vfgItem.ScrollBars = flexScrollBarBoth
    If err.Number = 0 And gcnOracle.Errors.Count > 0 Then
        gcnOracle.Errors.Clear
    End If
    If vfgItem.Rows = vfgItem.FixedRows Then
        vfgItem.Rows = vfgItem.FixedRows + 1
    End If
    
    '�����Ե���
    vfgItem.ColAlignment(0) = 4
    vfgItem.Cell(flexcpAlignment, 0, 0, 0, vfgItem.Cols - 1) = 4
    vfgItem.RowHeight(0) = vfgItem.RowHeightMin
    
    '��Ƭ������ݼ���
    '------------------------------------------------------------------------
    For i = 1 To mrsItem.RecordCount
        vfgItem.TextMatrix(i, 0) = i
        vfgItem.RowHeight(i) = vfgItem.RowHeightMin
        vfgItem.RowData(i) = Val(mrsItem!������ĿID)
        mrsItem.MoveNext
    Next
    
    '���ݽ�����������������һЩ����Ҫ����
    For i = 1 To vfgItem.Cols - 1
        If InStr(1, ",KEYID,���ID,�շ�ϸĿID,����,����ְ��ID,", "," & vfgItem.TextMatrix(0, i) & ",") <> 0 Then vfgItem.ColHidden(i) = True
    Next
    
    '�к��п��
    vfgItem.ColWidth(0) = Me.TextWidth(vfgItem.TextMatrix(vfgItem.Rows - 1, 0) & " ")
    If vfgItem.ColWidth(0) < 380 Then vfgItem.ColWidth(0) = 380
    
    vfgItem.FrozenCols = 0
    vfgItem.Editable = flexEDNone
    vfgItem.SheetBorder = vfgItem.BackColor
    
    vfgItem.Row = vfgItem.FixedRows: vfgItem.Col = vfgItem.FixedCols
    vfgItem.Redraw = flexRDDirect
        
    Call Form_Resize
    
    Screen.MousePointer = 0
    FillList = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ShowMe(ByVal frmParent As Form, strIDs As String, strNames As String) As Boolean
    mblnOK = False
    mstrIDs = strIDs
    mstrNAMEs = strNames
    Me.Show 1, frmParent
    ShowMe = mblnOK
    strIDs = mstrIDs
    strNames = mstrNAMEs
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdBound_Click()
    Dim node As MSComctlLib.node
    Dim strSel As String, strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim blnAdd As Boolean
    
    On Error GoTo ErrHand
    Set node = tvw_s.Nodes(1)
    strSel = GetSelNodes(node)
    
    If strSel <> "" Then
        strSQL = " Select /*+ Rule */ a.Id, b.���� ���,a.����,a.����" & vbNewLine & _
                " From ������ĿĿ¼ A,������Ŀ��� B,(Select Column_Value From Table(f_Num2list([1]))) C" & vbNewLine & _
                " Where  a.��� = b.���� And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And A.����Id = C.Column_Value" & vbNewLine & _
                " Order By a.���, a.����"

        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ŀ", strSel)
        Do While Not rsTemp.EOF
            blnAdd = True
            For lngRow = vfgBound.FixedRows To vfgBound.Rows - 1
                If Val(vfgBound.TextMatrix(lngRow, 1)) = Val(rsTemp!ID) Then
                    blnAdd = False
                    Exit For
                End If
            Next lngRow
            If blnAdd = True Then
                If Val(vfgBound.TextMatrix(vfgBound.Rows - 1, 1)) > 0 Or vfgBound.Rows = vfgBound.FixedRows Then
                    vfgBound.Rows = vfgBound.Rows + 1
                End If
                vfgBound.TextMatrix(vfgBound.Rows - 1, 0) = vfgBound.Rows - vfgBound.FixedRows
                vfgBound.TextMatrix(vfgBound.Rows - 1, 1) = Val(rsTemp!ID)
                vfgBound.TextMatrix(vfgBound.Rows - 1, 2) = CStr(Nvl(rsTemp!���))
                vfgBound.TextMatrix(vfgBound.Rows - 1, 3) = CStr(Nvl(rsTemp!����))
                vfgBound.TextMatrix(vfgBound.Rows - 1, 4) = CStr(Nvl(rsTemp!����))
                vfgBound.RowData(vfgBound.Rows - 1) = Val(rsTemp!ID)
                vfgBound.Row = vfgBound.Rows - 1
                vfgBound.TopRow = vfgBound.Rows - 1
                If vfgBound.ColWidth(0) < 380 Then vfgBound.ColWidth(0) = 380
            End If
            rsTemp.MoveNext
        Loop
        vfgBound.ColWidth(0) = Me.TextWidth(vfgBound.TextMatrix(vfgBound.Rows - 1, 0) & " ")
    Else
        MsgBox "������Ҫ��ѡһ��Ҫ��ӵĽڵ㣬��������б��н��й�ѡ��", vbInformation, gstrSysName
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdClear_Click()
    '������а󶨵���Ŀ
    Call FillBoundItem("")
End Sub

Private Sub cmdFind_Click()
'����:�ʾ����
    Dim strText As String, strMatch As String
    Dim strFind As String, strSQL As String
    Dim lngRow As Long, lngTypeID As Long
    
    On Error GoTo ErrHand
    
    If mrsFind.State = adStateOpen Then
        If Not mrsFind.EOF Then mrsFind.MoveNext
        Call LocaItem
        Exit Sub
    End If
    
    If Trim(txtFind.Text) = "" Then
        If txtFind.Enabled And txtFind.Visible Then txtFind.SetFocus
        Exit Sub
    End If
    
    If InStr(1, txtFind.Text, "'") > 0 Then
        MsgBox "��������ݰ����Ƿ��ַ� ' ,����!", vbInformation, gstrSysName
        If txtFind.Enabled And txtFind.Visible Then txtFind.SetFocus
        Exit Sub
    End If
    
    If Not tvw_s.SelectedItem Is Nothing Then
        lngTypeID = Val(Mid(tvw_s.SelectedItem.Key, 2))
    Else
        lngTypeID = 0
    End If
    
    strText = mstrLike & txtFind.Text & "%"
    If ZLCommFun.IsCharChinese(txtFind.Text) Then
        strFind = " And A.���� Like '" & strText & "'"
    ElseIf IsNumeric(txtFind.Text) Then
        strFind = " And A.���� Like '" & strText & "'"
    Else
        strFind = " And zlspellcode(A.����) Like '" & UCase(strText) & "'"
    End If
    
    '���������������ȡƥ��Ĵʾ�
    strSQL = " Select a.����id, a.Id,b.�ϼ�ID" & vbNewLine & _
            " From ������ĿĿ¼ a, ���Ʒ���Ŀ¼ b" & vbNewLine & _
            " Where a.����id = b.Id And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
            "      (b.����ʱ�� Is Null Or b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & strFind & _
            " Order by" & IIf(lngTypeID = 0, "", " DECODE(A.����ID," & lngTypeID & ",0,1),") & " b.����,b.����,a.����"
    Set mrsFind = zlDatabase.OpenSQLRecord(strSQL, "��Ŀ����")

    Call LocaItem
        
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub LocaItem()
    Dim lngRow As Long
    
    If mrsFind.RecordCount = 0 Then
        lblInfo.Caption = "û���ҵ�������������Ϣ"
        lblInfo.ForeColor = &HFF&
        Exit Sub
    End If
    
    If mrsFind.EOF = True Then
        lblInfo.Caption = "�Ѿ�������ж�λ����������������"
        lblInfo.ForeColor = &HFF&
        Exit Sub
    End If
    lblInfo.Caption = "���ҵ�" & mrsFind.RecordCount & "��,��ǰ�ǵ�" & mrsFind.AbsolutePosition & "��"
    lblInfo.ForeColor = &H8000000D
    
    If mrsFind.RecordCount > 0 Then
        If mrsFind.RecordCount <> mrsFind.AbsolutePosition Then
            cmdFind.Caption = "��һ��(&L)"
        Else
            cmdFind.Caption = "��λ(&L)"
            lblInfo.Caption = "�Ѿ������һ������������������"
        End If
    End If
    On Error Resume Next
    err.Clear: err = 0
    '��ʼ���ж�λ
    tvw_s.Nodes("_" & mrsFind!����id).Selected = True
    If err <> 0 Then err.Clear: Exit Sub
    Call tvw_s_NodeClick(tvw_s.Nodes("_" & mrsFind!����id))
    
    For lngRow = vfgItem.FixedRows To vfgItem.Rows - 1
        If Val(vfgItem.RowData(lngRow)) = Val(mrsFind!ID) Then
            vfgItem.Row = lngRow
            vfgItem.TopRow = lngRow
            Exit For
        End If
    Next lngRow
End Sub

Private Sub cmdOK_Click()
    Dim lngRow As Long
    Dim strIDs As String, strNames As String
    With vfgBound
        For lngRow = .FixedRows To .Rows - 1
            If Val(.TextMatrix(lngRow, 1)) > 0 Then
                strIDs = strIDs & "," & Val(.TextMatrix(lngRow, 1))
                strNames = strNames & "," & .TextMatrix(lngRow, 4)
            End If
        Next
    End With
    
    mstrIDs = Mid(strIDs, 2)
    mstrNAMEs = Mid(strNames, 2)
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    mstrPreNode = ""
    mblnDown = False
    mstrLike = IIf(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    Call FillTree
    Call FillBoundItem(mstrIDs)
End Sub

Private Sub FillBoundItem(ByVal strIDs As String)
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
    Dim strSQL As String
    
    strSQL = "Select /*+ Rule */ a.Id, b.���� ���, a.����, a.����" & vbNewLine & _
            " From ������ĿĿ¼ a, ������Ŀ��� b, (Select Column_Value From Table(f_Num2list([1]))) c" & vbNewLine & _
            " Where a.��� = b.���� And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And a.Id = c.Column_Value" & vbNewLine & _
            " Order By a.���, a.����"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�Ѿ���������Ŀ", strIDs)
    
    With vfgBound
        .Redraw = flexRDNone
        
        '����ͳ�Ƴ�����Ŀʱ����Ϊ��0��0��
        .Rows = 2
        .FixedRows = 1
        .Cols = 5
        .FixedCols = 1
        .RowData(1) = 0
        .ScrollBars = flexScrollBarNone
        Set .DataSource = rsTemp
        .ScrollBars = flexScrollBarBoth
        .ColHidden(1) = True
        If .Rows = .FixedRows Then
            .Rows = .FixedRows + 1
        End If
        
        '�����Ե���
        .ColAlignment(0) = 4
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
        .RowHeight(0) = .RowHeightMin
        
        '��Ƭ������ݼ���
        '------------------------------------------------------------------------
        For i = 1 To rsTemp.RecordCount
            .TextMatrix(i, 0) = i
            .RowHeight(i) = .RowHeightMin
            .RowData(i) = Val(rsTemp!ID)
             rsTemp.MoveNext
        Next
        
        '�к��п��
        .ColWidth(0) = Me.TextWidth(.TextMatrix(.Rows - 1, 0) & " ")
        If .ColWidth(0) < 380 Then .ColWidth(0) = 380
        .ColWidth(2) = 800
        .ColWidth(3) = 1000
        .ColWidth(4) = 2000
        
        .FrozenCols = 0
        .Editable = flexEDNone
        .SheetBorder = .BackColor
        
        .Row = .FixedRows: .Col = .FixedCols
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    With picSplit
        .Top = 0
        .Height = Me.ScaleHeight - picBottom.Height
    End With
    With tvw_s
        .Height = picSplit.Height
        .Width = picSplit.Left
    End With
    With vfgItem
        .Left = picSplit.Left + picSplit.Width
        .Width = Me.ScaleWidth - .Left
        .Height = picSplit.Height - vfgBound.Height - cmdClear.Height - 200
    End With
    
    With cmdClear
        .Top = vfgItem.Height + 100
        .Left = Me.ScaleWidth - .Width
    End With
    
    With cmdBound
        .Top = cmdClear.Top
        .Left = vfgItem.Left
    End With
    
    With vfgBound
        .Top = cmdClear.Top + cmdClear.Height + 100
        .Left = vfgItem.Left
        .Width = vfgItem.Width
    End With
    cmdCancel.Left = picBottom.Width - cmdCancel.Width - 150
    cmdOK.Top = cmdCancel.Top
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 150
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsItem = Nothing
    Set mrsFind = Nothing
End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnDown = (Button = 1)
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mblnDown Then Exit Sub
    Dim blnAdjust As Boolean
    
    blnAdjust = True
    If picSplit.Left + X < 3000 Then picSplit.Left = 3000: blnAdjust = False
    If picSplit.Left + X > Me.ScaleWidth - 2000 Then picSplit.Left = Me.ScaleWidth - 2000: blnAdjust = False
    If blnAdjust Then
        picSplit.Left = picSplit.Left + X
        Call Form_Resize
    End If
End Sub

Private Sub picSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnDown = False
    Call Form_Resize
End Sub

Private Sub tvw_s_NodeCheck(ByVal node As MSComctlLib.node)
    '�Զ���ѡ�¼����
    Call NodeCheck(node, node.Checked, True)
    Call NodeSelAll(node)
End Sub

Private Sub tvw_s_NodeClick(ByVal node As MSComctlLib.node)
    If node.Key = mstrPreNode Then Exit Sub
    '���ı�ʱ,���浱ǰ˳��(������)
    mstrPreNode = node.Key
    
    Call FillList
End Sub

Private Sub NodeCheck(ByVal node As MSComctlLib.node, ByVal blnSel As Boolean, Optional ByVal blnParent As Boolean = False)
    '�ݹ����,ѭ��ѡ�������ӽ��
    node.Checked = blnSel
    If node.Children > 0 Then Call NodeCheck(node.Child, blnSel)
    If blnParent Then Exit Sub
    If Not node.Next Is Nothing Then Call NodeCheck(node.Next, blnSel)
End Sub

Private Function NodeSelAll(ByVal node As MSComctlLib.node) As Boolean
    '���ͬ��(ֻҪѡ����һ���ӽ��,����㶼Ӧ�ù�ѡ;һ���ӽ�㶼ûѡ,����㲻��Ҫ��ѡ)
    Dim intCount As Integer
    Dim nodSource As MSComctlLib.node
    
    Set nodSource = node
    If Not node.Parent Is Nothing Then Set node = node.Parent.Child     '�����ǰ���Ǹ���㣬����Ϊ��һ���ӽ��
    If node.Checked Then intCount = 1
    Do While True
        If Not node.Next Is Nothing Then
            If node.Next.Checked Then intCount = intCount + 1
            If intCount > 0 Then Exit Do
            Set node = node.Next
        Else
            Exit Do
        End If
    Loop
    
    '���ϻ���
    Set node = nodSource
    Do While True
        If Not node.Parent Is Nothing Then
            node.Parent.Checked = intCount
            Set node = node.Parent
        Else
            Exit Do
        End If
    Loop
End Function

Private Function GetSelNodes(ByVal node As MSComctlLib.node) As String
    Dim strSel As String
    Dim strReturn As String
    
    '��ȡ����ѡ�����ĩ�����
    If node.Checked Then
        If node.Children > 0 Then
            strSel = GetSelNodes(node.Child)
            If strSel <> "" Then strReturn = strReturn & IIf(strReturn <> "", ",", "") & strSel
        Else
            strReturn = strReturn & IIf(strReturn <> "", ",", "") & Mid(node.Key, 2)
        End If
    End If
    If Not node.Next Is Nothing Then
        strSel = GetSelNodes(node.Next)
        If strSel <> "" Then strReturn = strReturn & IIf(strReturn <> "", ",", "") & strSel
    End If
    GetSelNodes = strReturn
End Function

Private Sub txtFind_Change()
    If Trim(txtFind.Text) = "" Then
        lblInfo.Caption = "�������������"
        lblInfo.ForeColor = &H8000&
    Else
        lblInfo.Caption = "�����λ��ɴʾ����"
        lblInfo.ForeColor = &H8000000D
    End If
    
    cmdFind.Caption = "��λ(&L)"
    Set mrsFind = New ADODB.Recordset
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmdFind.SetFocus
        Call cmdFind_Click
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub vfgBound_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
    With vfgBound
        If KeyCode = vbKeyDelete And .Row >= .FixedRows Then
            If .RowData(.Row) > 0 Then
                If .Row = .FixedRows And .Row = .Rows - 1 Then
                    .Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
                    .RowData(.Row) = 0
                Else
                    If .Row < .Rows - 1 Then
                        .RowPosition(.Row) = .Rows - 1
                        For lngRow = .Row To .Rows - 1
                            .TextMatrix(lngRow, 0) = lngRow + .FixedRows - 1
                        Next
                    End If
                    .RemoveItem .Rows - 1
                End If
                vfgBound.ColWidth(0) = Me.TextWidth(vfgBound.TextMatrix(vfgBound.Rows - 1, 0) & " ")
                If vfgBound.ColWidth(0) < 380 Then vfgBound.ColWidth(0) = 380
            End If
        End If
    End With
End Sub

Private Sub vfgItem_DblClick()
    Dim i As Long
    Dim blnAdd As Boolean
    With vfgItem
        blnAdd = True
        If .Row >= .FixedRows And .Cols >= .FixedCols Then
            If Val(.TextMatrix(.Row, .FixedCols)) <= 0 Then Exit Sub
            For i = vfgBound.FixedRows To vfgBound.Rows - 1
                If Val(.TextMatrix(.Row, 3)) = Val(vfgBound.TextMatrix(i, 1)) Then
                    blnAdd = False
                    Exit For
                End If
            Next i
            
            If blnAdd = True Then
                If Val(vfgBound.TextMatrix(vfgBound.Rows - 1, 1)) > 0 Or vfgBound.Rows = vfgBound.FixedRows Then
                    vfgBound.Rows = vfgBound.Rows + 1
                End If
                vfgBound.TextMatrix(vfgBound.Rows - 1, 0) = vfgBound.Rows - vfgBound.FixedRows
                vfgBound.TextMatrix(vfgBound.Rows - 1, 1) = .TextMatrix(.Row, 3)
                vfgBound.TextMatrix(vfgBound.Rows - 1, 2) = .TextMatrix(.Row, 5)
                vfgBound.TextMatrix(vfgBound.Rows - 1, 3) = .TextMatrix(.Row, 7)
                vfgBound.TextMatrix(vfgBound.Rows - 1, 4) = .TextMatrix(.Row, 8)
                vfgBound.RowData(vfgBound.Rows - 1) = Val(.TextMatrix(.Row, 3))
                vfgBound.Row = vfgBound.Rows - 1
                vfgBound.TopRow = vfgBound.Rows - 1
                vfgBound.ColWidth(0) = Me.TextWidth(vfgBound.TextMatrix(vfgBound.Rows - 1, 0) & " ")
                If vfgBound.ColWidth(0) < 380 Then vfgBound.ColWidth(0) = 380
            End If
        End If
    End With
End Sub
