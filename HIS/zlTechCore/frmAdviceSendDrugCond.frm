VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAdviceSendDrugCond 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "frmAdviceSendDrugCond.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab tabCond 
      Height          =   5535
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      TabHeight       =   564
      WordWrap        =   0   'False
      TabCaption(0)   =   "��������(&1)"
      TabPicture(0)   =   "frmAdviceSendDrugCond.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "imgLogo(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblTip(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDetail(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "��ҩ;��(&2)"
      TabPicture(1)   =   "frmAdviceSendDrugCond.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDetail(1)"
      Tab(1).Control(1)=   "lblTip(1)"
      Tab(1).Control(2)=   "imgLogo(1)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "ҩ���û�(&3)"
      TabPicture(2)   =   "frmAdviceSendDrugCond.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraDetail(2)"
      Tab(2).Control(1)=   "lblTip(2)"
      Tab(2).Control(2)=   "imgLogo(2)"
      Tab(2).ControlCount=   3
      Begin VB.Frame fraDetail 
         Height          =   4425
         Index           =   2
         Left            =   -74835
         TabIndex        =   28
         Top             =   975
         Width           =   5400
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   4080
            Left            =   1170
            TabIndex        =   20
            Top             =   210
            Width           =   4095
            _cx             =   7223
            _cy             =   7197
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   6
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmAdviceSendDrugCond.frx":05DE
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
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺҩ��(&G)"
            Height          =   180
            Left            =   135
            TabIndex        =   19
            Top             =   300
            Width           =   990
         End
      End
      Begin VB.Frame fraDetail 
         Height          =   4410
         Index           =   1
         Left            =   -74835
         TabIndex        =   26
         Top             =   975
         Width           =   5400
         Begin VB.CommandButton cmdAllWay 
            Caption         =   "ȫѡ"
            Height          =   330
            Left            =   180
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl + A"
            Top             =   3525
            Width           =   870
         End
         Begin VB.CommandButton cmdNoWay 
            Caption         =   "ȫ��"
            Height          =   330
            Left            =   180
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl + R"
            Top             =   3900
            Width           =   870
         End
         Begin MSComctlLib.ListView lvwWay 
            Height          =   4050
            Left            =   1170
            TabIndex        =   16
            Top             =   210
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   7144
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "��ҩ;��"
               Object.Width           =   6526
            EndProperty
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ҩ;��(&W)"
            Height          =   180
            Left            =   135
            TabIndex        =   15
            Top             =   300
            Width           =   990
         End
      End
      Begin VB.Frame fraDetail 
         Height          =   4425
         Index           =   0
         Left            =   165
         TabIndex        =   24
         Top             =   975
         Width           =   5400
         Begin VB.CheckBox chkLimit 
            Caption         =   "��ҩ;��ִ���Է��͵Ľ���ʱ��Ϊ׼����"
            Height          =   360
            Left            =   3300
            TabIndex        =   5
            Top             =   510
            Width           =   1920
         End
         Begin VB.CheckBox chk�Ӱ�Ӽ� 
            Caption         =   "ִ�мӰ�Ӽ�(&V)"
            Height          =   195
            Left            =   3615
            TabIndex        =   6
            Top             =   585
            Visible         =   0   'False
            Width           =   1650
         End
         Begin VB.ComboBox cboUnit 
            Height          =   300
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1260
            Width           =   4095
         End
         Begin VB.CommandButton cmdNoPati 
            Caption         =   "ȫ��"
            Height          =   330
            Left            =   180
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl + R"
            Top             =   3570
            Width           =   870
         End
         Begin VB.CommandButton cmdAllPati 
            Caption         =   "ȫѡ"
            Height          =   330
            Left            =   180
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl + A"
            Top             =   3195
            Width           =   870
         End
         Begin VB.OptionButton opt��Ч 
            Caption         =   "����(&T)"
            Height          =   180
            Index           =   1
            Left            =   2145
            TabIndex        =   2
            Top             =   255
            Width           =   930
         End
         Begin VB.OptionButton opt��Ч 
            Caption         =   "����(&L)"
            Height          =   180
            Index           =   0
            Left            =   1170
            TabIndex        =   1
            Top             =   255
            Value           =   -1  'True
            Width           =   930
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   1170
            TabIndex        =   4
            Top             =   540
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            CalendarTitleBackColor=   8388608
            CalendarTitleForeColor=   16777215
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   64290819
            CurrentDate     =   37953
         End
         Begin MSComctlLib.ListView lvwPati 
            Height          =   2310
            Left            =   1170
            TabIndex        =   12
            Top             =   1620
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   4075
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "����"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "סԺ��"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "����"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "ʣ���"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "סԺҽʦ"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "�ѱ�"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "����ȼ�"
               Object.Width           =   2028
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "����"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "��Ժ����"
               Object.Width           =   2857
            EndProperty
         End
         Begin VB.ComboBox cboҩ�� 
            Height          =   300
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   900
            Width           =   4110
         End
         Begin MSComctlLib.Toolbar tbrAutoSel 
            Height          =   360
            Left            =   1170
            TabIndex        =   30
            Top             =   3990
            Width           =   3045
            _ExtentX        =   5371
            _ExtentY        =   635
            ButtonWidth     =   5318
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Style           =   1
            TextAlignment   =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "���������������ſ�Ƿ�Ѳ���   "
                  Object.ToolTipText     =   "Ctrl + Q"
               EndProperty
            EndProperty
            BorderStyle     =   1
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ����(&P)"
            Height          =   180
            Left            =   135
            TabIndex        =   11
            Top             =   1695
            Width           =   990
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��(&E)"
            Height          =   180
            Left            =   135
            TabIndex        =   3
            Top             =   600
            Width           =   990
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ����(&U)"
            Height          =   180
            Left            =   135
            TabIndex        =   9
            Top             =   1320
            Width           =   990
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ҩҩ��(&R)"
            Height          =   180
            Left            =   135
            TabIndex        =   7
            Top             =   960
            Width           =   990
         End
      End
      Begin VB.Label lblTip 
         Caption         =   "����ʵ���������ҽ��ԭ����ִ��ҩ��ָ��Ϊ�µ�ҩ����ҽ������ʱ���ᷢ�͵��µ�ҩ��ִ�С�"
         Height          =   375
         Index           =   2
         Left            =   -73785
         TabIndex        =   29
         Top             =   585
         Width           =   4170
      End
      Begin VB.Image imgLogo 
         Height          =   480
         Index           =   2
         Left            =   -74535
         Picture         =   "frmAdviceSendDrugCond.frx":0633
         Top             =   480
         Width           =   480
      End
      Begin VB.Label lblTip 
         Caption         =   "����ͨ��ѡ��ͬ��ҩƷ��ҩ;��������������Ҫ������ЩҩƷ��"
         Height          =   375
         Index           =   1
         Left            =   -73785
         TabIndex        =   27
         Top             =   585
         Width           =   4170
      End
      Begin VB.Image imgLogo 
         Height          =   480
         Index           =   1
         Left            =   -74535
         Picture         =   "frmAdviceSendDrugCond.frx":0EFD
         Top             =   480
         Width           =   480
      End
      Begin VB.Label lblTip 
         Caption         =   "����ҩƷҽ�����͵���Ҫ������Ҫ���͵�ʱ�䣬ҽ�����ͣ��Լ�Ҫ����ҽ���ľ��岡�ˡ�"
         Height          =   375
         Index           =   0
         Left            =   1215
         TabIndex        =   25
         Top             =   585
         Width           =   4170
      End
      Begin VB.Image imgLogo 
         Height          =   480
         Index           =   0
         Left            =   465
         Picture         =   "frmAdviceSendDrugCond.frx":17C7
         Top             =   480
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   555
      TabIndex        =   23
      Top             =   5715
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4335
      TabIndex        =   22
      Top             =   5715
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3105
      TabIndex        =   21
      Top             =   5715
      Width           =   1100
   End
End
Attribute VB_Name = "frmAdviceSendDrugCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mstrPrivs As String 'IN
Public mlng����ID As Long 'IN/OUT
Public mlng����ID As Long 'IN
Public mblnOK As Boolean 'OUT:�Ƿ�ȷ��
Public mstrEnd As String 'OUT:����ʱ��
Public mint��Ч As Integer 'OUT:0-����,1-����
Public mstr����IDs As String 'OUT:����ID��
Public mstr��ҩIDs As String 'OUT:��ҩ;��ID��
Public mblnLimit As Boolean 'OUT:��ҩ;�������ͽ���ʱ�����Ƽ���
Public mlngҩ��ID As Long 'OUT:ָ����ҩ��
Public mrsҩ�� As ADODB.Recordset 'IN/OUT:ҩƷ�滻��(�ɸ���)

Private mrsWarn As ADODB.Recordset

Private Sub cboUnit_Click()
'���ܣ���ȡָ����Χ�ڵĲ����б�
    Dim rsTmp As New ADODB.Recordset
    Dim objItem As ListItem, strSQL As String
    Dim i As Integer, j As Integer, k As Integer
    Dim str����IDs As String, lng����ID As Long
        
    lvwPati.ListItems.Clear
    
    On Error GoTo errH
    
    strSQL = "Select ���ò���,��������,����ֵ From ���ʱ����� Where ����ID=[1]"
    Set mrsWarn = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboUnit.ItemData(cboUnit.ListIndex))
    
    str����IDs = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ҩ�����Ͳ���", "")
    If str����IDs <> "" And InStr(str����IDs, ":") > 0 Then
        lng����ID = Val(Split(str����IDs, ":")(0))
        str����IDs = Split(str����IDs, ":")(1)
    End If
        
    '��Ժ����:��Ժ���˽�ֹ��ҽ��,����ҽ��
    strSQL = _
        "Select A.����ID,A.����,A.סԺ��,B.��Ժ���� as ����," & _
        " Nvl(E.Ԥ�����,0)-Nvl(E.�������,0)+Decode(B.����,Null,0,Nvl(F.���,0)) as ʣ���," & _
        " A.������,Decode(X.����,'1',1,Decode(B.����,Null,0,1)) as ҽ��,B.����," & _
        " B.סԺҽʦ,B.�ѱ�,D.���� as ����ȼ�,C.���� as ����,B.��Ժ����" & _
        " From ������Ϣ A,������ҳ B,���ű� C,�շ���ĿĿ¼ D,������� E,ҽ�Ƹ��ʽ X," & _
        " (Select ����ID,��ҳID,Sum(���) As ��� From ����ģ����� Group By ����ID,��ҳID) F" & _
        " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And B.��Ժ����ID=C.ID" & _
        " And A.����ID=E.����ID(+) And E.����(+)=1 And B.����ID=F.����ID(+) And B.��ҳID=F.��ҳID(+)" & _
        " And B.��Ժ���� is NULL and Nvl(B.״̬,0)<>3 And B.����ȼ�ID=D.ID(+) And B.ҽ�Ƹ��ʽ=X.����(+)" & _
        IIF(cboUnit.ItemData(cboUnit.ListIndex) > 0, " And B.��ǰ����ID=[1]", "") & _
        IIF(cboUnit.ItemData(cboUnit.ListIndex) = 0, " Order by A.סԺ�� Desc", " Order by LPAD(B.��Ժ����,10,' ')")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboUnit.ItemData(cboUnit.ListIndex))
  
    For i = 1 To rsTmp.RecordCount
        Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!����ID, rsTmp!����)
        objItem.SubItems(1) = IIF(IsNull(rsTmp!סԺ��), "", rsTmp!סԺ��)
        objItem.SubItems(2) = IIF(IsNull(rsTmp!����), "", rsTmp!����)
        objItem.SubItems(3) = Format(Nvl(rsTmp!ʣ���, 0), "0.00")
        objItem.SubItems(4) = IIF(IsNull(rsTmp!סԺҽʦ), "", rsTmp!סԺҽʦ)
        objItem.SubItems(5) = IIF(IsNull(rsTmp!�ѱ�), "", rsTmp!�ѱ�)
        objItem.SubItems(6) = IIF(IsNull(rsTmp!����ȼ�), "", rsTmp!����ȼ�)
        objItem.SubItems(7) = IIF(IsNull(rsTmp!����), "", rsTmp!����)
        objItem.SubItems(8) = Format(rsTmp!��Ժ����, "yyyy-MM-dd HH:mm")
        
        '������Ϣ
        objItem.ListSubItems(1).Tag = Nvl(rsTmp!ҽ��, 0)
        objItem.ListSubItems(2).Tag = Nvl(rsTmp!������, 0)
        
        '���ղ����ú�ɫ��ʾ
        If Not IsNull(rsTmp!����) Then
            objItem.ForeColor = vbRed
            For j = 1 To objItem.ListSubItems.Count
                objItem.ListSubItems(j).ForeColor = vbRed
            Next
        End If
        
        '�ϴ��Ƿ�ѡ��
        If cboUnit.ItemData(cboUnit.ListIndex) = lng����ID And str����IDs <> "" Then
            If InStr("," & str����IDs & ",", "," & rsTmp!����ID & ",") > 0 Then
                objItem.Checked = True
                If k = 0 Then 'Ϊ�˿�����ѡ���
                    objItem.EnsureVisible
                    objItem.Selected = True
                    k = 1
                End If
            End If
        ElseIf rsTmp!����ID = mlng����ID Then
            objItem.Checked = True 'ȱʡֻѡ��ǰ����
            objItem.EnsureVisible
            objItem.Selected = True
        End If
        rsTmp.MoveNext
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboUnit_GotFocus()
    tabCond.Tab = 0
End Sub

Private Sub chkLimit_Click()
    chkLimit.ForeColor = IIF(chkLimit.Value = 1, &HC0&, Me.ForeColor)
End Sub

Private Sub cmdAllPati_Click()
    Call SelectLVW(lvwPati, True)
    lvwPati.SetFocus
End Sub

Private Sub SelectLVW(objLVW As Object, ByVal blnCheck As Boolean)
    Dim i As Long
    For i = 1 To objLVW.ListItems.Count
        objLVW.ListItems(i).Checked = blnCheck
    Next
End Sub

Private Sub cmdAllPati_GotFocus()
    tabCond.Tab = 0
End Sub

Private Sub cmdAllWay_Click()
    Call SelectLVW(lvwWay, True)
    lvwWay.SetFocus
End Sub

Private Sub cmdAllWay_GotFocus()
    tabCond.Tab = 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNoPati_Click()
    Call SelectLVW(lvwPati, False)
    lvwPati.SetFocus
End Sub

Private Sub cmdNoPati_GotFocus()
    tabCond.Tab = 0
End Sub

Private Sub cmdNoWay_Click()
    Call SelectLVW(lvwWay, False)
    lvwWay.SetFocus
End Sub

Private Sub cmdNoWay_GotFocus()
    tabCond.Tab = 1
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    
    If cboUnit.ListIndex = -1 Then
        MsgBox "��ѡ��һ��������", vbInformation, gstrSysName
        cboUnit.SetFocus: Exit Sub
    End If
    mlng����ID = cboUnit.ItemData(cboUnit.ListIndex)
    
    'ʱ�����Ч
    mint��Ч = IIF(opt��Ч(1).Value, 1, 0)
    If opt��Ч(0).Value Then
        mstrEnd = Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:ss")
    Else
        mstrEnd = ""
    End If
    
    '��ҩ;�����㷽ʽ
    mblnLimit = chkLimit.Value = 1
    
    '��ҩҩ��
    mlngҩ��ID = cboҩ��.ItemData(cboҩ��.ListIndex)
    
    'סԺ����
    mstr����IDs = ""
    For i = 1 To lvwPati.ListItems.Count
        If lvwPati.ListItems(i).Checked Then
            mstr����IDs = mstr����IDs & "," & Mid(lvwPati.ListItems(i).Key, 2)
        End If
    Next
    mstr����IDs = Mid(mstr����IDs, 2)
    If mstr����IDs = "" Then
        MsgBox "������ѡ��һ����Ҫ����ҽ�����ˡ�", vbInformation, gstrSysName
        tabCond.Tab = 0: lvwPati.SetFocus: Exit Sub
    End If
        
    '��ҩ;��
    mstr��ҩIDs = ""
    For i = 1 To lvwWay.ListItems.Count
        If lvwWay.ListItems(i).Checked Then
            mstr��ҩIDs = mstr��ҩIDs & "," & Mid(lvwWay.ListItems(i).Key, 2)
        End If
    Next
    mstr��ҩIDs = Mid(mstr��ҩIDs, 2)
    If mstr��ҩIDs = "" Then
        MsgBox "������ѡ��һ�ָ�ҩ;����", vbInformation, gstrSysName
        tabCond.Tab = 1: lvwWay.SetFocus: Exit Sub
    End If
    If UBound(Split(mstr��ҩIDs, ",")) + 1 = lvwWay.ListItems.Count Then
        mstr��ҩIDs = ""
    End If
    
    gbln�Ӱ�Ӽ� = chk�Ӱ�Ӽ�.Value = 1
    
    mblnOK = True
    Unload Me
End Sub

Private Sub dtpEnd_GotFocus()
    tabCond.Tab = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If tabCond.Tab = 0 Then
            Call cmdAllPati_Click
        Else
            Call cmdAllWay_Click
        End If
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If tabCond.Tab = 0 Then
            Call cmdNoPati_Click
        Else
            Call cmdNoWay_Click
        End If
    ElseIf KeyCode = 13 Then
        If Not ActiveControl Is vsDept _
            And Not ActiveControl Is tabCond Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf KeyCode = vbKeyQ And Shift = vbCtrlMask Then
        If tbrAutoSel.Visible Then
            Call tbrAutoSel_ButtonClick(tbrAutoSel.Buttons(1))
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not ActiveControl Is vsDept _
            And Not ActiveControl Is tabCond Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim curDate As Date
    Dim strTmp As String, lngTmp As Long
    
    Call RestoreListViewState(Me.lvwPati, App.ProductName & Me.Name, "")
    
    mblnOK = False
    
    '�Է��ͽ���ʱ��Ϊ׼
    chkLimit.Value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "���ƽ���ʱ��", 0))
    
    'ȱʡҽ����Ч
    lngTmp = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ҩ��ҽ����Ч", 0))
    opt��Ч(lngTmp).Value = True
    
    '������һ���ſ��ܽ���
    If InStr(mstrPrivs, "����ҩ������") = 0 Then
        opt��Ч(0).Value = True
        opt��Ч(1).Enabled = False
    ElseIf InStr(mstrPrivs, "����ҩ�Ƴ���") = 0 Then
        opt��Ч(1).Value = True
        opt��Ч(0).Enabled = False
    End If
   
    'ȱʡ����ʱ��
    curDate = zlDatabase.Currentdate
    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ҩ������ʱ��", "23:59:59")
    lngTmp = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ҩ��ʱ����", 0))
    dtpEnd.Value = Format(curDate + lngTmp, "yyyy-MM-dd " & strTmp)
    
    '����/����
    Call zlControl.LvwFlatColumnHeader(lvwPati)
    Call InitUnits
                        
    '��ҩҩ��
    Call Loadҩ��
    
    '��ҩ;��
    Call Load��ҩ;��
    
    'ҩ���û�
    Call Showҩ��
End Sub

Private Function Loadҩ��() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    cboҩ��.AddItem "����ҩ��"
    cboҩ��.ListIndex = 0
    
    On Error GoTo errH
    
    strSQL = _
        "Select Distinct A.ID,A.����,A.����" & _
        " From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " AND B.����ID=A.ID And B.������� IN(2,3) and B.�������� in('��ҩ��','��ҩ��','��ҩ��')" & _
        " Order by A.����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        cboҩ��.AddItem rsTmp!���� & "-" & rsTmp!����
        cboҩ��.ItemData(cboҩ��.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop
    Loadҩ�� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Showҩ��()
    Dim strTmp As String, i As Long, j As Long
    Dim str�û� As String, arr�û� As Variant
    
    mrsҩ��.Filter = 0
    If Not mrsҩ��.EOF Then
        vsDept.Rows = vsDept.FixedRows + mrsҩ��.RecordCount
        For i = 1 To mrsҩ��.RecordCount
            vsDept.Cell(flexcpData, i, 0) = CLng(mrsҩ��!ID)
            vsDept.TextMatrix(i, 0) = mrsҩ��!���� & "-" & mrsҩ��!����
            strTmp = strTmp & "|#" & mrsҩ��!ID & ";" & mrsҩ��!���� & "-" & mrsҩ��!����
            mrsҩ��.MoveNext
        Next
        
        str�û� = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ҩ��ҩ���û�", "")
        arr�û� = Split(str�û�, ",")
        For i = 1 To vsDept.Rows - 1
            mrsҩ��.Filter = "ID=" & CLng(vsDept.Cell(flexcpData, i, 0))
            For j = 0 To UBound(arr�û�)
                If arr�û�(j) Like mrsҩ��!ID & "-*" Then Exit For
            Next
            If j <= UBound(arr�û�) Then
                mrsҩ��.Filter = "ID=" & Val(Split(arr�û�(j), "-")(1))
                If Not mrsҩ��.EOF Then
                    vsDept.Cell(flexcpData, i, 1) = CLng(mrsҩ��!ID)
                    mrsҩ��.Filter = "ID=" & CLng(vsDept.Cell(flexcpData, i, 0))
                    mrsҩ��!��ID = CLng(vsDept.Cell(flexcpData, i, 1))
                    mrsҩ��.Update
                Else
                    vsDept.Cell(flexcpData, i, 1) = CLng(mrsҩ��!��ID)
                End If
            Else
                vsDept.Cell(flexcpData, i, 1) = CLng(mrsҩ��!��ID)
            End If
            
            mrsҩ��.Filter = "ID=" & CLng(vsDept.Cell(flexcpData, i, 1))
            vsDept.TextMatrix(i, 1) = mrsҩ��!���� & "-" & mrsҩ��!����
        Next
        If strTmp <> "" Then vsDept.ColComboList(1) = Mid(strTmp, 2)
    Else
        vsDept.Rows = vsDept.FixedRows + 1
        vsDept.Editable = flexEDNone
    End If
    vsDept.Row = vsDept.FixedRows: vsDept.Col = 1
    Call vsDept_AfterRowColChange(-1, -1, vsDept.Row, vsDept.Col)
End Sub

Private Function Load��ҩ;��() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objItem As ListItem, str��ҩIDs As String
    
    On Error GoTo errH
    
    str��ҩIDs = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ҩ����ҩ;��", "")
    
    strSQL = "Select ID,����,���� From ������ĿĿ¼ Where ���='E' And ��������='2' Order by ����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        Set objItem = lvwWay.ListItems.Add(, "_" & rsTmp!ID, rsTmp!���� & "-" & rsTmp!����)
        
        If str��ҩIDs <> "" Then
            If InStr("," & str��ҩIDs & ",", "," & rsTmp!ID & ",") > 0 Then
                objItem.Checked = True
            End If
        Else
            objItem.Checked = True
        End If
        rsTmp.MoveNext
    Next
    Load��ҩ;�� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function InitUnits() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, strSQL As String
    
    On Error GoTo errH
    
    '��������۲���
    If InStr(mstrPrivs, "ȫԺ����") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where A.ID=B.����ID And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Order by A.����"
    Else
        '����Ȩ������ֱ�����ڲ���+���ڿ�����������
        strSQL = _
            " Select A.ID,A.����,A.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
            " From ���ű� A,��������˵�� B,������Ա C" & _
            " Where A.ID=B.����ID And A.ID=C.����ID And C.��ԱID=[1]" & _
            " And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        If Not gbln�������Ҷ��� Then
            strSQL = strSQL & IIF(strSQL <> "", " Union ", "") & _
                " Select C.ID,C.����,C.����,Nvl(B.ȱʡ,0) as ȱʡ" & _
                " From ��λ״����¼ A,������Ա B,���ű� C" & _
                " Where A.����ID=C.ID And B.����ID=A.����ID And B.��ԱID=[1]" & _
                " And (C.����ʱ�� is NULL or Trunc(C.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        End If
        strSQL = "Select ID,����,����,Max(ȱʡ) as ȱʡ From (" & strSQL & ") Group by ID,����,���� Order by ����"
    End If
    
    cboUnit.Clear
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!���� & "-" & rsTmp!����
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If rsTmp!ID = mlng����ID Then cboUnit.ListIndex = cboUnit.NewIndex
            rsTmp.MoveNext
        Next
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long, strTmp As String
    
    '������������
    If mblnOK Then
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "���ƽ���ʱ��", chkLimit.Value
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ҩ������ʱ��", Format(dtpEnd.Value, "HH:mm:ss")
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ҩ��ʱ����", Int(CDate(Format(dtpEnd.Value, "yyyy-MM-dd")) - CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd")))
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ҩ��ҽ����Ч", IIF(opt��Ч(1).Value, 1, 0)
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ҩ����ҩ;��", mstr��ҩIDs
        
        '���ˣ�ѡ���˽�Ϊ��ǰ����ʱ,������
        If UBound(Split(mstr����IDs, ",")) = 0 And Val(mstr����IDs) = mlng����ID Then
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ҩ�����Ͳ���", ""
        Else
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ҩ�����Ͳ���", cboUnit.ItemData(cboUnit.ListIndex) & ":" & mstr����IDs
        End If
        
        'ҩ���û�
        mrsҩ��.Filter = 0
        For i = 1 To mrsҩ��.RecordCount
            strTmp = strTmp & "," & mrsҩ��!ID & "-" & mrsҩ��!��ID
            mrsҩ��.MoveNext
        Next
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ҩ��ҩ���û�", Mid(strTmp, 2)
    End If
    
    '�ͷ�˽�м�IN����
    mstrPrivs = ""
    mlng����ID = 0
    Set mrsWarn = Nothing
    
    Call SaveListViewState(Me.lvwPati, App.ProductName & Me.Name, "")
End Sub

Private Sub lvwPati_GotFocus()
    tabCond.Tab = 0
End Sub

Private Sub lvwWay_GotFocus()
    tabCond.Tab = 1
End Sub

Private Sub opt��Ч_Click(Index As Integer)
    dtpEnd.Enabled = opt��Ч(0).Value
    chkLimit.Visible = opt��Ч(0).Value
End Sub

Private Sub opt��Ч_GotFocus(Index As Integer)
    tabCond.Tab = 0
End Sub

Private Sub tabCond_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If tabCond.Tab = 0 Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf tabCond.Tab = 1 Then
            lvwWay.SetFocus
        ElseIf tabCond.Tab = 2 Then
            vsDept.SetFocus
        End If
    End If
End Sub

Private Sub vsDept_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    vsDept.Cell(flexcpData, Row, Col) = CLng(vsDept.ComboData)
    mrsҩ��.Filter = "ID=" & CLng(vsDept.Cell(flexcpData, Row, 0))
    mrsҩ��!��ID = CLng(vsDept.Cell(flexcpData, Row, Col))
    mrsҩ��.Update
End Sub

Private Sub vsDept_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If vsDept.Editable <> flexEDNone And NewCol = 1 Then
        vsDept.FocusRect = flexFocusSolid
    Else
        vsDept.FocusRect = flexFocusLight
    End If
End Sub

Private Sub vsDept_GotFocus()
    tabCond.Tab = 2
End Sub

Private Sub vsDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If vsDept.Col = 1 Then
            If vsDept.Row + 1 <= vsDept.Rows - 1 Then
                vsDept.Row = vsDept.Row + 1
            Else
                Call zlCommFun.PressKey(vbKeyTab)
                vsDept.Row = vsDept.FixedRows + 1
            End If
        Else
            vsDept.Col = 1
        End If
        Call vsDept.ShowCell(vsDept.Row, vsDept.Col)
    End If
End Sub

Private Sub vsDept_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If vsDept.ComboIndex <> -1 Then
            Call vsDept_KeyPress(13)
        End If
    End If
End Sub

Private Sub vsDept_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub tbrAutoSel_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i As Long
    
    If mrsWarn Is Nothing Then Exit Sub
    
    With lvwPati
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked Then
                'ֻ�����ۼƱ����������д���
                mrsWarn.Filter = "��������=1 And ���ò���=" & Val(.ListItems(i).ListSubItems(1).Tag) + 1
                If Not mrsWarn.EOF Then
                    If Val(.ListItems(i).SubItems(3)) + Val(.ListItems(i).ListSubItems(2).Tag) < Nvl(mrsWarn!����ֵ, 0) Then
                        .ListItems(i).Checked = False
                    End If
                End If
            End If
        Next
    End With
End Sub
