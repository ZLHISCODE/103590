VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCheckClass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҩƷ����ѡ��"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frmCheckClass.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "ҩƷ����(&1)"
      TabPicture(0)   =   "frmCheckClass.frx":6852
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tvw����"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "�ⷿ��λ(&2)"
      TabPicture(1)   =   "frmCheckClass.frx":686E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "vsfStock"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "chk��λ"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.CheckBox chk��λ 
         Caption         =   "����ʾ��ǰ�ⷿ�ѷ���Ļ�λ"
         Height          =   255
         Left            =   -74880
         TabIndex        =   4
         Top             =   480
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin MSComctlLib.TreeView tvw���� 
         Height          =   5175
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   9128
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   476
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "img16"
         Appearance      =   1
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfStock 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   5
         Top             =   840
         Width           =   4455
         _cx             =   7858
         _cy             =   8070
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
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorSel    =   14737632
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   3
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmCheckClass.frx":688A
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
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
         ExplorerBar     =   5
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   1
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "�˳�(&E)"
      Height          =   350
      Left            =   4920
      TabIndex        =   1
      Top             =   960
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&Q)"
      Height          =   350
      Left            =   4920
      TabIndex        =   0
      Top             =   480
      Width           =   1100
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
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckClass.frx":68FA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCheckClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngDrugId As Long '�ⷿ��λ
Private mstr��λ As String '��λ
Private mstrSelRows As String
Private mstrҩƷid As String
Private mBlnClick As Boolean    '������¼�Ƿ�����ȷ����ť

Private Sub chk��λ_Click()
    Load�ⷿ��λ
End Sub

Private Sub cmdCancel_Click()
'    Call frmCheckCourseCard.getҩƷid("", 0)
    mBlnClick = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim str����ID As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim intItem As Integer
    
    On Error GoTo errHandl
    mBlnClick = True
    
    mstr��λ = ""
    mstrҩƷid = ""
    str����ID = ""
    
    'ȡ��λ
    With vsfStock
        For intItem = 1 To .rows - 1
            If .Cell(flexcpChecked, intItem, .ColIndex("ѡ��")) = flexChecked Then
                mstr��λ = .TextMatrix(intItem, .ColIndex("����")) & "," & mstr��λ
            End If
        Next
    End With
    
    'ȡ��ҩƷ���ࣨ��ѡ�����ʾ���з��ࣩ
    For intItem = 1 To tvw����.Nodes.count
        If tvw����.Nodes(intItem).Key = "Root" And tvw����.Nodes(intItem).Checked = True Then
            
            gstrSQL = "select ID from �շ���ĿĿ¼ where ��� in('5','6','7') and  (����ʱ�� IS NULL OR ����ʱ��= to_date('3000-01-01','YYYY-MM-DD'))"
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ѯ����ҩƷ")
            Do While Not rsTemp.EOF
                mstrҩƷid = rsTemp!id & "," & mstrҩƷid
                rsTemp.MoveNext
            Loop
            Exit For
        ElseIf tvw����.Nodes(intItem).Key <> "Root" And _
            tvw����.Nodes(intItem).Key <> "_�г�ҩ" And _
            tvw����.Nodes(intItem).Key <> "_�в�ҩ" And _
            tvw����.Nodes(intItem).Key <> "_����ҩ" And _
            tvw����.Nodes(intItem).Checked Then
            str����ID = str����ID & "," & Mid(tvw����.Nodes(intItem).Key, 2)
        End If
    Next

    If str����ID <> "" Then
        str����ID = Mid(str����ID, 2)
    
        gstrSQL = "Select ҩƷid" & _
                  "  From ҩƷ��� " & _
                   " Where ҩ��id In (Select ID" & _
                    "               From ������ĿĿ¼ A " & _
                    "               Where a.����id  in (select * from Table(Cast(f_Num2list([1]) As Zltools.t_Numlist))))"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, str����ID)
        If rsTemp Is Nothing Then
            Unload Me
            Exit Sub
        Else
            mstrҩƷid = ""
            For i = 1 To rsTemp.RecordCount
                mstrҩƷid = rsTemp!ҩƷID & "," & mstrҩƷid
                rsTemp.MoveNext
            Next
        End If
    End If
    
    Unload Me
    Exit Sub
errHandl:
If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    'ҩƷ����
    Dim objnode As Node
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select Level as ��,ID,�ϼ�ID,����,DECODE(����,1,'����ҩ',2,'�г�ҩ','�в�ҩ') As ���� " & _
        " From ���Ʒ���Ŀ¼" & _
        " Where ���� in (1,2,3)" & _
        " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
        " Order by Level,����"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡҩƷ����")

    Set objnode = tvw����.Nodes.Add(, , "Root", "������;", 1)
    Set objnode = tvw����.Nodes.Add("Root", 4, "_����ҩ", "����ҩ", 1)
    Set objnode = tvw����.Nodes.Add("Root", 4, "_�в�ҩ", "�в�ҩ", 1)
    Set objnode = tvw����.Nodes.Add("Root", 4, "_�г�ҩ", "�г�ҩ", 1)

    Do While Not rsTemp.EOF
        If rsTemp!�� = 1 Then
            Set objnode = tvw����.Nodes.Add("_" & rsTemp!����, 4, "_" & rsTemp!id, rsTemp!����, 1)
        Else
            Set objnode = tvw����.Nodes.Add("_" & rsTemp!�ϼ�ID, 4, "_" & rsTemp!id, rsTemp!����, 1)
        End If
        rsTemp.MoveNext
    Loop
    tvw����.Nodes("Root").Selected = True
    tvw����.Nodes("Root").Expanded = True
    
    Call Load�ⷿ��λ
End Sub

Private Sub Load�ⷿ��λ()
    Dim rsData As ADODB.Recordset
    Dim lngRow As Long
    Dim intLevel As Integer
    
    On Error GoTo errHandle
    If chk��λ.Value = 1 Then
        gstrSQL = "Select a.����, a.���� From ҩƷ�ⷿ��λ A " & _
            " Where a.�ⷿid = [1] And Exists (Select 1 From ҩƷ��λ���� B Where b.�ⷿid = a.�ⷿid And b.��λid = a.Id) " & _
            " Order By ���� "
    Else
        gstrSQL = "Select ����, ���� From  ҩƷ�ⷿ��λ Where �ⷿid = [1] Order By ���� "
    End If
    
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ����ҩƷ�ⷿ��λ", mlngDrugId)
    
    If rsData.RecordCount = 0 Then
        vsfStock.rows = 1
        Exit Sub
    End If
    
    With vsfStock
        .Redraw = flexRDNone
        .rows = 1
        
        Do While Not rsData.EOF
            .rows = .rows + 1
            
            .TextMatrix(.rows - 1, .ColIndex("ѡ��")) = 0
            .TextMatrix(.rows - 1, .ColIndex("����")) = rsData!����
            .TextMatrix(.rows - 1, .ColIndex("����")) = rsData!����
            
            rsData.MoveNext
        Loop
        
        .Redraw = flexRDDirect
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetSubItem(ByVal str�ϼ����� As String, ByVal rsData As ADODB.Recordset)
    '�õݹ��㷨�����������Ŀ
    Dim rsClone As ADODB.Recordset
    
    Set rsClone = rsData.Clone
    
    rsClone.Filter = "�ϼ�='" & str�ϼ����� & "'"
    rsClone.Sort = "����"
    
    'û�ҵ���һ��ʱһ��Ҫ�˳�
    If rsClone.RecordCount = 0 Then Exit Sub
    
    With vsfStock
        .Redraw = flexRDNone
        
        Do While Not rsClone.EOF
            .rows = .rows + 1
            
            .TextMatrix(.rows - 1, .ColIndex("����")) = rsClone!���� + 1
            .TextMatrix(.rows - 1, .ColIndex("�ϼ�")) = rsClone!�ϼ�
            .TextMatrix(.rows - 1, .ColIndex("����")) = rsClone!����
            .TextMatrix(.rows - 1, .ColIndex("ѡ��")) = 0
            .TextMatrix(.rows - 1, .ColIndex("����")) = rsClone!����
            
            '����һ������Ŀ
            Call GetSubItem(rsClone!����, rsData)
            
            rsClone.MoveNext
        Loop
        
        .Redraw = flexRDDirect
    End With
End Sub

Public Sub ShowME(ByVal frmCard As frmNewCheckCourseCard, ByVal lng�ⷿID As Long, ByRef str��λ As String, ByRef strҩƷID As String, ByRef int�˳� As Integer)
    mBlnClick = False
    
    mlngDrugId = lng�ⷿID
    Me.Show vbModal, frmCard
    
    str��λ = mstr��λ
    strҩƷID = mstrҩƷid
    
    If mBlnClick = True Then
        int�˳� = 1
    Else
        int�˳� = 0
    End If
End Sub

Private Sub tvw����_NodeCheck(ByVal Node As MSComctlLib.Node)
    CheckNode Node, Node.Checked
End Sub

Private Function CheckNode(ByVal Node As Object, blnCheck As Boolean)
    Dim intIdx As Integer

    If Node.Children > 0 Then
        Set Node = Node.Child
        Do While Not Node Is Nothing
            Node.Checked = blnCheck
            If Node.Children > 0 Then
                CheckNode Node, blnCheck
            End If
            Set Node = Node.Next
        Loop
    Else
        Node.Checked = blnCheck
    End If
End Function





Private Sub CheckItem(ByVal str���� As String, ByVal intChecked As Integer)
    Dim lngRows As Long
    
    With vsfStock
        If str���� = "����" Then
            '��ǰ�Ǹ��ڵ㣬���пⷿ����Ϊȫѡ��ȫ��ѡ
            For lngRows = 2 To .rows - 1
                .Cell(flexcpChecked, lngRows, .ColIndex("ѡ��")) = intChecked
            Next
        Else
            '��ǰ���ӽڵ㣬��Ҫ�ݹ鴦���¼��ӽڵ�
            For lngRows = 2 To .rows - 1
                If .TextMatrix(lngRows, .ColIndex("�ϼ�")) = str���� Then
                    .Cell(flexcpChecked, lngRows, .ColIndex("ѡ��")) = intChecked
                    
                    Call CheckItem(.TextMatrix(lngRows, .ColIndex("����")), intChecked)
                End If
            Next
        End If
    End With
End Sub


Private Sub vsfStock_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, Cancel As Boolean)
    Dim lngRow As Long
    Dim strFlag As String

    With vsfStock
        If .MouseRow <= 0 Then Exit Sub
        If .MouseCol <> .ColIndex("ѡ��") Then Exit Sub
        
        mstrSelRows = ""

        For lngRow = 1 To .rows - 1
            If .IsSelected(lngRow) Then
                mstrSelRows = IIf(mstrSelRows = "", "", mstrSelRows & ",") & lngRow
            End If
        Next
    End With
End Sub


Private Sub vsfStock_Click()
    Dim IntCheck As Integer
    Dim lngRow As Long
    Dim lngRows As Long
    
    With vsfStock
        If .MouseRow <= 0 Then Exit Sub
        If .MouseCol <> .ColIndex("ѡ��") Then Exit Sub
        
        lngRow = .MouseRow
        
        IntCheck = .Cell(flexcpChecked, lngRow, .ColIndex("ѡ��"))
        
'        Call CheckItem(.TextMatrix(lngRow, .ColIndex("����")), IntCheck)

        If InStr(1, mstrSelRows, ",") > 0 Then
            '��ѡ����
            For lngRows = 1 To .rows - 1
                If InStr(1, "," & mstrSelRows & ",", "," & lngRows & ",") > 0 Then
                    .Cell(flexcpChecked, lngRows, .ColIndex("ѡ��")) = IntCheck
                       
                End If
            Next
            
            mstrSelRows = ""
        Else
            '��ѡʱ����ǰ�ڵ�
            .Cell(flexcpChecked, lngRow, .ColIndex("ѡ��")) = IntCheck
        End If
        
    End With
End Sub
