VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDrugSelect 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "ҩƷѡ����"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11670
   Icon            =   "frmDrugSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   11670
   StartUpPosition =   1  '����������
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   5835
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   10292
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ils16"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.PictureBox picVsf 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   3600
      ScaleHeight     =   5895
      ScaleWidth      =   8055
      TabIndex        =   1
      Top             =   120
      Width           =   8055
      Begin VSFlex8Ctl.VSFlexGrid vsItem 
         Height          =   5835
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   7980
         _cx             =   14076
         _cy             =   10292
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
         FormatString    =   $"frmDrugSelect.frx":058A
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
   Begin MSComctlLib.ImageList ils16 
      Left            =   6120
      Top             =   3600
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
            Picture         =   "frmDrugSelect.frx":0617
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSelect.frx":0BB1
            Key             =   "Expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSelect.frx":114B
            Key             =   "��ҩ"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSelect.frx":16E5
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSelect.frx":1C7F
            Key             =   "��ҩ"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSelect.frx":2219
            Key             =   "����"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdEsc 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "�˳�"
      Height          =   180
      Left            =   0
      TabIndex        =   3
      Top             =   -900
      Width           =   90
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   0
      X2              =   10000
      Y1              =   5460
      Y2              =   5460
   End
End
Attribute VB_Name = "frmDrugSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrҩƷ���� As String, mlng��Ŀid   As Long
Private mbytOK As Byte
Private mrsItem As ADODB.Recordset
Private mstr����Tmp As String

Public Function ShowSelect(frmParent As Object, bytOK As Byte, Optional ByVal strҩƷ���� As String, Optional ByVal lng��Ŀid As Long) As ADODB.Recordset
'���ܣ���ʾҩƷѡ����
'������strҩƷ����=���ڶ�λ����
'      lng��Ŀid=���ڶ�λ��Ŀ


    mstrҩƷ���� = strҩƷ����
    mlng��Ŀid = lng��Ŀid
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    
    bytOK = mbytOK
    Set ShowSelect = IIF(bytOK = 1, mrsItem, Nothing)
End Function





Private Sub cmdEsc_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    InitDockPannel '��ʼ���϶��ؼ�
    Call FillTree
    
    mstr����Tmp = ""
    mbytOK = 0
End Sub



'���ò��ֿؼ�
Public Function DockPannelInit(ByRef dkpMain As DockingPane) As Boolean
On Error GoTo errH
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True
    DockPannelInit = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Function



'InitDockPannel��ʼ���򻮷�
Private Sub InitDockPannel()
    Dim objPane As Pane
On Error GoTo errH
    Set objPane = dkpMain.CreatePane(1, 200, 500, DockLeftOf, objPane)
    objPane.Title = "ҩƷ����Ŀ¼"
    objPane.Options = PaneNoCaption
    Set objPane = dkpMain.CreatePane(2, 550, 500, DockRightOf, objPane)
    objPane.Title = "����"
    objPane.Options = PaneNoCaption

    Call DockPannelInit(dkpMain)
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Sub


Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
On Error GoTo errH
    Select Case Item.ID
        Case 1
            Item.Handle = tvw_s.hwnd
        Case 2
            Item.Handle = picVsf.hwnd

    End Select
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Sub




Private Function FillTree() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objNode As Node
    
    On Error GoTo errH

    strSQL = _
        " Select 0 as ��,����,-���� as ID,-Null as �ϼ�ID,����||'' as ����," & _
        " ����||'.'||Decode(����,1,'����ҩ',2,'�г�ҩ',3,'�в�ҩ',4,'��ҩ�䷽') as ����" & _
        " From ���Ʒ���Ŀ¼ Where ���� in (1,2,3,4) And ����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Group by ����"
    strSQL = strSQL & " Union ALL " & _
        " Select Level as ��,����,ID,Nvl(�ϼ�ID,-����) as �ϼ�ID,����,���� From ���Ʒ���Ŀ¼" & _
        " Where ���� in (1,2,3,4) And ����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD')" & _
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



Private Sub Form_Resize()
    On Error Resume Next
    Call PicVsf_Resize
End Sub

Private Sub PicVsf_Resize()
    vsItem.Top = 10: vsItem.Left = 10: vsItem.Width = picVsf.Width - 10: vsItem.Height = picVsf.Height - 10
End Sub

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Key = mstr����Tmp Then Exit Sub
    Call FillList
    mstr����Tmp = Node.Key
End Sub


Private Function FillList() As Boolean
    Dim strSub As String, strSQL As String
    Dim objNode As Node, str��� As String
    Dim i As Long
    
    On Error GoTo errH
    Set objNode = tvw_s.SelectedItem '����ΪNothing
    If objNode Is Nothing Then Exit Function
    If Not mrsItem Is Nothing Then mrsItem.Filter = ""
    
    '��ʾ�¼�����Ŀ
    If Val(Mid(objNode.Key, 2)) < 0 Then
        strSub = " And A.����ID IN(" & _
            " Select ID From ���Ʒ���Ŀ¼ Where ����=[1] And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " )"
    Else
        strSub = " And A.����ID IN(" & _
            " Select ID From ���Ʒ���Ŀ¼ Where ����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD')" & _
            " Start With ID=[3] Connect by Prior ID=�ϼ�ID)"
    End If
    
    '�����е�����ȷ�����
    If Val(objNode.Tag) > 0 Then str��� = Choose(Val(objNode.Tag), "5", "6", "7", "8", "", "9", "4")
    If str��� <> "" Then strSub = strSub & " And A.���=[2]"
    
    strSQL = "Select a.Id As ������Ŀid, b.Id As �շ�ϸĿid,decode(a.���,'5','����ҩ','6','�г�ҩ','7','�в�ҩ','8','�䷽') as ���, a.����, b.���, a.���㵥λ, d.ҩƷ����,C.סԺ��λ as ������λ" & _
                " From ������ĿĿ¼ A, �շ���ĿĿ¼ B, ҩƷ��� C, ҩƷ���� D" & _
                " Where c.ҩƷid= b.Id(+)   And a.Id =c.ҩ��id (+) And c.ҩ��id = d.ҩ��id(+) and (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & strSub
    
    Set mrsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(objNode.Tag), str���, Val(Mid(objNode.Key, 2)))
    
    '������
    '--------------------------------------------------------------------------
    vsItem.Redraw = flexRDNone
    
    '����ͳ�Ƴ�����Ŀʱ����Ϊ��0��0��
    If vsItem.FixedRows = 0 Then
        vsItem.Rows = 2
        vsItem.FixedRows = 1
    End If
    If vsItem.FixedCols = 0 Then
        vsItem.Cols = 2
        vsItem.FixedCols = 1
    End If
    
    vsItem.ScrollBars = flexScrollBarNone
    Set vsItem.DataSource = mrsItem
    vsItem.ScrollBars = flexScrollBarBoth
    If err.Number = 0 And gcnOracle.Errors.Count > 0 Then
        gcnOracle.Errors.Clear
    End If
    If vsItem.Rows = vsItem.FixedRows Then
        vsItem.Rows = vsItem.FixedRows + 1
    End If
    
    '�����Ե���
    vsItem.ColAlignment(0) = 4
    vsItem.Cell(flexcpAlignment, 0, 0, 0, vsItem.Cols - 1) = 4
    vsItem.RowHeight(0) = vsItem.RowHeightMin
    For i = 1 To vsItem.Cols - 1
        vsItem.ColAlignment(i) = 1
        
        If vsItem.TextMatrix(0, i) Like "*ID" Then
            vsItem.ColHidden(i) = True
            vsItem.ColWidth(i) = 0
        ElseIf vsItem.ColWidth(i) > 2800 Then
            vsItem.ColWidth(i) = 2800
        ElseIf mrsItem.RecordCount = 0 Then
            vsItem.ColWidth(i) = 1000
        End If
        vsItem.ColData(i) = i '��¼ԭʼ�к�,���ڴ�����˳��
    Next
    vsItem.Redraw = flexRDBuffered
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Sub vsItem_DblClick()
    If vsItem.Row = -1 Then Exit Sub
    If vsItem.Row >= vsItem.FixedRows Then
        If mrsItem.RecordCount = 1 Then
            mbytOK = 1
        Else
            mbytOK = 0
        End If

        Unload Me
    End If
End Sub

Private Sub vsItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= vsItem.FixedRows Then
        mrsItem.Filter = IIF(vsItem.TextMatrix(NewRow, GetCol("���")) = "�䷽", "������Ŀid =" & Val(vsItem.TextMatrix(NewRow, GetCol("������Ŀid"))), "������Ŀid =" & Val(vsItem.TextMatrix(NewRow, GetCol("������Ŀid"))) & " And �շ�ϸĿid =" & Val(vsItem.TextMatrix(NewRow, GetCol("�շ�ϸĿid"))))
    End If
End Sub


Private Sub vsItem_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsItem.ColDataType(Col) = flexDTBoolean Then Cancel = True
End Sub

Private Function GetCol(ByVal strName As String) As Long
    Dim i As Long
    For i = 1 To vsItem.Cols - 1
        If UCase(vsItem.TextMatrix(0, i)) = UCase(strName) Then
            GetCol = i: Exit Function
        End If
    Next
End Function



