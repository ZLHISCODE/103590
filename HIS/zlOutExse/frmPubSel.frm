VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPubSel 
   AutoRedraw      =   -1  'True
   Caption         =   "ѡ����"
   ClientHeight    =   5784
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8676
   Icon            =   "frmPubSel.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5784
   ScaleWidth      =   8676
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList img 
      Left            =   120
      Top             =   1800
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubSel.frx":058A
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubSel.frx":6DEC
            Key             =   "Expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubSel.frx":D64E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubSel.frx":13EB0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   4560
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   3045
      _ExtentX        =   5376
      _ExtentY        =   8043
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img"
      Appearance      =   1
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3210
      Left            =   2145
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3216
      ScaleWidth      =   48
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   540
      Width           =   45
   End
   Begin VSFlex8Ctl.VSFlexGrid vsItem 
      Height          =   4560
      Left            =   3120
      TabIndex        =   7
      Top             =   600
      Width           =   5445
      _cx             =   9604
      _cy             =   8043
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.8
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
      RowHeightMin    =   280
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPubSel.frx":1A712
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
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      ScaleHeight     =   456
      ScaleWidth      =   8676
      TabIndex        =   5
      Top             =   0
      Width           =   8670
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѡ��һ����Ŀ,Ȼ����ȷ��"
         Height          =   180
         Left            =   180
         TabIndex        =   6
         Top             =   157
         Width           =   2430
      End
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   0
      ScaleHeight     =   516
      ScaleWidth      =   8676
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5268
      Width           =   8670
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   7335
         TabIndex        =   2
         Top             =   105
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   6210
         TabIndex        =   1
         Top             =   105
         Width           =   1100
      End
   End
End
Attribute VB_Name = "frmPubSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mfrmParent As Object
Private mstrKey As String
Private mstrPrivs As String
Private mint���� As Integer
Private mint������Դ As Integer
Private mstr���� As String
Private mint��ҩ��̬ As Integer
Private mrsItem As ADODB.Recordset
Private mblnOK As Boolean
Private mstrLike As String
Private mstrSaveTag As String
Private mlng����ID As Long
Private mlngҩ��id As Long
Private mlng��������ID As Long
Private mlngX As Long
Private mlngY As Long
Private mlngH As Long

Public Function ShowSelect(ByVal frmParent As Object, ByVal strPrivs As String, _
                                           ByVal int������Դ As Integer, ByVal int���� As Integer, _
                                           ByVal lng����ID As Long, ByVal lngҩ��ID As Long, _
                                           ByVal lng��������ID As Long, ByRef blnCancel As Boolean, _
                                           Optional ByVal str���� As String, _
                                           Optional ByVal int��ҩ��̬ As Integer = -1, _
                                           Optional ByVal lngX As Long, _
                                           Optional ByVal lngY As Long, _
                                           Optional ByVal lngH As Long) As ADODB.Recordset
'���ܣ��๦��ѡ����
'������
'���:int������Դ=ָ������Դ,1-����,2-סԺ
    '     blnҩ����λ=�Ƿ�ҩ����λ��ʾ���ͼ۸�
    '     str����=����ƥ�������,���û����Ϊѡ������ʽ,����Ϊ�б�ʽ
    '     int��ҩ��̬:-1��ʾ��������ҩ��̬,0-ֻ��ʾɢװ��̬����ҩ,1-ֻ��ʾ��Ƭ��̬����ҩ;2-ֻ��ʾ�����̬����ҩ
'����:blnCancel-�Ƿ�Ϊ�û�ȡ������
'���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��

    Set mfrmParent = frmParent
    
    mstrPrivs = strPrivs: mstr���� = str����
    mlngҩ��id = lngҩ��ID: mlng����ID = lng����ID
    mint������Դ = int������Դ: mint���� = int����:
    mlngX = lngX: mlngY = lngY: mlngH = lngH
    mlng��������ID = lng��������ID: mint��ҩ��̬ = int��ҩ��̬

    mstrSaveTag = IIf(mstr���� <> "", 1, 0)

    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    
    If mblnOK Then
        If mrsItem Is Nothing Then Exit Function
        Set ShowSelect = mrsItem
        If ShowSelect.RecordCount = 0 Then Set ShowSelect = Nothing
    End If
    blnCancel = Not mblnOK
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Screen.MousePointer = 0
End Function

Private Sub cmdCancel_Click()
    Set mrsItem = Nothing 'ȡ����־
    Call SaveWinState(Me, App.ProductName, Me.Caption)
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngID As Long
    If mrsItem Is Nothing Then Exit Sub
    If mrsItem.RecordCount < 1 Then Exit Sub
    Call SaveWinState(Me, App.ProductName, Me.Caption)
    If mrsItem.RecordCount > 1 Then
        With vsItem
            If Val(.TextMatrix(.Row, .ColIndex("id"))) > 0 Then lngID = Val(.TextMatrix(.Row, .ColIndex("id")))
        End With
        mrsItem.Filter = "id=" & lngID
    End If
    mblnOK = True: Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If vsItem.Visible Then
        If vsItem.Row = 0 And tvw_s.Visible = True Then
            tvw_s.SetFocus
        Else
            vsItem.SetFocus
        End If
    Else
        tvw_s.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngIdx As Long
    If KeyCode = 13 And cmdOK.Enabled Then
        cmdOK_Click
    ElseIf KeyCode = vbKeyEscape And cmdCancel.Enabled Then
        cmdCancel_Click
    End If
End Sub

Private Sub Form_Load()
    Dim lngScrW As Long, lngScrH As Long, lngColW As Long
    Dim vRect As RECT, strIDs As String, i As Long
    Dim lngUpH As Long, lngDnH As Long

    Call RestoreWinState(Me, App.ProductName, mstrSaveTag)
    
    mblnOK = False
    mstrLike = gstrLike
    mstrKey = ""
    If mstr���� = "" Then
        '��ȡ���ʧ��,����ʾ,��ȡ���˳�
        If Not FillTree Then
            mblnOK = True: Unload Me: Exit Sub
        End If
        '�����,��ʾ,��ȡ���˳�
        If tvw_s.Nodes.Count = 0 Then
            MsgBox "û����������շ���Ŀ���,���ȵ��շ���Ŀ���������á�", vbInformation, gstrSysName
            mblnOK = True: Unload Me: Exit Sub
        End If
    Else
        tvw_s.Visible = False
        pic.Visible = False
        cmdOK.Visible = False
        cmdCancel.Visible = False

        '���ƥ������
        Call FillList(strIDs)
        If mrsItem Is Nothing Then
            Unload Me: Exit Sub
        ElseIf mrsItem.RecordCount = 1 Then
            'ֻ��һ����Ŀʱ,ֱ�ӷ���
            mblnOK = True: Unload Me: Exit Sub
        ElseIf mrsItem.RecordCount > 0 Then
            '������ͬһ����Ŀʱ,ֱ�ӷ���
            If mstr���� <> "" Then
                If UBound(Split(strIDs, ",")) = 1 Then
                    mblnOK = True: Unload Me: Exit Sub
                End If
            End If
            
            vsItem.Appearance = flexFlat
            Call zlControl.FormSetCaption(Me, False, False)
            Me.Left = mlngX: Me.Height = 3240
            lngScrH = GetSystemMetrics(SM_CYFULLSCREEN) * 15 '��Ļ���ø߶�
            If mlngY + mlngH + Me.Height > lngScrH Then
                Me.Top = mlngY - Me.Height
            Else
                Me.Top = mlngY + mlngH
            End If
            
            Call Form_Resize
        Else
            mblnOK = True: Unload Me: Exit Sub
        End If
    End If
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If mstr���� = "" Then
        
         tvw_s.Left = 0
        tvw_s.Top = picInfo.Top + picInfo.Height + 30
        tvw_s.Height = Me.ScaleHeight - picCmd.Height - tvw_s.Top
        
        pic.Top = tvw_s.Top
        pic.Left = tvw_s.Left + tvw_s.Width
        pic.Height = tvw_s.Height
         
        vsItem.Top = tvw_s.Top
        vsItem.Left = pic.Left + pic.Width
        vsItem.Width = Me.ScaleWidth - tvw_s.Width - pic.Width
        vsItem.Height = tvw_s.Height
       
        cmdCancel.Top = cmdOK.Top
        
        If Me.ScaleWidth - cmdCancel.Width * 1.5 < 4100 Then
            cmdCancel.Left = 4100
        Else
            cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 200
        End If
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 100
    Else
        vsItem.Left = 0
        vsItem.Top = 0
        vsItem.Width = Me.ScaleWidth
        vsItem.Height = Me.ScaleHeight
    End If
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveColPosition
    Call SaveColWidth
    Call SaveWinState(Me, App.ProductName, mstrSaveTag)
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If tvw_s.Width + X < 1000 Or vsItem.Width - X < 1000 Then Exit Sub
        pic.Left = pic.Left + X
        tvw_s.Width = tvw_s.Width + X
        vsItem.Left = vsItem.Left + X
        vsItem.Width = vsItem.Width - X
        Me.Refresh
    End If
End Sub

Private Function FillTree() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objNode As Node, strTmp As String
    Dim str���� As String
 
    strSQL = _
    "Select 0 As ��, To_Number('99999999' || ����) As ID, -null As �ϼ�id, '�в�ҩ' As ����" & vbNewLine & _
    "From ���Ʒ���Ŀ¼" & vbNewLine & _
    "Where ���� =3 And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)" & vbNewLine & _
    "Group By ����"
    strSQL = strSQL & " Union ALL " & _
    "Select Level As ��,-id As ID, Nvl(-�ϼ�id, To_Number('99999999' || ����)) As �ϼ�id, '[' || ���� || ']' || ���� As ����" & vbNewLine & _
    "From ���Ʒ���Ŀ¼" & vbNewLine & _
    "Where  ����=3 And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)" & vbNewLine & _
    "Start With �ϼ�id Is Null" & vbNewLine & _
    "Connect By Prior ID = �ϼ�id"

    strSQL = strSQL & " Order by ��,����"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name)
    tvw_s.Visible = True
    For i = 1 To rsTmp.RecordCount
        If IsNull(rsTmp!�ϼ�ID) Then
            Set objNode = tvw_s.Nodes.Add(, , "_" & rsTmp!ID, rsTmp!����, "Close")
        Else
            Set objNode = tvw_s.Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!ID, rsTmp!����, "Close")
        End If
        objNode.Tag = 3 '��ŷ�������:0-��ҩƷ������,1-����ҩ,2-�г�ҩ,3-�в�ҩ,7-��������
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
        'tvw_s.Nodes(1).Selected = True
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

Private Sub FillList(Optional strIDs As String)
'���ܣ����ݵ�ǰ��������װ��������ĿĿ¼
'������blnClass=�Ƿ��ؽ����࿨(Ӧ��������Ŀ�ı�ʱ���ؽ�)
'          strIDs=��ȡ����ĿID��,�����ж�����ʱ�Ƿ������ͬ��ͬһ���շ���Ŀ
    Dim objTab As MSComctlLib.Tab
    Dim objNode As Node, objItem As ListItem
    Dim arrClass As Variant, strClass As String
    Dim strInput As String
    Dim str����ID As String
    Dim lngҩ��ID As Long, strStock As String
    Dim strҩ����λ As String, strҩ����װ As String
    Dim strMain As String, strSQL As String
    Dim strTmp As String, strSQLItem As String
    Dim i As Long
    Dim strWherePriceGrade As String
    Dim cllTemp As Collection, bln��ʾ��� As Boolean
    Dim bln���ҩƷ��� As Boolean
    
    strIDs = ""
    Set objNode = tvw_s.SelectedItem '����ƥ��ʱ,ΪNothing

    '�����Ŀ�嵥�����࿨Ƭ
    '------------------------------------------------------------------------
    vsItem.Rows = vsItem.FixedRows
    vsItem.Rows = vsItem.FixedRows + 1
    Me.Refresh
    

    On Error GoTo errH
    Screen.MousePointer = 11
    '��ȡ��ҩ��Ϣ��¼��
    Set mrsItem = GetChineDrugRecordset(mstr����)

    '������
    '--------------------------------------------------------------------------
    vsItem.Redraw = flexRDNone
    vsItem.ScrollBars = flexScrollBarNone
    Set vsItem.DataSource = mrsItem
    vsItem.ScrollBars = flexScrollBarBoth
    If Err.Number = 0 And gcnOracle.Errors.Count > 0 Then
        gcnOracle.Errors.Clear
    End If
    If vsItem.Rows = vsItem.FixedRows Then
        vsItem.Rows = vsItem.FixedRows + 1
    End If
    
    '�����Ե���
    vsItem.ColAlignment(0) = 4
    vsItem.Cell(flexcpAlignment, 0, 0, 0, vsItem.COLS - 1) = 4
    vsItem.RowHeight(0) = vsItem.RowHeightMin
    For i = 1 To vsItem.COLS - 1
        vsItem.ColKey(i) = UCase(Trim(vsItem.TextMatrix(0, i)))
        If InStr("����,���", vsItem.TextMatrix(0, i)) > 0 Then
            vsItem.ColAlignment(i) = 7
        Else
            vsItem.ColAlignment(i) = 1
        End If
        If UCase(Trim(vsItem.TextMatrix(0, i))) Like "*ID" Then
            vsItem.ColHidden(i) = True
            vsItem.ColWidth(i) = 0
        ElseIf Trim(vsItem.TextMatrix(0, i)) = "����ϵ��" Then
            vsItem.ColHidden(i) = True
            vsItem.ColWidth(i) = 0
        ElseIf vsItem.ColWidth(i) > 2800 Then
            vsItem.ColWidth(i) = 2800
        ElseIf mrsItem.RecordCount = 0 Then
            vsItem.ColWidth(i) = 1000
        End If
        vsItem.ColData(i) = i '��¼ԭʼ�к�,���ڴ�����˳��
    Next
    
    '�ָ���˳��:Ӧ����������֮ǰ
    Call RestoreColPosition
    Call RestoreColWidth
    '������:������,�Ա���洦���к�
    Call RestoreColSort
    
     '���÷��񣬸�����ؼ������ӡ���桱����Ϣ
    If mint��ҩ��̬ = 0 And mlngҩ��id > 0 Then
        If mrsItem.RecordCount > 0 Then
            bln��ʾ��� = InStr(1, mstrPrivs, ";��ʾ���;") > 0
            bln���ҩƷ��� = InStr(mstrPrivs, ";�������;") = 0
            Call gobjPublicExpense.zlLoadStockFromService(vsItem, mrsItem, 0, 0, mlngҩ��id, 0, bln��ʾ���, _
                    bln���ҩƷ���, False, False)
        End If
     End If

    '��Ƭ������ݼ���
    '------------------------------------------------------------------------
    With vsItem
        For i = 1 To vsItem.Rows - 1
            .TextMatrix(i, 0) = i
            .RowHeight(i) = vsItem.RowHeightMin
            '�ռ���ĿID:ֻ�ռ����2��
            If mstr���� <> "" Then
                If UBound(Split(strIDs, ",")) < 2 And Val(.TextMatrix(i, .ColIndex("ID"))) > 0 Then
                    If InStr(strIDs & ",", "," & Val(.TextMatrix(i, .ColIndex("ID"))) & ",") = 0 Then
                        strIDs = strIDs & "," & Val(.TextMatrix(i, .ColIndex("ID")))
                    End If
                End If
            End If
        Next
    End With

    '�к��п��
    vsItem.ColWidth(0) = Me.TextWidth(vsItem.TextMatrix(vsItem.Rows - 1, 0) & " ")
    If vsItem.ColWidth(0) < 380 Then vsItem.ColWidth(0) = 380
    
    vsItem.Row = vsItem.FixedRows: vsItem.Col = vsItem.FixedCols
    Call vsItem_AfterRowColChange(-1, -1, vsItem.Row, vsItem.Col)
    
    'ѡ������Ҳ������
    If CreatePlugIn(0) Then
        On Error Resume Next
        Call gobjPlugIn.AfterSelectorReady(99, "��ҩѡ����", vsItem, mfrmParent)
        Call zlPlugInErrH(Err, "AfterSelectorReady")
        Err.Clear: On Error GoTo errH
    End If
    
    vsItem.Redraw = flexRDDirect

    Call Form_Resize
    
    If Val(vsItem.TextMatrix(1, vsItem.ColIndex("id"))) = 0 Then mrsItem.Filter = "id=-1"
    If mrsItem.RecordCount > 0 Then mrsItem.MoveFirst

    Screen.MousePointer = 0
    Exit Sub
errH:
    LockWindowUpdate 0
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    cmdOK.Enabled = False
End Sub

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    
    If mstrKey = Node.Key Then Exit Sub
    mstrKey = Node.Key
    Call FillList
    
End Sub

Private Function GetSubTree(ByVal objNode As Node) As String
'���ܣ�����һ��������������Key(���ý��)
    Dim strKeys As String
    Dim objTmp As Node
    
    strKeys = "," & Mid(objNode.Key, 2) & strKeys
    Set objTmp = objNode.Child
    Do While Not objTmp Is Nothing
        If objTmp.Children > 0 Then
            strKeys = "," & GetSubTree(objTmp) & strKeys
        Else
            strKeys = "," & Mid(objTmp.Key, 2) & strKeys
        End If
        Set objTmp = objTmp.Next
    Loop
    GetSubTree = Mid(strKeys, 2)
End Function

Private Function GetChineDrugRecordset(ByVal strInput As String) As ADODB.Recordset
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡ��ҩ��Ϣ��¼��
    '��Σ�strInput-Ҫ���ҵ�ֵ
    '���Σ�
    '���أ���¼��
    '------------------------------------------------------------------------------------------------------------------------
    Dim str���� As String, str��� As String, str��׼��Ŀ As String
    Dim strSQL As String, str����ID As String
    Dim str����ʱ�� As String, strWhere As String
    Dim lng����id As Long
    Dim objNode As Node
    Dim bln��ʾ�¼� As Boolean
    
    On Error GoTo errHandle

    '����ҩƷȨ��
    str���� = ""
    If InStr(mstrPrivs, ";����ҩƷ����;") = 0 Then str���� = str���� & " And E.�������<>'����ҩ'"
    If InStr(mstrPrivs, ";����ҩƷ����;") = 0 Then str���� = str���� & " And E.�������<>'����ҩ'"
    If InStr(mstrPrivs, ";����ҩƷ����;") = 0 Then str���� = str���� & " And E.��ֵ���� Not IN('����','����')"
    bln��ʾ�¼� = False
    Set objNode = tvw_s.SelectedItem '����ƥ��ʱ,ΪNothing
    If mstr���� = "" Then
        lng����id = -1 * Val(Mid(objNode.Key, 2))
         '�����еķ���ID
        If bln��ʾ�¼� Then
            '��ʾ�¼�����Ŀ
            If Mid(objNode.Key, 2) = "99999999" & objNode.Tag Then
                str����ID = " And A.����ID IN(Select ID From ���Ʒ���Ŀ¼ Where ����=3)"
            Else
                str����ID = " And A.����ID IN(Select ID From ���Ʒ���Ŀ¼ Start With ID=[9]Connect by Prior ID=�ϼ�ID)"
            End If
        Else
                str����ID = " And A.����ID=[9]"
        End If
    Else
        If Len(mstr����) < 2 Then mstrLike = "" '�Ż�
    End If
    
    If mint��ҩ��̬ = 0 Then
        str��� = _
        "   And Nvl(C.��ҩ��̬,0) = [6] And (D.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or D.����ʱ�� IS NULL) And D.������� IN([7],3)" & _
        "   And (D.վ��='" & gstrNodeNo & "' Or D.վ�� is Null) "
    Else
         str��� = " And Exists(Select 1 From ҩƷ��� C Where C.ҩ��ID=E.ҩ��ID And Nvl(C.��ҩ��̬,0) = [6])"
    End If
    
    str����ʱ�� = "" & _
        "   And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " & _
        "   And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        "   And A.������� IN([7],3)"
    
    str��׼��Ŀ = ""
    If mint��ҩ��̬ = 0 Then
        If mint���� <> 0 Then
            '���˺�:24862
            If zl_Check��׼��Ŀ(gclsInsure, mint����, mlng����ID, False) Then str��׼��Ŀ = Get������׼��Ŀ(mlng����ID, "D.ID")
        End If
    End If
        
    If strInput <> "" Then
        strWhere = " And (A.���� Like [1] And B.����=[3] Or B.���� Like [2] And B.����=[3] Or B.���� Like upper([2]) And B.���� IN([3],3))"
        If IsNumeric(strInput) Then                         '10,11.����ȫ������ʱֻƥ�����'����ҩƷ,��Ҫƥ�����(����Ϊ3��������)
            If Mid(gstrMatchMode, 1, 1) = "1" Then strWhere = " And (A.���� Like [1] And B.����=[3] Or B.���� Like Upper([2]) And B.����=3)"
        ElseIf zlCommFun.IsCharAlpha(strInput) Then         '01,11.����ȫ����ĸʱֻƥ�����
            If Mid(gstrMatchMode, 2, 1) = "1" Then strWhere = " And B.���� Like Upper([2]) And B.����=[3]"
        ElseIf zlCommFun.IsCharChinese(strInput) Then
            strWhere = " And B.���� Like [2] And B.����=[3]"
        End If
         '��ɢװʱ��Ʒ����ʾ���Ҳ���ʾ���
        strSQL = "" & _
        "   Select  distinct A.ID,A.����,A.����,A.���㵥λ" & _
        "   From ������ĿĿ¼ A,������Ŀ���� B" & _
        "   Where A.ID=B.������ĿID  And A.���='7' " & str����ʱ�� & strWhere
        
        If mint��ҩ��̬ = 0 Then
            'ɢװ����ʾ�����,����ԭ������
            strSQL = _
            " Select distinct  A.ID as ҩ��ID,C.ҩƷID as ID,C.ҩƷID,D.����,A.����,D.���,A.���㵥λ as ������λ," & _
                    IIf(gblnҩ����λ, "C." & gstrҩ����λ, "D.���㵥λ") & " as ��λ,D.����,D.��������,d.ִ�п��� AS ִ�п���_ID," & IIf(mint���� <> 0, "N.���� ҽ������,", "") & _
            "       Decode(D.�Ƿ���,1,'ʱ��',LTrim(To_Char(Sum(F.�ּ�)" & _
                    IIf(gblnҩ����λ, "*Nvl(C." & gstrҩ����װ & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "'))) as ����," & _
            "       NULL as ���," & IIf(gblnҩ����λ, "C." & gstrҩ����װ, "1 ") & " as ����ϵ��,7 as ���id" & _
            " From ҩƷ���� E,ҩƷ��� C,�շ���ĿĿ¼ D,�շѼ�Ŀ F, " & vbNewLine & _
                        IIf(mint���� <> 0, "����֧����Ŀ M,����֧������ N,", "") & vbNewLine & _
            "          (" & strSQL & ") A " & vbNewLine & _
            " Where   A.ID=E.ҩ��ID And A.ID=C.ҩ��ID And C.ҩƷID=D.ID  " & vbNewLine & _
            "        And D.ID=F.�շ�ϸĿID " & vbNewLine & _
                     IIf(mint���� <> 0, " And C.ҩƷID=M.�շ�ϸĿID(+) And M.����(+)=[5] And M.����ID=N.ID(+)" & vbNewLine, "") & _
            "        And exists(Select 1 From �շ�ִ�п��� A1 Where A1.�շ�ϸĿID=C.ҩƷID And A1.ִ�п���ID=[4]   And (A1.������Դ is NULL Or A1.������Դ=[7]) and (A1.��������ID is null or A1.��������ID=[8])  ) " & vbNewLine & _
            "        And Sysdate Between F.ִ������ and Nvl(F.��ֹ����,TO_DATE('3000-01-01','YYYY-MM-DD'))" & _
                     str��� & str���� & str��׼��Ŀ & _
            " Group by A.ID,C.ҩƷID,A.���㵥λ,D.����,A.����,D.���,D.����,D.��������,d.ִ�п���,D.�Ƿ���," & IIf(mint���� <> 0, "N.����,", "") & _
                    IIf(gblnҩ����λ, "C.���ﵥλ,C.�����װ", "D.���㵥λ") & _
            " Order by D.����"
        Else
             '��ɢװʱ��Ʒ����ʾ���Ҳ���ʾ���
            strSQL = strSQL & _
            "        And exists(Select 1 From ����ִ�п��� A1 Where A1.������ĿID=A.ID And A1.ִ�п���ID=[4]   And (A1.������Դ is NULL Or A1.������Դ=[7]) and (A1.��������ID is null or A1.��������ID=[8])  ) " & vbNewLine
            strSQL = _
                " Select Distinct A.ID,A.ID as ҩ��ID,A.����,A.����,A.���㵥λ as ��λ" & _
                " From ҩƷ���� E,(" & strSQL & ") A" & _
                " Where A.ID=E.ҩ��ID  " & _
                "         And Exists(Select 1 From ҩƷ��� C Where C.ҩ��ID=E.ҩ��ID And Nvl(C.��ҩ��̬,0) = [6])" & _
                "         And Rownum<=100" & _
                " Order by A.����"
        End If
    Else
        If mint��ҩ��̬ = 0 Then
            'ɢװ����ʾ�����,����ԭ������
            strSQL = "" & _
            "  Select ҩƷID As Id,ҩ��ID,�ϼ�ID,����,����,���,������λ,��λ,����,��������,ִ�п���_ID,����,���,ҩƷID,����ϵ��,7 as ���id " & _
            "  From ( " & _
            " Select A.ID,A.ID as ҩ��ID,A.����ID as �ϼ�ID,D.����,D.����,D.���,A.���㵥λ as ������λ," & _
                        IIf(gblnҩ����λ, " C." & gstrҩ����λ, "D.���㵥λ") & " as ��λ,D.����,D.��������,d.ִ�п��� as ִ�п���_ID" & IIf(mint���� = 0, "", ",N.���� ҽ������") & "," & _
            "           Decode(D.�Ƿ���,1,'ʱ��',LTrim(To_Char(Sum(F.�ּ�)" & _
                        IIf(gblnҩ����λ, "*Nvl(C." & gstrҩ����װ & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "'))) as ����," & _
            "           NULL as ���,C.ҩƷID, " & IIf(gblnҩ����λ, "C." & gstrҩ����װ, "1 ") & " as ����ϵ��" & _
            " From  ������ĿĿ¼ A,ҩƷ���� E,ҩƷ��� C,�շ���ĿĿ¼ D,�շѼ�Ŀ F" & _
                        IIf(mint���� = 0, "", ",����֧����Ŀ M,����֧������ N") & _
            " Where A.ID=E.ҩ��ID And A.ID=C.ҩ��ID And C.ҩƷID=D.ID And C.ҩƷID =F.�շ�ϸĿID And A.���='7'  " & _
                    IIf(mint���� = 0, "", "       And C.ҩƷID=M.�շ�ϸĿID(+) And   M.����(+)=" & mint���� & " And M.����ID=N.ID(+)") & _
            "        And exists(Select 1 From �շ�ִ�п��� A1 Where A1.�շ�ϸĿID=C.ҩƷID And A1.ִ�п���ID=[4]   And (A1.������Դ is NULL Or A1.������Դ=[7]) and (A1.��������ID is null or A1.��������ID=[8])  ) " & vbNewLine & _
            "       And Sysdate Between F.ִ������ and Nvl(F.��ֹ����,TO_DATE('3000-01-01','YYYY-MM-DD'))" & _
            "       And D.������� IN(" & mint������Դ & ",3)" & str��׼��Ŀ & str��� & str����ʱ�� & str����ID & _
            " Group by A.ID,A.���㵥λ ,A.����ID,D.����,D.����,D.���,D.����,D.��������,d.ִ�п���" & IIf(mint���� = 0, "", ",N.����") & ",D.�Ƿ���,C.ҩƷID," & _
                 IIf(gblnҩ����λ, "C.���ﵥλ,C.�����װ", "D.���㵥λ") & _
            ")"
        Else
            '��ɢװʱ��Ʒ����ʾ���Ҳ���ʾ���
            strSQL = "" & _
            "Select Distinct A.ID,ID as ҩ��ID,A.����ID as �ϼ�ID,A.����,A.����,A.���㵥λ as ��λ,E.����ְ�� as ����ְ��ID" & _
            " From ������ĿĿ¼ A,ҩƷ���� E" & _
            " Where A.ID=E.ҩ��ID" & str���� & str����ʱ�� & str��� & str����ID & _
            "        And exists(Select 1 From ����ִ�п��� A1 Where A1.������ĿID=A.ID And A1.ִ�п���ID=[4]   And (A1.������Դ is NULL Or A1.������Դ=[7]) and (A1.��������ID is null or A1.��������ID=[8])  ) " & vbNewLine
        End If
    End If
    Set GetChineDrugRecordset = zlDatabase.OpenSQLRecord(strSQL, "�в�ҩ", strInput & "%", mstrLike & strInput & "%", gbytCode + 1, mlngҩ��id, mint����, mint��ҩ��̬, mint������Դ, mlng��������ID, lng����id)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SaveColPosition(Optional ByVal strType As String)
'���ܣ�������˳��:�к�,˳��|...
'˵����Ӧ����SaveWinState֮ǰ,���ڲ�ʹ�ø��Ի�ʱ��ע������
    Dim strPos As String, i As Long
        
    If Not gblnMyStyle Then Exit Sub
    
    With vsItem
        For i = 0 To .COLS - 1
            strPos = strPos & "|" & .ColData(i) & "," & i
        Next
        
        If mstr���� = "" And strType = "" And Not tvw_s.SelectedItem Is Nothing Then strType = tvw_s.SelectedItem.Tag
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColPosition", Mid(strPos, 2)
    End With
End Sub

Private Sub SaveColWidth(Optional ByVal strType As String)
'���ܣ������п��
'˵����Ӧ����SaveWinState֮ǰ,���ڲ�ʹ�ø��Ի�ʱ��ע������
    Dim strPos As String, i As Long
        
    If Not gblnMyStyle Then Exit Sub
    If mstr���� = "" And strType = "" And Not tvw_s.SelectedItem Is Nothing Then strType = tvw_s.SelectedItem.Tag
    Call SaveFlexState(vsItem, App.ProductName & Me.Name & strType)
End Sub

Private Sub RestoreColWidth()
'���ܣ��ָ��п��
'˵����Ӧ���ڻָ�����֮��
    Dim strType As String
    
    If Not gblnMyStyle Then Exit Sub
    
    If mstr���� = "" Then strType = tvw_s.SelectedItem.Tag
    Call RestoreFlexState(vsItem, App.ProductName & Me.Name & strType)
End Sub


Private Sub RestoreColPosition()
'���ܣ��ָ���˳��
'˵����Ӧ����������֮ǰ
    Dim rsPos As New ADODB.Recordset
    Dim strType As String, strPos As String
    Dim i As Long, j As Long
    
    If Not gblnMyStyle Then Exit Sub
    
    With vsItem
        If mstr���� = "" Then strType = tvw_s.SelectedItem.Tag
        strPos = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColPosition", "")
        If strPos <> "" Then
            rsPos.Fields.Append "Col", adBigInt
            rsPos.Fields.Append "Position", adBigInt
            rsPos.CursorLocation = adUseClient
            rsPos.LockType = adLockOptimistic
            rsPos.CursorType = adOpenStatic
            rsPos.Open
            
            For i = 0 To UBound(Split(strPos, "|"))
                rsPos.AddNew
                rsPos!Col = Split(Split(strPos, "|")(i), ",")(0)
                rsPos!Position = Split(Split(strPos, "|")(i), ",")(1)
                rsPos.Update
            Next
            rsPos.Sort = "Position"
            
            'ColPosition:>=0,ReadOnly,�ı������к�Ҳ�ı�
            For i = 1 To rsPos.RecordCount
                For j = i - 1 To .COLS - 1
                    If .ColData(j) = rsPos!Col Then Exit For
                Next
                If j <= .COLS - 1 Then
                    .ColPosition(j) = rsPos!Position
                End If
                rsPos.MoveNext
            Next
        End If
    End With
End Sub

Private Sub RestoreColSort()
'���ܣ�������
    Dim strType As String, strSort As String, i As Long
        
    With vsItem
        Set .Cell(flexcpPicture, 0, 0, 0, .COLS - 1) = Nothing
        .Cell(flexcpPictureAlignment, 0, 0, 0, .COLS - 1) = 7
        If gblnMyStyle Then
            If mstr���� = "" Then strType = tvw_s.SelectedItem.Tag
            strSort = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColSort", "")
            If strSort <> "" Then
                '��Ϊ���ܵ�����˳��,���Բ�����ʵ��������
                For i = 0 To .COLS - 1
                    If .ColData(i) = Val(Split(strSort, ",")(0)) Then Exit For
                Next
                If i <= .COLS - 1 Then
                    .Col = i
                    .Sort = Val(Split(strSort, ",")(1))
                    
                    If Val(Split(strSort, ",")(1)) Mod 2 = 1 Then
                        .Cell(flexcpPicture, 0, i) = img.ListImages(3).Picture
                    Else
                        .Cell(flexcpPicture, 0, i) = img.ListImages(4).Picture
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub vsItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= vsItem.FixedRows Then
        cmdOK.Enabled = Val(vsItem.TextMatrix(NewRow, 1)) <> 0
    Else
        cmdOK.Enabled = False
    End If
End Sub

Private Sub vsItem_AfterSort(ByVal Col As Long, Order As Integer)
    Dim strType As String, i As Long
    
    With vsItem
        .Cell(flexcpPicture, 0, 0, 0, .COLS - 1) = Nothing
        
        If Order Mod 2 = 1 Then
            .Cell(flexcpPicture, 0, Col) = img.ListImages(3).Picture
        Else
            .Cell(flexcpPicture, 0, Col) = img.ListImages(4).Picture
        End If
        
        If Val(.TextMatrix(.Row, 1)) <> 0 Then
            .Redraw = flexRDNone
            For i = 1 To .Rows - 1
                .TextMatrix(i, 0) = i
            Next
            .Redraw = flexRDDirect
            Call vsItem_AfterRowColChange(-1, -1, .Row, .Col)
        End If
            
        '��Ϊ������˳��ı�,���Ա���ԭʼ�к�
        If mstr���� = "" Then strType = tvw_s.SelectedItem.Tag
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColSort", .ColData(Col) & "," & Order
    End With
End Sub

Private Sub vsItem_BeforeSort(ByVal Col As Long, Order As Integer)
    'ǿ�Ʊ����а��ַ�������
    If vsItem.TextMatrix(0, Col) = "����" Then
        If Order = 1 Then Order = 7
        If Order = 2 Then Order = 8
    End If
End Sub

Private Sub vsItem_DblClick()
    If vsItem.MouseRow >= vsItem.FixedRows Then
        Call vsItem_KeyPress(13)
    End If
End Sub

Private Sub vsItem_KeyPress(KeyAscii As Integer)
    Static strIdx As String
    Static sngTim As Single
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cmdOK.Enabled Then cmdOK_Click
    Else
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            If Abs(Timer - sngTim) > 0.5 Then
                strIdx = ""
            End If
            sngTim = Timer
            strIdx = strIdx & Chr(KeyAscii)
            KeyAscii = 0
            
            If Len(strIdx) > 4 Then strIdx = Left(strIdx, 4)
            
            If vsItem.Rows - 1 >= CInt(strIdx) And CInt(strIdx) > 0 Then
                vsItem.Row = Val(strIdx)
                vsItem.ShowCell vsItem.Row, vsItem.Col
            End If
        End If
    End If
End Sub


