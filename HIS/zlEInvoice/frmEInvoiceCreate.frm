VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEInvoiceCreate 
   BorderStyle     =   0  'None
   Caption         =   "����Ʊ�ݿ���"
   ClientHeight    =   10860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13770
   LinkTopic       =   "Form1"
   ScaleHeight     =   10860
   ScaleWidth      =   13770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picMain 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   8748
      Left            =   384
      ScaleHeight     =   8745
      ScaleWidth      =   12945
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   12948
      Begin VB.PictureBox picFilter 
         BorderStyle     =   0  'None
         Height          =   468
         Left            =   168
         ScaleHeight     =   465
         ScaleWidth      =   12690
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   408
         Width           =   12684
         Begin VB.ComboBox cboҵ������ 
            Height          =   276
            Left            =   912
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   96
            Width           =   1812
         End
         Begin VB.ComboBox cbo�շ�Ա 
            Height          =   276
            Left            =   3720
            TabIndex        =   6
            Top             =   96
            Width           =   1812
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "ˢ��(&R)"
            Height          =   300
            Left            =   11640
            TabIndex        =   11
            Top             =   84
            Width           =   1000
         End
         Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
            Height          =   276
            Left            =   6648
            TabIndex        =   8
            Top             =   96
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   476
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   139329539
            CurrentDate     =   43941
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   276
            Left            =   9168
            TabIndex        =   10
            Top             =   96
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   476
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   139329539
            CurrentDate     =   43941
         End
         Begin VB.Label lblҵ������ 
            AutoSize        =   -1  'True
            Caption         =   "ҵ������"
            Height          =   180
            Left            =   144
            TabIndex        =   3
            Top             =   144
            Width           =   720
         End
         Begin VB.Label lbl�շ�Ա 
            AutoSize        =   -1  'True
            Caption         =   "�շ�Ա"
            Height          =   180
            Left            =   3120
            TabIndex        =   5
            Top             =   144
            Width           =   540
         End
         Begin VB.Label lbl�շ�ʱ�� 
            AutoSize        =   -1  'True
            Caption         =   "�շ�ʱ��"
            Height          =   180
            Left            =   5880
            TabIndex        =   7
            Top             =   144
            Width           =   720
         End
         Begin VB.Label lblҵ��ʱ��_ 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   8928
            TabIndex        =   9
            Top             =   144
            Width           =   180
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfExse 
         Height          =   1356
         Left            =   432
         TabIndex        =   12
         Top             =   2184
         Width           =   6108
         _cx             =   10774
         _cy             =   2392
         Appearance      =   2
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
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
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
   End
   Begin VB.Shape shpBorder 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000C&
      Height          =   1032
      Left            =   24
      Top             =   900
      Width           =   528
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      Height          =   300
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   10845
      _Version        =   589884
      _ExtentX        =   19129
      _ExtentY        =   529
      _StockProps     =   6
      Caption         =   "��������Ʊ��"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmEInvoiceCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form, mlngSys As Long, mlngModule As Long, mstrDBUser As String, mstrEInvPrivs As String
Private mcbsMain   As Object          'CommandBar�ؼ�
Private mobjEInvoice As clsEInvoiceModule
Private mobjPubEInvoice As Object 'zlPublicExpense.clsPubEInvoice
Private mblnPrinting As Boolean
Private mrs�շ�Ա As ADODB.Recordset

Public Event ShowPopupMenu(ByVal blnAddOutPutExcel As Boolean)
Public Event ShowInfo(ByVal strInfo As String)

Public Sub InitCommVariable(frmParent As Form, cbsThis As Object, ByVal lngSys As Long, lngModule As Long, ByVal strDBUser As String, _
    ByVal strEInvPrivs As String, objEInvoice As Object, objPubEInvoice As Object)
    '��ʼ������
    Set mfrmMain = frmParent
    Set mcbsMain = cbsThis
    mstrDBUser = strDBUser
    mlngSys = lngSys: mlngModule = lngModule
    mstrEInvPrivs = strEInvPrivs
    Set mobjEInvoice = objEInvoice
    Set mobjPubEInvoice = objPubEInvoice
End Sub

Public Sub zlDefCommandBars(Optional ByVal blnInsideTools As Boolean)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom

    '�ļ��˵�
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With cbrMenuBar.CommandBar.Controls
        '���������Excel֮��
        Set cbrControl = .Find(, conMenu_File_Excel)
    End With

    '�༭�˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If

    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", cbrMenuBar.Index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_EInvoice, "��Ʊ(&N)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SelAll, "ȫѡ(&A)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClsAll, "ȫ��(&C)")
    End With

    '�鿴�˵�
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) 'ˢ����ǰ(���ʱע�ⷴ��)
    End With

    '����������
    '-----------------------------------------------------
    Set cbrToolBar = mcbsMain(2)
    For Each cbrControl In cbrToolBar.Controls '�����ǰ������һ��Control
        If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.Index - 1): Exit For
        End If
    Next
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_EInvoice, "��Ʊ", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SelAll, "ȫѡ", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClsAll, "ȫ��", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��", cbrControl.Index + 1): cbrControl.BeginGroup = True
    End With

    '����Ŀ����
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_SelAll
        .Add FCONTROL, Asc("C"), conMenu_Edit_ClsAll
    End With
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Select Case Control.ID
    Case conMenu_File_Preview 'Ԥ��
        Call OutputList(2)
    Case conMenu_File_Print '��ӡ
        Call OutputList(1)
    Case conMenu_File_Excel '�����Excel��
        Call OutputList(3)
    Case conMenu_Edit_EInvoice '��Ʊ
        Call CreateEInvoice
    Case conMenu_Edit_SelAll 'ȫѡ
        Call Grid_SelAllRecord(vsfExse, True)
    Case conMenu_Edit_ClsAll 'ȫ��
        Call Grid_SelAllRecord(vsfExse, False)
    Case conMenu_View_Refresh 'ˢ��
        Call cmdRefresh_Click
    End Select
End Sub

Private Sub CreateEInvoice()
    Dim cllSwapData As Collection, strErrMsg As String
    Dim lng����ID As Long, strDate As String, bln������ As Boolean
    Dim i As Long, byt���� As Byte, lng����ID As Long, blnInit As Boolean
    Dim lngCount As Long, blnChecked As Boolean
    
    On Error GoTo ErrHandler
    With vsfExse
        blnChecked = False
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = vbChecked Then blnChecked = True: Exit For
        Next
    
        .Cell(flexcpForeColor, .FixedRows, 0, .Rows - 1, .Cols - 1) = vbBlack
        For i = .FixedRows To .Rows - 1
            lng����ID = Val(.RowData(i)): lng����ID = 0
            If lng����ID <> 0 And (.Cell(flexcpChecked, i, .ColIndex("ѡ��")) = vbChecked Or Not blnChecked And i = .Row) Then
                lngCount = lngCount + 1
                
                byt���� = Val(NVL(.Cell(flexcpData, i, .ColIndex("ҵ������")))) 'Array("1-�շ�", "2-Ԥ��", "3-����", "4-�Һ�", "5-���￨")
                If byt���� = 1 Or byt���� = 4 Then
                    bln������ = Val(NVL(.Cell(flexcpData, i, .ColIndex("���ݺ�")))) = 1
                ElseIf byt���� = 12 Then '����˿�
                    lng����ID = Val(NVL(.Cell(flexcpData, i, .ColIndex("���ݺ�"))))
                End If
                
                byt���� = byt���� Mod 10
                If blnInit = False Then
                    If GetPubEInvoiceObject(Me, mlngSys, mlngModule, mobjPubEInvoice, byt����) = False Then Exit Sub
                    blnInit = True
                End If
                
                If mobjPubEInvoice.zlGetEInvoiceIDFromBalanceID(byt����, lng����ID) <> 0 Then
                        .TextMatrix(i, .ColIndex("��Ʊ���")) = "��Ʊʧ��"
                        .TextMatrix(i, .ColIndex("��Ʊ˵��")) = "���ν����ѿ��ߵ���Ʊ�ݣ���ˢ�º����ԡ�"
                        .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                Else
                    If GetSwapCollectFromBalanceID(byt����, lng����ID, cllSwapData, bln������, lng����ID, False, strErrMsg) = False Then
                        .TextMatrix(i, .ColIndex("��Ʊ���")) = "��Ʊʧ��"
                        .TextMatrix(i, .ColIndex("��Ʊ˵��")) = strErrMsg
                        .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                    Else
                        If mobjPubEInvoice.zlOnlyCreateEinvoice(Me, byt����, cllSwapData, Nothing, False, strErrMsg) Then
                            .TextMatrix(i, .ColIndex("��Ʊ���")) = "��Ʊ�ɹ�"
                            .TextMatrix(i, .ColIndex("��Ʊ˵��")) = ""
                        Else
                            .TextMatrix(i, .ColIndex("��Ʊ���")) = "��Ʊʧ��"
                            .TextMatrix(i, .ColIndex("��Ʊ˵��")) = strErrMsg
                            .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                        End If
                    End If
                End If
            
                If Not blnChecked Then Exit For
            End If
        Next
    End With
    
    If lngCount = 0 Then
        MsgBox "��ѡ����Ҫ���ߵ���Ʊ�ݵķ��ü�¼��", vbInformation, gstrSysName
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    If Not Me.Visible Then Exit Sub
    On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel 'Ԥ��,��ӡ,�����Excel��
        Control.Enabled = vsfExse.TextMatrix(1, 1) <> ""
    
    Case conMenu_Edit_EInvoice '��Ʊ
        Control.Visible = zlStr.IsHavePrivs(mstrEInvPrivs, "���ߵ���Ʊ��")
        Control.Enabled = Control.Visible And vsfExse.TextMatrix(1, 1) <> ""
    
    Case conMenu_Edit_SelAll 'ȫѡ
        Control.Enabled = vsfExse.TextMatrix(1, 1) <> ""
    Case conMenu_Edit_ClsAll 'ȫ��
        Control.Enabled = vsfExse.TextMatrix(1, 1) <> ""
    End Select
End Sub

Private Sub cbo�շ�Ա_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
     
    If cbo�շ�Ա.ListIndex <> -1 Then
        '�����б�ʱ,�����ı�������������
        If UCase(cbo�շ�Ա.Text) <> UCase(cbo�շ�Ա.List(cbo�շ�Ա.ListIndex)) Then Call zlControl.CboSetIndex(cbo�շ�Ա.hWnd, -1)
    End If
    
    If cbo�շ�Ա.Text = "" Then
        cbo�շ�Ա.ListIndex = -1
    ElseIf cbo�շ�Ա.ListIndex = -1 Then
        If Select�շ�Ա(Me, mlngSys, mlngModule, cbo�շ�Ա, mrs�շ�Ա) = False Then
            KeyAscii = 0: zlControl.TxtSelAll cbo�շ�Ա: Exit Sub
        End If
    End If
    
    If cbo�շ�Ա.ListIndex = -1 Then cbo�շ�Ա.Text = ""
End Sub

Private Sub cbo�շ�Ա_LostFocus()
    If cbo�շ�Ա.Text <> "" And cbo�շ�Ա.ListIndex < 0 Then cbo�շ�Ա.Text = ""
End Sub

Private Sub cmdRefresh_Click()
    '����δ���ߵ���Ʊ�ݵķ�������
    Dim dtBegin As Date, dtEnd As Date
    
    '1.���ݼ��
    On Error GoTo ErrHandler
    dtBegin = dtp��ʼʱ��.Value: dtEnd = dtp����ʱ��.Value
    If dtp��ʼʱ�� > dtp����ʱ�� Then
        MsgBox "���õĿ�ʼʱ�䲻�ܴ��ڽ���ʱ�䣡", vbInformation, gstrSysName
        zlControl.ControlSetFocus dtp����ʱ��:  Exit Sub
    End If
    
    If DateDiff("m", dtp��ʼʱ��, dtp����ʱ��) > 6 Then
        If MsgBox("�Ե�ǰ����ʱ�䷶Χ�ڵ����ݽ��в�ѯ������Ҫ�ϳ�ʱ�䣬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    '2.��ȡ����
    Dim rsExse As ADODB.Recordset
    If GetExseData(zlStr.NeedCode(cboҵ������.Text), zlStr.NeedName(cbo�շ�Ա.Text), _
        dtp��ʼʱ��.Value, dtp����ʱ��.Value, rsExse) = False Then Exit Sub
        
    Dim lngOldRow As Long, lngOldCol As Long
    lngOldRow = vsfExse.Row: lngOldCol = vsfExse.Col
    
    'ѡ��,ҵ������,NO,����,�Ա�,����,�����,סԺ��,�շ�Ա,�շ�ʱ��,���ý��,��Ʊ���,��Ʊ˵��
    vsfExse.Clear 1
    vsfExse.Rows = vsfExse.FixedRows + 1
    
    With vsfExse
        .Redraw = flexRDNone
        Do While Not rsExse.EOF
            If .TextMatrix(.Rows - 1, .ColIndex("ҵ������")) <> "" Then .Rows = .Rows + 1
            .RowData(.Rows - 1) = Val(NVL(rsExse!����ID))
            .TextMatrix(.Rows - 1, .ColIndex("ҵ������")) = Decode(Val(NVL(rsExse!ҵ������)) Mod 10, 1, "�շ�", 2, "Ԥ��", 3, "����", 4, "�Һ�", 5, "���￨")
            .Cell(flexcpData, .Rows - 1, .ColIndex("ҵ������")) = Val(NVL(rsExse!ҵ������))
            .TextMatrix(.Rows - 1, .ColIndex("���ݺ�")) = NVL(rsExse!NO)
            Select Case Val(NVL(rsExse!ҵ������)) Mod 10
            Case 1
                .Cell(flexcpData, .Rows - 1, .ColIndex("���ݺ�")) = NVL(rsExse!������)
            Case 2
                .Cell(flexcpData, .Rows - 1, .ColIndex("���ݺ�")) = NVL(rsExse!����ID)
            Case 4
                .Cell(flexcpData, .Rows - 1, .ColIndex("���ݺ�")) = NVL(rsExse!������)
            End Select
            .TextMatrix(.Rows - 1, .ColIndex("����")) = NVL(rsExse!����)
            .TextMatrix(.Rows - 1, .ColIndex("�Ա�")) = NVL(rsExse!�Ա�)
            .TextMatrix(.Rows - 1, .ColIndex("����")) = NVL(rsExse!����)
            .TextMatrix(.Rows - 1, .ColIndex("�����")) = NVL(rsExse!�����)
            .TextMatrix(.Rows - 1, .ColIndex("סԺ��")) = NVL(rsExse!סԺ��)
            .TextMatrix(.Rows - 1, .ColIndex("�շ�Ա")) = NVL(rsExse!����Ա����)
            .TextMatrix(.Rows - 1, .ColIndex("�շ�ʱ��")) = Format(NVL(rsExse!�տ�ʱ��), "yyyy-MM-dd HH:mm:ss")
            .TextMatrix(.Rows - 1, .ColIndex("���ý��")) = FormatEx(Val(NVL(rsExse!���)), 2, , , 6)
        
            rsExse.MoveNext
        Loop
            
        .Cell(flexcpFontBold, .FixedRows, .ColIndex("��Ʊ���"), .Rows - 1, .ColIndex("��Ʊ˵��")) = True
        
        If .Rows > .FixedRows And .Cols > .FixedCols Then     'ȱʡ��λ��
            .Row = -1 '��֤��ѡ���в���������Ҳ����RowColChange�¼�
            .Row = IIf(lngOldRow < .FixedRows Or lngOldRow > .Rows - 1, IIf(.Rows - 1 > .FixedRows, .FixedRows + 1, .FixedRows), lngOldRow)
            .Col = IIf(lngOldCol = 0 Or lngOldCol > .Cols - 1, .FixedCols, lngOldCol)
            .ShowCell .Row, .Col  '������ʾ��ָ����Ԫ
        End If
        
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHandler:
    vsfExse.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Call InitExseGrid
    
    Dim varData As Variant, i As Integer
    varData = Array("1-�շ�", "2-Ԥ��", "3-����", "4-�Һ�", "5-���￨")
    cboҵ������.Clear
    For i = 0 To UBound(varData)
        cboҵ������.AddItem varData(i)
    Next
    cboҵ������.ListIndex = 0
    
    Call Load�շ�Ա(cbo�շ�Ա, mrs�շ�Ա)

    dtp����ʱ��.Value = zlDatabase.Currentdate
    dtp��ʼʱ��.Value = Format(DateAdd("d", -1, dtp����ʱ��.Value), "yyyy-MM-dd 00:00:00")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    sccTitle.Move 8, 8, shpBorder.Width - 20
    picMain.Move sccTitle.Left, sccTitle.Top + sccTitle.Height, Me.ScaleWidth - 2 * sccTitle.Left, Me.ScaleHeight - (2 * sccTitle.Top + sccTitle.Height)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmMain = Nothing
    Set mcbsMain = Nothing
    Set mobjEInvoice = Nothing
    
    Set mrs�շ�Ա = Nothing
End Sub

Private Function InitExseGrid() As Boolean
    '��ʼ��VSFGrid���ؼ�
    Dim strHead As String, varData As Variant
    Dim i As Integer

    On Error GoTo ErrHandler
    '����1,���뷽ʽ1,�п�1|����2,���뷽ʽ2,�п�2|...
    strHead = "ѡ��,4,500|ҵ������,1,900|���ݺ�,1,1000|����,1,1000|�Ա�,1,600|����,1,600|�����,1,1000|סԺ��,1,1000" & _
                    "|�շ�Ա,1,800|�շ�ʱ��,4,2000|���ý��,7,1200" & _
                    "|��Ʊ���,1,1000|��Ʊ˵��,1,5000"
    With vsfExse
        .Redraw = flexRDNone '��ͣ�����ʾˢ��
        .Clear
        .Rows = 2
        .FixedRows = 1: .FixedCols = 0

        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColKey(i) = Split(varData(i), ",")(0)  '����Keyֵ,���ڸ��� ColIndex() ȷ����
            .ColWidth(i) = Split(varData(i), ",")(2)
            If .ColWidth(i) = 0 Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = Split(varData(i), ",")(1)
        Next

        .AllowSelection = False '�������ѡ
        .AllowBigSelection = False '���������̶���/��ѡ������/����
        .SelectionMode = flexSelectionByRow '����ѡ��
        .AllowUserResizing = flexResizeColumns '�����û������п�
        .BackColorSel = &HE0E0E0
        .ForeColorSel = vbBlack
        
        .Editable = flexEDKbdMouse

        .RowHeightMin = 300
        .ColDataType(.ColIndex("ѡ��")) = flexDTBoolean

        .Redraw = flexRDBuffered 'ˢ�±����ʾ
    End With
    InitExseGrid = True
    Exit Function
ErrHandler:
    vsfExse.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub picMain_Resize()
    On Error Resume Next
    picFilter.Move 0, 0, picMain.ScaleWidth
    vsfExse.Move 0, picFilter.Top + picFilter.Height, picMain.ScaleWidth, picMain.ScaleHeight - (picFilter.Top + picFilter.Height)
End Sub

Private Sub vsfExse_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnPrinting Then Exit Sub
    If OldRow = NewRow Then Exit Sub
    
    On Error Resume Next
    vsfExse.ForeColorSel = vsfExse.CellForeColor
End Sub

Private Sub vsfExse_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> vsfExse.ColIndex("ѡ��") Or vsfExse.TextMatrix(Row, 1) = "" Then Cancel = True: Exit Sub
End Sub

Private Sub vsfExse_GotFocus()
    Call SetActiveList(vsfExse)
End Sub

Private Sub vsfExse_LostFocus()
    Call SetActiveList(vsfExse, False)
End Sub

Private Sub SetActiveList(vsfGrid As VSFlexGrid, Optional ByVal blnGetFocus As Boolean = True)
    '���ÿؼ�ѡ���б�������ɫ
    If blnGetFocus Then
        vsfExse.BackColorSel = &HE0E0E0

        If vsfGrid Is Nothing Then Exit Sub
        vsfGrid.BackColorSel = &H8000000D '&HC0C0C0
    Else
        If vsfGrid Is Nothing Then Exit Sub
        vsfGrid.BackColorSel = &HE0E0E0
    End If
End Sub

Private Sub vsfExse_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (Button = vbRightButton) Then Exit Sub
    RaiseEvent ShowPopupMenu(False)
End Sub

Private Sub OutputList(bytStyle As Byte)
'���ܣ�������б�
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte
    Dim intCurrentRow As Integer, vsfGrid As VSFlexGrid
    
    On Error GoTo ErrHandler
    '��ͷ
    Set objOut = New zlPrint1Grd
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    Set vsfGrid = vsfExse
    objOut.Title.Text = "δ���ߵ���Ʊ�ݷ����嵥"
    
    '����
    Set objRow = New zlTabAppRow
    objRow.Add "ҵ�����ͣ�" & cboҵ������.Text
    objRow.Add "�շ�Ա��" & cbo�շ�Ա.Text
    objRow.Add "����ʱ�䣺" & Format(dtp��ʼʱ��, "yyyy-mm-dd") & " �� " & Format(dtp����ʱ��, "yyyy-mm-dd")
    objOut.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    vsfGrid.Redraw = False
    intCurrentRow = vsfGrid.Row
    mblnPrinting = True
    
    '����
    Set objOut.Body = vsfGrid
    '���
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    mblnPrinting = False
    vsfGrid.Row = intCurrentRow
    vsfGrid.Redraw = True
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
    mblnPrinting = False
    vsfGrid.Row = intCurrentRow
    vsfGrid.Redraw = True
End Sub
