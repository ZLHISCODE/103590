VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CC0839AF-B32F-436B-8884-BE2BB3B4C73F}#4.1#0"; "zlIDKind.ocx"
Begin VB.Form frmEInvoicePrint 
   BorderStyle     =   0  'None
   Caption         =   "ֽ��Ʊ�ݹ���"
   ClientHeight    =   10350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13785
   LinkTopic       =   "Form2"
   ScaleHeight     =   10350
   ScaleWidth      =   13785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picMain 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   8748
      Left            =   840
      ScaleHeight     =   8745
      ScaleWidth      =   12780
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Width           =   12780
      Begin VB.Frame fraSplit 
         BorderStyle     =   0  'None
         Height          =   50
         Left            =   1632
         MousePointer    =   7  'Size N S
         TabIndex        =   14
         Top             =   3864
         Width           =   1005
      End
      Begin VB.PictureBox picFilter 
         BorderStyle     =   0  'None
         Height          =   444
         Left            =   24
         ScaleHeight     =   450
         ScaleWidth      =   12690
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   168
         Width           =   12684
         Begin VB.TextBox txtPatient 
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   1248
            MaxLength       =   100
            TabIndex        =   5
            ToolTipText     =   "��λ:F6,����:-����ID,*�����,+סԺ��,.�Һŵ���,����:*2536��ʾ������Ų���"
            Top             =   84
            Width           =   1470
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "ˢ��(&R)"
            Height          =   300
            Left            =   11568
            TabIndex        =   12
            Top             =   84
            Width           =   1000
         End
         Begin VB.ComboBox cboƱ������ 
            Height          =   276
            Left            =   3744
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   96
            Width           =   1812
         End
         Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
            Height          =   276
            Left            =   6648
            TabIndex        =   9
            Top             =   96
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   476
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   73269251
            CurrentDate     =   43941
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   276
            Left            =   9168
            TabIndex        =   11
            Top             =   96
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   476
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   73269251
            CurrentDate     =   43941
         End
         Begin zlIDKind.IDKindNew IDKind 
            Height          =   300
            Left            =   600
            TabIndex        =   4
            Top             =   84
            Width           =   636
            _ExtentX        =   1111
            _ExtentY        =   529
            Appearance      =   2
            IDKindStr       =   $"frmEInvoicePrint.frx":0000
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontSize        =   9
            FontName        =   "����"
            IDKind          =   -1
            ShowPropertySet =   -1  'True
            DefaultCardType =   "0"
            NotContainFastKey=   "F1;CTRL+F1;F2;F3;CTRL+F4;F5;F6;F7;CTRL+F7;F8;F9;F10;F11;F12;CTRL+F12;CTRL+S;CTRL+A;CTRL+R;CTRL+D;CTRL+Q;ESC;ALT+?"
            AllowAutoICCard =   -1  'True
            AllowAutoIDCard =   -1  'True
            MustSelectItems =   "����,���￨"
            BackColor       =   -2147483633
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Left            =   168
            TabIndex        =   3
            Top             =   144
            Width           =   360
         End
         Begin VB.Label lblҵ��ʱ��_ 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   8928
            TabIndex        =   10
            Top             =   144
            Width           =   180
         End
         Begin VB.Label lbl�շ�ʱ�� 
            AutoSize        =   -1  'True
            Caption         =   "�շ�ʱ��"
            Height          =   180
            Left            =   5880
            TabIndex        =   8
            Top             =   150
            Width           =   720
         End
         Begin VB.Label lblƱ������ 
            AutoSize        =   -1  'True
            Caption         =   "Ʊ������"
            Height          =   180
            Left            =   2976
            TabIndex        =   6
            Top             =   144
            Width           =   720
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfEInvoice 
         Height          =   1356
         Left            =   360
         TabIndex        =   13
         Top             =   1392
         Width           =   6108
         _cx             =   1983064598
         _cy             =   1983056216
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
      Begin VSFlex8Ctl.VSFlexGrid vsfExse 
         Height          =   1404
         Left            =   720
         TabIndex        =   15
         Top             =   4656
         Width           =   4404
         _cx             =   1983061592
         _cy             =   1983056300
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
      Height          =   1035
      Left            =   0
      Top             =   900
      Width           =   525
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
      Caption         =   "����Ʊ�ݴ�ӡ"
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
Attribute VB_Name = "frmEInvoicePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form, mlngSys As Long, mlngModule As Long, mstrDBUser As String, mstrEInvPrivs As String
Private mcbsMain   As Object          'CommandBar�ؼ�
Private mobjEInvoice As clsEInvoiceModule
Private mobjPubEInvoice As Object ' zlPublicExpense.clsPubEInvoice
Private mblnPrinting As Boolean
Attribute mblnPrinting.VB_VarHelpID = -1
Private mrsInfo As ADODB.Recordset
Private mobjSquareCard As Object
Private mcllResult As Collection

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
        With cbrMenuBar.CommandBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&R)", cbrControl.index + 1): cbrControl.BeginGroup = True
        End With
    End If

    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", cbrMenuBar.index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PrintEInvoice, "��ӡ���ӷ�Ʊ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SendMsg, "��Ϣ����(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PrintNotice, "��ӡ��֪��(&N)"): cbrControl.IconId = 103
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_TurnPaper, "����ֽ��Ʊ��(&T)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ReTurnPaper, "���»���Ʊ��(&R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CancelTurnPaper, "����ֽ��Ʊ��(&C)")
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
            Set cbrControl = cbrToolBar.Controls(cbrControl.index - 1): Exit For
        End If
    Next
    With cbrToolBar.Controls
    Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PrintEInvoice, "��ӡ���ӷ�Ʊ", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SendMsg, "��Ϣ����", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PrintNotice, "��ӡ��֪��", cbrControl.index + 1): cbrControl.IconId = 103
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SelAll, "ȫѡ", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClsAll, "ȫ��", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��", cbrControl.index + 1): cbrControl.BeginGroup = True
    End With

    '����Ŀ����
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_SelAll
        .Add FCONTROL, Asc("C"), conMenu_Edit_ClsAll
    End With
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim frmParaSet As frmEInvoiceParaSet
    
    Select Case Control.ID
    Case conMenu_File_Preview 'Ԥ��
        Call OutputList(2)
    Case conMenu_File_Print '��ӡ
        Call OutputList(1)
    Case conMenu_File_Excel '�����Excel��
        Call OutputList(3)
    
    Case conMenu_File_Parameter '��������
        Set frmParaSet = New frmEInvoiceParaSet
        Call frmParaSet.ShowMe(Me, mlngSys, 1145)
        
    Case conMenu_Edit_PrintEInvoice '��ӡ���ӷ�Ʊ
        Call ExcutePrintEInvoice(zlStr.NeedCode(cboƱ������.Text))
    Case conMenu_Edit_PrintNotice '��ӡ��֪��
        Call ExecutePrintNotice(zlStr.NeedCode(cboƱ������.Text))
    Case conMenu_Edit_SendMsg '��Ϣ����
        Call ExecuteSendMsg(zlStr.NeedCode(cboƱ������.Text))
    Case conMenu_Edit_TurnPaper '����ֽ��Ʊ��
        Call ExcuteTurnPaper(zlStr.NeedCode(cboƱ������.Text), False)
    Case conMenu_Edit_ReTurnPaper '���»���Ʊ��
        Call ExcuteTurnPaper(zlStr.NeedCode(cboƱ������.Text), True)
    Case conMenu_Edit_CancelTurnPaper '����ֽ��Ʊ��
        Call ExcuteCancelTurnPaper(zlStr.NeedCode(cboƱ������.Text))
        
    Case conMenu_Edit_SelAll 'ȫѡ
        Call Grid_SelAllRecord(vsfEInvoice, True)
    Case conMenu_Edit_ClsAll 'ȫ��
        Call Grid_SelAllRecord(vsfEInvoice, False)
        
    Case conMenu_View_Refresh 'ˢ��
        Call cmdRefresh_Click
    End Select
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    If Not Me.Visible Then Exit Sub
    On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel 'Ԥ��,��ӡ,�����Excel��
        If Me.ActiveControl Is vsfExse Then
            Control.Enabled = vsfExse.TextMatrix(1, 1) <> ""
        Else
            Control.Enabled = vsfEInvoice.TextMatrix(1, 1) <> ""
        End If
    
    Case conMenu_Edit_PrintEInvoice '��ӡ���ӷ�Ʊ
        Control.Enabled = Control.Visible And vsfEInvoice.TextMatrix(1, 1) <> ""
    Case conMenu_Edit_SendMsg '��Ϣ����
        Control.Enabled = Control.Visible And vsfEInvoice.TextMatrix(1, 1) <> ""
    Case conMenu_Edit_PrintNotice '��ӡ��֪��
        Control.Enabled = Control.Visible And vsfEInvoice.TextMatrix(1, 1) <> ""
    
    Case conMenu_Edit_TurnPaper '����ֽ��Ʊ��
        Control.Visible = zlStr.IsHavePrivs(mstrEInvPrivs, "����ֽ��Ʊ��")
        Control.Enabled = Control.Visible And vsfEInvoice.TextMatrix(1, 1) <> ""
    Case conMenu_Edit_ReTurnPaper '���»���Ʊ��
        Control.Visible = zlStr.IsHavePrivs(mstrEInvPrivs, "���»���Ʊ��")
        Control.Enabled = Control.Visible And vsfEInvoice.TextMatrix(1, 1) <> ""
    Case conMenu_Edit_CancelTurnPaper '����ֽ��Ʊ��
        Control.Visible = zlStr.IsHavePrivs(mstrEInvPrivs, "����ֽ��Ʊ��")
        Control.Enabled = Control.Visible And vsfEInvoice.TextMatrix(1, 1) <> ""
    
    Case conMenu_Edit_SelAll 'ȫѡ
        Control.Enabled = Control.Visible And vsfEInvoice.TextMatrix(1, 1) <> ""
    Case conMenu_Edit_ClsAll 'ȫ��
        Control.Enabled = Control.Visible And vsfEInvoice.TextMatrix(1, 1) <> ""
    End Select
End Sub

Private Sub cboƱ������_Click()
    Dim byt���� As Byte
    
    If cboƱ������.Tag = cboƱ������.Text Then Exit Sub
    cboƱ������.Tag = cboƱ������.Text
    
    byt���� = zlStr.NeedCode(cboƱ������.Text)
    vsfExse.Visible = byt���� <> 2
    fraSplit.Visible = byt���� <> 2
    vsfExse.Tag = IIf(byt���� = 2, "ExseGridHidden", "")
    Call picMain_Resize
    Call cmdRefresh_Click
End Sub

Private Sub cmdRefresh_Click()
    Set mcllResult = Nothing
    Call LoadEInvoiceData
End Sub

Private Function LoadEInvoiceData() As Boolean
    '��ʾ����Ʊ������
    Dim dtBegin As Date, dtEnd As Date
    
    '1.���ݼ��
    On Error GoTo ErrHandler
    dtBegin = dtp��ʼʱ��.Value: dtEnd = dtp����ʱ��.Value
    If dtp��ʼʱ�� > dtp����ʱ�� Then
        MsgBox "���õĿ�ʼʱ�䲻�ܴ��ڽ���ʱ�䣡", vbInformation, gstrSysName
        zlControl.ControlSetFocus dtp����ʱ��:  Exit Function
    End If
    
    If DateDiff("m", dtp��ʼʱ��, dtp����ʱ��) > 6 Then
        If MsgBox("�Ե�ǰ����ʱ�䷶Χ�ڵ����ݽ��в�ѯ������Ҫ�ϳ�ʱ�䣬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    
    '2.��ȡ����
    Dim rsEInvoice As ADODB.Recordset, bytQueryType As Byte, varQueryValue As Variant
    'bytQueryType ��ѯ���ͣ�0-���У�1-������ID��ѯ��2-�����õ��ݺŲ�ѯ��3-������Ʊ�ݺŲ�
    Select Case IDKind.GetCurCard.����
    Case "�շѵ��ݺ�"
        If Trim(txtPatient.Text) <> "" Then
            bytQueryType = 2
            varQueryValue = Trim(txtPatient.Text)
        End If
    Case "����Ʊ�ݺ�"
        If Trim(txtPatient.Text) <> "" Then
            bytQueryType = 3
            varQueryValue = Trim(txtPatient.Text)
        End If
    Case Else
        If Val(txtPatient.Tag) <> 0 Then
            bytQueryType = 1
            varQueryValue = Val(txtPatient.Tag)
        End If
    End Select
    If GetEInvoiceData(zlStr.NeedCode(cboƱ������.Text), dtp��ʼʱ��.Value, dtp����ʱ��.Value, rsEInvoice, 3, 1, bytQueryType, varQueryValue) = False Then Exit Function
    
    Dim lngOldRow As Long, lngOldCol As Long
    lngOldRow = vsfEInvoice.Row: lngOldCol = vsfEInvoice.Col
    vsfEInvoice.Clear 1
    vsfEInvoice.Rows = 1 '���Dataֵ
    vsfEInvoice.Rows = vsfEInvoice.FixedRows + 1
    
    vsfExse.Clear 1
    vsfExse.Rows = 1 '
    vsfExse.Rows = vsfExse.FixedRows + 1
    
    With vsfEInvoice
        .Redraw = flexRDNone
        'ѡ��,NO,����,�Ա�,����,�����,סԺ��,Ʊ������,Ʊ�ݴ���,Ʊ�ݺ���,������,Ʊ�ݽ��,��Ʊ��,��Ʊʱ��,����ֽ�ʷ�Ʊ,ֽ�ʷ�Ʊ��
        Do While Not rsEInvoice.EOF
            If .TextMatrix(.Rows - 1, .ColIndex("���ݺ�")) <> "" Then .Rows = .Rows + 1
            .RowData(.Rows - 1) = Val(Nvl(rsEInvoice!����ID))
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = Val(Nvl(rsEInvoice!ID))
            .TextMatrix(.Rows - 1, .ColIndex("���ݺ�")) = Nvl(rsEInvoice!No)
            .TextMatrix(.Rows - 1, .ColIndex("����")) = Nvl(rsEInvoice!����)
            .TextMatrix(.Rows - 1, .ColIndex("�Ա�")) = Nvl(rsEInvoice!�Ա�)
            .TextMatrix(.Rows - 1, .ColIndex("����")) = Nvl(rsEInvoice!����)
            .TextMatrix(.Rows - 1, .ColIndex("�����")) = Nvl(rsEInvoice!�����)
            .TextMatrix(.Rows - 1, .ColIndex("סԺ��")) = Nvl(rsEInvoice!סԺ��)
            
            .TextMatrix(.Rows - 1, .ColIndex("Ʊ������")) = Decode(Val(Nvl(rsEInvoice!Ʊ��)), 1, "�շ�", 2, "Ԥ��", 3, "����", 4, "�Һ�", 5, "���￨")
            .Cell(flexcpData, .Rows - 1, .ColIndex("Ʊ������")) = Val(Nvl(rsEInvoice!Ʊ��))
            .TextMatrix(.Rows - 1, .ColIndex("Ʊ�ݴ���")) = Nvl(rsEInvoice!Ʊ�ݴ���)
            .TextMatrix(.Rows - 1, .ColIndex("Ʊ�ݺ���")) = Nvl(rsEInvoice!Ʊ�ݺ���)
            .TextMatrix(.Rows - 1, .ColIndex("������")) = Nvl(rsEInvoice!������)
            .TextMatrix(.Rows - 1, .ColIndex("Ʊ�ݽ��")) = FormatEx(Val(Nvl(rsEInvoice!Ʊ�ݽ��)), 2, , , 6)
            .TextMatrix(.Rows - 1, .ColIndex("��Ʊ��")) = Nvl(rsEInvoice!��Ʊ��)
            .TextMatrix(.Rows - 1, .ColIndex("��Ʊʱ��")) = Format(Nvl(rsEInvoice!��Ʊʱ��), "yyyy-MM-dd HH:mm:ss")
            .TextMatrix(.Rows - 1, .ColIndex("����ֽ�ʷ�Ʊ")) = IIf(Val(Nvl(rsEInvoice!�Ƿ񻻿�)) = 1, "��", "")
            .TextMatrix(.Rows - 1, .ColIndex("ֽ�ʷ�Ʊ��")) = Nvl(rsEInvoice!ֽ�ʷ�Ʊ��)
            
            Select Case Val(Nvl(rsEInvoice!Ʊ��))
            Case 2
                .Cell(flexcpData, .Rows - 1, .ColIndex("���ݺ�")) = Val(Nvl(rsEInvoice!�˿�ID))
            Case 1, 4
                .Cell(flexcpData, .Rows - 1, .ColIndex("���ݺ�")) = Val(Nvl(rsEInvoice!������))
            End Select
            
            If Not mcllResult Is Nothing Then
                If CollectionExitsValue(mcllResult, "_" & Val(Nvl(rsEInvoice!ID))) Then
                    .TextMatrix(.Rows - 1, .ColIndex("��ӡ���")) = mcllResult("_" & Val(Nvl(rsEInvoice!ID)))(0)
                    .TextMatrix(.Rows - 1, .ColIndex("��ӡ˵��")) = mcllResult("_" & Val(Nvl(rsEInvoice!ID)))(1)
                    If mcllResult("_" & Val(Nvl(rsEInvoice!ID)))(0) Like "*ʧ��*" Then
                        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
                    End If
                End If
            End If
        
            rsEInvoice.MoveNext
        Loop
        
        .Cell(flexcpFontBold, .FixedRows, .ColIndex("��ӡ���"), .Rows - 1, .ColIndex("��ӡ˵��")) = True
        
        If .Rows > .FixedRows And .Cols > .FixedCols Then     'ȱʡ��λ��
            .Row = -1 '��֤��ѡ���в���������Ҳ����RowColChange�¼�
            .Row = IIf(lngOldRow < .FixedRows Or lngOldRow > .Rows - 1, IIf(.Rows - 1 > .FixedRows, .FixedRows + 1, .FixedRows), lngOldRow)
            .Col = IIf(lngOldCol = 0 Or lngOldCol > .Cols - 1, .FixedCols, lngOldCol)
            .ShowCell .Row, .Col  '������ʾ��ָ����Ԫ
        End If
        
        .Redraw = flexRDBuffered
    End With
    LoadEInvoiceData = True
    Exit Function
ErrHandler:
    vsfExse.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Load()
    Dim strIDKindStr As String
    
    Call CreateSquareCardObject(Me, mlngModule)
    strIDKindStr = "��|��������￨|0;ҽ|ҽ����|0;��|���֤��|0;IC|IC����|1;��|�����|0;ס|סԺ��|0;��|�ֻ���|0;��|�շѵ��ݺ�|0;Ʊ|����Ʊ�ݺ�|0"
    Call IDKind.zlInit(Me, mlngSys, mlngModule, gcnOracle, mstrDBUser, mobjSquareCard, strIDKindStr, txtPatient)
    
    Call InitEInvoiceGrid
    Call InitExseGrid
    
    Dim varData As Variant, i As Integer
    varData = Array("1-�շ�", "2-Ԥ��", "3-����", "4-�Һ�", "5-���￨")
    cboƱ������.Clear
    For i = 0 To UBound(varData)
        cboƱ������.AddItem varData(i)
    Next
    cboƱ������.ListIndex = 0
    
    dtp����ʱ��.Value = zlDatabase.Currentdate
    dtp��ʼʱ��.Value = Format(DateAdd("d", -7, dtp����ʱ��.Value), "yyyy-MM-dd 00:00:00")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    sccTitle.Move 8, 8, shpBorder.Width - 20
    picMain.Move sccTitle.Left, sccTitle.Top + sccTitle.Height, Me.ScaleWidth - 2 * sccTitle.Left, Me.ScaleHeight - (2 * sccTitle.Top + sccTitle.Height)
End Sub

Private Sub fraSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    If Button <> vbLeftButton Then Exit Sub
    If vsfEInvoice.Height + Y < 1200 Or vsfExse.Height - Y < 1200 Then Exit Sub

    fraSplit.Top = fraSplit.Top + Y
    
    vsfEInvoice.Height = vsfEInvoice.Height + Y
    vsfExse.Top = vsfExse.Top + Y
    vsfExse.Height = vsfExse.Height - Y
    Me.Refresh
End Sub

Private Sub picMain_Resize()
    On Error Resume Next
    picFilter.Move 0, 0, picMain.ScaleWidth
    If vsfExse.Tag = "ExseGridHidden" Then
        vsfEInvoice.Move 0, picFilter.Top + picFilter.Height, picMain.ScaleWidth, picMain.ScaleHeight - (picFilter.Top + picFilter.Height)
    Else
        vsfEInvoice.Move 0, picFilter.Top + picFilter.Height, picMain.ScaleWidth, picMain.ScaleHeight * 2 / 3
        fraSplit.Move 0, vsfEInvoice.Top + vsfEInvoice.Height, picMain.ScaleWidth
        vsfExse.Move 0, fraSplit.Top + fraSplit.Height, picMain.ScaleWidth, picMain.ScaleHeight - (fraSplit.Top + fraSplit.Height)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmMain = Nothing
    Set mcbsMain = Nothing
    Set mobjEInvoice = Nothing
End Sub

Private Function InitEInvoiceGrid() As Boolean
    '��ʼ��VSFGrid���ؼ�
    Dim strHead As String, varData As Variant
    Dim i As Integer

    On Error GoTo ErrHandler
    '����1,���뷽ʽ1,�п�1|����2,���뷽ʽ2,�п�2|...
    strHead = "ѡ��,4,500|ID,1,0|���ݺ�,1,1000|����,1,1000|�Ա�,1,1000|����,1,1000|�����,1,1000|סԺ��,1,1000" & _
                    "|Ʊ������,1,1000|Ʊ�ݴ���,1,2000|Ʊ�ݺ���,1,2000|������,1,2000|Ʊ�ݽ��,7,1000" & _
                    "|��Ʊ��,1,1000|��Ʊʱ��,4,2000|����ֽ�ʷ�Ʊ,3,2000|ֽ�ʷ�Ʊ��,1,2000" & _
                    "|��ӡ���,1,1000|��ӡ˵��,1,5000"
    With vsfEInvoice
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
    InitEInvoiceGrid = True
    Exit Function
ErrHandler:
    vsfEInvoice.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitExseGrid() As Boolean
    '��ʼ��VSFGrid���ؼ�
    Dim strHead As String, varData As Variant
    Dim i As Integer

    On Error GoTo ErrHandler
    '����1,���뷽ʽ1,�п�1|����2,���뷽ʽ2,�п�2|...
    strHead = "���ݺ�,1,1000|��������,1,1000|������,1,800|�ѱ�,1,500|���,4,800|����,1,3000|��Ʒ��,1,3000" & _
                    "|���,1,1200|��λ,4,1000|����,7,800|����,7,1000|Ӧ�ս��,7,1000|ʵ�ս��,7,1000|ִ�п���,4,1500"

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

        .RowHeightMin = 300

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

Private Sub vsfEInvoice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnPrinting Then Exit Sub
    If OldRow = NewRow Or NewRow < vsfEInvoice.FixedRows Then Exit Sub
    
    Call ShowEInvoiceExse(zlStr.NeedCode(cboƱ������.Text), Val(vsfEInvoice.RowData(NewRow)))
    
    On Error Resume Next
    vsfEInvoice.ForeColorSel = vsfEInvoice.CellForeColor
End Sub

Private Function ShowEInvoiceExse(ByVal byt���� As Byte, ByVal lngEInvoice As Long) As Boolean
    '��ʾ������ϸ
    '��Σ�
    '   byt���� 1-�շѣ�2-Ԥ����3-���ʣ�4-�Һţ�5-���￨
    Dim rsExse As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If byt���� = 2 Then ShowEInvoiceExse = True: Exit Function
    
    With vsfExse
        .Clear 1
        .Rows = .FixedRows + 1
        
        If GetEInvoiceExse(byt����, lngEInvoice, rsExse) = False Then Exit Function
        
        'NO,��������,������,�ѱ�,���,����,��Ʒ��,���,��λ,����,����,Ӧ�ս��,ʵ�ս��,ִ�п���
        .Redraw = flexRDNone
        Do While Not rsExse.EOF
            If .TextMatrix(.Rows - 1, .ColIndex("���ݺ�")) <> "" Then .Rows = .Rows + 1
            .RowData(.Rows - 1) = Val(Nvl(rsExse!���))
            .TextMatrix(.Rows - 1, .ColIndex("���ݺ�")) = Nvl(rsExse!No)
            .TextMatrix(.Rows - 1, .ColIndex("��������")) = Nvl(rsExse!��������)
            .TextMatrix(.Rows - 1, .ColIndex("������")) = Nvl(rsExse!������)
            .TextMatrix(.Rows - 1, .ColIndex("�ѱ�")) = Nvl(rsExse!�ѱ�)
            .TextMatrix(.Rows - 1, .ColIndex("���")) = Nvl(rsExse!���)
            .TextMatrix(.Rows - 1, .ColIndex("����")) = Nvl(rsExse!����)
            .TextMatrix(.Rows - 1, .ColIndex("��Ʒ��")) = Nvl(rsExse!��Ʒ��)
            .TextMatrix(.Rows - 1, .ColIndex("���")) = Nvl(rsExse!���)
            .TextMatrix(.Rows - 1, .ColIndex("��λ")) = Nvl(rsExse!��λ)
            .TextMatrix(.Rows - 1, .ColIndex("����")) = Nvl(rsExse!����)
            .TextMatrix(.Rows - 1, .ColIndex("����")) = FormatEx(Val(Nvl(rsExse!����)), 2, , , 6)
            .TextMatrix(.Rows - 1, .ColIndex("Ӧ�ս��")) = FormatEx(Val(Nvl(rsExse!Ӧ�ս��)), 2, , , 6)
            .TextMatrix(.Rows - 1, .ColIndex("ʵ�ս��")) = FormatEx(Val(Nvl(rsExse!ʵ�ս��)), 2, , , 6)
            .TextMatrix(.Rows - 1, .ColIndex("ִ�п���")) = Nvl(rsExse!ִ�п���)
        
            rsExse.MoveNext
        Loop
        
        .Redraw = flexRDBuffered
    End With
    ShowEInvoiceExse = True
    Exit Function
ErrHandler:
    vsfExse.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfEInvoice_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> vsfEInvoice.ColIndex("ѡ��") Or vsfEInvoice.TextMatrix(Row, 1) = "" Then Cancel = True: Exit Sub
End Sub

Private Sub vsfEInvoice_GotFocus()
    Call SetActiveList(vsfEInvoice)
End Sub

Private Sub vsfEInvoice_LostFocus()
    Call SetActiveList(vsfEInvoice, False)
End Sub

Private Sub vsfEInvoice_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Not (Me.ActiveControl Is vsfEInvoice And Button = vbRightButton) Then Exit Sub
    RaiseEvent ShowPopupMenu(False)
End Sub

Private Sub vsfExse_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnPrinting Then Exit Sub
    If OldRow = NewRow Then Exit Sub
    
    On Error Resume Next
    vsfExse.ForeColorSel = vsfExse.CellForeColor
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
        vsfEInvoice.BackColorSel = &HE0E0E0
        vsfExse.BackColorSel = &HE0E0E0

        If vsfGrid Is Nothing Then Exit Sub
        vsfGrid.BackColorSel = &H8000000D '&HC0C0C0
    Else
        If vsfGrid Is Nothing Then Exit Sub
        vsfGrid.BackColorSel = &HE0E0E0
    End If
End Sub

Private Sub ExcutePrintEInvoice(ByVal byt���� As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡ����Ʊ��(A4ֽ)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllSwapData As Collection, strErrMsg As String
    Dim lngEInvoiceID As Long, strDate As String, bln������ As Boolean
    Dim i As Long, lng����ID As Long, blnInit As Boolean
    Dim lngCount As Long, blnChecked As Boolean
    
    On Error GoTo ErrHandler
    If GetPubEInvoiceObject(Me, mlngSys, mlngModule, mobjPubEInvoice, byt����) = False Then Exit Sub
    With vsfEInvoice
        blnChecked = False
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = vbChecked Then blnChecked = True: Exit For
        Next
        
        .Cell(flexcpText, 0, .ColIndex("��ӡ���"), 1, .ColIndex("��ӡ˵��")) = "��ӡ���" & vbTab & "��ӡ˵��"
        .Cell(flexcpText, .FixedRows, .ColIndex("��ӡ���"), .Rows - 1, .ColIndex("��ӡ˵��")) = ""
        .Cell(flexcpForeColor, .FixedRows, 0, .Rows - 1, .Cols - 1) = vbBlack
        
        For i = .FixedRows To .Rows - 1
            lngEInvoiceID = Val(.TextMatrix(i, .ColIndex("ID")))
            If lngEInvoiceID <> 0 And (.Cell(flexcpChecked, i, .ColIndex("ѡ��")) = vbChecked Or Not blnChecked And i = .Row) Then
                lngCount = lngCount + 1
                
                If mobjPubEInvoice.zlPrintEInvoice(Me, lngEInvoiceID, False, strErrMsg) Then
                    .TextMatrix(i, .ColIndex("��ӡ���")) = "��ӡ�ɹ�"
                    .TextMatrix(i, .ColIndex("��ӡ˵��")) = ""
                Else
                    .TextMatrix(i, .ColIndex("��ӡ���")) = "��ӡʧ��"
                    .TextMatrix(i, .ColIndex("��ӡ˵��")) = strErrMsg
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                End If
            
                If Not blnChecked Then Exit For
            End If
        Next
    End With
    
    If lngCount = 0 Then
        MsgBox "��ѡ����Ҫ��ӡ����Ʊ�ݵļ�¼��", vbInformation, gstrSysName
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ExcuteTurnPaper(ByVal byt���� As Byte, ByVal bln���»��� As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ��ֽ�ʷ�Ʊ���������»���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngEInvoiceID As Long, strEInvoiceCode As String, strEInvoiceNO As String, strInvoiceNO_Out As String
    Dim strNO As String, lngԭ����ID As Long, blnHaveEInvoice As Boolean
    Dim rsTemp As ADODB.Recordset, i As Long
    Dim cllSwapData As Collection, int����״̬ As Integer, strUseDate As String
    Dim bln������ As Boolean, lng����ID As Long, strErrMsg As String
    Dim lngCount As Long, blnChecked As Boolean, blnFirst As Boolean
    Dim cllPati As Collection, cllBalance As Collection
    
    On Error GoTo ErrHandler
    Set mcllResult = New Collection
    If GetPubEInvoiceObject(Me, mlngSys, mlngModule, mobjPubEInvoice, byt����) = False Then Exit Sub
    With vsfEInvoice
        blnChecked = False
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = vbChecked Then blnChecked = True: Exit For
        Next
        
        .Cell(flexcpText, 0, .ColIndex("��ӡ���"), 1, .ColIndex("��ӡ˵��")) = "�������" & vbTab & "����˵��"
        .Cell(flexcpText, .FixedRows, .ColIndex("��ӡ���"), .Rows - 1, .ColIndex("��ӡ˵��")) = ""
        
        blnFirst = True
        For i = .FixedRows To .Rows - 1
            lngEInvoiceID = Val(.TextMatrix(i, .ColIndex("ID")))
            If lngEInvoiceID <> 0 And (.Cell(flexcpChecked, i, .ColIndex("ѡ��")) = vbChecked Or Not blnChecked And i = .Row) Then
                lngCount = lngCount + 1
                
                lngԭ����ID = Val(.RowData(i))
                strEInvoiceCode = .TextMatrix(i, .ColIndex("Ʊ�ݴ���"))
                strEInvoiceNO = .TextMatrix(i, .ColIndex("Ʊ�ݺ���"))
                
                lng����ID = 0: bln������ = False
                If byt���� = 2 Then
                    lng����ID = Val(Nvl(.Cell(flexcpData, i, .ColIndex("���ݺ�"))))
                ElseIf byt���� = 1 Or byt���� = 4 Then
                    bln������ = Val(Nvl(.Cell(flexcpData, i, .ColIndex("���ݺ�")))) = 1
                End If
            
                If .TextMatrix(i, .ColIndex("����ֽ�ʷ�Ʊ")) = "" And bln���»��� Then
                    strErrMsg = "��ֽ��Ʊ�ݻ�����¼����ִ��[����ֽ��Ʊ��]��"
                    mcllResult.Add Array("����ʧ��", strErrMsg), "_" & lngEInvoiceID
                ElseIf .TextMatrix(i, .ColIndex("����ֽ�ʷ�Ʊ")) <> "" And Not bln���»��� Then
                    strErrMsg = "�ѻ���ֽ��Ʊ�ݣ���ִ��[���»���Ʊ��]��"
                    mcllResult.Add Array("����ʧ��", strErrMsg), "_" & lngEInvoiceID
                        
                ElseIf GetSwapCollectFromBalanceID(byt����, lngԭ����ID, cllSwapData, bln������, lng����ID, False, strErrMsg) = False Then
                    mcllResult.Add Array("����ʧ��", strErrMsg), "_" & lngEInvoiceID
                Else
                    Set cllPati = cllSwapData("_PatiInfo")
                    strInvoiceNO_Out = GetNextPaperInvoice(Me, cllPati, 0, byt����, blnFirst)
                    If strInvoiceNO_Out = "" Then
                        strErrMsg = "��ȡ��һ����ЧƱ�ݺ�ʧ��"
                        mcllResult.Add Array("����ʧ��", strErrMsg), "_" & lngEInvoiceID
                        If blnFirst Then Call LoadEInvoiceData: Exit Sub
                        Exit For
                    End If
                    blnFirst = False
                    
                    Set cllBalance = cllSwapData("_BalanceInfo")
                    cllBalance.Remove "_��Ʊ��"
                    cllBalance.Add strInvoiceNO_Out, "_��Ʊ��"
                    
                    If mobjPubEInvoice.zlTurnPaperInvoice(Me, byt����, cllSwapData, lngEInvoiceID, _
                        strEInvoiceCode, strEInvoiceNO, strInvoiceNO_Out, int����״̬, , False, strErrMsg) Then
                        mcllResult.Add Array("�����ɹ�", IIf(bln���»���, "����Ʊ�ݺţ�" & .TextMatrix(i, .ColIndex("ֽ�ʷ�Ʊ��")), "")), "_" & lngEInvoiceID
                    Else
                        mcllResult.Add Array("����ʧ��", strErrMsg), "_" & lngEInvoiceID
                    End If
                End If
            
                If Not blnChecked Then Exit For
            End If
        Next
    End With
    
    If lngCount = 0 Then
        MsgBox "��ѡ����Ҫ" & IIf(bln���»���, "����", "") & "����ֽ�ʷ�Ʊ�ļ�¼��", vbInformation, gstrSysName
    End If
    
    Call LoadEInvoiceData
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetNextPaperInvoice(ByVal frmMain As Object, ByVal cllPatiInfo As Collection, _
    ByRef lng����ID As Long, ByVal byt���� As Byte, ByVal blnFirst As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡһ���ŷ�Ʊ��
    '���:
    '   byt���ϣ�1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
    '����:��Ʊ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceNO As String, strErrMsg_Out As String
     
    '����Ʊ�����ö�ȡ
    On Error GoTo ErrHandler
    If mobjPubEInvoice.zlGetNextInvoiceNo(frmMain, byt����, strInvoiceNO, cllPatiInfo, lng����ID, False, strErrMsg_Out) = False Then Exit Function
    
    If strInvoiceNO = "" Then
        If frmInputBox.InputBox(frmMain, "��Ʊ��ȷ��", "�޷���ȡ��Ҫʹ�õķ�Ʊ�ţ�" & _
                        vbCrLf & "�������뻻����Ҫʹ�õķ�Ʊ���룺", 30, 1, False, False, strInvoiceNO) = False Then Exit Function
    ElseIf blnFirst Then
        If frmInputBox.InputBox(frmMain, "��Ʊ��ȷ��", "��ȷ�ϻ�����Ҫʹ�õķ�Ʊ�ţ�", 30, 1, False, False, strInvoiceNO) = False Then Exit Function
    End If
    GetNextPaperInvoice = strInvoiceNO
    Exit Function
ErrHandler:
    Err.Clear
End Function

Private Sub ExcuteCancelTurnPaper(ByVal byt���� As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ��ֽ�ʷ�Ʊ���ϲ���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceNO As String, blnHaveEInvoice As Boolean
    Dim strNO As String, lngԭ����ID As Long, i As Long
    Dim rsTemp As ADODB.Recordset, lngEInvoiceID As Long
    Dim bln������ As Boolean, strErrMsg As String
    Dim lngCount As Long, blnChecked As Boolean
     
    On Error GoTo ErrHandler
    Set mcllResult = New Collection
    If GetPubEInvoiceObject(Me, mlngSys, mlngModule, mobjPubEInvoice, byt����) = False Then Exit Sub
    With vsfEInvoice
        blnChecked = False
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = vbChecked Then blnChecked = True: Exit For
        Next
        
        .Cell(flexcpText, 0, .ColIndex("��ӡ���"), 1, .ColIndex("��ӡ˵��")) = "���Ͻ��" & vbTab & "����˵��"
        .Cell(flexcpText, .FixedRows, .ColIndex("��ӡ���"), .Rows - 1, .ColIndex("��ӡ˵��")) = ""
        
        For i = .FixedRows To .Rows - 1
            lngEInvoiceID = Val(.TextMatrix(i, .ColIndex("ID")))
            If lngEInvoiceID <> 0 And (.Cell(flexcpChecked, i, .ColIndex("ѡ��")) = vbChecked Or Not blnChecked And i = .Row) Then
                lngCount = lngCount + 1
                
                lngԭ����ID = Val(.RowData(i))
                If .TextMatrix(i, .ColIndex("����ֽ�ʷ�Ʊ")) = "" Then
                    strErrMsg = "��ֽ��Ʊ�ݻ�����¼��"
                    mcllResult.Add Array("����ʧ��", strErrMsg), "_" & lngEInvoiceID
                Else
                    bln������ = False
                    If byt���� = 1 Or byt���� = 4 Then
                        bln������ = Val(Nvl(.Cell(flexcpData, i, .ColIndex("���ݺ�")))) = 1
                    End If
                    
                    strInvoiceNO = .TextMatrix(i, .ColIndex("ֽ�ʷ�Ʊ��"))
                    If mobjPubEInvoice.zlCancelPaperInvoice(Me, byt����, strInvoiceNO, lngԭ����ID, lngEInvoiceID, , , , bln������, , False, strErrMsg) Then
                        mcllResult.Add Array("���ϳɹ�", ""), "_" & lngEInvoiceID
                    Else
                        mcllResult.Add Array("����ʧ��", strErrMsg), "_" & lngEInvoiceID
                    End If
                End If
            
                If Not blnChecked Then Exit For
            End If
        Next
    End With
    
    If lngCount = 0 Then
        MsgBox "��ѡ����Ҫ����ֽ�ʷ�Ʊ�ļ�¼��", vbInformation, gstrSysName
    End If
    
    Call LoadEInvoiceData
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ExecutePrintNotice(ByVal byt���� As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡ��֪��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllSwapData As Collection, strErrMsg As String
    Dim lngEInvoiceID As Long, strDate As String, bln������ As Boolean
    Dim i As Long, lng����ID As Long, blnInit As Boolean
    Dim lngCount As Long, blnChecked As Boolean
    
    On Error GoTo ErrHandler
    If GetPubEInvoiceObject(Me, mlngSys, mlngModule, mobjPubEInvoice, byt����) = False Then Exit Sub
    With vsfEInvoice
        blnChecked = False
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = vbChecked Then blnChecked = True: Exit For
        Next
        
        .Cell(flexcpText, 0, .ColIndex("��ӡ���"), 1, .ColIndex("��ӡ˵��")) = "��ӡ���" & vbTab & "��ӡ˵��"
        .Cell(flexcpText, .FixedRows, .ColIndex("��ӡ���"), .Rows - 1, .ColIndex("��ӡ˵��")) = ""
        .Cell(flexcpForeColor, .FixedRows, 0, .Rows - 1, .Cols - 1) = vbBlack
        
        For i = .FixedRows To .Rows - 1
            lngEInvoiceID = Val(.TextMatrix(i, .ColIndex("ID")))
            If lngEInvoiceID <> 0 And (.Cell(flexcpChecked, i, .ColIndex("ѡ��")) = vbChecked Or Not blnChecked And i = .Row) Then
                lngCount = lngCount + 1
                
                If mobjPubEInvoice.zlPrintNotice(Me, byt����, lngEInvoiceID, False, strErrMsg) Then
                    .TextMatrix(i, .ColIndex("��ӡ���")) = "��ӡ�ɹ�"
                    .TextMatrix(i, .ColIndex("��ӡ˵��")) = ""
                Else
                    .TextMatrix(i, .ColIndex("��ӡ���")) = "��ӡʧ��"
                    .TextMatrix(i, .ColIndex("��ӡ˵��")) = strErrMsg
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                End If
            
                If Not blnChecked Then Exit For
            End If
        Next
    End With
    
    If lngCount = 0 Then
        MsgBox "��ѡ����Ҫ��ӡ��֪���ļ�¼��", vbInformation, gstrSysName
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ExecuteSendMsg(ByVal byt���� As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ϣ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllSwapData As Collection, strErrMsg As String
    Dim lngEInvoiceID As Long, strDate As String, bln������ As Boolean
    Dim i As Long, lng����ID As Long, blnInit As Boolean
    Dim lngCount As Long, blnChecked As Boolean
    
    On Error GoTo ErrHandler
    If GetPubEInvoiceObject(Me, mlngSys, mlngModule, mobjPubEInvoice, byt����) = False Then Exit Sub
    With vsfEInvoice
        blnChecked = False
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = vbChecked Then blnChecked = True: Exit For
        Next
        
        .Cell(flexcpText, 0, .ColIndex("��ӡ���"), 1, .ColIndex("��ӡ˵��")) = "���ͽ��" & vbTab & "����˵��"
        .Cell(flexcpText, .FixedRows, .ColIndex("��ӡ���"), .Rows - 1, .ColIndex("��ӡ˵��")) = ""
        .Cell(flexcpForeColor, .FixedRows, 0, .Rows - 1, .Cols - 1) = vbBlack
        
        For i = .FixedRows To .Rows - 1
            lngEInvoiceID = Val(.TextMatrix(i, .ColIndex("ID")))
            If lngEInvoiceID <> 0 And (.Cell(flexcpChecked, i, .ColIndex("ѡ��")) = vbChecked Or Not blnChecked And i = .Row) Then
                lngCount = lngCount + 1
                
                If mobjPubEInvoice.zlSendEinvoiceMsg(Me, lngEInvoiceID, False, strErrMsg) Then
                    .TextMatrix(i, .ColIndex("��ӡ���")) = "���ͳɹ�"
                    .TextMatrix(i, .ColIndex("��ӡ˵��")) = ""
                Else
                    .TextMatrix(i, .ColIndex("��ӡ���")) = "����ʧ��"
                    .TextMatrix(i, .ColIndex("��ӡ˵��")) = strErrMsg
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                End If
            
                If Not blnChecked Then Exit For
            End If
        Next
    End With
    
    If lngCount = 0 Then
        MsgBox "��ѡ����Ҫ������Ϣ�ļ�¼��", vbInformation, gstrSysName
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
    
    If Me.ActiveControl Is vsfExse Then
        Set vsfGrid = vsfExse
        objOut.Title.Text = "����Ʊ����ϸ�嵥"
    Else
        Set vsfGrid = vsfEInvoice
        objOut.Title.Text = "����Ʊ�ݷ�����ϸ�嵥"
    End If
    
    '����
    If Me.ActiveControl Is vsfExse Then
        Set objRow = New zlTabAppRow
        objRow.Add "Ʊ�ݴ��룺" & vsfEInvoice.TextMatrix(vsfEInvoice.Row, vsfEInvoice.ColIndex("Ʊ�ݴ���"))
        objRow.Add "Ʊ�ݺ��룺" & vsfEInvoice.TextMatrix(vsfEInvoice.Row, vsfEInvoice.ColIndex("Ʊ�ݺ���"))
        objOut.UnderAppRows.Add objRow
    Else
        Set objRow = New zlTabAppRow
        objRow.Add "Ʊ�����ͣ�" & cboƱ������.Text
        objRow.Add "����ʱ�䣺" & Format(dtp��ʼʱ��, "yyyy-mm-dd") & " �� " & Format(dtp����ʱ��, "yyyy-mm-dd")
        objOut.UnderAppRows.Add objRow
    End If
    
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

Private Sub txtPatient_Change()
    txtPatient.Tag = ""
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, blnCancel As Boolean
     
    On Error GoTo ErrHandler
    If txtPatient.Locked Then Exit Sub

    If IDKind.GetCurCard.���� Like "����*" Then
        '103563,ֻҪ����ĵ�һ���ַ��ǡ�-+*����������ȫ���֣�����Ϊ����ˢ��
        If Not (InStr("-+*", Left(txtPatient.Text, 1)) > 0 And IsNumeric(Mid(txtPatient.Text, 2))) Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        End If
    ElseIf IDKind.GetCurCard.���� = "�����" Or IDKind.GetCurCard.���� = "סԺ��" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        txtPatient.IMEMode = 0
    End If

    If Not (blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 _
        Or KeyAscii = 13 And Trim(txtPatient.Text) <> "") Then Exit Sub

    If KeyAscii <> 13 Then
        txtPatient.Text = txtPatient.Text & Chr(KeyAscii): txtPatient.SelStart = Len(txtPatient.Text)
    End If
    KeyAscii = 0
    Call FindPati(IDKind.GetCurCard, blnCard)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ҳ���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, rsPatient As ADODB.Recordset
    Dim strInput As String
    
    RaiseEvent ShowInfo("")
    strInput = Trim(txtPatient.Text)
    If Val(txtPatient.Tag) <> 0 Then strInput = "-" & txtPatient.Tag
    
    If objCard.���� = "�շѵ��ݺ�" Or objCard.���� = "����Ʊ�ݺ�" Then
        '
    Else
        If Not GetPatient(objCard, strInput, rsPatient, blnCancel, blnCard) Then
            If blnCancel Then 'ȡ������
                txtPatient.Text = ""
                zlControl.ControlSetFocus txtPatient
                Exit Sub
            End If
            RaiseEvent ShowInfo("δ�ҵ��ò��ˣ�������������!")
            If blnCard Then
                txtPatient.Text = ""
                '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
                txtPatient.PasswordChar = ""
                txtPatient.IMEMode = 0
            Else
                zlControl.TxtSelAll txtPatient
            End If
            zlControl.ControlSetFocus txtPatient
            Exit Sub
        End If
        
        txtPatient.Text = Nvl(rsPatient!����)
        txtPatient.Tag = Val(Nvl(rsPatient!����ID))
    End If
    
    Call cmdRefresh_Click
    
    zlControl.ControlSetFocus txtPatient
    zlControl.TxtSelAll txtPatient
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, _
    ByRef rsPatient As ADODB.Recordset, ByRef blnCancel As Boolean, Optional ByVal blnCard As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:
    '   objCard-ָ���Ŀ����
    '   strInput-�����ֵ
    '   blnCard-�Ƿ�ˢ��
    '����:
    '   rsPatient-������Ϣ���ֶΣ�����ID,����
    '   blnCancel-�Ƿ�ȡ������
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strPati As String, strSQL As String
    Dim vRect As RECT, i As Integer, lng����ID As Long, strPassWord As String, strErrMsg As String
    Dim strWhere As String, lng�����ID As Long

    On Error GoTo ErrHandler
    blnCancel = False
    strWhere = ""
    If blnCard And objCard.���� Like "����*" And InStr("-+*", Left(strInput, 1)) = 0 Then  '103563
        If IDKind.Cards.��ȱʡ������ And Not IDKind.GetfaultCard Is Nothing Then
            lng�����ID = IDKind.GetfaultCard.�ӿ����
        Else
            lng�����ID = "-1"
        End If
        If mobjSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg, lng�����ID) = False Then Exit Function
        If lng����ID <= 0 Then Exit Function
        strInput = "-" & lng����ID
        strWhere = strWhere & " And A.����ID=[1] "
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then  '����ID
        strWhere = strWhere & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then  'סԺ��(��ס(��)Ժ�Ĳ���)
        strWhere = strWhere & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [1])"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����(�������ﲡ��)
        strWhere = strWhere & " And A.�����=[1]"
        '75087,���ﲡ���շ�ʱ,����Ҫ���������������,ֻ��Ҫ��������ŵ����˳��ż����ҵ��������Ĳ�����Ϣ������
        strInput = "*" & zlCommFun.GetFullNo(Mid(strInput, 2), 3)
    Else '��������
        Select Case objCard.����
            Case "����", "��������￨"
                strPati = _
                    " Select A.����ID as ID,A.����ID,A.����,A.�Ա�,A.����,A.סԺ��,B.���� as ����,A.��ǰ���� as ����,A.��������,A.���֤��,A.��ͥ��ַ" & _
                    " From ������Ϣ A,���ű� B" & _
                    " Where A.ͣ��ʱ�� is NULL And A.��ǰ����ID=B.ID(+) And A.���� Like [1]" & _
                    " Order by A.����"
                vRect = zlControl.GetControlRect(txtPatient.hWnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%")
                If rsTmp Is Nothing Then Exit Function
                strInput = rsTmp!����ID
                strWhere = strWhere & " And A.����ID=[2]"
            Case "ҽ����"
                strInput = UCase(strInput)
                strWhere = strWhere & " And A.ҽ����=[2]"
            Case "���֤��", "�������֤", "���֤"
                strInput = UCase(strInput)
                If mobjSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strInput = "-" & lng����ID
                strWhere = strWhere & " And A.����ID=[1]"
            Case "IC����", "IC��"
                strInput = UCase(strInput)
                If mobjSquareCard.zlGetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strInput = "-" & lng����ID
                strWhere = strWhere & " And A.����ID=[1]"
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.�����=[2]"
                strInput = zlCommFun.GetFullNo(strInput, 3)
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [2])"
            Case Else
                '��������,��ȡ��صĲ���ID
                If objCard.�ӿ���� > 0 Then
                    lng�����ID = objCard.�ӿ����
                    If mobjSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then Exit Function
                Else
                    If mobjSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then Exit Function
                End If
                If lng����ID <= 0 Then Exit Function
                strWhere = strWhere & " And A.����ID=[1]"
                strInput = "-" & lng����ID
        End Select
    End If
    
    strSQL = "Select A.����ID,A.���� From ������Ϣ A Where A.ͣ��ʱ�� is NULL" & strWhere
    Set rsPatient = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    If rsPatient.EOF Then Exit Function
    
    GetPatient = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub IDKind_ItemClick(index As Integer, objCard As Card)
    txtPatient.IMEMode = 0
    txtPatient.Text = ""
    zlControl.ControlSetFocus txtPatient: zlControl.TxtSelAll txtPatient
End Sub

'10.35.130:Private Sub IDKind_ReadCard(ByVal objCard As Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
'10.35.140�Ժ�:
Private Sub IDKind_ReadCard(ByVal objCard As Card, objPatiInfor As clsPatientInfo, blnCancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.����
    If txtPatient.Text <> "" Then Call FindPati(objCard, False)
End Sub

Private Function CreateSquareCardObject(ByRef frmMain As Object, ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������㿨����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    
    On Error Resume Next
    If mobjSquareCard Is Nothing Then
        Set mobjSquareCard = CreateObject("zlOneCardComLib.clsOneCardComLib")
        If Err <> 0 Then Exit Function
    End If
    
    'Public Function zlInitComponents(ByVal frmMain As Object, _
        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
        ByVal cnOracle As ADODB.Connection, _
        Optional blnDeviceSet As Boolean = False, _
        Optional strExpand As String) As Boolean
    '����:zlInitComponents (��ʼ���ӿڲ���)
    '���: frmMain-���õ�������
    '        lngModule-HIS����ģ���
    '       lngSys-�����ϵͳ��
    '       strDBUser-���ݿ��û���
    '       cnOracle -HIS/��������
    '       blnDeviceSet-�豸���õ��ó�ʼ��
    '       strExpand-��չ��Ϣ(��ѡ����:�����ID-����ʱ,��ʾȫ����ʼ��,����ʱ,ֻ��ʼ��ָ���Ľӿ�)
    '����:��������True:���óɹ�,False:����ʧ��
    If mobjSquareCard.zlInitComponents(frmMain, lngModule, mlngSys, mstrDBUser, gcnOracle, False, strExpend) = False Then
         '��ʼ�������ɹ�,����Ϊ�����ڴ���
         Set mobjSquareCard = Nothing
         Exit Function
    End If
    CreateSquareCardObject = True
End Function
