VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.UserControl ctlDockExpense 
   ClientHeight    =   7020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10560
   ScaleHeight     =   7020
   ScaleWidth      =   10560
   Begin VB.PictureBox picExpense 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   2424
      Left            =   150
      ScaleHeight     =   2430
      ScaleWidth      =   7620
      TabIndex        =   2
      Top             =   2280
      Width           =   7620
      Begin VSFlex8Ctl.VSFlexGrid vsExpense 
         Height          =   1440
         Left            =   96
         TabIndex        =   3
         Top             =   108
         Width           =   6960
         _cx             =   12277
         _cy             =   2540
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
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   2
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
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "������Ϣ:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   4
         Top             =   2160
         Width           =   1020
      End
   End
   Begin VB.PictureBox picAdvice 
      BorderStyle     =   0  'None
      Height          =   1668
      Left            =   408
      ScaleHeight     =   1665
      ScaleWidth      =   7335
      TabIndex        =   0
      Top             =   504
      Width           =   7332
      Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
         Height          =   1272
         Left            =   24
         TabIndex        =   1
         Top             =   84
         Width           =   6960
         _cx             =   12277
         _cy             =   2244
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
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   2
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
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "ctlDockExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------------------
'˵��:
'   �����ʹ�ÿؼ�(DockingPance�ؼ�),����󶨴�����ģ̬����,�������,
'   ����ÿؼ���ʽ,�򲻻�����������,���,���˵���Ϊ�ؼ���ʽ
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'��α���
Private mstrAdviceIDAndPayNums As String
Private mstrAdviceIDFull As String  '������ҽ�������Ϣ��,ҽ��ID�ͷ��ͺźͶ���ִ�б�־(ҽ��ID1:���ͺ�1:����ִ��,ҽ��ID2:���ͺ�2:����ִ��,...)
Private mstrNos As String, mbyt��¼���� As Byte, mbyt������Դ As Byte '�����ݺŲ�ʱ
Private mbytFun As Byte '0-��ҽ������;1-�����ݺŲ���
Private mblnMoved As Boolean
Private mlngModule As Long
Private mlngִ�п���ID As Long
Private mobjSquareCard As Object
'-----------------------------------------------------------------------------------------
Private mobjPubAdvice As Object  '����ҽ������
Private mfrmParent As Object
Private mobjSaveData As Object
Private mstrPrivsAnnexFee As String
Private mbytFocus As Byte

Private Enum mPaneIdx
    Pan_AdviceList = 1  'ҽ���б�
    Pan_FeeList = 2     '�����б�
End Enum
Private Type ty_adviceProperty 'ҽ����Ϣ
    lngҽ��ID As Long
    lng���ͺ� As Long
    bln����ִ�� As Boolean
    lng����ID  As Long
    lng��ҳId   As Long
    str�Һŵ�   As String
    lng���˿���ID   As Long
    lng���˲���ID   As Long
    lng��������ID   As Long
    int��¼����   As Integer
    int��˱�־   As Integer
    int����ģʽ   As Integer
    int������Դ As Integer
    intִ��״̬ As Integer

    lng���ID  As Long
    str�������  As String
    strNO As String
    dat����ʱ��  As Date
    str�ѱ�   As String
    lng�Ƽ�����  As Long
    bln������� As Boolean
    str�Ʒ�״̬ As String
    strFeeTab As String

End Type
Private mTYAdviceProperty As ty_adviceProperty
Private mrsPrice As ADODB.Recordset 'ҽ���Ƽ۹�ϵ

'ȱʡ����ֵ:
Const m_def_COLOR_FOCUS = &HFFCC99
Const m_def_COLOR_LOST = &HFFEBD7
Const m_def_Tittle = "ҽ�����ѹ���"
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
'���Ա���:
Dim m_COLOR_FOCUS As OLE_COLOR
Dim m_COLOR_LOST As OLE_COLOR
Dim m_Tittle As String
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
'Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
'�¼�����:
Event Click()
Attribute Click.VB_Description = "���û���һ�������ϰ��²��ͷ���갴ťʱ������"
Event DblClick()
Attribute DblClick.VB_Description = "���û���һ�������ϰ��²��ͷ���갴ť���ٴΰ��²��ͷ���갴ťʱ������"
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "���û���ӵ�н���Ķ����ϰ��������ʱ������"
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "���û����º��ͷ� ANSI ��ʱ������"
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "���û���ӵ�н���Ķ������ͷż�ʱ������"
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "���û���ӵ�н���Ķ����ϰ�����갴ťʱ������"
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "���û��ƶ����ʱ������"
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "���û���ӵ�н���Ķ������ͷ���귢����"

'---------------------------------------------------------------------------------------------------------
'����¼�
Event Activate() '���Ѽ���ʱ
Event RequestRefresh() 'Ҫ��������ˢ��
Event StatusTextUpdate(ByVal bytType As Byte, ByVal Text As String) 'Ҫ�����������״̬������
'bytType:1-����ִ��,2-����ȡ��

Event zlPopupMenu(lngҽ��ID As Long, lng���ͺ� As Long, strNO As String, int��¼���� As Integer, X As Single, Y As Single)
Private mblnNotFirstSel As Boolean '�ǵ�һ��ѡ��

Private Sub InitPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ʼ����
    '����:���˺�
    '����:2014-05-26 10:30:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single
    Dim strReg As String
    Dim panThis As Pane
    Set panThis = dkpMan.CreatePane(Pan_AdviceList, 200, 580, DockLeftOf, Nothing)
    panThis.Title = "ҽ����Ϣ"
    panThis.Handle = picAdvice.hWnd
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Tag = Pan_AdviceList

    Set panThis = dkpMan.CreatePane(Pan_FeeList, 250, 580, DockBottomOf, panThis)
    panThis.Title = "������Ϣ"
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picExpense.hWnd
    panThis.Tag = Pan_FeeList
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
    'zlRestoreDockPanceToReg  Me, dkpMan, "����"
End Sub
Private Sub picAdvice_Resize()
    Err = 0: On Error Resume Next
    With picAdvice
        vsAdvice.Move .ScaleLeft, .ScaleTop, .ScaleWidth, .ScaleHeight
    End With
End Sub
Private Sub picExpense_Resize()
    Err = 0: On Error Resume Next
    With picExpense
        If lblInfo.Visible Then
            vsExpense.Move .ScaleLeft, .ScaleTop, .ScaleWidth, .ScaleHeight - lblInfo.Height - 120
            lblInfo.Top = .ScaleHeight - lblInfo.Height - 30
        Else
            vsExpense.Move .ScaleLeft, .ScaleTop, .ScaleWidth, .ScaleHeight
        End If
    End With
End Sub
Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case Pan_AdviceList
        Item.Handle = picAdvice.hWnd
    Case Pan_FeeList
        Item.Handle = picExpense.hWnd
    End Select
End Sub

Private Sub UserControl_Resize()
    dkpMan.RecalcLayout
    Call picAdvice_Resize
    Call picExpense_Resize
End Sub

Private Sub vsAdvice_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsAdvice, Tittle, "ҽ����Ϣ", True
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
    Call RefreshExpenseData
End Sub
Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vsAdvice.AutoSizeMode = flexAutoSizeRowHeight
    Call vsAdvice.AutoSize(vsAdvice.ColIndex("ҽ������"))
    zl_vsGrid_Para_Save mlngModule, vsAdvice, Tittle, "ҽ����Ϣ", True
End Sub
Private Sub vsAdvice_GotFocus()
    vsAdvice.BackColorSel = COLOR_FOCUS
    mbytFocus = 1
End Sub
Private Sub vsAdvice_LostFocus()
    vsAdvice.BackColorSel = COLOR_LOST
End Sub
Private Sub vsExpense_GotFocus()
    mbytFocus = 2
    vsExpense.BackColorSel = COLOR_FOCUS
End Sub
Private Sub vsExpense_LostFocus()
    vsExpense.BackColorSel = COLOR_LOST
End Sub

Private Sub vsExpense_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsExpense, Tittle, "������Ϣ", True
End Sub
Private Sub vsExpense_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
     If OldRow = NewRow Then Exit Sub
     
     With vsExpense
        If NewRow <= 0 Or NewRow >= .Rows Then Exit Sub
        If NewCol <= 0 Or NewCol >= .Cols Then Exit Sub
        Call Load������Ϣ(NewRow)
        .ForeColorSel = .Cell(flexcpForeColor, NewRow, NewCol)
     End With
End Sub

Private Sub Load������Ϣ(ByVal lngRow As Long)
    Dim strSql As String, rsInfo As ADODB.Recordset
    On Error GoTo errH
    strSql = "Select ����" & vbNewLine & _
            "From ҩƷ�շ���¼" & vbNewLine & _
            "Where ���� = 21 And ����id In (Select Max(ID) From " & IIf(mTYAdviceProperty.int������Դ = 1, "������ü�¼", "סԺ���ü�¼") & " Where NO = [1] And ��¼���� = [2] And ��� = [3])"
    Set rsInfo = gobjDatabase.OpenSQLRecord(strSql, "Load������Ϣ", vsExpense.TextMatrix(lngRow, vsExpense.ColIndex("���ݺ�")), _
                                            Val(vsExpense.TextMatrix(lngRow, vsExpense.ColIndex("��¼����"))), Val(vsExpense.TextMatrix(lngRow, vsExpense.ColIndex("���"))))
    If rsInfo.EOF Then
        lblInfo.Visible = False
        vsExpense.Move picExpense.ScaleLeft, picExpense.ScaleTop, picExpense.ScaleWidth, picExpense.ScaleHeight
        Exit Sub
    End If
    lblInfo.Visible = True
    vsExpense.Move picExpense.ScaleLeft, picExpense.ScaleTop, picExpense.ScaleWidth, picExpense.ScaleHeight - lblInfo.Height - 120
    lblInfo.Caption = "������Ϣ:" & rsInfo!���� & "(" & vsExpense.TextMatrix(lngRow, vsExpense.ColIndex("����")) & _
                        vsExpense.TextMatrix(lngRow, vsExpense.ColIndex("���㵥λ")) & ")"
    lblInfo.Top = picExpense.ScaleHeight - lblInfo.Height - 30
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Sub

Private Sub vsExpense_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsExpense, Tittle, "������Ϣ", True
End Sub
Private Sub UserControl_Initialize()
    mlngModule = pҽ�����ѹ���  'ҽ�����ѹ���
    Call InitPancel
    Call InitGridHead(True)
End Sub

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=8,0,0,0
Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "����/���ö������ı���ͼ�εı���ɫ��"
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "����/���ö������ı���ͼ�ε�ǰ��ɫ��"
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "����/����һ��ֵ������һ�������Ƿ���Ӧ�û������¼���"
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property
 

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "ָ�� Label �� Shape �ı�����ʽ��͸���Ļ��ǲ�͸���ġ�"
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "����/���ö���ı߿���ʽ��"
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "ǿ����ȫ�ػ�һ������"
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ˢ������
    '����:���˺�
    '����:2014-05-30 14:48:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mbytFun <> 1 Then '��ҽ��ID����ҽ������
        If LoadAdviceData(mstrAdviceIDAndPayNums) = False Then Exit Sub
    Else
        '�����ݼ���ҽ������
        If LoadFeeListFromNos(mbyt��¼����, mstrNos, mbyt������Դ, mblnMoved) = False Then Exit Sub
    End If
End Sub
Public Property Get Isδ�Ʒ�() As Boolean
    Isδ�Ʒ� = InStr(mTYAdviceProperty.str�Ʒ�״̬, ",-1,")
End Property
Public Property Get IsHaveExpenseData() As Boolean
    IsHaveExpenseData = Get���ݺ� <> ""
End Property

'Ϊ�û��ؼ���ʼ������
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    m_Tittle = m_def_Tittle
    Set UserControl.Font = Ambient.Font
    m_COLOR_FOCUS = m_def_COLOR_FOCUS
    m_COLOR_LOST = m_def_COLOR_LOST
End Sub
'�Ӵ������м�������ֵ
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_Tittle = PropBag.ReadProperty("Tittle", m_def_Tittle)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.FontSize = PropBag.ReadProperty("FontSize", 9)
    m_COLOR_FOCUS = PropBag.ReadProperty("COLOR_FOCUS", m_def_COLOR_FOCUS)
    m_COLOR_LOST = PropBag.ReadProperty("COLOR_LOST", m_def_COLOR_LOST)
End Sub
Private Sub UserControl_Terminate()
    Err = 0: On Error Resume Next
    Set mobjPubAdvice = Nothing '�ͷ�ҽ��������������
    If gcnOracle Is Nothing Then Exit Sub
    If gcnOracle.State = 0 Then Exit Sub
    If gobjDatabase Is Nothing Then Exit Sub
    
    zlSaveDockPanceToReg Me, dkpMan, "����"
    zl_vsGrid_Para_Save mlngModule, vsAdvice, Tittle, "ҽ����Ϣ", True
    zl_vsGrid_Para_Save mlngModule, vsExpense, Tittle, "������Ϣ", True
End Sub

'������ֵд���洢��
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Tittle", m_Tittle, m_def_Tittle)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontSize", UserControl.FontSize, 9)
    Call PropBag.WriteProperty("COLOR_FOCUS", m_COLOR_FOCUS, m_def_COLOR_FOCUS)
    Call PropBag.WriteProperty("COLOR_LOST", m_COLOR_LOST, m_def_COLOR_LOST)
End Sub
Public Function zlRefresh(ByVal frmMain As Object, ByVal lng����id As Long, ByVal strAdviceIdAndPayNums As String, _
    Optional ByVal blnMoved As Boolean = False, Optional ByVal strNos As String, _
    Optional ByVal byt��¼���� As Byte, Optional ByVal byt������Դ As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������ˢ��
    '���:lng����id-����ID
    '     strAdviceIdAndPayNums-ҽ��ID�ͷ��ͺźͶ���ִ�б�־(ҽ��ID1:���ͺ�1:����ִ��,ҽ��ID2:���ͺ�2:����ִ��,...)
    '     strNos:���ݺ�(�������ʱ,�ö��ŷ���)
    '     byt��¼����:ҽ��ID����ʱ,�Ŵ���,��������(1-�շѵ�;2-���ʵ�)
    '     byt������Դ-1-����;2-סԺ
    '     blnMoved -�ò��˵������Ƿ���ת��
    '     bln����ִ��-���ڼ�����Ŀ��һ���ɼ���һ����Ŀ���Ƿ�������е�ĳһ������ִ��
    '����:
    '����:���˺�
    '����:2014-04-10 11:02:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = pҽ�����ѹ���  'ҽ�����ѹ���
    mstrNos = strNos: mbyt��¼���� = byt��¼����: mbyt������Դ = byt������Դ
    mlngִ�п���ID = lng����id
    mblnMoved = blnMoved
    mbytFun = IIf(strAdviceIdAndPayNums = "", 1, 0)
    mstrAdviceIDFull = strAdviceIdAndPayNums
    mstrPrivsAnnexFee = GetInsidePrivs(pҽ�����ѹ���)
    Set mfrmParent = frmMain
    mstrAdviceIDAndPayNums = GetAdviceIDAndPayNums(strAdviceIdAndPayNums)
    If mblnMoved = False Then
        mblnMoved = gobjDatabase.TableDataMoved("����ҽ������", " (ҽ��ID,���ͺ�) IN", " (Select C1 As ҽ��id, C2 As ���ͺ� From Table(f_Num2list2('" & mstrAdviceIDAndPayNums & "')))")
    End If
    Call VisiblePancel  '��ʾ������ҽ���б�
    If mbytFun <> 1 Then '��ҽ��ID����ҽ������
        If LoadAdviceData(mstrAdviceIDAndPayNums) = False Then Exit Function
        If mblnNotFirstSel = False Then Call SetDefalutFocus(True)
    Else
        '�����ݼ���ҽ������
        If LoadFeeListFromNos(mbyt��¼����, mstrNos, byt������Դ, blnMoved) = False Then Exit Function
        If mblnNotFirstSel = False Then Call SetDefalutFocus(False)
    End If
    mblnNotFirstSel = True
    zlRefresh = True
End Function
Private Function Get����ִ��״̬(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ִ��״̬
    '���:lngҽ��ID-ҽ��ID
    '     lng���ͺ�-���ͺ�
    '����:
    '����:��ȡ����ִ��״̬
    '����:���˺�
    '����:2014-05-27 11:12:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, i As Long, varTemp As Variant
    On Error GoTo errHandle
    varData = Split(mstrAdviceIDFull, ",")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ":::", ":")
        If lngҽ��ID = Val(varTemp(0)) And lng���ͺ� = Val(varTemp(1)) Then
            Get����ִ��״̬ = Val(varTemp(2)): Exit For
        End If
    Next
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetAdviceIDAndPayNums(ByVal strAdviceIDFull As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ��ID�ͷ��ͺŵ��ַ���
    '���:strAdviceIDFull(ҽ��ID�ͷ��ͺźͶ���ִ�б�־(ҽ��ID1:���ͺ�1:����ִ��,ҽ��ID2:���ͺ�2:����ִ��,...))
    '����:������ҽ��ID�ͷ��ͺ�Ϊ��ʽ�Ĵ�(ҽ��ID:���ͺ�,ҽ��ID:���ͺ�....)
    '����:���˺�
    '����:2014-05-26 15:37:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, i As Long, varTemp As Variant
    Dim strAdvice As String
    varData = Split(strAdviceIDFull, ",")
    strAdvice = ""
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & "::", ":")
        strAdvice = strAdvice & "," & varTemp(0) & ":" & varTemp(1)
    Next
    If strAdvice <> "" Then strAdvice = Mid(strAdvice, 2)
    GetAdviceIDAndPayNums = strAdvice
End Function
Private Sub VisiblePancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ������ҽ���б�
    '����:���˺�
    '����:2014-05-26 16:17:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim panThis As Pane
    If dkpMan Is Nothing Then Exit Sub
    Set panThis = dkpMan.FindPane(Pan_AdviceList)
    If panThis Is Nothing Then Exit Sub
    
    If mbytFun = 1 Then
        panThis.Close
    Else
        panThis.Closed = False
    End If
    dkpMan.RecalcLayout
End Sub

Private Sub vsExpense_DblClick()
    '˫���鿴
    If Get���ݺ� = "" Then Exit Sub
    If vsExpense.IsSubtotal(vsExpense.Row) = True Then Exit Sub
    
    Call frmTechnicExpense.EditCard(mfrmParent, mstrPrivsAnnexFee, 1, mTYAdviceProperty.lngҽ��ID, mTYAdviceProperty.lng���ͺ�, mTYAdviceProperty.lng����ID, mTYAdviceProperty.lng��ҳId, _
         IIf(mTYAdviceProperty.int������Դ = 2, 2, 1), Val(vsExpense.TextMatrix(vsExpense.Row, vsExpense.ColIndex("��¼����"))), mTYAdviceProperty.lng��������ID, mTYAdviceProperty.lng���˿���ID, 0, "", mTYAdviceProperty.strNO, Get���ݺ�)
End Sub

Private Function LoadAdviceData(ByVal strAdviceIdAndPayNums As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҽ������
    '���:strAdviceIDAndPayNums:ҽ��ID�ͷ��ͺ��ַ�����ҽ��ID1:���ͺ�1,ҽ��ID2:���ͺ�2
    '����:ҽ�����ݼ��سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-05-26 10:51:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, i As Long
    Dim rsTotal As ADODB.Recordset

    If mobjPubAdvice Is Nothing Then
        Call InitGridHead(True): LoadAdviceData = True
        Exit Function
    End If
    If GetAdviceMoney(strAdviceIdAndPayNums, rsTotal) = False Then Set rsTotal = Nothing
     
    'GetExecAdviceRecord:
    '  strIDsAndNos ҽ��ID�ͷ��ͺ��ַ�����ҽ��ID1:���ͺ�1,ҽ��ID2:���ͺ�2
    '  rsReturn:���صļ�¼��,�����ļ�¼����Ϣ��:
    '    ҽ��ID,���ID,���ͺ�,����ID,��ҳID,��ʼʱ��,ҽ������,����,Ӧ�ս��,ʵ�ս��,ҽ������,����ҽ��,����ʱ��
    On Error GoTo errHandle
    If mobjPubAdvice.GetExecAdviceRecord(strAdviceIdAndPayNums, rsTemp) = False Then Exit Function
    With vsAdvice
        .Redraw = flexRDNone
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        i = 1
        Do While Not rsTemp.EOF

            .TextMatrix(i, .ColIndex("ҽ��ID")) = Val(Nvl(rsTemp!ҽ��ID))
            .TextMatrix(i, .ColIndex("���ID")) = Val(Nvl(rsTemp!���ID))
            .TextMatrix(i, .ColIndex("���ͺ�")) = Val(Nvl(rsTemp!���ͺ�))
            .TextMatrix(i, .ColIndex("����ID")) = Val(Nvl(rsTemp!����ID))
            .TextMatrix(i, .ColIndex("����ִ��")) = Get����ִ��״̬(Val(Nvl(rsTemp!ҽ��ID)), Val(Nvl(rsTemp!���ͺ�)))
            .TextMatrix(i, .ColIndex("��ҳID")) = Val(Nvl(rsTemp!��ҳID))
            .TextMatrix(i, .ColIndex("��ʼʱ��")) = Format(rsTemp!��ʼʱ��, "yyyy-mm-dd HH:MM")
            .TextMatrix(i, .ColIndex("ҽ������")) = Nvl(rsTemp!ҽ������)
            .TextMatrix(i, .ColIndex("����")) = Nvl(rsTemp!����)
            If Not rsTotal Is Nothing Then
                If rsTotal.State = 1 Then
                    rsTotal.Filter = "ҽ��ID=" & Val(Nvl(rsTemp!ҽ��ID)) & " and ���ͺ�=" & Val(Nvl(rsTemp!���ͺ�))
                    If rsTotal.EOF = False Then
                        .TextMatrix(i, .ColIndex("Ӧ�ս��")) = Format(Val(Nvl(rsTotal!Ӧ�ս��)), gSysPara.Money_Decimal.strFormt_VB)
                        .TextMatrix(i, .ColIndex("ʵ�ս��")) = Format(Val(Nvl(rsTotal!ʵ�ս��)), gSysPara.Money_Decimal.strFormt_VB)
                    End If
                End If
            End If
            .TextMatrix(i, .ColIndex("ҽ������")) = Nvl(rsTemp!ҽ������)
            .TextMatrix(i, .ColIndex("����ҽ��")) = Nvl(rsTemp!����ҽ��)
            .TextMatrix(i, .ColIndex("����ʱ��")) = Format(rsTemp!����ʱ��, "yyyy-mm-dd HH:MM")
            i = i + 1
            rsTemp.MoveNext
        Loop
        .WordWrap = False
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        If .ColWidth(.ColIndex("ҽ������")) >= 6000 Then .ColWidth(.ColIndex("ҽ������")) = 6000
        .Redraw = flexRDBuffered
        '�ָ�������
        zl_vsGrid_Para_Restore mlngModule, vsAdvice, Tittle, "ҽ����Ϣ", True
        '��ҽ������,�����и�
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        Call .AutoSize(.ColIndex("ҽ������"))
        If .Rows > 1 Then
            Call vsAdvice_AfterRowColChange(-1, 0, .Row, .Col)
        End If
        .Redraw = flexRDBuffered
    End With
    
    LoadAdviceData = True
    Exit Function
errHandle:
    vsAdvice.Redraw = flexRDBuffered
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub InitGridHead(Optional ByRef blnInitHead As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������ؼ��ı�ͷ����Ϣ
    '���:blnInitHead-�Ƿ��ʼ����ͷ��Ϣ
    '����:���˺�
    '����:2014-05-26 10:40:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strHeand As String, varData As Variant
    On Error GoTo errHandle

    With vsAdvice
        .Redraw = flexRDNone
        If blnInitHead Then
            strHeand = "" & _
            "ҽ��ID,���ID,����ִ��,���ͺ�,����ID,��ҳID, ��ʼʱ��,ҽ������,����,Ӧ�ս��,ʵ�ս��,ҽ������,����ҽ��,����ʱ��"
            varData = Split(strHeand, ",")
            .Clear 1
            .Cols = UBound(varData) + 1
            .Rows = 2
           For i = 0 To UBound(varData)
                .TextMatrix(0, i) = varData(i)
           Next
        End If

        For i = 0 To .Cols - 1
            .TextMatrix(0, i) = varData(i)
            .ColKey(i) = Trim(UCase(.TextMatrix(0, i)))
            If i = .ColIndex("ҽ������") Then .ColWidth(i) = 2500
            'ColData:����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
            If .ColKey(i) Like "*ID" Or .ColKey(i) = "���ͺ�" Or .ColKey(i) = "����ִ��" Then
               .ColHidden(i) = True
               .ColData(i) = "-1||1"
               .ColAlignment(i) = flexAlignLeftCenter
            ElseIf .ColKey(i) Like "*ʱ��" Or .ColKey(i) Like "*����" _
                Or .ColKey(i) Like "*��" Or .ColKey(i) Like "*��" Then
                .ColAlignment(i) = flexAlignCenterCenter
                .ColData(i) = "0||0"
            ElseIf .ColKey(i) Like "*��*" Or .ColKey(i) Like "*��" _
                Or .ColKey(i) Like "*��" Then
                .ColAlignment(i) = flexAlignRightCenter
                .ColData(i) = "0||0"
            Else
                .ColAlignment(i) = flexAlignLeftCenter
                .ColData(i) = "0||0"
            End If
            If .ColKey(i) = "ҽ������" Then
                .ColData(i) = "1||0"
            End If
        Next
        .Redraw = flexRDBuffered
    End With
    With vsExpense
        If blnInitHead Then
            strHeand = "��������,��¼����,�շѱ�־,��������,���ݺ�,�շ�ϸĿID,�ѱ�,��������,������,���,���,��Ŀ,����,����,���㵥λ,Ӧ�ս��,ʵ�ս��,ִ�в���,ִ�����,ִ��״̬,�շ����,�Ǽ�ʱ��,����Ա����"
            varData = Split(strHeand, ",")
            .Clear 1
            .Cols = UBound(varData) + 1
            .Rows = 2
           For i = 0 To UBound(varData)
                .TextMatrix(0, i) = varData(i)
           Next
        End If
        For i = 0 To .Cols - 1
            .ColKey(i) = Trim(UCase(.TextMatrix(0, i)))
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignLeftCenter
            .MergeCol(i) = False
            Select Case .ColKey(i)
            Case "����", "Ӧ�ս��", "ʵ�ս��"
                .ColAlignment(i) = flexAlignRightCenter
            Case "���ݺ�", "��������", "�ѱ�", "��������", "������", "��¼״̬", "ִ��״̬", "�շ����"
                 'If .ColKey(i) <> "��������" Then
                 .ColHidden(i) = True
                If .ColKey(i) <> "��������" Then
                    .ColAlignment(i) = flexAlignCenterCenter
                End If
            Case Else
                If .ColKey(i) Like "*ID" Then
                    .ColHidden(i) = True
                ElseIf .ColKey(i) Like "*ʱ��" Or .ColKey(i) Like "*����" Then
                    .ColAlignment(i) = flexAlignCenterCenter
                End If
            End Select
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .AllowBigSelection = False
        .AllowSelection = False
        .AllowUserFreezing = flexFreezeNone
        
        .OutlineBar = flexOutlineBarComplete
        .ExplorerBar = flexExSortShow
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    vsAdvice.Redraw = flexRDBuffered
    vsExpense.Redraw = flexRDBuffered
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function CreatePubAdvice() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҽ���Ĺ�������
    '����:�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-05-26 10:44:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjPubAdvice Is Nothing Then
        Err = 0: On Error Resume Next
        Set mobjPubAdvice = CreateObject("zlPublicAdvice.clsPublicAdvice")
        If Err <> 0 Then
            'Call MsgBox("����ҽ��������ʧ,ҽ����Ϣ����ʾ�쳣,����ϵͳ����Ա��ϵ!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName)
            Err = 0: On Error GoTo 0
            Exit Function
        End If
        Err = 0: On Error GoTo Errhand:
        Call mobjPubAdvice.InitCommon(gcnOracle, glngSys)
    End If
    CreatePubAdvice = True
    Exit Function
Errhand:
    If gobjComlib Is Nothing Then Exit Function
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function
Private Function LoadFeeListFromNos(ByVal byt��¼���� As Byte, ByVal strNos As String, _
    ByVal byt������Դ As Byte, ByVal blnMoved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ��ݺźͼ�¼���ʼ�������
    '���:byt��¼����:(1-�շ�;2-����)
    '     strNos:���ݺ�,����ö��ŷ���
    '     byt������Դ-1-����;2-סԺ
    '     blnMoved-�Ƿ�ת������ʷ��ռ�
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-05-26 16:23:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String, i As Long, r As Long
    Dim strFeeTab As String
    Dim blnҩ����λ As Boolean, strҩ����λ As String, strҩ����װ As String

    On Error GoTo errHandle
    'ҩƷ��λ
    blnҩ����λ = Val(gobjDatabase.GetPara("ҩƷ��λ", glngSys, pҽ�����ѹ���)) <> 0
    If byt������Դ = 1 Then
        strҩ����λ = "���ﵥλ": strҩ����װ = "�����װ"
    Else
        strҩ����λ = "סԺ��λ": strҩ����װ = "סԺ��װ"
    End If

    strFeeTab = IIf(byt������Դ = 1, "������ü�¼", "סԺ���ü�¼")
    strFeeTab = IIf(blnMoved, "H", "") & strFeeTab

    strSql = "" & _
    "   Select mod(A.��¼����,10) as ��¼����,A.��¼״̬,A.NO as ���ݺ�," & _
    "          A.�ѱ�,Nvl(A.�۸񸸺�,A.���) as ���,A.�շ�ϸĿID, " & _
    "          avg(Nvl(A.����,1)*A.����) as ����,sum(A.��׼����) as ��׼����,Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��," & _
    "          Max(decode(A.��¼״̬,2,0,A.ִ��״̬)) as ִ��״̬,A.�շ����, " & _
    "          Max(decode(A.��¼״̬,2,NULL,decode(A.��¼����,11,NULL,to_char(a.�Ǽ�ʱ��,'yyyy-mm-dd hh24:mi:ss')))) as �Ǽ�ʱ��," & _
    "          Max(decode(A.��¼״̬,2,NULL,decode(A.��¼����,11,NULL,to_char(a.����ʱ��,'yyyy-mm-dd hh24:mi:ss')))) as ����ʱ��," & _
    "          Max(decode(A.��¼״̬,2,NULL,decode(A.��¼����,11,NULL,A.����Ա����))) as ����Ա����," & _
    "          A.ִ�в���ID,A.��������ID,A.������" & _
    "   From " & strFeeTab & " A,Table(f_str2list([2])) B" & _
    "   Where  mod(A.��¼����,10)=[1]  And  A.NO=B.Column_Value" & _
    "   Group by mod(A.��¼����,10),A.��¼״̬,A.NO,A.�ѱ�,Nvl(A.�۸񸸺�,A.���),A.�շ�ϸĿID,A.�շ����," & _
    "           A.ִ�в���ID,A.��������ID,A.������"
    
    strSql = "" & _
    "   Select /*+ RULE */ '' as ��������,mod(A.��¼����,10) as ��¼����,decode(nvl(max(a.��¼״̬),0),1,1,0) as �շѱ�־," & _
    "       Decode( a.��¼����, 1, '�շ�', 2, '����', 3, '����', 4, '�Һ�', '5', 'ҽ�ƿ�', 'δ֪') As ��������," & _
    "       A.���ݺ�, " & _
    "       A.�շ�ϸĿID,A.�ѱ�,M.���� as ��������,A.������,C.���� as ���,A.��� ," & _
    "       Nvl(F.����,B.����)||Decode(B.���,NULL,NULL,' '||B.���) as ��Ŀ," & _
    "       sum(A.��׼����" & IIf(blnҩ����λ, "*Nvl(E." & strҩ����װ & ",1)", "") & ") as ����," & _
    "       Sum(Nvl(A.����,1)" & IIf(blnҩ����λ, "/Nvl(E." & strҩ����װ & ",1)", "") & ") as ����," & _
            IIf(blnҩ����λ, "Decode(E.ҩƷID,NULL,B.���㵥λ,E." & strҩ����λ & ")", "B.���㵥λ") & " as ���㵥λ," & _
    "       Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��,D.���� as ִ�в���," & _
    "       Decode(Max(Nvl(A.ִ��״̬,0)),0,'δִ��',1,'��ȫִ��',2,'����ִ��') as ִ�����," & _
    "       Max(Nvl(A.ִ��״̬,0)) as ִ��״̬,a.�շ����, " & _
    "       max(A.�Ǽ�ʱ��) as �Ǽ�ʱ��,Max(A.����Ա����) as ����Ա����" & _
    " From  (" & strSql & ") A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,ҩƷ��� E,�շ���Ŀ���� F,���ű� M" & _
    " Where A.�շ�ϸĿID=B.ID   And A.�շ����=C.���� And a.��������id = M.Id(+) And A.ִ�в���ID=D.ID(+)" & _
    "       And B.ID=E.ҩƷID(+) And A.�շ�ϸĿID=F.�շ�ϸĿID(+)" & _
    "       And F.����(+)=1 And F.����(+)=[3] " & _
    " Group by   a.��¼����,A.���ݺ�,A.�ѱ�,M.����,A.������,A.���,C.����, A.�շ�ϸĿID ,Nvl(F.����,B.����),B.���,B.���㵥λ,D.����," & _
    "       a.�շ����,E.ҩƷID,Nvl(E." & strҩ����װ & ",1),E." & strҩ����λ & "" & _
    "      " & _
    " Having Nvl(Sum(A.Ӧ�ս��),0)<>0 Or Nvl(Sum(A.ʵ�ս��),0)<>0" & _
    " Order by ��������,���ݺ�,���"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Tittle, byt��¼����, strNos, IIf(gSysPara.bytҩƷ������ʾ = 0, 1, 3))
    With vsExpense
        .Redraw = flexRDNone
        .Cols = 1: .FixedCols = 0
        .Rows = 2
        .MergeRow(1) = False
        Set .DataSource = rsTemp
    End With
    Call SetExpenseGridProperty '������������
    vsExpense.Redraw = flexRDBuffered
    LoadFeeListFromNos = True
    Exit Function
errHandle:
    vsExpense.Redraw = flexRDBuffered
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SetExpenseGridProperty()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���÷���������������
    '����:���˺�
    '����:2014-05-26 17:40:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long

    On Error GoTo errHandle
    With vsExpense
        For i = 0 To .Cols - 1
            .ColKey(i) = Trim(UCase(.TextMatrix(0, i)))
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignLeftCenter
            .MergeCol(i) = False
            Select Case .ColKey(i)
            Case "����", "Ӧ�ս��", "ʵ�ս��"
                .ColAlignment(i) = flexAlignRightCenter
            Case "���", "���ݺ�", "��������", "�ѱ�", "��������", "������", "��¼״̬", "ִ��״̬", "�շ����"
                 'If .ColKey(i) <> "��������" Then
                 .ColHidden(i) = True
                If .ColKey(i) <> "��������" Then
                    .ColAlignment(i) = flexAlignCenterCenter
                End If
            Case Else
                If .ColKey(i) Like "*ID" Then
                    .ColHidden(i) = True
                ElseIf .ColKey(i) Like "*ʱ��" Or .ColKey(i) Like "*����" Then
                    .ColAlignment(i) = flexAlignCenterCenter
                End If
            End Select
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        zl_vsGrid_Para_Restore mlngModule, vsExpense, Tittle, "������Ϣ", True
    End With
    '���鴦��
    Call ExpenseSplitGroup
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ExpenseSplitGroup()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Է����б���Ϣ���з�����ʾ
    '����:���˺�
    '����:2014-05-26 16:58:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim strTemp As String

    On Error GoTo errHandle
    With vsExpense
        For i = 0 To .Cols - 1
            If i < .ColIndex("���") And i <> .ColIndex("��������") Then
                If i <> .ColIndex("��������") Or mbytFun = 1 Then
                    .ColHidden(i) = True
                End If
            End If
        Next
        
        .OutlineBar = flexOutlineBarComplete
        .Subtotal flexSTClear
        .MultiTotals = True
        '&H8000000F
        .Subtotal flexSTSum, .ColIndex("���ݺ�"), .ColIndex("ʵ�ս��"), , &H8000000F, , True, "%s", , True
        .Subtotal flexSTSum, .ColIndex("���ݺ�"), .ColIndex("Ӧ�ս��"), , &H8000000F, , True, "%s", , True
        .SubtotalPosition = flexSTAbove
        If mbytFun = 1 Then
            .Outline .ColIndex("���")
            .OutlineCol = .ColIndex("���")
        Else
            .Outline .ColIndex("��������")
            .OutlineCol = .ColIndex("��������")
        End If
        
        For i = 1 To .Rows - 1
            .MergeRow(i) = False
            If .IsSubtotal(i) Then
                .IsCollapsed(i) = flexOutlineExpanded
                strTemp = .Cell(flexcpTextDisplay, i, 0)
                .RowHeight(i) = 450
                '�����ݺ���ʾ����Ŀ��������
                If mbytFun = 1 Then
                    .Cell(flexcpText, i, .ColIndex("���")) = strTemp
                Else
                    .Cell(flexcpText, i, .ColIndex("��������")) = strTemp
                End If
                 strTemp = .Cell(flexcpTextDisplay, i + 1, .ColIndex("��������"))
                 strTemp = strTemp & Space(2) & "�ѱ�:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("�ѱ�"))
                 strTemp = strTemp & Space(2) & "��������:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("��������"))
                 strTemp = strTemp & Space(2) & "������:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("������"))
                 .MergeRow(i) = True
                 .MergeCells = flexMergeRestrictRows
                 If Val(.TextMatrix(i + 1, .ColIndex("��¼����"))) Mod 10 = 1 _
                        And Val(.TextMatrix(i + 1, .ColIndex("�շѱ�־"))) = 1 Then   '���շ���ɫ��ʾ
                        .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &HC00000 '����
                        .ForeColorSel = &HC00000
                 Else
                        .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlack
                        .ForeColorSel = vbBlack
                 End If
                 For j = 0 To .Cols - 1
                    If j > .ColIndex("��������") And j < .ColIndex("Ӧ�ս��") Then
                        If mbytFun = 1 Then
                            If j > .ColIndex("���") Then
                                .Cell(flexcpText, i, j) = strTemp
                                .Cell(flexcpFontBold, i, j) = True
                            End If
                        Else
                            .Cell(flexcpText, i, j) = strTemp
                            .Cell(flexcpFontBold, i, j) = True
                        End If
                       '82582:���ϴ�,2015/2/10,ȥ������еĶ��ŷָ���
                    ElseIf .ColIndex("ʵ�ս��") = j Then
                        .TextMatrix(i, j) = Format(Val(zlFormatNum(.TextMatrix(i, j))), gSysPara.Money_Decimal.strFormt_VB)
                    ElseIf .ColIndex("Ӧ�ս��") = j Then
                        .TextMatrix(i, j) = " " & Format(Val(zlFormatNum(.TextMatrix(i, j))), gSysPara.Money_Decimal.strFormt_VB)
                    End If
                 Next
            Else
                .TextMatrix(i, .ColIndex("����")) = Format(Val(zlFormatNum(.TextMatrix(i, .ColIndex("����")))), gSysPara.Price_Decimal.strFormt_VB)
                .TextMatrix(i, .ColIndex("Ӧ�ս��")) = Format(Val(zlFormatNum(.TextMatrix(i, .ColIndex("Ӧ�ս��")))), gSysPara.Money_Decimal.strFormt_VB)
                .TextMatrix(i, .ColIndex("ʵ�ս��")) = Format(Val(zlFormatNum(.TextMatrix(i, .ColIndex("ʵ�ս��")))), gSysPara.Money_Decimal.strFormt_VB)
            End If
        Next
        If mbytFun = 1 Then
            Call .AutoSize(.ColIndex("���"))
        Else
            Call .AutoSize(.ColIndex("��������"))
        End If
        Call .AutoSize(.ColIndex("����"))
        For j = 0 To .Cols - 1
            If j > .ColIndex("��Ŀ") And j < .ColIndex("Ӧ�ս��") Then
                .MergeCol(j) = True
            Else
                .MergeCol(j) = False
            End If
        Next
        
    End With
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

'-------------------------------------------
Private Function SetAdviceProperty(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, _
    ByVal bln����ִ�� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҽ�����������
    '����:���˺�
    '����:2014-05-27 10:03:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tyAdviceProperty As ty_adviceProperty
    Dim strSql As String, rsTemp As ADODB.Recordset


    On Error GoTo errHandle
    With tyAdviceProperty
        .lngҽ��ID = lngҽ��ID
        .lng���ͺ� = lng���ͺ�
        .bln����ִ�� = bln����ִ��
    End With
    mTYAdviceProperty = tyAdviceProperty

    strSql = _
    " Select A.����ID,A.��ҳID,A.�Һŵ�,A.���˿���ID,D.��ǰ����id,A.��������ID,A.������Դ,C.����ģʽ,A.�������,E.�Ƽ�����," & _
    "       Decode(A.�������,'D',Nvl(A.���ID,A.ID),A.���ID) as ���ID,B.NO,B.��¼����,Nvl(B.�������,0) as �������," & _
    "       B.ִ��״̬,B.����ʱ��,Nvl(D.�ѱ�,C.�ѱ�) as �ѱ�,d.��˱�־ " & _
    " From ������Ϣ C,������ҳ D," & IIf(mblnMoved, "H", "") & "����ҽ����¼ A," & IIf(mblnMoved, "H", "") & "����ҽ������ B,������ĿĿ¼ E" & _
    " Where A.ID=B.ҽ��ID And A.ID=[1] And B.���ͺ�=[2] And A.������ĿID=E.ID" & _
    " And A.����ID=C.����ID And A.����ID=D.����ID(+) And A.��ҳID=D.��ҳID(+)"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Tittle, mTYAdviceProperty.lngҽ��ID, mTYAdviceProperty.lng���ͺ�)
    If rsTemp.RecordCount = 0 Then SetAdviceProperty = True: Exit Function


    With mTYAdviceProperty
        .lng���ID = Val(Nvl(rsTemp!���ID))
        .lng����ID = Val(Nvl(rsTemp!����ID))
        .lng��ҳId = Val(Nvl(rsTemp!��ҳID))
        .str�Һŵ� = Nvl(rsTemp!�Һŵ�)
        .lng���˿���ID = Val(Nvl(rsTemp!���˿���id))
        .lng���˲���ID = Val(Nvl(rsTemp!��ǰ����ID))
        .lng��������ID = Val(Nvl(rsTemp!��������id))
        .int��¼���� = Val(Nvl(rsTemp!��¼����))
        .int��˱�־ = Val(Nvl(rsTemp!��˱�־))
        .int����ģʽ = Val(Nvl(rsTemp!����ģʽ))
        .str������� = Nvl(rsTemp!�������)
        .strNO = Nvl(rsTemp!NO)
        .intִ��״̬ = Val(Nvl(rsTemp!ִ��״̬))
        .dat����ʱ�� = rsTemp!����ʱ��
        .str�ѱ� = Nvl(rsTemp!�ѱ�)
        .lng�Ƽ����� = Val(Nvl(rsTemp!�Ƽ�����))
        .str�Ʒ�״̬ = GetSendFeeState()
        .int������Դ = Val(Nvl(rsTemp!������Դ))
        .bln������� = Val(Nvl(rsTemp!�������))
         '�����סԺҽ��վ�ɷ���������ʣ�����������ü�¼��
         '��ǰ������ҽ��վ����Ϊ�������ʱ��rsTemp!������ʵ�ֵΪ�գ�δ������ʷ����
        .strFeeTab = "������ü�¼"
        If .int������Դ = 2 Then
            If .bln������� Then
                .int������Դ = 1    '�������ﲡ��(�������һ�����������۲���)
            Else
                .strFeeTab = "סԺ���ü�¼"
            End If
        End If
        '������ϺͶಿλ�����Ŀ���ۺ�ִ��״̬
        If (.str������� = "C" Or .str������� = "D") And Not .bln����ִ�� Then
            strSql = "" & _
            "   Select ִ��״̬ From ����ҽ������ " & _
            "   Where ���ͺ�=[1]  And ҽ��ID IN(Select ID From " & IIf(mblnMoved, "H", "") & "����ҽ����¼ Where (ID=[2] Or ���ID=[2]) And ������� In('C','D'))"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Tittle, lng���ͺ�, IIf(.lng���ID <> 0, .lng���ID, .lngҽ��ID))
            strSql = ""
            Do While Not rsTemp.EOF
                If InStr(strSql, Nvl(rsTemp!ִ��״̬, 0)) = 0 Then
                    strSql = strSql & Nvl(rsTemp!ִ��״̬, 0)
                End If
                rsTemp.MoveNext
            Loop
            .intִ��״̬ = IIf(Len(strSql) = 1, Val(strSql), 3)
        End If
    End With
    SetAdviceProperty = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetSendFeeState() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ��ҽ��ĳ�η���֮��ļƷ�״̬
    '���:lngҽ��ID-�����������Ŀ,��������Ŀ,��һ��������Ŀ��ҽ��ID(����ҽ��վ����ʾ����Ŀ��)
    '     lng���ͺ�-���ͺ�
    '     bln����ִ��-�����Ŀ�Ƿ����ִ��
    '����:
    '����:",-1,0,1,"������-1=����Ʒ�,1=�ѼƷ�,0=δ�Ʒ�,�������ﵥ�ݣ�2=�����շ�,3=ȫ���շ�
    '����:���˺�
    '����:2014-05-27 09:33:50
    '˵��:��ȡָ��ҽ��ĳ�η���֮��ļƷ�״̬����Ҫ����һЩ���ҽ���ж��ּƷѵ�״̬
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset, strSql As String

    On Error GoTo errH

    If mTYAdviceProperty.bln����ִ�� Then
        strSql = "Select Distinct �Ʒ�״̬ From ����ҽ������ Where ҽ��ID=[1] And ���ͺ�=[2]"
    Else
        strSql = "Select ID From ����ҽ����¼ Where (ID=[3] Or ���ID=[3]) And �������=[4]"
        strSql = "Select Distinct �Ʒ�״̬ From ����ҽ������ Where ҽ��ID IN(" & strSql & ") And ���ͺ�=[2]"
    End If
    If mblnMoved Then
        strSql = Replace(strSql, "����ҽ����¼", "H����ҽ����¼")
        strSql = Replace(strSql, "����ҽ������", "H����ҽ������")
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlCISKernel", mTYAdviceProperty.lngҽ��ID, mTYAdviceProperty.lng���ͺ�, IIf(mTYAdviceProperty.lng���ID <> 0, mTYAdviceProperty.lng���ID, mTYAdviceProperty.lngҽ��ID), mTYAdviceProperty.str�������)
    strSql = ""
    Do While Not rsTmp.EOF
        strSql = strSql & "," & IIf(Val("" & rsTmp!�Ʒ�״̬) > 1, 1, Val("" & rsTmp!�Ʒ�״̬))
        rsTmp.MoveNext
    Loop
    If strSql <> "" Then GetSendFeeState = strSql & ","
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
Private Function LoadFeeDataFromAdvice(ByVal lngҽ��ID As Long, _
    ByVal lng���ͺ� As Long, ByVal bln����ִ�� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ��ҽ������Ҫ���ü����ӷ���
    '���:lngҽ��ID-��ǰҽ��ID
    '     lng���ͺ�-���ͺ�
    '     bln����ִ��-�����Ŀ�Ƿ����ִ��
    '����:���˺�
    '����:2014-05-27 11:02:09
    '˵����1.����ҽ������������ü����ӷ���,�����ÿ�����δ����
    '      2.Ŀǰ�����ݲ�֧�ֲ����˷�,�����嵥��ֻ�����ʾ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnDataMoved As Boolean
    Dim dblӦ�� As Double, dblʵ�� As Double
    Dim strIF As String
    Dim tyAdviceProperty As ty_adviceProperty
    
    On Error GoTo errHandle

    '����ָ��ҽ������ϸ����
    Set mrsPrice = Nothing
    If lngҽ��ID = 0 Or lng���ͺ� = 0 Then
        vsExpense.Rows = 2: vsExpense.Clear 1
        vsExpense.Subtotal flexSTClear
        mTYAdviceProperty = tyAdviceProperty '115514,���ҽ����Ϣ
        LoadFeeDataFromAdvice = True: Exit Function
    End If
        
        
    '����ָ��ҽ�����������
    Call SetAdviceProperty(lngҽ��ID, lng���ͺ�, bln����ִ��)

    blnDataMoved = mblnMoved
    If Not blnDataMoved Then
        blnDataMoved = gobjDatabase.DateMoved(mTYAdviceProperty.dat����ʱ��)
    End If
    
    If LoadDataFromAdvices(blnDataMoved) = False Then Exit Function
    LoadFeeDataFromAdvice = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetNotFeeSQL(ByVal blnDataMoved As Boolean, ByRef strҽ��IDs As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:׷��δ�ƷѲ��ֵķ���
    '����:strҽ��IDs-�����漰��ҽ��IDs
    '����:����SQL
    '����:���˺�
    '����:2014-05-27 15:05:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String, rsHead As ADODB.Recordset
    Dim int��Դ As Integer, lngִ�в���ID As Long
    Dim str�Ǽ�ʱ�� As String

    On Error GoTo errHandle

    strҽ��IDs = ""

    '����δ�Ʒ�״̬,ֱ�Ӷ�ȡ�շѹ�ϵ��ʾ
    If InStr(mTYAdviceProperty.str�Ʒ�״̬, ",0,") = 0 Then Exit Function

    Call LoadAdvicePrice(mblnMoved)

    If mrsPrice Is Nothing Then Exit Function
    If mrsPrice.State <> 1 Then Exit Function

    int��Դ = IIf(mTYAdviceProperty.int������Դ = 2, 2, 1)
    strSql = "" & _
    "   Select A.��������ID,B.���� as ��������,Nvl(ͣ��ҽ��,����ҽ��) as ����ҽ��,������Դ,��ʼִ��ʱ��  " & _
    "   From ����ҽ����¼ A,���ű� B  " & _
    "   Where A.��������ID=B.ID And A.ID=[1]"
    If blnDataMoved Then
        strSql = Replace(strSql, "����ҽ����¼", "H����ҽ����¼")
    End If

    Set rsHead = gobjDatabase.OpenSQLRecord(strSql, Tittle, mTYAdviceProperty.lngҽ��ID)
    If rsHead.EOF Then Exit Function
    str�Ǽ�ʱ�� = Format(Nvl(rsHead!��ʼִ��ʱ��), "yyyy-mm-dd HH:MM:SS")
    If str�Ǽ�ʱ�� = "" Then str�Ǽ�ʱ�� = Format(gobjDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Dim lng��Ŀid As Long, int���� As Long, i As Long
    strSql = ""
    With mrsPrice
        If mrsPrice.RecordCount <> 0 Then .MoveFirst
        For i = 1 To .RecordCount
            If lng��Ŀid <> !�շ�ϸĿID Then int���� = i

            lngִ�в���ID = Get�շ�ִ�п���ID(mTYAdviceProperty.lng����ID, mTYAdviceProperty.lng��ҳId, !���, !�շ�ϸĿID, !ִ�п���, mTYAdviceProperty.lng���˿���ID, Nvl(!��������id, 0), int��Դ)

            strSql = strSql & IIf(strSql <> "", " Union ALL ", "") & _
            " Select 1 as ��������,'[δ�Ʒ�]' as NO," & mTYAdviceProperty.int��¼���� & " as ��¼����,1 as ��¼״̬," & _
            "       '" & mTYAdviceProperty.str�ѱ� & "' as �ѱ�," & i & " as ���," & IIf(int���� = i, "-NULL", int����) & " as �۸񸸺�," & _
                    !ҽ��ID & " as ҽ�����,'" & !��� & "' as �շ����," & !�շ�ϸĿID & " as �շ�ϸĿID," & _
                  rsHead!��������id & " as ��������ID ," & "'" & Nvl(rsHead!����ҽ��) & "'  as ������," & lngִ�в���ID & " as ִ�в���ID," & _
            "       0 as ִ��״̬," & !������ĿID & " as ������ĿID,1 as ����," & !���� & " as ����," & !���� & " as ��׼����," & _
                     !Ӧ�� & " as Ӧ�ս��," & !ʵ�� & " as ʵ�ս��, " & _
            "       To_Date('" & str�Ǽ�ʱ�� & "','YYYY-MM-DD HH24:MI:SS') as �Ǽ�ʱ��, " & _
            "       To_Date('" & str�Ǽ�ʱ�� & "','YYYY-MM-DD HH24:MI:SS') as ����ʱ��, '" & _
                    IIf(Val(Nvl(rsHead!������Դ)) = 3, Nvl(rsHead!����ҽ��), UserInfo.����) & "' as ����Ա , 0 as ���շ�" & _
            " From Dual"
            lng��Ŀid = !�շ�ϸĿID
             strҽ��IDs = strҽ��IDs & "," & !ҽ��ID
            .MoveNext
        Next
        If strSql = "" Then Exit Function
        strҽ��IDs = Mid(strҽ��IDs, 2) 'ȡ����������漰��ҽ��ID
    End With
    GetNotFeeSQL = " Union ALL " & strSql
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadDataFromAdvices(ByVal blnDataMoved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:׷��ҽ����Ӧ�������ú͸��ӷ���
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-05-27 16:10:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, strSQL1 As String, strSQL2 As String
    Dim strFeeTab As String, strIDs As String
    Dim blnҩ����λ As String, strҩ����λ As String, strҩ����װ As String
    Dim strWith As String, int��Դ As Integer
    Dim strҽ��IDs As String, rsTemp As ADODB.Recordset

    On Error GoTo errHandle
    strSql = ""
    strFeeTab = mTYAdviceProperty.strFeeTab
    If strFeeTab = "" Then Exit Function
     'ҩƷ��λ
    int��Դ = IIf(mTYAdviceProperty.int������Դ = 2, 2, 1)
    blnҩ����λ = Val(gobjDatabase.GetPara("ҩƷ��λ", glngSys, pҽ�����ѹ���)) <> 0
    If int��Դ = 1 Then
        strҩ����λ = "���ﵥλ": strҩ����װ = "�����װ"
    Else
        strҩ����λ = "סԺ��λ": strҩ����װ = "סԺ��װ"
    End If

    '������鲿λ������������������ϵķ���
    If mTYAdviceProperty.bln����ִ�� Then
        strҽ��IDs = "Select [1] From Dual"
    Else
        '�ಿλ��飬�������в�λ�ͷ�����
        '1.���;����
        '2.����
        strҽ��IDs = _
        " Select ID From ����ҽ����¼ Where ID=[1] Or (���ID=[1] And ������� IN('F','D'))" & _
        " Union ALL " & _
        " Select ID From ����ҽ����¼ Where �������='C' And ���ID=[2]"
    End If
    If mblnMoved Then
        strҽ��IDs = Replace(strҽ��IDs, "����ҽ����¼", "H����ҽ����¼")
    ElseIf blnDataMoved Then
        strҽ��IDs = strҽ��IDs & " Union ALL " & Replace(strҽ��IDs, "����ҽ����¼", "H����ҽ����¼")
    End If

    '�����ѼƷ�״̬,Ӧ�ÿ���ֱ�Ӷ�ȡ�����ò���
    'ֻ��һ�ŵ���,���ܺ�����ҽ������;
    '��ʾԭʼ������Ϣ��ʣ�ಿ�ݽ��
    If InStr(mTYAdviceProperty.str�Ʒ�״̬, ",1,") > 0 Then
        '������鲿λ������������������ϵķ���
        strSQL1 = _
        " Select 1 as ��������,A.��¼����,Decode(B.��¼״̬,0,0,1) as ���շ�," & _
        "       A.NO,B.�ѱ�,Sum(B.Ӧ�ս��) as Ӧ�ս��,Sum(B.ʵ�ս��) as ʵ�ս��," & _
        "       C.���� as ��������,B.������,Max(Decode(Floor(x.��¼����/10), 0, x.�Ǽ�ʱ��, Null)) As �Ǽ�ʱ��, " & _
        "       Max(Decode(Floor(x.��¼����/10), 0, Nvl(x.����Ա����, x.������), Null)) As ����Ա " & _
        " From ����ҽ������ A," & strFeeTab & " B,���ű� C," & strFeeTab & " X" & _
        " Where A.ҽ��ID IN(" & strҽ��IDs & ") And A.���ͺ�=[3]" & _
        "   And A.NO=B.NO And A.��¼����=Decode(B.��¼����,11,1,B.��¼����)" & _
        "   And A.ҽ��ID=B.ҽ�����+0 And B.��������ID=C.ID" & _
        "   And B.NO=X.NO And B.��¼����=X.��¼���� And B.���=X.��� And X.��¼״̬ IN(0,1,3)" & _
        " Group by A.��¼����,Decode(B.��¼״̬,0,0,1),A.NO,B.�ѱ�,C.����,B.������"
        If mblnMoved Then
            strSQL1 = Replace(strSQL1, "����ҽ������", "H����ҽ������")
            strSQL1 = Replace(strSQL1, strFeeTab, "H" & strFeeTab)
        ElseIf blnDataMoved Then
            strSQL2 = Replace(strSQL1, "����ҽ������", "H����ҽ������")
            strSQL2 = Replace(strSQL2, strFeeTab, "H" & strFeeTab)
            strSQL1 = strSQL1 & " Union ALL " & strSQL2
            strSQL1 = _
                " Select A.��������,A.��¼����,A.���շ�,A.NO,A.�ѱ�," & _
                "       Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��," & _
                " A.��������,A.������,A.�Ǽ�ʱ��,A.����Ա From (" & strSQL1 & ") A" & _
                " Group by A.��������,A.��¼����,A.���շ�,A.NO,A.�ѱ�,A.��������,A.������,A.�Ǽ�ʱ��,A.����Ա"
        End If
        strSql = strSql & IIf(strSql <> "", " Union ALL ", "") & strSQL1
    End If

    '�����ò���(��ʾԭʼ������Ϣ��ʣ�ಿ�ݽ��)
    strSQL1 = _
    " Select 2 as ��������,A.��¼����,Decode(B.��¼״̬,0,0,1) as ���շ�," & _
    "       A.NO,B.�ѱ�,Sum(B.Ӧ�ս��) as Ӧ�ս��,Sum(B.ʵ�ս��) as ʵ�ս��," & _
    "       C.���� as ��������,B.������,Max(Decode(Floor(x.��¼����/10), 0, x.�Ǽ�ʱ��, Null)) As �Ǽ�ʱ��, " & _
    "       Max(Decode(Floor(x.��¼����/10), 0, Nvl(x.����Ա����, x.������), Null)) As ����Ա " & _
    " From ����ҽ������ A," & strFeeTab & " B,���ű� C," & strFeeTab & " X" & _
    " Where A.ҽ��ID IN(" & strҽ��IDs & ") And A.���ͺ�=[3]" & _
    "       And A.NO=B.NO And A.��¼����=Decode(B.��¼����,11,1,B.��¼����)" & _
    "       And A.ҽ��ID=B.ҽ����� And B.��������ID=C.ID" & _
    "       And B.NO=X.NO And B.��¼����=X.��¼���� And B.���=X.��� And X.��¼״̬ IN(0,1,3)" & _
    " Group by A.��¼����,Decode(B.��¼״̬,0,0,1),A.NO,B.�ѱ�,C.����,B.������"

    If mblnMoved Then
        strSQL1 = Replace(strSQL1, "����ҽ������", "H����ҽ������")
        strSQL1 = Replace(strSQL1, strFeeTab, "H" & strFeeTab)
    ElseIf blnDataMoved Then
        strSQL2 = Replace(strSQL1, "����ҽ������", "H����ҽ������")
        strSQL2 = Replace(strSQL2, strFeeTab, "H" & strFeeTab)
        strSQL1 = strSQL1 & " Union ALL " & strSQL2
        strSQL1 = _
        " Select A.��������,A.��¼����,A.���շ�,A.NO,A.�ѱ�," & _
        "       Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��," & _
        "       A.��������,A.������,A.�Ǽ�ʱ��,A.����Ա From (" & strSQL1 & ") A" & _
        " Group by A.��������,A.��¼����,A.���շ�,A.NO,A.�ѱ�,A.��������,A.������,A.�Ǽ�ʱ��,A.����Ա"
    End If
    strSql = strSql & IIf(strSql <> "", " Union ALL ", "") & strSQL1
    strWith = "With �����б� as (" & strSql & ") "

    '85232:���ϴ�,2015/5/29,��ȡ���˷��ü�¼ʱ�ų������ϵļ�¼
    '��ԭʼ��¼Ϊ׼��(��˵ĵǼ�ʱ���ִ��״̬)
    strSql = _
        " Select C.��������,A.NO,Decode(A.��¼����,11,1,A.��¼����) As ��¼����,A.��¼״̬,A.�ѱ�, " & _
        "        A.���,A.�۸񸸺�,A.ҽ�����,A.�շ����,A.�շ�ϸĿID,A.��������ID,A.������,A.ִ�в���ID, " & _
        "        Max(a.ִ��״̬) Over(Partition By a.��¼����,a.No,a.���) As ִ��״̬, " & _
        "        A.������ĿID,A.����,A.����,A.��׼����,A.Ӧ�ս��,A.ʵ�ս��, " & _
        "        A.����ʱ��,C.�Ǽ�ʱ��," & _
        "        C.����Ա as ����Ա����,C.���շ�" & _
        " From " & strFeeTab & " A,�����б� C" & _
        " Where Decode(A.��¼����,11,1,A.��¼����)= C.��¼���� And A.NO=C.NO "

    If mblnMoved Then
        strSql = Replace(strSql, strFeeTab, "H" & strFeeTab)
    ElseIf blnDataMoved Then
        strSql = strSql & " Union ALL " & Replace(strSql, strFeeTab, "H" & strFeeTab)
    End If
    strSql = strSql & GetNotFeeSQL(blnDataMoved, strIDs)

    strSql = strWith & vbCrLf & strSql
    '��ɾ�����˷�����,����ʾ
    '80752,Ƚ����,2014-12-23,��ʾ����ɾ�����˷����ʼ�¼
    '    " Having Nvl(Sum(A.Ӧ�ս��),0)<>0 Or Nvl(Sum(A.ʵ�ս��),0)<>0"
    strSql = "" & _
    " Select��decode(nvl(A.��������,1),1,'������','���ӷ���') as ��������, " & _
    "       A.��¼����,A.���շ� as �շѱ�־,decode(A.��¼����,1,'�շѵ�','���ʵ�') as ��������, " & _
    "       A.NO as ���ݺ�,A.�շ�ϸĿID," & _
    "       A.�ѱ�,L.���� as ��������,A.������, " & _
    "       C.���� as ���,Nvl(A.�۸񸸺�,A.���) as ���," & _
    "       Nvl(F.����,B.����)||Decode(B.���,NULL,NULL,' '||B.���) as ��Ŀ," & _
    "       A.��׼����" & IIf(blnҩ����λ, "*Nvl(E." & strҩ����װ & ",1)", "") & " as ����," & _
    "       Sum(Nvl(A.����,1)*A.����" & IIf(blnҩ����λ, "/Nvl(E." & strҩ����װ & ",1)", "") & ") as ����," & _
            IIf(blnҩ����λ, "Decode(E.ҩƷID,NULL,B.���㵥λ,E." & strҩ����λ & ")", "B.���㵥λ") & " as ���㵥λ," & _
    "       Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��,D.���� as ִ�в���," & _
    "       Decode(Nvl(A.ִ��״̬,0),0,'δִ��',1,'��ȫִ��',2,'����ִ��') as ִ�����, " & _
    "       Nvl(A.ִ��״̬,0) as ִ��״̬,a.�շ����," & _
    "       to_Char(A.�Ǽ�ʱ��,'yyyy-mm-dd hh24:mi:ss') as �Ǽ�ʱ��,A.����Ա����" & _
    " From (" & strSql & ") A,�շ���ĿĿ¼ B,�շ���Ŀ��� C," & _
    "      ���ű� D,���ű� L,ҩƷ��� E,�շ���Ŀ���� F" & _
    " Where A.�շ�ϸĿID=B.ID And A.�շ����=C.����  And A.��������ID=L.ID(+) And A.ִ�в���ID=D.ID(+)" & _
    "       And B.ID=E.ҩƷID(+) And A.�շ�ϸĿID=F.�շ�ϸĿID(+)" & _
    "       And F.����(+)=1 And F.����(+)=[4] And A.ҽ�����+0 IN(" & strҽ��IDs & ")" & _
    " Group by A.�շ�ϸĿID,A.��������,A.��¼����,A.���շ�,A.NO,A.�ѱ�,L.����,A.������,A.����Ա����, " & _
    "       to_Char(A.�Ǽ�ʱ��,'yyyy-mm-dd hh24:mi:ss'),Nvl(A.�۸񸸺�,A.���),C.���� , " & _
    "       Nvl(F.����,B.����),B.���,B.���㵥λ,D.����," & _
    "       A.��׼����,Nvl(A.ִ��״̬,0),a.�շ����,E.ҩƷID,Nvl(E." & strҩ����װ & ",1),E." & strҩ����λ & _
    " Order by  �������� Desc,��¼����,�Ǽ�ʱ�� Desc, No, ���"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Tittle, mTYAdviceProperty.lngҽ��ID, _
        mTYAdviceProperty.lng���ID, mTYAdviceProperty.lng���ͺ�, IIf(gSysPara.bytҩƷ������ʾ = 0, 1, 3))

    With vsExpense
        .Redraw = flexRDNone
        .Cols = 1: .FixedCols = 0
        .Rows = 2
        .MergeRow(1) = False
        Set .DataSource = rsTemp
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        If rsTemp.RecordCount = 0 Then
            .Subtotal flexSTClear
            .Rows = 2
            .Clear 1
        Else
            Call SetExpenseGridProperty '������������
        End If
    End With
    
    
    zl_vsGrid_Para_Save mlngModule, vsExpense, Tittle, "������Ϣ", True
    
    vsExpense.Redraw = flexRDBuffered
    LoadDataFromAdvices = True
    Exit Function
errHandle:
    vsExpense.Redraw = flexRDBuffered
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadAdvicePrice(ByVal blnMoved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ��ҽ���ļƼ۹�ϵ����ʱ��¼��
    '����:��ȡ�Ƽ۹�ϵ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-05-27 11:59:19
    '˵��:Ҫ�������ĿӦ�ò��Ƕ���,Ժ��ִ��,����Ʒ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsAdvice As New ADODB.Recordset
    Dim dbl���� As Double, bln�������� As Boolean
    Dim strSql As String, strIF As String, strPrice As String
    Dim blnHaveSub As Boolean, lng������ID As Long
    Dim cur�ϼ� As Currency, i As Long, j As Long
    Dim strҩƷ�۸�ȼ� As String, str���ļ۸�ȼ� As String, str��ͨ�۸�ȼ� As String
    Dim strWherePriceGrade As String

    Set mrsPrice = New ADODB.Recordset
    mrsPrice.Fields.Append "ҽ��ID", adBigInt
    mrsPrice.Fields.Append "��������ID", adBigInt
    mrsPrice.Fields.Append "���", adVarChar, 10
    mrsPrice.Fields.Append "�շ�ϸĿID", adBigInt
    mrsPrice.Fields.Append "���㵥λ", adVarChar, 100, adFldIsNullable
    mrsPrice.Fields.Append "��������", adInteger
    mrsPrice.Fields.Append "ִ�п���", adInteger
    mrsPrice.Fields.Append "������ĿID", adBigInt
    mrsPrice.Fields.Append "�վݷ�Ŀ", adVarChar, 50, adFldIsNullable
    mrsPrice.Fields.Append "����", adDouble
    mrsPrice.Fields.Append "����", adDouble
    mrsPrice.Fields.Append "Ӧ��", adCurrency
    mrsPrice.Fields.Append "ʵ��", adCurrency
    mrsPrice.Fields.Append "����", adInteger

    mrsPrice.CursorLocation = adUseClient
    mrsPrice.LockType = adLockOptimistic
    mrsPrice.CursorType = adOpenStatic
    mrsPrice.Open

    On Error GoTo errH

    '��ȡҪ���������õ�ҽ����¼
    '������鲿λ������������������ϵķ���
    If mTYAdviceProperty.bln����ִ�� Then
        strIF = " And B.ID=[1]"
    Else
        'F,D:���,����
        'C:����
        strIF = " And (B.ID=[1] Or (B.���ID=[1] And B.������� IN('F','D')) Or (B.���ID=[2] And B.�������='C'))"
    End If

    strSql = _
    " Select B.���,A.ҽ��ID,B.���ID,B.�������,B.������ĿID,B.��������ID,A.ִ�в���ID," & _
    "        Nvl(A.��������,Sum(Nvl(C.��������,0))) as ����,B.�걾��λ,B.��鷽��,B.ִ�б��,b.����ID,b.��ҳID" & _
    " From ����ҽ������ A,����ҽ����¼ B,����ҽ��ִ�� C" & _
    " Where Nvl(A.�Ʒ�״̬,0)=0 " & strIF & _
    "   And A.ҽ��ID=B.ID And A.���ͺ�=[3]" & _
    "   And C.ҽ��ID(+)=A.ҽ��ID And C.���ͺ�(+)=A.���ͺ�" & _
    " Group by B.���,A.ҽ��ID,B.���ID,B.�������,B.������ĿID,B.��������ID," & _
    "       A.ִ�в���ID,A.��������,B.�걾��λ,B.��鷽��,B.ִ�б��,b.����ID,b.��ҳID" & _
    " Having Nvl(A.��������,Sum(Nvl(C.��������,0)))<>0" & _
    " Order by ���"
    If blnMoved Then
        strSql = Replace(strSql, "����ҽ����¼", "H����ҽ����¼")
        strSql = Replace(strSql, "����ҽ������", "H����ҽ������")
        strSql = Replace(strSql, "����ҽ��ִ��", "H����ҽ��ִ��")
    End If
    Set rsAdvice = gobjDatabase.OpenSQLRecord(strSql, Tittle, mTYAdviceProperty.lngҽ��ID, mTYAdviceProperty.lng���ID, mTYAdviceProperty.lng���ͺ�)
    
    If rsAdvice.RecordCount > 0 Then
        'ҽ��IDֻ��һ�����϶���ͬһ������
        If GetPriceGradeStartType() > 0 Then
            Call GetPriceGrade(gstrNodeNo, Val(Nvl(rsAdvice!����ID)), Val(Nvl(rsAdvice!��ҳID)), "", strҩƷ�۸�ȼ�, str���ļ۸�ȼ�, str��ͨ�۸�ȼ�)
        End If
        If strҩƷ�۸�ȼ� <> "" Or str���ļ۸�ȼ� <> "" Or str��ͨ�۸�ȼ� <> "" Then
            strWherePriceGrade = _
                "      And ((Instr(';5;6;7;', ';' || c.��� || ';') > 0 And b.�۸�ȼ� = [8])" & vbNewLine & _
                "            Or (Instr(';4;', ';' || c.��� || ';') > 0 And b.�۸�ȼ� = [9])" & vbNewLine & _
                "            Or (Instr(';4;5;6;7;', ';' || c.��� || ';') = 0 And b.�۸�ȼ� = [10])" & vbNewLine & _
                "            Or (b.�۸�ȼ� Is Null" & vbNewLine & _
                "                And Not Exists (Select 1" & vbNewLine & _
                "                                From �շѼ�Ŀ" & vbNewLine & _
                "                                Where b.�շ�ϸĿid = �շ�ϸĿid And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
                "                                      And ((Instr(';5;6;7;', ';' || c.��� || ';') > 0 And �۸�ȼ� = [8])" & vbNewLine & _
                "                                            Or (Instr(';4;', ';' || c.��� || ';') > 0 And �۸�ȼ� = [9])" & vbNewLine & _
                "                                            Or (Instr(';4;5;6;7;', ';' || c.��� || ';') = 0 And �۸�ȼ� = [10])))))"
        Else
            strWherePriceGrade = " And b.�۸�ȼ� Is Null "
        End If
    End If
    
    For i = 1 To rsAdvice.RecordCount
        dbl���� = Nvl(rsAdvice!����, 0)

        '��ȡ��Ӧ���շѼ�Ŀ:ֻ��ȡ�̶�����,�Ҳ��Ǳ�۵Ķ���
        bln�������� = (rsAdvice!������� = "F" And Not IsNull(rsAdvice!���ID))
        '����û�мӲ�λ������������Ҫ��Distinct
        strPrice = "" & _
        "Select * From (" & _
        "   Select Distinct C.������ĿID,C.�շ���ĿID,C.��鲿λ,C.��鷽��,C.��������,C.�շ�����,C.���ж���,C.������Ŀ,C.�շѷ�ʽ,c.���ÿ���id" & _
        "               ,Max(Nvl(c.���ÿ���id, 0)) Over(Partition By c.������Ŀid, c.��鲿λ, c.��鷽��, c.��������) As Top" & _
        "   From �����շѹ�ϵ C Where C.������ĿID=[1]" & _
        "           And (C.���ÿ���ID is Null And C.������Դ = 0 or C.���ÿ���ID = [6] And C.������Դ = [7])" & _
        "   ) Where Nvl(���ÿ���id, 0) = Top"

        strSql = _
        " Select Nvl(A.������Ŀ,0) as ����,A.�շ���ĿID,A.�շ�����,B.������ĿID,D.�վݷ�Ŀ," & _
        "       C.���,C.���㵥λ,C.ִ�п���,Decode(C.�Ƿ���,1,B.ȱʡ�۸�,B.�ּ�) as ����,C.���ηѱ�," & _
                IIf(bln��������, "Nvl(B.�����շ���,100)/100", "1") & " as ������" & _
        " From (" & strPrice & ") A,�շѼ�Ŀ B,�շ���ĿĿ¼ C,������Ŀ D," & _
        "      (Select [1] as ������ĿID,Decode([2],0,Null,[2]) as ���ID," & _
        "               Decode([3],'None',Null,[3]) as �걾��λ,Decode([4],'None',Null,[4]) as ��鷽��,[5] as ִ�б�� From Dual " & _
        "       ) X" & _
        " Where A.������ĿID=X.������ĿID" & _
        "       And (   X.���ID is Null And X.ִ�б�� IN(1,2) And A.��������=1" & _
        "               Or X.�걾��λ=A.��鲿λ And X.��鷽��=A.��鷽�� And Nvl(A.��������,0)=0" & _
        "               Or X.��鷽�� is Null And Nvl(A.��������,0)=0 And A.��鲿λ is Null And A.��鷽�� is Null)" & _
        "       And A.�շ���ĿID=B.�շ�ϸĿID And A.�շ���ĿID=C.ID And B.������ĿID=D.ID" & _
        "       And (C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                strWherePriceGrade & vbNewLine & _
        "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
        "       And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
        "       And Nvl(A.���ж���,0)=1 And Nvl(C.�Ƿ���,0)=0" & _
        " Order By �շ���ĿID,����"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Tittle, Val(rsAdvice!������ĿID), Val(Nvl(rsAdvice!���ID, 0)), _
            CStr(Nvl(rsAdvice!�걾��λ, "None")), CStr(Nvl(rsAdvice!��鷽��, "None")), Val(Nvl(rsAdvice!ִ�б��, 0)), _
            mlngִ�п���ID, mTYAdviceProperty.int������Դ, strҩƷ�۸�ȼ�, str���ļ۸�ȼ�, str��ͨ�۸�ȼ�)

        blnHaveSub = False: lng������ID = 0: cur�ϼ� = 0
        If Not rsTemp.EOF And gSysPara.bln��������ۿ� Then
            rsTemp.Filter = "����=1"
            If Not rsTemp.EOF Then blnHaveSub = True
            rsTemp.Filter = "����=0"
            If Not rsTemp.EOF Then lng������ID = rsTemp!������ĿID
            rsTemp.Filter = 0
        End If

        For j = 1 To rsTemp.RecordCount
            mrsPrice.AddNew
            mrsPrice!ҽ��ID = rsAdvice!ҽ��ID
            mrsPrice!��������id = rsAdvice!��������id
            mrsPrice!��� = rsTemp!���
            mrsPrice!�շ�ϸĿID = rsTemp!�շ���ĿID
            mrsPrice!���㵥λ = Nvl(rsTemp!���㵥λ)
            mrsPrice!�������� = IIf(bln��������, 1, 0)
            mrsPrice!ִ�п��� = Nvl(rsTemp!ִ�п���, 0)
            mrsPrice!������ĿID = rsTemp!������ĿID
            mrsPrice!�վݷ�Ŀ = rsTemp!�վݷ�Ŀ
            mrsPrice!���� = Format(Nvl(rsTemp!����, 0), gSysPara.Price_Decimal.strFormt_VB)
            mrsPrice!���� = Format(Nvl(rsTemp!�շ�����, 0) * dbl����, "0.00000")
            mrsPrice!Ӧ�� = Format(mrsPrice!���� * mrsPrice!���� * rsTemp!������, gSysPara.Money_Decimal.strFormt_VB)
            mrsPrice!���� = rsTemp!����
            If gSysPara.bln��������ۿ� And blnHaveSub Then
                mrsPrice!ʵ�� = mrsPrice!Ӧ��
                cur�ϼ� = cur�ϼ� + mrsPrice!ʵ��
            ElseIf Nvl(rsTemp!���ηѱ�, 0) = 0 Then
                mrsPrice!ʵ�� = Format(ActualMoney(mTYAdviceProperty.str�ѱ�, mrsPrice!������ĿID, mrsPrice!Ӧ��, _
                    rsTemp!�շ���ĿID, Nvl(rsAdvice!ִ�в���ID, 0), Nvl(rsTemp!�շ�����, 0) * dbl����, 0), gSysPara.Money_Decimal.strFormt_VB)
            Else
                mrsPrice!ʵ�� = mrsPrice!Ӧ��
            End If
            mrsPrice.Update
            rsTemp.MoveNext
        Next

        If gSysPara.bln��������ۿ� And blnHaveSub And lng������ID <> 0 Then
            cur�ϼ� = Format(ActualMoney(mTYAdviceProperty.str�ѱ�, lng������ID, cur�ϼ�), gSysPara.Money_Decimal.strFormt_VB) - cur�ϼ�
            mrsPrice.Filter = "����=0"
            mrsPrice!ʵ�� = Nvl(mrsPrice!ʵ��, 0) + cur�ϼ�
            mrsPrice.Update
            mrsPrice.Filter = 0
        End If
        rsAdvice.MoveNext
    Next
    If mrsPrice.RecordCount > 0 Then mrsPrice.MoveFirst
    LoadAdvicePrice = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
    Set mrsPrice = Nothing
End Function

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=13,0,0,ҽ�����ѹ���
Public Property Get Tittle() As String
    Tittle = m_Tittle
End Property

Public Property Let Tittle(ByVal New_Tittle As String)
    m_Tittle = New_Tittle
    PropertyChanged "Tittle"
End Property


'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "����һ�� Font ����"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property


Public Sub zlPrintData(bytStyle As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������嵥
    '���:bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-05-29 17:00:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte, i As Long
    Dim lngRow As Long, lngCol As Long
    Dim strWidth As String
    
    If mTYAdviceProperty.lng����ID = 0 Then Exit Sub
    
    strSql = "" & _
    "Select Nvl(Nvl(B.����, C.����), A.����) ����, Nvl(Nvl(B.�Ա�, C.�Ա�), A.�Ա�) �Ա�, Nvl(Nvl(B.����, C.����), A.����) ����, A.�����, B.סԺ��" & vbNewLine & _
    "From ������Ϣ A, ������ҳ B, ���˹Һż�¼ C" & vbNewLine & _
    "Where A.����id = B.����id(+) And B.��ҳid(+) = [2] And A.����id = C.����id(+) And A.����� = C.�����(+) And A.����id = [1]"

    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Tittle, mTYAdviceProperty.lng����ID, mTYAdviceProperty.lng��ҳId)
    If rsTmp.EOF Then Exit Sub
    
    '��ͷ
    objOut.Title.Text = IIf(mbytFun = 1, "ҽԺ�����嵥", "ҽԺ�����嵥")
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    Set objRow = New zlTabAppRow
    objRow.Add "���ˣ�" & Nvl(rsTmp!����) & " �Ա�" & Nvl(rsTmp!�Ա�) & " ���䣺" & Nvl(rsTmp!����)
    If mTYAdviceProperty.lng��ҳId <> 0 Then
        objRow.Add "סԺ�ţ�" & Nvl(rsTmp!סԺ��)
    Else
        objRow.Add "����ţ�" & Nvl(rsTmp!�����)
    End If
    objOut.UnderAppRows.Add objRow
    
    '����
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(gobjDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    '����
    Set objOut.Body = vsExpense
    
    '���
    vsExpense.Redraw = False
    lngRow = vsExpense.Row: lngCol = vsExpense.Col
        
    strWidth = ""
    For i = 0 To vsExpense.Cols - 1
        strWidth = strWidth & "," & vsExpense.ColWidth(i)
        If i <= vsExpense.FixedCols - 1 Or vsExpense.ColHidden(i) Then
            vsExpense.ColWidth(i) = 0
        End If
    Next
        
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    strWidth = Mid(strWidth, 2)
    For i = 0 To vsExpense.Cols - 1
        vsExpense.ColWidth(i) = Split(strWidth, ",")(i)
    Next
    vsExpense.Row = lngRow: vsExpense.Col = lngCol
    vsExpense.Redraw = True
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Public Function zlBuildMainExpense(Optional ByVal frmMain As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������
    '����:���ɳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 13:37:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objParent As Object
    If mbytFun = 1 Then Exit Function
    Set objParent = frmMain
    If frmMain Is Nothing Then Set objParent = mfrmParent
    If InStr(mTYAdviceProperty.str�Ʒ�״̬, ",-1,") > 0 Then
        zlBuildMainExpense = FuncFeeMainAppend(objParent)
    Else
        zlBuildMainExpense = FuncFeeMain
    End If
End Function

Private Function FuncFeeMainAppend(ByVal frmMain As Object, Optional strOutNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������
    '����:strOutNos-����ɹ��ĵ��ݺ�
    '����:���ɳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 13:42:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mTYAdviceProperty.lngҽ��ID = 0 Then Exit Function
    If mTYAdviceProperty.intִ��״̬ = 1 Then
        MsgBox "��ִ����Ŀ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
        Exit Function
    End If
    If mblnMoved Then
        MsgBox "�ò��˵ı���" & IIf(mTYAdviceProperty.int������Դ = 2, "סԺ", "����") & "�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
            "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Function
    End If
    If frmTechnicExpense.EditCard(frmMain, mstrPrivsAnnexFee, 0, mTYAdviceProperty.lngҽ��ID, mTYAdviceProperty.lng���ͺ�, mTYAdviceProperty.lng����ID, mTYAdviceProperty.lng��ҳId, _
         IIf(mTYAdviceProperty.int������Դ = 2, 2, 1), mTYAdviceProperty.int��¼����, mTYAdviceProperty.lng��������ID, mTYAdviceProperty.lng���˿���ID, 0, "", mTYAdviceProperty.strNO, "", "", , , , strOutNos, _
         , , mobjSquareCard) Then
         FuncFeeMainAppend = True
         Call RefreshExpenseData
    End If
End Function

Private Function FuncFeeMain(Optional strOutNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������
    '����:strOutNos-����ɹ��ĵ��ݺ�
    '����:���ɳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 13:45:09
    '---------------------------------------------------------------------------------------------------------------------------------------------

 
    Dim rsPati As New ADODB.Recordset
    Dim rsAdvice As New ADODB.Recordset
    
    Dim int��Դ As Integer, lngҽ��ID As Long
    Dim int�۸񸸺� As Integer, lng��Ŀid As Long, lngִ�в���ID As Long
    Dim lng���˲���ID As Long, lng���˿���ID As Long, lng���ID As Long
    Dim arrSQL As Variant, arrCountSQL As Variant, strSql As String, strDate As String, i As Long, j As Long
    Dim int������Ŀ�� As Integer, lng���մ���ID As Long, str���ձ��� As String, curͳ���� As Currency, str�������� As String
    Dim lng��������ID As Long, str����ҽ�� As String, int��� As Integer, strMsg As String
    Dim int����� As Integer, strTmp As String
    Dim blnTrans  As Boolean
    
    If mTYAdviceProperty.lngҽ��ID = 0 Then Exit Function
    If mrsPrice Is Nothing Then Exit Function
    If mrsPrice.RecordCount = 0 Then
        MsgBox "��ִ����Ŀû�п��ԼƷѵ������á�" & vbCrLf & "�����Ҫ��������ֹ����丽�ӷ��á�", vbInformation, gstrSysName
        Exit Function
    End If
    If mTYAdviceProperty.intִ��״̬ = 1 Then
        MsgBox "��ִ����Ŀ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
        Exit Function
    End If
    If mblnMoved Then
        MsgBox "�ò��˵ı���" & IIf(mTYAdviceProperty.int������Դ = 2, "סԺ", "����") & "�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
            "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mTYAdviceProperty.int��¼���� = 1 Then
        If BillExistBalance(mTYAdviceProperty.strNO) Then
            MsgBox "���� " & mTYAdviceProperty.strNO & " �Ѿ��շѣ��������������ŵ��ݵ������á�" & vbCrLf & "�����Ҫ��������ֹ����丽�ӷ��á�", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf mTYAdviceProperty.int��¼���� = 2 Then
        'סԺ��Ժ���˷�������
        If mTYAdviceProperty.int������Դ = 2 Then
            If Not PatiCanBilling(mTYAdviceProperty.lng����ID, mTYAdviceProperty.lng��ҳId, mstrPrivsAnnexFee, pҽ�����ѹ���) Then Exit Function
        End If
    End If
    
    If MsgBox("ȷʵҪ���ɸ���Ŀ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Function
    End If
            
    int��Դ = IIf(mTYAdviceProperty.int������Դ = 2, 2, 1)
    
    Screen.MousePointer = 11
    
    '��ȡ���˵���Ϣ
    strSql = "Select Nvl(Nvl(B.����, C.����), A.����) ����, Nvl(Nvl(B.�Ա�, C.�Ա�), A.�Ա�) �Ա�, Nvl(Nvl(B.����, C.����), A.����) ����," & vbNewLine & _
            "       Nvl(B.�ѱ�, A.�ѱ�) As �ѱ�, A.�����, B.סԺ��, Nvl(A.��ǰ����, B.��Ժ����) As ����, Nvl(A.��ǰ����id, B.��ǰ����id) As ���˲���id," & vbNewLine & _
            "       Nvl(A.��ǰ����id, B.��Ժ����id) As ���˿���id, Nvl(B.����, A.����) As ����, D.���� As ������" & vbNewLine & _
            "From ������Ϣ A, ������ҳ B, ���˹Һż�¼ C, ҽ�Ƹ��ʽ D" & vbNewLine & _
            "Where A.����id = B.����id(+) And B.��ҳid(+) = [2] And A.����id = C.����id(+) And A.����� = C.�����(+) And A.ҽ�Ƹ��ʽ = D.����(+) And A.����id=[1]"

    On Error GoTo errH
    Set rsPati = gobjDatabase.OpenSQLRecord(strSql, Tittle, mTYAdviceProperty.lng����ID, mTYAdviceProperty.lng��ҳId)
    
    '���ܶ��շ���ΪҩƷ����
    If mTYAdviceProperty.int��¼���� = 1 Then
        lng���ID = ExistIOClass(8) '���ﻮ�۵�
    Else
        lng���ID = ExistIOClass(9) '����/סԺ���ʵ�
    End If
    
    '���ܷ���ʱ���Զ������˲���������,�������ֹ�����ʣ�ಿ�ݡ�
    '1.��Ϊ���ݺ���ͬ,����Ҫ�����������
    '2.����������շѻ��۵���Ҫ��֤һ�ŵ����еǼ�ʱ����ͬ(��Ȼ�շ��޷�����)
    '3.��2����������������������Ѿ��շѣ�������������������
    int��� = GetBillMax���(mTYAdviceProperty.strNO, mTYAdviceProperty.int��¼����, strDate, IIf(mTYAdviceProperty.int������Դ = 2, 2, 1))
    If mTYAdviceProperty.int��¼���� = 2 Or strDate = "" Then
        strDate = "To_Date('" & Format(gobjDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    Else
        strDate = "To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS')"
    End If
    
    arrSQL = Array()
    arrCountSQL = Array()
    With mrsPrice
        .MoveFirst
        For i = 1 To .RecordCount
            '��ȡ��Ӧ��ҽ����Ϣ
            If lngҽ��ID <> !ҽ��ID Then
                strSql = "Select ҽ����Ч,���˿���ID,��������ID,����ҽ��,Ӥ��,ִ��Ƶ��,�Ƽ����� From ����ҽ����¼ Where ID=[1]"
                Set rsAdvice = gobjDatabase.OpenSQLRecord(strSql, Tittle, Val(!ҽ��ID))
                
                '����ǰ�����Ʒ�ҽ�����Ϊ�ѼƷ�
                ReDim Preserve arrCountSQL(UBound(arrCountSQL) + 1)
                arrCountSQL(UBound(arrCountSQL)) = "ZL_����ҽ������_�Ʒ�(" & !ҽ��ID & "," & mTYAdviceProperty.lng���ͺ� & ")"
                
                int����� = 0
            End If
            lngҽ��ID = !ҽ��ID
            
            '���˲�������
            lng���˲���ID = Nvl(rsPati!���˲���ID, 0)
            lng���˿���ID = Nvl(rsPati!���˿���id, 0)
            If lng���˿���ID = 0 Then
                lng���˲���ID = Nvl(rsAdvice!���˿���id, 0)
                lng���˿���ID = Nvl(rsAdvice!���˿���id, 0)
            End If
            If lng���˿���ID = 0 Then
                lng���˲���ID = UserInfo.����ID
                lng���˿���ID = UserInfo.����ID
            End If
            
            '�������Ҽ�������
            lng��������ID = rsAdvice!��������id
            str����ҽ�� = rsAdvice!����ҽ��
            
            'ÿ���շ���Ŀ�Ĵ���
            If lng��Ŀid <> !�շ�ϸĿID Then
                int�۸񸸺� = int��� '��ȡ�۸񸸺�
                If !���� = 0 Then int����� = int���  '��ȡPriceʱ�Ƿ�ҽ������������ǰ���
                lngִ�в���ID = Get�շ�ִ�п���ID(mTYAdviceProperty.lng����ID, mTYAdviceProperty.lng��ҳId, !���, !�շ�ϸĿID, !ִ�п���, Nvl(rsAdvice!���˿���id, 0), Nvl(rsAdvice!��������id, 0), int��Դ)
                            
                '��ȡ������Ŀ��Ϣ
                If int��Դ = 2 And Not IsNull(rsPati!����) Then
                    strMsg = gclsInsure.GetItemInsure(mTYAdviceProperty.lng����ID, !�շ�ϸĿID, !ʵ��, False, rsPati!����, "||" & !����)
                    If strMsg <> "" Then
                        int������Ŀ�� = Val(Split(strMsg, ";")(0))
                        lng���մ���ID = Val(Split(strMsg, ";")(1))
                        curͳ���� = Format(Val(Split(strMsg, ";")(2)), gSysPara.Money_Decimal.strFormt_VB)
                        str���ձ��� = CStr(Split(strMsg, ";")(3))
                        If UBound(Split(strMsg, ";")) >= 5 Then
                            If Split(strMsg, ";")(5) <> "" Then
                                str�������� = Split(strMsg, ";")(5)
                            End If
                        End If
                    End If
                End If
            End If
            lng��Ŀid = !�շ�ϸĿID
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If int��Դ = 1 Then
                If mTYAdviceProperty.int��¼���� = 1 Then
                    '�������ﻮ�۵���
                    arrSQL(UBound(arrSQL)) = lng��Ŀid & ";" & _
                        "zl_���ﻮ�ۼ�¼_Insert('" & mTYAdviceProperty.strNO & "'," & int��� & "," & mTYAdviceProperty.lng����ID & ",NULL," & _
                        IIf(IsNull(rsPati!�����), "NULL", "'" & rsPati!����� & "'") & ",'" & Nvl(rsPati!������) & "','" & Nvl(rsPati!����) & "'," & _
                        "'" & Nvl(rsPati!�Ա�) & "','" & Nvl(rsPati!����) & "','" & Nvl(rsPati!�ѱ�) & "',NULL," & _
                        lng���˿���ID & "," & lng��������ID & ",'" & str����ҽ�� & "'," & _
                        IIf(Val(Nvl(!����)) = 1, ZVal(int�����), "NULL") & "," & lng��Ŀid & ",'" & !��� & "','" & !���㵥λ & "',NULL,1," & !���� & "," & _
                        !�������� & "," & ZVal(lngִ�в���ID) & "," & IIf(int�۸񸸺� = int���, "NULL", int�۸񸸺�) & "," & _
                        !������ĿID & ",'" & Nvl(!�վݷ�Ŀ) & "'," & !���� & "," & !Ӧ�� & "," & !ʵ�� & "," & _
                        strDate & "," & strDate & ",NULL,'" & UserInfo.���� & "',NULL," & _
                        !ҽ��ID & ",'" & Nvl(rsAdvice!ִ��Ƶ��) & "',NULL,NULL," & Nvl(rsAdvice!ҽ����Ч, 0) & "," & _
                        Nvl(rsAdvice!�Ƽ�����, 0) & ",1,'" & str���ձ��� & "','" & str�������� & "'," & int������Ŀ�� & "," & ZVal(lng���մ���ID) & "," & _
                        "NULL,0,NULL,NULL," & ZVal(lng���˲���ID) & ")"
                Else
                    '����������ʵ���
                    arrSQL(UBound(arrSQL)) = lng��Ŀid & ";" & _
                        "zl_������ʼ�¼_Insert('" & mTYAdviceProperty.strNO & "'," & int��� & "," & mTYAdviceProperty.lng����ID & "," & _
                        IIf(IsNull(rsPati!�����), "NULL", "'" & rsPati!����� & "'") & ",'" & Nvl(rsPati!����) & "','" & Nvl(rsPati!�Ա�) & "'," & _
                        "'" & Nvl(rsPati!����) & "','" & Nvl(rsPati!�ѱ�) & "',NULL," & ZVal(rsAdvice!Ӥ��) & "," & _
                        lng���˿���ID & "," & lng��������ID & "," & _
                        "'" & str����ҽ�� & "'," & IIf(!���� = 1, ZVal(int�����), "NULL") & "," & lng��Ŀid & ",'" & !��� & "'," & _
                        "'" & !���㵥λ & "',1," & !���� & "," & !�������� & "," & ZVal(lngִ�в���ID) & "," & _
                        IIf(int�۸񸸺� = int���, "NULL", int�۸񸸺�) & "," & !������ĿID & ",'" & Nvl(!�վݷ�Ŀ) & "'," & !���� & "," & _
                        !Ӧ�� & "," & !ʵ�� & "," & strDate & "," & strDate & ",NULL,NULL,'" & UserInfo.��� & "'," & _
                        "'" & UserInfo.���� & "',NULL,NULL," & !ҽ��ID & "," & _
                        "'" & Nvl(rsAdvice!ִ��Ƶ��) & "',NULL,NULL," & Nvl(rsAdvice!ҽ����Ч, 0) & "," & _
                        Nvl(rsAdvice!�Ƽ�����, 0) & ",1,NULL,0,NULL," & ZVal(mTYAdviceProperty.lng��ҳId) & "," & ZVal(lng���˲���ID) & ")"
                End If
            Else
                '����סԺ���ʵ���
                arrSQL(UBound(arrSQL)) = lng��Ŀid & ";" & _
                    "zl_סԺ���ʼ�¼_Insert('" & mTYAdviceProperty.strNO & "'," & int��� & "," & mTYAdviceProperty.lng����ID & "," & ZVal(mTYAdviceProperty.lng��ҳId) & "," & _
                    IIf(IsNull(rsPati!סԺ��), "NULL", "'" & rsPati!סԺ�� & "'") & ",'" & Nvl(rsPati!����) & "','" & Nvl(rsPati!�Ա�) & "'," & _
                    "'" & Nvl(rsPati!����) & "','" & Nvl(rsPati!����) & "','" & Nvl(rsPati!�ѱ�) & "'," & _
                    lng���˲���ID & "," & lng���˿���ID & ",NULL," & ZVal(rsAdvice!Ӥ��) & "," & _
                    lng��������ID & ",'" & str����ҽ�� & "'," & IIf(!���� = 1, ZVal(int�����), "NULL") & "," & lng��Ŀid & ",'" & !��� & "'," & _
                    "'" & !���㵥λ & "'," & int������Ŀ�� & "," & ZVal(lng���մ���ID) & ",'" & str���ձ��� & "'," & _
                    "1," & !���� & "," & !�������� & "," & ZVal(lngִ�в���ID) & "," & _
                    IIf(int�۸񸸺� = int���, "NULL", int�۸񸸺�) & "," & !������ĿID & ",'" & Nvl(!�վݷ�Ŀ) & "'," & !���� & "," & _
                    !Ӧ�� & "," & !ʵ�� & "," & curͳ���� & "," & strDate & "," & strDate & ",NULL,NULL," & _
                    "'" & UserInfo.��� & "','" & UserInfo.���� & "',NULL," & ZVal(lng���ID) & ",NULL,NULL,NULL," & _
                    !ҽ��ID & ",'" & Nvl(rsAdvice!ִ��Ƶ��) & "',NULL,NULL," & Nvl(rsAdvice!ҽ����Ч, 0) & "," & _
                    Nvl(rsAdvice!�Ƽ�����, 0) & ",NULL,'" & str�������� & "')"
            End If
            
            int��� = int��� + 1
            
            .MoveNext
        Next
    End With
    
     '��SQL���а��շ�ϸĿID����
    For i = 0 To UBound(arrSQL) - 1
        For j = i + 1 To UBound(arrSQL)
            If CLng(Split(arrSQL(j), ";")(0)) < CLng(Split(arrSQL(i), ";")(0)) Then
                strTmp = CStr(arrSQL(j))
                arrSQL(j) = arrSQL(i)
                arrSQL(i) = strTmp
            End If
        Next
    Next
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    
    For i = 0 To UBound(arrCountSQL)
        Call gobjDatabase.ExecuteProcedure(CStr(arrCountSQL(i)), Tittle)
    Next
    
    For i = 0 To UBound(arrSQL)
        Call gobjDatabase.ExecuteProcedure(CStr(Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1)), Tittle)
    Next
    
    '���ύǰ����ҽ������
    If int��Դ = 2 And Not IsNull(rsPati!����) Then
        If gclsInsure.GetCapability(support�����ϴ�, mTYAdviceProperty.lng����ID, rsPati!����) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, mTYAdviceProperty.lng����ID, rsPati!����) Then
            strMsg = ""
            If Not gclsInsure.TranChargeDetail(2, mTYAdviceProperty.strNO, 2, 1, strMsg, , rsPati!����) Then
                gcnOracle.RollbackTrans
                If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                Screen.MousePointer = 0: Exit Function
            End If
        End If
    End If
    
    gcnOracle.CommitTrans: blnTrans = False
    strOutNos = mTYAdviceProperty.strNO
    
    '���ύ�����ҽ������
    If int��Դ = 2 And Not IsNull(rsPati!����) Then
        If gclsInsure.GetCapability(support�����ϴ�, mTYAdviceProperty.lng����ID, rsPati!����) And gclsInsure.GetCapability(support������ɺ��ϴ�, mTYAdviceProperty.lng����ID, rsPati!����) Then
            strMsg = ""
            If Not gclsInsure.TranChargeDetail(2, mTYAdviceProperty.strNO, 2, 1, strMsg, , rsPati!����) Then
                If strMsg <> "" Then
                    MsgBox strMsg, vbInformation, gstrSysName
                Else
                    MsgBox "����""" & mTYAdviceProperty.strNO & """��������ҽ������ʧ��,�õ����ѱ��棡", vbInformation, gstrSysName
                End If
            End If
        End If
    End If
    On Error GoTo 0
    Screen.MousePointer = 0
    FuncFeeMain = True
    MsgBox "ִ����Ŀ�����������ɳɹ���", vbInformation, gstrSysName
    Call RefreshExpenseData
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function


Public Function zlFuncFeeNewPrice(ByVal frmMain As Object, Optional strOutNos As String, _
    Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���丽���շѵ���
    '���:objMain-���õ�������
    '����:strOutNos-�ɹ����շѵ���
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 14:08:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mTYAdviceProperty.lngҽ��ID = 0 Then Exit Function
    
    If mTYAdviceProperty.intִ��״̬ = 1 Then
        MsgBox "��ִ����Ŀ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
        Exit Function
    End If
    If mblnMoved Then
        MsgBox "�ò��˵ı���" & IIf(mTYAdviceProperty.int������Դ = 2, "סԺ", "����") & "�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
            "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Function
    End If
    
    If frmTechnicExpense.EditCard(frmMain, mstrPrivsAnnexFee, 0, mTYAdviceProperty.lngҽ��ID, mTYAdviceProperty.lng���ͺ�, mTYAdviceProperty.lng����ID, mTYAdviceProperty.lng��ҳId, _
         IIf(mTYAdviceProperty.int������Դ = 2, 2, 1), 1, mlngִ�п���ID, mTYAdviceProperty.lng���˿���ID, 0, "", "", "", "", , , , strOutNos, , objSaveData, mobjSquareCard) Then
         zlFuncFeeNewPrice = True
         Call Refresh
         Call RefreshExpenseData
    End If

End Function
Public Function zlFuncFeeNewBilling(ByVal frmMain As Object, Optional strOutNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ʵ���
    '����:strOutNos-���سɹ����ɵĵ���
    '����:���ѳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 16:03:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mTYAdviceProperty.lngҽ��ID = 0 Then Exit Function
    
    If mTYAdviceProperty.intִ��״̬ = 1 Then
        MsgBox "��ִ����Ŀ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mblnMoved Then
        MsgBox "�ò��˵ı���" & IIf(mTYAdviceProperty.int������Դ = 2, "סԺ", "����") & "�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
            "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Function
    End If
    
    If frmTechnicExpense.EditCard(frmMain, mstrPrivsAnnexFee, 0, mTYAdviceProperty.lngҽ��ID, mTYAdviceProperty.lng���ͺ�, mTYAdviceProperty.lng����ID, mTYAdviceProperty.lng��ҳId, _
         IIf(mTYAdviceProperty.int������Դ = 2, 2, 1), 2, mlngִ�п���ID, mTYAdviceProperty.lng���˿���ID, 0, "", "", "", "", , , , strOutNos, , , mobjSquareCard) Then
         zlFuncFeeNewBilling = True
         Call Refresh
         Call RefreshExpenseData
    End If
End Function
Public Function zlFuncFeeNewNull(ByVal frmMain As Object, Optional strOutNos As String, _
    Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ķ���
    '����:strOutNOs-����ɹ��ĵ��ݺ�
    '����:���ѳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 16:02:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mTYAdviceProperty.lngҽ��ID = 0 Then Exit Function
    If mTYAdviceProperty.intִ��״̬ = 1 Then
        MsgBox "��ִ����Ŀ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
        Exit Function
    End If
    If mblnMoved Then
        MsgBox "�ò��˵ı���" & IIf(mTYAdviceProperty.int������Դ = 2, "סԺ", "����") & "�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
            "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Function
    End If
    
    If frmTechnicExpense.EditCard(frmMain, mstrPrivsAnnexFee, 0, mTYAdviceProperty.lngҽ��ID, mTYAdviceProperty.lng���ͺ�, mTYAdviceProperty.lng����ID, mTYAdviceProperty.lng��ҳId, _
         IIf(mTYAdviceProperty.int������Դ = 2, 2, 1), 2, mlngִ�п���ID, mTYAdviceProperty.lng���˿���ID, 0, "", "", "", "", , , True, strOutNos, , objSaveData, mobjSquareCard) Then
         zlFuncFeeNewNull = True
         Call Refresh
        Call RefreshExpenseData
    End If
End Function

Public Function zlFuncFeeModi(Optional frmMain As Object, Optional int������Դ As Integer, Optional int��¼���� As Integer, Optional strNO As String, Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ķ�
    '���:int������Դ-1-����;2-סԺ
    '     int��¼����-��¼����
    '     strNO-��ָ���ĵ��ݽ��иķ�
    '����:
    '����:�޸ĳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 17:48:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln��� As Boolean
    Dim strFeeTab As String
    If strNO <> "" Then
      '�����ݺŽ��иķ�
      If int������Դ = 1 Then   '����
        strFeeTab = "������ü�¼"
      Else
        strFeeTab = "סԺ���ü�¼"
      End If
        If gobjDatabase.NOMoved(strFeeTab, strNO, "��¼����=", int��¼����) Then
            MsgBox "���õ��� " & strNO & " �Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
            Exit Function
        End If
 
        If int��¼���� = 2 Then
            If Not BillIdentical(strNO, IIf(int������Դ = 2, 2, 1)) Then
                MsgBox "����""" & strNO & """�а�������δ��˻�ֶ����˵����ݣ��������޸ġ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        If BillExistDelete(strNO, int��¼����, IIf(int������Դ = 2, 2, 1)) Then
            MsgBox "�õ��ݰ�����" & IIf(int��¼���� = 1, "�˷�", "����") & "����,�������޸ģ�", vbInformation, gstrSysName
            Exit Function
        End If
        
        '�����������ִ�л�ȫ��ִ�е���Ŀ,��һ������ȫ������,�������޸�
        If HaveExecute(strNO, int��¼����, False, IIf(int������Դ = 2, 2, 1)) Then
            MsgBox "�õ����а�����ȫִ�л򲿷�ִ�е���Ŀ,�������޸ģ�", vbInformation, gstrSysName
            Exit Function
        End If
        
        If int��¼���� = 2 Then
           bln��� = BillisZeroLog(strNO, IIf(int������Դ = 2, 2, 1))
        End If
        
        If zlIs��������(strNO, int��¼����) Then
            If frmStuffCharge.zlBillEdit(frmMain, 0, pҽ�����ѹ���, mstrPrivsAnnexFee, int��¼����, strNO, IIf(int������Դ = 2, 2, 1), 0, 0, _
                0, 0, , , , , 0, 0, , , , , , objSaveData, mobjSquareCard) = False Then Exit Function
            zlFuncFeeModi = True
            Exit Function
        End If
        
        If Not frmTechnicExpense.EditCard(frmMain, mstrPrivsAnnexFee, 0, 0, 0, 0, 0, _
             IIf(int������Դ = 2, 2, 1), int��¼����, 0, 0, 0, "", "", strNO, "", , , bln���, , , objSaveData, mobjSquareCard) Then Exit Function
        zlFuncFeeModi = True
        Exit Function
    End If
    If mTYAdviceProperty.lngҽ��ID = 0 Then Exit Function
    
    If vsExpense.TextMatrix(vsExpense.Row, 0) = "������" And mTYAdviceProperty.lng�Ƽ����� <> 1 Then
        MsgBox "ִ����Ŀ�������ò����޸ġ������Ҫ��������ֹ����丽�ӷ��á�", vbInformation, gstrSysName
        Exit Function
    End If
    If mTYAdviceProperty.intִ��״̬ = 1 Then
        MsgBox "��ִ����Ŀ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
        Exit Function
    End If
    If mblnMoved Then
        MsgBox "�ò��˵ı���" & IIf(mTYAdviceProperty.int������Դ = 2, "סԺ", "����") & "�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
            "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Function
    End If

    With vsExpense
        '78225:���ϴ�,2014/9/24,��ȡ��ȷ�ĵ��ݺźͼ�¼����
        strNO = Get���ݺ�
        If strNO = "" Or strNO = "[δ�Ʒ�]" Then Exit Function
        int��¼���� = Get��¼����
        
        If gobjDatabase.DateMoved(mTYAdviceProperty.dat����ʱ��) Then
            If gobjDatabase.NOMoved(mTYAdviceProperty.strFeeTab, strNO, "��¼����=", int��¼����) Then
                MsgBox "���õ��� " & strNO & " �Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                    "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        If int��¼���� = 2 Then
            If Not BillIdentical(strNO, IIf(mTYAdviceProperty.int������Դ = 2, 2, 1)) Then
                MsgBox "����""" & strNO & """�а�������δ��˻�ֶ����˵����ݣ��������޸ġ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        If Val(.TextMatrix(.Row, .ColIndex("��¼����"))) Mod 10 = 1 _
                        And Val(.TextMatrix(.Row, .ColIndex("�շѱ�־"))) = 1 Then
            MsgBox "�õ����Ѿ��շѣ��������޸ġ�", vbInformation, gstrSysName
            Exit Function
        End If
        If BillExistDelete(strNO, int��¼����, IIf(mTYAdviceProperty.int������Դ = 2, 2, 1)) Then
            MsgBox "�õ��ݰ�����" & IIf(int��¼���� = 1, "�˷�", "����") & "����,�������޸ģ�", vbInformation, gstrSysName
            Exit Function
        End If
        '�����������ִ�л�ȫ��ִ�е���Ŀ,��һ������ȫ������,�������޸�
        If HaveExecute(strNO, int��¼����, False, IIf(mTYAdviceProperty.int������Դ = 2, 2, 1)) Then
            MsgBox "�õ����а�����ȫִ�л򲿷�ִ�е���Ŀ,�������޸ģ�", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    If int��¼���� = 2 Then
       bln��� = BillisZeroLog(strNO, IIf(mTYAdviceProperty.int������Դ = 2, 2, 1))
    End If
    
    If zlIs��������(strNO, int��¼����) Then
        If frmStuffCharge.zlBillEdit(frmMain, 0, pҽ�����ѹ���, mstrPrivsAnnexFee, int��¼����, strNO, IIf(mTYAdviceProperty.int������Դ = 2, 2, 1), mTYAdviceProperty.lng����ID, mTYAdviceProperty.lng��ҳId, _
            mlngִ�п���ID, mTYAdviceProperty.lng���˿���ID, , , , , mTYAdviceProperty.lngҽ��ID, mTYAdviceProperty.lng���ͺ�, , , , , , objSaveData, mobjSquareCard) Then
            zlFuncFeeModi = True
            Call Refresh
            Call RefreshExpenseData
        End If
        Exit Function
    End If
    '78225:���ϴ�,2014/9/24,��ȡ��ȷ�ĵ��ݺźͼ�¼����
    If frmTechnicExpense.EditCard(frmMain, mstrPrivsAnnexFee, 0, mTYAdviceProperty.lngҽ��ID, mTYAdviceProperty.lng���ͺ�, mTYAdviceProperty.lng����ID, mTYAdviceProperty.lng��ҳId, _
        IIf(mTYAdviceProperty.int������Դ = 2, 2, 1), int��¼����, mlngִ�п���ID, mTYAdviceProperty.lng���˿���ID, 0, "", "", strNO, "", , , bln���, , , objSaveData, mobjSquareCard) Then
        zlFuncFeeModi = True
        Call Refresh
        Call RefreshExpenseData
    End If
End Function

Public Function zlFuncFeeDel(Optional frmMain As Object, Optional int������Դ As Integer, _
    Optional int��¼���� As Integer, Optional strNO As String, Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɾ������
    '���:int������Դ-1-����;2-סԺ
    '     int��¼����-��¼����
    '     strNO-��ָ���ĵ��ݽ��иķ�
    '����:
    '����:ɾ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 18:06:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFeeTab As String
    If strNO <> "" Then
        '�����ݺŽ��иķ�
        If int������Դ = 1 Then   '����
          strFeeTab = "������ü�¼"
        Else
          strFeeTab = "סԺ���ü�¼"
        End If
      
        '�����ݸķ�
        If gobjDatabase.NOMoved(strFeeTab, strNO, "��¼����=", int��¼����) Then
            MsgBox "���õ��� " & strNO & " �Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
            Exit Function
        End If
        
        If int��¼���� = 2 Then
            If Not BillIdentical(strNO, IIf(mTYAdviceProperty.int������Դ = 2, 2, 1)) Then
                MsgBox "����""" & strNO & """�а�������δ��˻�ֶ����˵����ݣ�������ɾ����", vbInformation, gstrSysName
                Exit Function
            End If
            'סԺ��Ժ���˷�������
            If int������Դ = 2 Then
                If Not PatiCanBilling(mTYAdviceProperty.lng����ID, mTYAdviceProperty.lng��ҳId, mstrPrivsAnnexFee, pҽ�����ѹ���) Then Exit Function
            End If
        End If
        If zlIs��������(strNO, int��¼����) Then
            If frmStuffCharge.zlBillEdit(frmMain, 3, pҽ�����ѹ���, mstrPrivsAnnexFee, int��¼����, strNO, IIf(mTYAdviceProperty.int������Դ = 2, 2, 1), _
                , , , , , , , , 0, , , , , , , objSaveData, mobjSquareCard) = False Then Exit Function
            zlFuncFeeDel = True
            Exit Function
        End If
        
        If Not frmTechnicExpense.EditCard(frmMain, mstrPrivsAnnexFee, 3, 0, 0, 0, 0, _
             IIf(int������Դ = 2, 2, 1), int��¼����, 0, 0, 0, "", "", strNO, "", , , False, , , objSaveData, mobjSquareCard) Then Exit Function
        zlFuncFeeDel = True
        Exit Function
    End If
    
    If mTYAdviceProperty.lngҽ��ID = 0 Then Exit Function
    If mTYAdviceProperty.intִ��״̬ = 1 Then
        MsgBox "��ִ����Ŀ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
        Exit Function
    End If
    If mblnMoved Then
        MsgBox "�ò��˵ı���" & IIf(mTYAdviceProperty.int������Դ = 2, "סԺ", "����") & "�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
            "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Function
    End If
    
    With vsExpense
        '78225:���ϴ�,2014/9/24,��ȡ��ȷ�ĵ��ݺźͼ�¼����
        strNO = .TextMatrix(.Row, .ColIndex("���ݺ�"))
        If strNO = "" Or strNO = "[δ�Ʒ�]" Then Exit Function
        int��¼���� = Get��¼����
        
       If gobjDatabase.DateMoved(mTYAdviceProperty.dat����ʱ��) Then
            If gobjDatabase.NOMoved(mTYAdviceProperty.strFeeTab, strNO, "��¼����=", int��¼����) Then
                MsgBox "���õ��� " & strNO & " �Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                    "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If int��¼���� = 2 Then
            If Not BillIdentical(strNO, IIf(mTYAdviceProperty.int������Դ = 2, 2, 1)) Then
                MsgBox "����""" & strNO & """�а�������δ��˻�ֶ����˵����ݣ�������ɾ����", vbInformation, gstrSysName
                Exit Function
            End If
            'סԺ��Ժ���˷�������
            If mTYAdviceProperty.int������Դ = 2 Then
                If Not PatiCanBilling(mTYAdviceProperty.lng����ID, mTYAdviceProperty.lng��ҳId, mstrPrivsAnnexFee, pҽ�����ѹ���) Then Exit Function
            End If
        End If
        If Val(.TextMatrix(.Row, .ColIndex("��¼����"))) Mod 10 = 1 _
                        And Val(.TextMatrix(.Row, .ColIndex("�շѱ�־"))) = 1 Then
            MsgBox "�õ����Ѿ��շѣ�������ɾ����", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    If vsExpense.TextMatrix(vsExpense.Row, vsExpense.ColIndex("��������")) = "������" Then
        If InStr(mstrPrivsAnnexFee, "ɾ��������") = 0 Then
            MsgBox "��û��ɾ�������õ�Ȩ�ޣ�����ɾ�������á�", vbInformation, gstrSysName
            Exit Function
        Else
            If MsgBox("������ɾ���������²���,��ȷʵҪɾ����Ŀ��������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    End If
    
    If zlIs��������(strNO, int��¼����) Then
    
        If frmStuffCharge.zlBillEdit(frmMain, 3, pҽ�����ѹ���, mstrPrivsAnnexFee, int��¼����, strNO, _
            IIf(mTYAdviceProperty.int������Դ = 2, 2, 1), , , , , , , , , IIf(mTYAdviceProperty.bln����ִ��, mTYAdviceProperty.lngҽ��ID, 0), _
            , , , , , , objSaveData, mobjSquareCard) = False Then
            Exit Function
        End If
        zlFuncFeeDel = True
        Call Refresh
        Call RefreshExpenseData
        Exit Function
    End If
    
    If frmTechnicExpense.EditCard(frmMain, mstrPrivsAnnexFee, 3, IIf(mTYAdviceProperty.bln����ִ��, mTYAdviceProperty.lngҽ��ID, 0), mTYAdviceProperty.lng���ͺ�, mTYAdviceProperty.lng����ID, mTYAdviceProperty.lng��ҳId, _
         IIf(mTYAdviceProperty.int������Դ = 2, 2, 1), int��¼����, 0, 0, 0, "", "", strNO, "", , , False, , , objSaveData, mobjSquareCard) Then
         zlFuncFeeDel = True
         Call Refresh
         Call RefreshExpenseData
    End If
    
End Function

Public Function zlFuncExtraFeeExe(ByVal frmMain As Object, ByVal bytType As Byte, ByVal strMainPrivs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ִ����ȡ��ִ��
    '���:bytType=0-ȡ��ִ��,1-ִ��
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 18:19:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim objParent As Object
    Dim strSql As String
    Dim strNO As String, int��¼���� As Integer, blnIsAbnormal As Boolean, blnTrans As Boolean
    Dim i As Long, blnDo As Boolean, str��� As String, strDate As String
    Dim arr���() As String, arrSQL As Variant
    Dim str��� As String, str����� As String, curMoney As Currency
    Dim blnRefresh As Boolean, lngRow As Long
    Dim blnJudge As Boolean, blnTrace As Boolean, blnHaveDrug As Boolean, blnHave���� As Boolean, strMsg As String
    
    '1.�������ȫ��ҩƷ����Զ����ϵĸ����������ģ��򲻴���ִ�У���Ϊ��Щ�Ƿ�ҩ�ű�ʾִ�С�
    '2.�������õ�������ִ��ʱ����ϵͳ���������Ƿ��Զ����ϡ�
    blnDo = False
    If mTYAdviceProperty.int��˱�־ >= 1 And gSysPara.byt������˷�ʽ = 1 Then
        MsgBox "�ò��˵ķ���������˽׶Σ����������ҽ���ͷ��á�", vbInformation, gstrSysName
        Exit Function
    End If
    '78789:���ϴ�,2014/10/22,����ִ����ȡ��ִ��ʱ��ֻ���ѡ��ĵ���
    strNO = Get���ݺ�
    int��¼���� = Get��¼����
    
    If strNO = "" Then
        MsgBox "��ǰ�޵�����Ϣ������" & IIf(bytType = 0, "ȡ��", "") & "ִ�еǼ�", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If

    With vsExpense
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpText, i, .ColIndex("���ݺ�")) = strNO Then
                lngRow = lngRow + 1
                If InStr(",����ҩ,�г�ҩ,�в�ҩ,", "," & .Cell(flexcpText, i, .ColIndex("���")) & ",") = 0 Then
                    blnJudge = True
                    If .Cell(flexcpText, i, .ColIndex("���")) = "����" And Not gSysPara.bln����ִ�з��� Then
                        '�ж��Ƿ����ĸ�������
                        strSql = " Select 1  From �������� where ����ID=[1] And ��������=1 "
                        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Tittle, Val(.Cell(flexcpText, i, .ColIndex("��Ŀ"))))
                        If rsTmp.RecordCount = 0 Then
                            blnJudge = True
                        Else
                            blnHave���� = True
                            blnJudge = False
                        End If
                    End If
                    If blnJudge Then
                        blnDo = True
                        '��ҩƷ���ݣ������ڲ��ּ�¼ִ�е������
                        If bytType = 1 Then
                            If .Cell(flexcpText, i, .ColIndex("ִ��״̬")) = 1 Then
                                MsgBox "�õ����Ѿ���ȫִ�У������ٴεǼ�ִ�С�", vbQuestion, gstrSysName
                                Exit Function
                            End If
                        Else
                            If .Cell(flexcpText, i, .ColIndex("ִ��״̬")) = 0 Then
                                MsgBox "�õ���δִ�У�����ȡ��ִ�С�", vbQuestion, gstrSysName
                                Exit Function
                            End If
                        End If
                        str��� = str��� & "," & .Cell(flexcpText, i, .ColIndex("���"))
                    End If
                Else
                    blnHaveDrug = True
                End If
            End If
        Next
        If blnDo = False Then
            If blnHaveDrug And blnHave���� Then
                strMsg = "ҩƷͨ����ҩ����ҩ������ִ�л�ȡ��ִ�У�" & vbNewLine & "�����ķ��Զ����������,����ͨ������������ִ�л�ȡ��ִ�У�" _
                        & vbNewLine & "������ֱ�ӵǼǻ�ȡ��ִ�С�"
            ElseIf blnHaveDrug Or blnHave���� Then
                strMsg = IIf(blnHaveDrug, "ҩƷͨ����ҩ����ҩ������ִ�л�ȡ��ִ�У�������ֱ�ӵǼǻ�ȡ��ִ�С�", _
                    " �����ķ��Զ����������,����ͨ������������ִ�л�ȡ��ִ�У�" & vbNewLine & "������ֱ�ӵǼǻ�ȡ��ִ�С�")
            End If
            'strMsg = ""Ϊֻ���и�����������,�������Զ�����
            If strMsg <> "" Then
                MsgBox strMsg, vbQuestion, gstrSysName
                Exit Function
            End If
        End If
        str��� = Mid(str���, 2)
    End With
    arrSQL = Array()
    
    If bytType = 1 Then
        If int��¼���� = 2 And gSysPara.blnִ�к���� Then
            curMoney = GetUnAuditBill(strNO, int��¼����, str���, str�����)
        End If
        
        If mTYAdviceProperty.int������Դ = 2 Then
            'סԺ���ʵ���
            
            '�����Զ�����ʱ�ķ��ü��
            If gSysPara.bln����ִ�з��� And Not gSysPara.blnִ�к���� Then
                If Not CheckStuffAudit(strNO, int��¼����) Then
                    MsgBox "�������ܼ�����" & vbCrLf & vbCrLf & "�����д�����δ��˵����ķ��ã�������ִ��֮���Զ����ϡ�", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            
            'ִ�к��Զ����ʱ���Գ�Ժ���˼��ǿ�Ƽ���Ȩ��
            If gSysPara.blnִ�к���� And curMoney <> 0 Then
                If Not PatiCanBilling(mTYAdviceProperty.lng����ID, mTYAdviceProperty.lng��ҳId, mstrPrivsAnnexFee, pҽ�����ѹ���) Then Exit Function
            End If
        Else
            '������ʵģ�����ִ�к��Զ���ˣ��������Ƚ��ѻ����
            If int��¼���� = 1 Or int��¼���� = 2 And Not gSysPara.blnִ�к���� Then
                If Not CheckFinishCharge(strNO, int��¼����, blnIsAbnormal) Then
                    If gSysPara.blnִ��ǰ�Ƚ��� And Not mobjSquareCard Is Nothing Then
                        If blnIsAbnormal Then
                            MsgBox "�ò��˻������쳣���ã����顣", vbInformation, gstrSysName
                            Exit Function
                        End If
                        '����һ��ͨ,��Ŀִ��ǰ�������շѻ��ȼ������,�������ݺţ�����ҽ��ID��ȡ����δ�շѵ��ݻ�δ��˵ļ��ʵ�
                        blnRefresh = mobjSquareCard.zlSquareAffirm(Me, pҽ�����ѹ���, strMainPrivs, mTYAdviceProperty.lng����ID, 0, False, int��¼����, strNO, "")
                        If Not blnRefresh Then
                            Exit Function
                        End If
                    Else
                        MsgBox "�ò��˻�����δ" & IIf(int��¼����, "�շ�", "��˼���") & "�ķ��ã����顣", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        End If
        
        '�Լ��ʷ��ý��б���
        If int��¼���� = 2 And curMoney <> 0 And gSysPara.blnִ�к���� Then
            If Not FinishBillingWarn(Me, mstrPrivsAnnexFee, mTYAdviceProperty.lng����ID, mTYAdviceProperty.lng��ҳId, mTYAdviceProperty.lng���˲���ID, curMoney, str���, str�����) Then Exit Function
        End If
        
        If MsgBox("��ȷ��Ҫ������""" & strNO & """�Ǽ�Ϊ��ִ����", vbQuestion + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then Exit Function
        
        strDate = "To_Date('" & Format(gobjDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        arr��� = Split(str���, ",")
        '�����ų���ҩƷ��
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_���˷��ü�¼_Execute('" & strNO & "'," & int��¼���� & IIf(UBound(arr���) + 1 = lngRow, ",Null,", ",'" & str��� & "',") & mTYAdviceProperty.int������Դ & ",Null,'" & UserInfo.���� & "'," & strDate & ")"
    Else
        If MsgBox("��ȷ��Ҫȡ������""" & strNO & """��ִ�еǼ���", vbQuestion + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then Exit Function
        arr��� = Split(str���, ",")
        '�����ų���ҩƷ��
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_���˷��ü�¼_UnExecute('" & strNO & "'," & int��¼���� & IIf(UBound(arr���) + 1 = lngRow, ",Null,", ",'" & str��� & "',") & mTYAdviceProperty.int������Դ & ")"
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call gobjDatabase.ExecuteProcedure(CStr(arrSQL(i)), Tittle)
        Next
    gcnOracle.CommitTrans: blnTrans = False
    lngRow = vsExpense.Row
    If blnRefresh Then
        Call LoadFeeDataFromAdvice(mTYAdviceProperty.lngҽ��ID, mTYAdviceProperty.lng���ͺ�, mTYAdviceProperty.bln����ִ��)
    End If
    vsExpense.Row = lngRow
    'ˢ�·�����ϸ�嵥
    Call LoadFeeDataFromAdvice(mTYAdviceProperty.lngҽ��ID, mTYAdviceProperty.lng���ͺ�, mTYAdviceProperty.bln����ִ��)
    RaiseEvent StatusTextUpdate(IIf(bytType = 0, 2, 1), IIf(bytType = 0, "ȡ��", "") & "ִ�в����ɹ���")
    zlFuncExtraFeeExe = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function GetUnAuditBill(ByVal strNO As String, ByVal int��¼���� As Integer, str��� As String, str����� As String) As Currency
'���ܣ���ȡδ��˵ļ��ʵ��ݵĽ���������ڼ��ʱ���
'������
'���أ�str���,str�����=���ڱ�����ʾ
'˵����
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, curMoney As Currency
    
    str��� = "": str����� = ""
    
    On Error GoTo errH
    strSql = _
        " Select B.����,B.����,Sum(A.ʵ�ս��) as ���" & _
        " From סԺ���ü�¼ A,�շ���Ŀ��� B" & _
        " Where A.NO = [1] And A.��¼���� = [2] And A.��¼״̬ = 0 And A.�շ����=B.����" & _
        " Group by B.����,B.����"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Tittle, strNO, int��¼����)
    
    curMoney = 0
    Do While Not rsTmp.EOF
        curMoney = curMoney + Nvl(rsTmp!���, 0)
        str��� = str��� & rsTmp!����
        str����� = str����� & "," & rsTmp!����
        rsTmp.MoveNext
    Loop
    
    str����� = Mid(str�����, 2)
    GetUnAuditBill = curMoney
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function CheckStuffAudit(ByVal strNO As String, ByVal int��¼���� As Integer) As Boolean
'���ܣ��жϵ����и������������Ƿ����δ��˵ļ��ʷ���
'������
'���أ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = " Select Nvl(Sum(A.ʵ�ս��),0) as ���" & _
        " From סԺ���ü�¼ A,�������� B" & _
        " Where NO = [1] And A.��¼���� = [2] And A.��¼״̬ = 0 And A.�շ����='4' And A.�շ�ϸĿID=B.����ID And B.��������=1"
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Tittle, strNO, int��¼����)
    If Not rsTmp.EOF Then
        CheckStuffAudit = rsTmp!��� = 0
    Else
        CheckStuffAudit = True
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function CheckFinishCharge(ByVal strNO As String, ByVal int��¼���� As Integer, ByRef blnIsAbnormal As Boolean) As Boolean
'���ܣ��ж�ָ���ĵ����Ƿ����շѣ��Լ��Ƿ�����쳣����
'������blnIsAbnormal=�Ƿ�����շ��쳣�ļ�¼

    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select ��¼״̬,ִ��״̬ From ������ü�¼ Where NO = [1] And ��¼���� = [2] And (��¼״̬ = 0 or ִ��״̬ = 9)"
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Tittle, strNO, int��¼����)
    rsTmp.Filter = "��¼״̬=0"
    If rsTmp.RecordCount > 0 Then
        CheckFinishCharge = False
    Else
        CheckFinishCharge = True
        rsTmp.Filter = "ִ��״̬=9"
        If rsTmp.RecordCount > 0 Then blnIsAbnormal = True
    End If
    
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function FuncExtraFeeMove(Optional frmMain As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ת��
    '���:frmMain-���õ�������
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 18:15:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String, int��¼���� As Integer
    Dim objParent As Object
    
    If mTYAdviceProperty.lngҽ��ID = 0 Then Exit Function
    If mblnMoved Then
        MsgBox "�ò��˵ı���" & IIf(mTYAdviceProperty.int������Դ = 2, "סԺ", "����") & "�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
            "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Function
    End If
    
    With vsExpense
        If Not Is���ӷ� Then
            MsgBox "��ǰѡ��ĵ��ݲ��ǲ���ĸ��ӷ��á�", vbInformation, gstrSysName
            Exit Function
        End If
        
        strNO = Get���ݺ�
        int��¼���� = Get��¼����
    End With
    If Not frmMain Is Nothing Then
        Set objParent = frmMain
    Else
        Set objParent = Me
    End If
    
    If frmExtraFeemove.ShowMe(objParent, mTYAdviceProperty.lng����ID, mTYAdviceProperty.lng��ҳId, mTYAdviceProperty.str�Һŵ�, mTYAdviceProperty.lng���ID, mTYAdviceProperty.str�������, mlngִ�п���ID, _
        strNO, int��¼����, IIf(mTYAdviceProperty.strFeeTab = "סԺ���ü�¼", 2, 1)) Then
        Call Refresh
    End If
End Function
Public Function zlFuncPlugIn(ByVal frmMain As Object, ByVal Control As CommandBarControl) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�����
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-05-29 17:50:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String
    If CreatePlugIn(pҽ�����ѹ���) = False Then Exit Function
    strNO = Get���ݺ�
    On Error Resume Next
    Call gobjPlugIn.ExecuteFunc(glngSys, pҽ�����ѹ���, Control.Parameter, _
        mTYAdviceProperty.lng����ID, IIf(mTYAdviceProperty.str�Һŵ� = "", mTYAdviceProperty.lng��ҳId, mTYAdviceProperty.str�Һŵ�), mTYAdviceProperty.lngҽ��ID, strNO)
    Call zlPlugInErrH(Err, "ExecuteFunc")
    Err.Clear: On Error GoTo 0
End Function
Public Function zlFuncAdviceReCharge(ByVal intType As Integer, Optional frmMain As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������������
    '���:intType=1-����;2-���
    '     frmMain-���õ�������
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 18:24:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnOK As Boolean
    Dim strCommon As String, intAtom As Integer
    Dim objParent As Object, strNO As String, lngAdviceID As Long
    Set objParent = frmMain
    If objParent Is Nothing Then Set objParent = mfrmParent
    
    '���÷��ò�������
    On Error Resume Next
    If gobjInExse Is Nothing Then
        Set gobjInExse = CreateObject("zl9InExse.clsInExse")
        If gobjInExse Is Nothing Then Exit Function
    End If
    Err.Clear: On Error GoTo 0
    
    '�������úϷ�������
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & AnalyseComputer
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", intAtom)
    
    If intType = 1 Then
        lngAdviceID = mTYAdviceProperty.lngҽ��ID
        strNO = Get���ݺ�
        
        Select Case mbytFocus
        Case 1
            '��ҽ��
            blnOK = gobjInExse.CallReCharge(objParent, gcnOracle, gstrDBUser, glngSys, 0, 1, mlngִ�п���ID, mstrPrivsAnnexFee, mTYAdviceProperty.lng����ID, , lngAdviceID)
        Case 2
            '������
            blnOK = gobjInExse.CallReCharge(objParent, gcnOracle, gstrDBUser, glngSys, 0, 1, mlngִ�п���ID, mstrPrivsAnnexFee, mTYAdviceProperty.lng����ID, strNO)
        Case Else
            '������
            blnOK = gobjInExse.CallReCharge(objParent, gcnOracle, gstrDBUser, glngSys, 0, 1, mlngִ�п���ID, mstrPrivsAnnexFee, mTYAdviceProperty.lng����ID)
        End Select
    ElseIf intType = 2 Then
        blnOK = gobjInExse.CallReCharge(objParent, gcnOracle, gstrDBUser, glngSys, 1, 1, mlngִ�п���ID, mstrPrivsAnnexFee, mTYAdviceProperty.lng����ID)
    End If
    Call GlobalDeleteAtom(intAtom)
    zlFuncAdviceReCharge = blnOK
    
    If blnOK And frmMain Is Nothing Then RaiseEvent RequestRefresh
End Function
 

Private Sub vsExpense_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strNO As String, int��¼���� As Integer, lng���ͺ� As Long, lngҽ��ID As Long
    If Button = 1 Then Exit Sub
    
    With vsExpense
        strNO = Get���ݺ�
        If strNO = "[δ�Ʒ�]" Then strNO = ""
        int��¼���� = Val(Get��¼����)
    End With
    RaiseEvent zlPopupMenu(mTYAdviceProperty.lngҽ��ID, mTYAdviceProperty.lng���ͺ�, strNO, int��¼����, X, Y)
End Sub

Public Function IsFunValied(ByVal bytType As Byte, ByVal bytPrivsCheck As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�м��ĳ�����Ƿ���Ч
    '���: bytType- 1-�޸ĸ���;2-ɾ������;3-����ת��;4-����ִ��;5-����ȡ��ִ��;6-��������;7-�������
    '      bytPrivsCheck -���Ȩ��:0-�����Ȩ��;1-������ݺ�Ȩ��;2-�����Ȩ��
    '����:
    '����:������Ч,����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 17:00:52
    '˵��:
    '   1.���ݸ����б��е�����,���ĳ����Ƿ���Ч
    '   2.����Ȩ�޼��ĳ����Ƿ���Ч
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnValied As Boolean, bln���ӷ��� As Boolean, bln������ As Boolean
    
    On Error GoTo errHandle
    
    bln������ = Is������: bln���ӷ��� = Is���ӷ�
    
    Select Case bytType
    Case 1 '�޸ĸ���
        blnValied = mTYAdviceProperty.lngҽ��ID <> 0 And (mTYAdviceProperty.intִ��״̬ = 0 Or mTYAdviceProperty.intִ��״̬ = 3) _
                   And (bln���ӷ��� Or mTYAdviceProperty.lng�Ƽ����� = 1)
        
        If bytPrivsCheck <> 0 Then
            blnValied = IIf(bytPrivsCheck = 2, True, blnValied) And InStr(mstrPrivsAnnexFee, ";�޸ķ���;") > 0
        End If
    
    Case 2 'ɾ������
        blnValied = mTYAdviceProperty.lngҽ��ID <> 0 And (mTYAdviceProperty.intִ��״̬ = 0 Or mTYAdviceProperty.intִ��״̬ = 3) _
            And (bln���ӷ��� Or bln������)
        If bytPrivsCheck <> 0 Then
            blnValied = IIf(bytPrivsCheck = 2, True, blnValied) And InStr(mstrPrivsAnnexFee, ";ɾ������;") > 0
        End If
    Case 3 '����ת��
         blnValied = mTYAdviceProperty.lngҽ��ID <> 0 And bln���ӷ���
        If bytPrivsCheck <> 0 Then
            blnValied = IIf(bytPrivsCheck = 2, True, blnValied) And InStr(mstrPrivsAnnexFee, ";���丽�ӷ���;") > 0
        End If
    Case 4 '����ִ��
         blnValied = mTYAdviceProperty.lngҽ��ID <> 0 And (mTYAdviceProperty.intִ��״̬ = 0 Or mTYAdviceProperty.intִ��״̬ = 3) And bln���ӷ���
        If bytPrivsCheck <> 0 Then
            blnValied = IIf(bytPrivsCheck = 2, True, blnValied)
        End If
    Case 5 '����ȡ��ִ��
        blnValied = mTYAdviceProperty.lngҽ��ID <> 0 And (mTYAdviceProperty.intִ��״̬ = 0 Or mTYAdviceProperty.intִ��״̬ = 3) And bln���ӷ���
        If bytPrivsCheck <> 0 Then
            blnValied = IIf(bytPrivsCheck = 2, True, blnValied)
        End If
    Case 6 '��������
        blnValied = mTYAdviceProperty.int������Դ = 2 And mlngִ�п���ID <> 0
        If bytPrivsCheck <> 0 Then
            blnValied = IIf(bytPrivsCheck = 2, True, blnValied) And Not (InStr(mstrPrivsAnnexFee, ";ҩƷ��������;") = 0 _
            And InStr(mstrPrivsAnnexFee, ";������������;") = 0 _
            And InStr(mstrPrivsAnnexFee, ";������������;") = 0)
        End If
    Case 7 '�������
        blnValied = mTYAdviceProperty.int������Դ = 2 And mlngִ�п���ID <> 0
        If bytPrivsCheck <> 0 Then
            blnValied = IIf(bytPrivsCheck = 2, True, blnValied) And InStr(mstrPrivsAnnexFee, ";�������;") > 0
        End If
    Case Else
        blnValied = True
    End Select
    IsFunValied = blnValied
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlFuncStuffCharge(ByVal frmMain As Object, ByVal int��¼���� As Integer, _
    Optional strOutNos As String, Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ļ���
    '���:int��¼����:1-�շ�(����),2-����(��/ס)
    '����:strOutNos-����ɹ��ı��������շѻ���ʵ��ݵ���
    '����:���˺�
    '����:2010-12-14 13:17:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mTYAdviceProperty.lngҽ��ID = 0 Then Exit Function
    If mTYAdviceProperty.intִ��״̬ = 1 Then
        MsgBox "��ִ����Ŀ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
        Exit Function
    End If
    If mblnMoved Then
        MsgBox "�ò��˵ı���" & IIf(mTYAdviceProperty.int������Դ = 2, "סԺ", "����") & "�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
            "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Function
    End If
    If frmStuffCharge.zlBillEdit(frmMain, 0, pҽ�����ѹ���, mstrPrivsAnnexFee, int��¼����, "", _
        IIf(mTYAdviceProperty.int������Դ = 2, 2, 1), mTYAdviceProperty.lng����ID, mTYAdviceProperty.lng��ҳId, mlngִ�п���ID, mTYAdviceProperty.lng���˿���ID, _
        0, "", False, "", mTYAdviceProperty.lngҽ��ID, mTYAdviceProperty.lng���ͺ�, "", , , strOutNos, , objSaveData, mobjSquareCard) = True Then
        zlFuncStuffCharge = True
        Call Refresh
        Call RefreshExpenseData
    End If
End Function

Public Function zlFuncExtraFeeMove(Optional frmMain As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ת��
    '���:frmMain-���õ�������
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 18:15:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String, int��¼���� As Integer
    Dim objParent As Object
    
    If mTYAdviceProperty.lngҽ��ID = 0 Then Exit Function
    If mblnMoved Then
        MsgBox "�ò��˵ı���" & IIf(mTYAdviceProperty.int������Դ = 2, "סԺ", "����") & "�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
            "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Function
    End If
    If Not Is���ӷ� Then
        MsgBox "��ǰѡ��ĵ��ݲ��ǲ���ĸ��ӷ��á�", vbInformation, gstrSysName
        Exit Function
    End If
    strNO = Get���ݺ�
    int��¼���� = Get��¼����
    If Not frmMain Is Nothing Then
        Set objParent = frmMain
    Else
        Set objParent = mfrmParent
    End If
    If frmExtraFeemove.ShowMe(objParent, mTYAdviceProperty.lng����ID, mTYAdviceProperty.lng��ҳId, mTYAdviceProperty.str�Һŵ�, mTYAdviceProperty.lng���ID, mTYAdviceProperty.str�������, mlngִ�п���ID, _
        strNO, int��¼����, IIf(mTYAdviceProperty.strFeeTab = "סԺ���ü�¼", 2, 1)) Then
        zlFuncExtraFeeMove = True
        Call Refresh
    End If
End Function

Public Sub zlUpdateCommandBars(ByVal cbsMain As Object, ByVal Control As CommandBarControl)
    Dim blnEnabled As Boolean
    If cbsMain Is Nothing Then Exit Sub
    
    If vsExpense.Redraw = flexRDNone Then Exit Sub
    '����Ȩ�����ð�ť�ɼ�״̬
    Call SetControlVisible(Control)
    If Not Control.Visible Then Exit Sub
    
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel '��������嵥
        Control.Enabled = IsHaveExpenseData
    Case conMenu_Edit_Append
        
        If Isδ�Ʒ� Then
            Control.Enabled = mTYAdviceProperty.lngҽ��ID <> 0 And (mTYAdviceProperty.intִ��״̬ = 0 Or mTYAdviceProperty.intִ��״̬ = 3)
            If Control.Parent Is cbsMain(2) Then
                Control.Caption = "��������"
            Else
                Control.Caption = "����������(&N)"
            End If
        Else
            blnEnabled = Not mrsPrice Is Nothing
            If blnEnabled Then blnEnabled = mrsPrice.RecordCount <> 0
            Control.Enabled = mTYAdviceProperty.lngҽ��ID <> 0 And (mTYAdviceProperty.intִ��״̬ = 0 Or mTYAdviceProperty.intִ��״̬ = 3) And InStr(mTYAdviceProperty.str�Ʒ�״̬, ",0,") > 0 And blnEnabled
            If Control.Parent Is cbsMain(2) Then
                Control.Caption = "��������"
            Else
                Control.Caption = "����������(&N)"
            End If
        End If
    Case conMenu_Edit_ExtraFeeMove  '����ת��
        Control.Enabled = mTYAdviceProperty.lngҽ��ID <> 0 And Is���ӷ�
    Case conMenu_Edit_ExtraFeeExe   '����ִ��
        Control.Enabled = mTYAdviceProperty.lngҽ��ID <> 0 And (mTYAdviceProperty.intִ��״̬ = 0 Or mTYAdviceProperty.intִ��״̬ = 3) _
            And Is���ӷ�
    Case conMenu_Edit_ExtraFeeUnExe '����ȡ��ִ��
        Control.Enabled = mTYAdviceProperty.lngҽ��ID <> 0 And (mTYAdviceProperty.intִ��״̬ = 0 Or mTYAdviceProperty.intִ��״̬ = 3) _
            And Is���ӷ�
    Case conMenu_Edit_NewItem
        Control.Enabled = mTYAdviceProperty.lngҽ��ID <> 0 And (mTYAdviceProperty.intִ��״̬ = 0 Or mTYAdviceProperty.intִ��״̬ = 3)
    '78929:���ϴ�,2014/10/27�����С��ʱ���ķѲ�����
    Case conMenu_Edit_Modify
        Control.Visible = InStr(mstrPrivsAnnexFee, ";�޸ķ���;") > 0
        Control.Enabled = Control.Visible And mTYAdviceProperty.lngҽ��ID <> 0 And (mTYAdviceProperty.intִ��״̬ = 0 Or mTYAdviceProperty.intִ��״̬ = 3) And _
        (Is���ӷ� Or mTYAdviceProperty.lng�Ƽ����� = 1) And Not IIf(vsExpense.Row < 0, True, vsExpense.IsSubtotal(vsExpense.Row))
    Case conMenu_Edit_Delete
        Control.Visible = InStr(mstrPrivsAnnexFee, ";ɾ������;") > 0
        Control.Enabled = Control.Visible And mTYAdviceProperty.lngҽ��ID <> 0 And (mTYAdviceProperty.intִ��״̬ = 0 Or mTYAdviceProperty.intִ��״̬ = 3) _
            And (Is���ӷ� Or Is������)
    Case conMenu_Edit_ChargeDelApply, conMenu_Edit_ChargeDelAudit '�����������
        Control.Enabled = mTYAdviceProperty.int������Դ = 2 And mlngִ�п���ID <> 0
    End Select
End Sub
Public Property Get Is���ӷ�() As Boolean
    If mbytFun = 1 Then Exit Function
    With vsExpense
        If .Rows <= 1 Then Exit Property
        If .Row < 0 Then Exit Property
        If .ColIndex("��������") Then Exit Property
        Is���ӷ� = .TextMatrix(.Row, .ColIndex("��������")) = "���ӷ���"
    End With
End Property
Public Property Get Is������() As Boolean
    If mbytFun = 1 Then Exit Property
    With vsExpense
        If .Rows <= 1 Then Exit Property
        If .Row < 0 Then Exit Property
        If .ColIndex("��������") Then Exit Property
        Is������ = .TextMatrix(.Row, .ColIndex("��������")) = "������"
    End With
End Property
Private Function Get���ݺ�() As String
    Dim lngRow As Long
    With vsExpense
        If .Rows <= 1 Then Exit Function
        If .Row <= 0 Then Exit Function
        If .ColIndex("���ݺ�") < 0 Then Exit Function
        If .IsSubtotal(.Row) Then
            lngRow = .Row + 1
            If lngRow > .Rows - 1 Then lngRow = .Row
            Get���ݺ� = .TextMatrix(lngRow, .ColIndex("���ݺ�"))
        Else
            Get���ݺ� = .TextMatrix(.Row, .ColIndex("���ݺ�"))
        End If
    End With
End Function
Private Function Get��¼����() As Integer
    Dim lngRow As Long, int��¼���� As Integer
    With vsExpense
        int��¼���� = mTYAdviceProperty.int��¼����
        
        If .Rows <= 1 Then Get��¼���� = int��¼����: Exit Function
        If .Row < 0 Then Get��¼���� = int��¼����: Exit Function
        If .ColIndex("��¼����") < 0 Then Get��¼���� = int��¼����: Exit Function
        
        If .IsSubtotal(.Row) Then
            lngRow = .Row + 1
            If lngRow > .Rows - 1 Then lngRow = .Row
            Get��¼���� = Val(.TextMatrix(lngRow, .ColIndex("��¼����")))
        Else
            Get��¼���� = Val(.TextMatrix(.Row, .ColIndex("��¼����")))
        End If
    End With
End Function

Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'���ܣ�����Ȩ�����ò˵��͹������Ŀɼ�״̬
    Dim blnVisible As Boolean
    
    'Ȩ��ֻ���ж�һ��,�Ѿ��жϹ�����������ж�
    If Control.Category = "���ж�" Then Exit Sub

    blnVisible = True
    Select Case Control.ID
    Case conMenu_Edit_Append, conMenu_Edit_NewItem, conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_ExtraFeeMove
        If InStr(mstrPrivsAnnexFee, ";���丽�ӷ���;") = 0 Then blnVisible = False
    Case conMenu_Edit_NewItem * 10# + 1  '�����շѷ���
        If mTYAdviceProperty.int������Դ = 2 Or InStr(mstrPrivsAnnexFee, ";�����շѷ���;") = 0 Or mTYAdviceProperty.int����ģʽ = 1 Then blnVisible = False
    Case conMenu_Edit_NewItem * 10# + 2     '������ʷ���
        If mTYAdviceProperty.int������Դ = 2 Then
            If InStr(mstrPrivsAnnexFee, ";����סԺ���ʷ���;") = 0 Then blnVisible = False
        Else
            If InStr(mstrPrivsAnnexFee, ";����������ʷ���;") = 0 Then blnVisible = False
        End If
    Case conMenu_Edit_NewItem * 10# + 3 '������ķ���
        If InStr(mstrPrivsAnnexFee, ";������ķ���;") = 0 Then blnVisible = False
    Case conMenu_Edit_NewItem * 10# + 5  '�������ļ��ʺ��շ�
        If mTYAdviceProperty.int������Դ = 2 Then
            If InStr(mstrPrivsAnnexFee, ";���䱸�����ķ���;") = 0 Or InStr(mstrPrivsAnnexFee, ";����סԺ���ʷ���;") = 0 Then
                blnVisible = False
            End If
        Else
            If InStr(mstrPrivsAnnexFee, ";���䱸�����ķ���;") = 0 Or InStr(mstrPrivsAnnexFee, ";����������ʷ���;") = 0 Then
                blnVisible = False
            End If
        End If
    Case conMenu_Edit_NewItem * 10# + 4     '
        If mTYAdviceProperty.int������Դ = 2 Then
            blnVisible = False
        Else
            If InStr(mstrPrivsAnnexFee, ";���䱸�����ķ���;") = 0 Or _
               InStr(mstrPrivsAnnexFee, ";�����շѷ���;") = 0 Then blnVisible = False
        End If
    Case conMenu_Edit_ChargeDelApply '��������
        '���˺� ����: 34873   ����:2010-12-22 13:50:07
        '55380
        If InStr(mstrPrivsAnnexFee, ";ҩƷ��������;") = 0 _
            And InStr(mstrPrivsAnnexFee, ";������������;") = 0 _
            And InStr(mstrPrivsAnnexFee, ";������������;") = 0 Then blnVisible = False
    Case conMenu_Edit_ChargeDelAudit '�������
        If InStr(mstrPrivsAnnexFee, ";�������;") = 0 Then blnVisible = False
    End Select
    
    Control.Visible = blnVisible
    Control.Category = "���ж�"
End Sub
Private Sub RefreshExpenseData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ˢ�·�������
    '����:���˺�
    '����:2014-05-30 14:54:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngҽ��ID As Long, lng���ͺ� As Long, bln����ִ�� As Boolean
    If mbytFun = 1 Then
        Call LoadFeeListFromNos(mbyt��¼����, mstrNos, mbyt������Դ, mblnMoved)
    Else
        With vsAdvice
            If .Row > 0 And .Row <= .Rows - 1 Then
                lngҽ��ID = Val(.TextMatrix(.Row, .ColIndex("ҽ��ID")))
                lng���ͺ� = Val(.TextMatrix(.Row, .ColIndex("���ͺ�")))
                bln����ִ�� = Val(.TextMatrix(.Row, .ColIndex("����ִ��"))) = 1
            End If
        End With
        Call LoadFeeDataFromAdvice(lngҽ��ID, lng���ͺ�, bln����ִ��)
    End If
End Sub

Private Function GetAdviceMoney(ByVal strAdviceIdAndPayNums As String, _
    ByRef rsTemp As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ҽ����Ӧ�ս���ʵ�ս��
    '���:strAdviceIdAndPayNums-ҽ��ID�ͷ��ͺ�(ҽ��ID:���ͺ�,...)
    '����:rsTemp-����:ҽ��ID,���ͺ�,Ӧ�ս��,ʵ�ս��
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-05-30 15:13:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, blnMoved As Boolean
    Dim strSQL1 As String, strSub As String
    Dim strWith As String, strWith1 As String, strTemp As String
    Dim rsSub As ADODB.Recordset, strҽ��IDs As String
    Dim strSubWith As String, strSubWith1 As String
    
    On Error GoTo errHandle
    '124405,���ϴ���2018/4/17����ȡ��ʷ����ʱ������������ͬ��with�Ӿ�
    blnMoved = gobjDatabase.TableDataMoved("����ҽ������", " (ҽ��ID,���ͺ�) IN", " (Select C1 As ҽ��id, C2 As ���ͺ� From Table(f_Num2list2('" & strAdviceIdAndPayNums & "')))")
    
    '85249:���ϴ�,2015/5/29,withԤ��Ƕ��ʹ��
    strWith = "" & _
    "   Select A.ҽ��ID,a.���ͺ�,A.��¼����,A.NO" & vbNewLine & _
    "   From ����ҽ������ A" & vbNewLine & _
    "   Where (A.ҽ��id,A.���ͺ�) IN (Select /*+cardinality(B,10)*/B.C1, B.C2 From Table(f_Num2list2([1])) B)" & vbNewLine & _
    "   Union ALL " & vbNewLine & _
    "   Select A.ҽ��ID,A.���ͺ�,A.��¼����,A.NO " & vbNewLine & _
    "   From ����ҽ������ A" & vbNewLine & _
    "   Where (A.ҽ��id,A.���ͺ�) IN (Select /*+cardinality(B,10)*/B.C1, B.C2 From Table(f_Num2list2([1])) B) "
    
    If blnMoved Then
        strWith1 = strWith
        strWith1 = Replace(strWith1, "����ҽ������", "H����ҽ������")
        strWith1 = Replace(strWith1, "����ҽ������", "H����ҽ������")
        strWith = strWith & " Union ALL " & strWith1
    End If
    strWith = "   With ҽ������ as ( " & strWith & " )"
    
    strTemp = "Select A.Id As ҽ��ID, a.���id As ��ID, a.������� From ����ҽ����¼ A Where A.���ID IN (Select /*+cardinality(B,10)*/B.C1 From Table(f_Num2list2([1])) B)"
    Set rsSub = gobjDatabase.OpenSQLRecord(strTemp, Tittle, strAdviceIdAndPayNums)
    If Not rsSub.EOF Then
        strSubWith = "" & _
        "   Select a.ҽ��id, a.���ͺ�, a.��¼����, a.No, m.���id As ��id" & vbNewLine & _
        "   From ����ҽ������ a, ����ҽ����¼ m" & vbNewLine & _
        "   Where a.ҽ��id = m.Id And (a.ҽ��id, a.���ͺ�, a.No) In" & vbNewLine & _
        "      (Select a.Id As ҽ��id, m.���ͺ�, m.No" & vbNewLine & _
        "                         From ����ҽ����¼ a, ����ҽ������ m" & vbNewLine & _
        "                         Where a.���id = m.ҽ��id And" & vbNewLine & _
        "                               (m.ҽ��id, m.���ͺ�) In (Select /*+cardinality(b,10) */" & vbNewLine & _
        "                                                    b.C1, b.C2" & vbNewLine & _
        "                                                   From Table(f_Num2list2([1])) b))" & vbNewLine & _
        "   Union ALL " & vbNewLine & _
        "   Select a.ҽ��id, a.���ͺ�, a.��¼����, a.No, m.���id As ��id" & vbNewLine & _
        "   From ����ҽ������ a, ����ҽ����¼ m" & vbNewLine & _
        "   Where a.ҽ��id = m.Id And" & vbNewLine & _
        "      (a.ҽ��id, a.���ͺ�) In (Select a.Id As ҽ��id, m.���ͺ�" & vbNewLine & _
        "                          From ����ҽ����¼ a, ����ҽ������ m" & vbNewLine & _
        "                          Where a.���id = m.ҽ��id And" & vbNewLine & _
        "                                (m.ҽ��id, m.���ͺ�) In (Select /*+cardinality(b,10) */" & vbNewLine & _
        "                                                     b.C1, b.C2" & vbNewLine & _
        "                                                    From Table(f_Num2list2([1])) b))"

        If blnMoved Then
            strSubWith1 = strSubWith
            strSubWith1 = Replace(strSubWith1, "����ҽ������", "H����ҽ������")
            strSubWith1 = Replace(strSubWith1, "����ҽ������", "H����ҽ������")
            strSubWith1 = Replace(strSubWith1, "����ҽ����¼", "H����ҽ����¼")
            strSubWith = strSubWith & " Union ALL " & strSubWith1
        End If
        strSubWith = "   ,ҽ���������� as ( " & strSubWith & " )"
        
        strSql = "" & _
        "   Select B.ҽ��ID,B.���ͺ�,A.Ӧ�ս��,A.ʵ�ս��" & vbNewLine & _
        "   From ������ü�¼ A,ҽ������ B " & vbNewLine & _
        "   Where mod(A.��¼����,10)=B.��¼����  And A.NO=B.NO And A.ҽ�����=B.ҽ��ID " & vbNewLine & _
        "   Union All" & _
        "   Select B.��ID As ҽ��ID,B.���ͺ�,A.Ӧ�ս��,A.ʵ�ս��" & vbNewLine & _
        "   From ������ü�¼ A,ҽ���������� B,����ҽ����¼ C " & vbNewLine & _
        "   Where mod(A.��¼����,10)=B.��¼���� And a.ҽ�����=c.id And c.������� In ('C','D','F') And A.NO=B.NO And A.ҽ�����=B.ҽ��ID "
    Else
        '87435,һ�ŵ��ݶ��ҽ��ʱ��ͨ��ҽ��IDͳ�Ʒ��ò���ȷ 'And A.ҽ�����=B.ҽ��ID
        strSql = "" & _
        "   Select B.ҽ��ID,B.���ͺ�,sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��" & vbNewLine & _
        "   From ������ü�¼ A,ҽ������ B " & vbNewLine & _
        "   Where mod(A.��¼����,10)=B.��¼���� And A.NO=B.NO And A.ҽ�����=B.ҽ��ID " & vbNewLine & _
        "   Group By B.ҽ��ID,B.���ͺ�"
    End If
    
    strSql = strSql & " UNION ALL " & vbNewLine & _
    Replace(Replace(strSql, "������ü�¼", "סԺ���ü�¼"), "mod(A.��¼����,10)", "A.��¼����")
    
    If blnMoved Then
        strSQL1 = strSql
        strSQL1 = Replace(strSQL1, "����ҽ������", "H����ҽ������")
        strSQL1 = Replace(strSQL1, "����ҽ������", "H����ҽ������")
        strSQL1 = Replace(strSQL1, "������ü�¼", "H������ü�¼")
        strSQL1 = Replace(strSQL1, "סԺ���ü�¼", "HסԺ���ü�¼")
        strSql = strSql & " Union ALL " & strSQL1
    End If
    
    strSql = strWith & strSubWith & vbCrLf & strSql
    
    strSql = "" & _
    "   Select ҽ��ID,���ͺ�,sum(Ӧ�ս��) as Ӧ�ս��,Sum(ʵ�ս��) as ʵ�ս��" & vbNewLine & _
    "   From (" & strSql & ")  " & vbNewLine & _
    "   Group By ҽ��ID,���ͺ�"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Tittle, strAdviceIdAndPayNums, strҽ��IDs)
    GetAdviceMoney = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "ָ���������ÿһ�г��ֵ������С(�԰�Ϊ��λ)��"
    FontSize = UserControl.FontSize
    Call ReSetFontSize
End Property
Public Property Let FontSize(ByVal New_FontSize As Single)
    UserControl.FontSize() = New_FontSize
    Call ReSetFontSize
    PropertyChanged "FontSize"
End Property
Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������С
    '����:���˺�
    '����:2012-06-18 16:52:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngFontSize As Single
    sngFontSize = UserControl.FontSize
    
    Err = 0: On Error Resume Next
    picAdvice.FontSize = sngFontSize
    picExpense.FontSize = sngFontSize
    dkpMan.PaintManager.CaptionFont.Size = sngFontSize
    dkpMan.PanelPaintManager.Font.Size = sngFontSize
    Call gobjControl.VSFSetFontSize(vsAdvice, sngFontSize)
    Call gobjControl.VSFSetFontSize(vsExpense, sngFontSize)
    With vsExpense
        .Cell(flexcpFontSize, 0, 0, .Rows - 1, .Cols - 1) = sngFontSize
    End With
    dkpMan.RecalcLayout
End Sub

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=10,0,0,&HFFCC99
Public Property Get COLOR_FOCUS() As OLE_COLOR
    COLOR_FOCUS = m_COLOR_FOCUS
End Property

Public Property Let COLOR_FOCUS(ByVal New_COLOR_FOCUS As OLE_COLOR)
    m_COLOR_FOCUS = New_COLOR_FOCUS
    PropertyChanged "COLOR_FOCUS"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=10,0,0,&HFFEBD7
Public Property Get COLOR_LOST() As OLE_COLOR
    COLOR_LOST = m_COLOR_LOST
End Property

Public Property Let COLOR_LOST(ByVal New_COLOR_LOST As OLE_COLOR)
    m_COLOR_LOST = New_COLOR_LOST
    PropertyChanged "COLOR_LOST"
End Property
Public Sub SetDefalutFocus(ByVal blnAdviceFocus As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡ�Ĺ���б�
    '���:blnAdviceFocus-�Ƿ�ָ��ҽ���б�(true-ȱʡָ��ҽ��;false-Ϊ��ϸ�б�)
    '����:���˺�
    '����:2014-06-03 11:29:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If blnAdviceFocus Then
        If vsAdvice.Enabled And vsAdvice.Visible Then vsAdvice.SetFocus
        Call vsAdvice_GotFocus
        Call vsExpense_LostFocus
    Else
        If vsExpense.Enabled And vsExpense.Visible Then vsExpense.SetFocus
        Call vsExpense_GotFocus
        Call vsAdvice_LostFocus
    End If
End Sub
Public Function zlInitCommon(ByRef objSquareCard As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�������ӿ�
    '����:��ʼ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-07-08 10:25:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If CreatePubAdvice = False Then Exit Function
    If objSquareCard Is Nothing Then
        Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If mobjSquareCard.zlInitComponents(Me, pҽ�����ѹ���, glngSys, gstrDBUser, gcnOracle, False) = False Then
            Set mobjSquareCard = Nothing
            MsgBox "ҽ�ƿ�������zl9CardSquare����ʼ��ʧ�ܣ�", vbInformation, gstrSysName
        End If
    Else
        Set mobjSquareCard = objSquareCard
    End If
    zlInitCommon = True
End Function

