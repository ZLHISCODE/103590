VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmLisPrinter 
   BorderStyle     =   0  'None
   ClientHeight    =   7230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrReturn 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   210
      Top             =   6660
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      ScaleHeight     =   1275
      ScaleWidth      =   4515
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4680
      Width           =   4575
      Begin VB.HScrollBar hsb 
         Height          =   330
         Left            =   240
         MousePointer    =   99  'Custom
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   900
         Width           =   3135
      End
      Begin VB.VScrollBar vsb 
         Height          =   975
         Left            =   4200
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   120
         Width           =   330
      End
      Begin zl9NewQuery.ctlQueryItem QueryItem 
         Height          =   735
         Left            =   1080
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1296
      End
      Begin VB.PictureBox picBack1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3360
         ScaleHeight     =   735
         ScaleWidth      =   975
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fraMain 
      Height          =   1735
      Left            =   30
      TabIndex        =   3
      Top             =   -30
      Width           =   10005
      Begin VB.Frame fratak 
         Height          =   105
         Left            =   0
         TabIndex        =   8
         Top             =   960
         Width           =   9675
      End
      Begin VB.TextBox TxtID 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   2730
         TabIndex        =   0
         Top             =   150
         Width           =   6855
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   6780
         TabIndex        =   9
         Top             =   1170
         Width           =   165
      End
      Begin VB.Label lblID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ɨ�����룺"
         BeginProperty Font 
            Name            =   "����"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   2700
      End
      Begin VB.Image imgTitle 
         Height          =   750
         Left            =   30
         Picture         =   "frmLisPrinter.frx":0000
         Stretch         =   -1  'True
         Top             =   150
         Width           =   2655
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   26.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   525
         Left            =   7320
         TabIndex        =   6
         Top             =   1170
         Width           =   1365
      End
      Begin VB.Label lbl�Ա� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   26.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   525
         Left            =   3720
         TabIndex        =   5
         Top             =   1170
         Width           =   1620
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   26.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   525
         Left            =   120
         TabIndex        =   4
         Top             =   1170
         Width           =   1620
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid msfMain 
      Height          =   2835
      Left            =   30
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1770
      Width           =   9690
      _cx             =   17092
      _cy             =   5001
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   15199202
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16633516
      ForeColorSel    =   16711680
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   16761024
      GridColorFixed  =   16761024
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   25
      Cols            =   30
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   450
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
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
      ExplorerBar     =   0
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
   Begin zl9NewQuery.ctlButton ctlClear 
      Height          =   540
      Left            =   8370
      TabIndex        =   2
      Top             =   5790
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   953
      Caption         =   "���"
      BackColor       =   16777215
      FontSize        =   21.75
      FontBold        =   -1  'True
      AutoSize        =   0   'False
      ButtonHeight    =   420
   End
   Begin zl9NewQuery.ctlButton ctlReturn 
      Height          =   540
      Left            =   5280
      TabIndex        =   10
      Top             =   5880
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   953
      Caption         =   "����"
      BackColor       =   16777215
      FontSize        =   21.75
      FontBold        =   -1  'True
      AutoSize        =   0   'False
      ButtonHeight    =   420
   End
   Begin VB.Shape shp 
      BorderColor     =   &H00FF0000&
      Height          =   1575
      Left            =   120
      Top             =   4560
      Width           =   4935
   End
End
Attribute VB_Name = "frmLisPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mInputType As Integer               '��������:0=����;1=סԺ;2=����;3=���￨;4=IC��;5=����ID
Private mstrSource As String                '������Դ
Private mstrNO As String                    '���ݱ���
Private Enum mCol
    ������Ŀ = 0
    ������
    ����ʱ��
    �����
    ���ʱ��
    ״̬
    ˵��
    ��ӡ����
    �걾ID
    ҽ��id
    ����id
End Enum

Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mblnFist As Boolean
Private mvarPageNo As Long
Private mvarSvrDept As String           '��������ҽ���Ŀ���
Private mvarSvrDuty As String           '��������ҽ����ְ��
Private mlngHelpPage As Long
Private mintPrintDelayed As Integer     '��ӡ��ʱʱ��
Private mintClear As Integer            '�������ʱ��
Private mintBack As Integer             'ʱ���ӡ����󷵻���ҳ

Private mvarLeftStart As Single
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private mintReturn As Integer

Private Sub ctlClear_CommandClick()
    Dim intRow As Integer, intCol As Integer
     '��ռ�¼
    Me.lbl���� = "����:"
    Me.lbl���� = "����:"
    Me.lbl�Ա� = "�Ա�:"
    Me.lbl��ʾ = ""
    With Me.msfMain
        For intRow = 1 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .TextMatrix(intRow, intCol) = ""
            Next
        Next
    End With
    Me.TxtID.Text = ""
    tmrReturn.Enabled = False
    Me.TxtID.SetFocus
End Sub

'Private Sub ctlRead_CommandClick()
'    If mobjICCard Is Nothing Then
'        Set mobjICCard = CreateObject("zlICCard.clsICCard")
'        Set mobjICCard.gcnOracle = gcnOracle
'    End If
'    If Not mobjICCard Is Nothing Then
'        TxtID.Text = mobjICCard.Read_Card()
'        If TxtID.Text <> "" Then
'            Call TxtID_KeyPress(vbKeyReturn)
'            mblnICCard = True
'        End If
'    End If
'End Sub

Private Sub ctlReturn_CommandClick()
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim lngPageNum As Long
    If mblnFist = False Then Exit Sub
    mblnFist = False
  
    DoEvents
    mlngHelpPage = Val(GetPara("������ӡ����ҳ��"))
    ctlReturn.Visible = Val(GetPara("������ӡ��ʾ���ذ�ť"))
    mintPrintDelayed = Val(GetPara("�����ӡ��ʱ", 0))
    
    If mlngHelpPage > 0 Then
        Call Form_Resize
        Call LoadPageItemList(mlngHelpPage)
        Call CalcVsb
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Dim intLoop As Integer
    
    mblnFist = True
    Me.ctlClear.ShowPicture = False
    'Me.ctlPrinter.ShowPicture = False
    Me.ctlReturn.ShowPicture = False
    
    With Me.msfMain
'        .ColWidth(mCol.������Ŀ) = 4500
'        .TextMatrix(0, mCol.������Ŀ) = "������Ŀ"
'
'        .ColWidth(mCol.������) = 2000
'        .TextMatrix(0, mCol.������) = "������"
'
'        .ColWidth(mCol.����ʱ��) = 2500
'        .TextMatrix(0, mCol.����ʱ��) = "����ʱ��"
'
'        .ColWidth(mCol.�����) = 2000
'        .TextMatrix(0, mCol.�����) = "�����"
'
'        .ColWidth(mCol.���ʱ��) = 2500
'        .TextMatrix(0, mCol.���ʱ��) = "���ʱ��"
'
'        .ColWidth(mCol.״̬) = 2000
'        .TextMatrix(0, mCol.״̬) = "״̬"
'
'        .ColWidth(mCol.˵��) = 4000
'        .TextMatrix(0, mCol.˵��) = "˵��"
        
        .ColWidth(mCol.������Ŀ) = 6000
        .TextMatrix(0, mCol.������Ŀ) = "������Ŀ"
        
        .ColWidth(mCol.������) = 0
        .TextMatrix(0, mCol.������) = "������"

        .ColWidth(mCol.����ʱ��) = 0
        .TextMatrix(0, mCol.����ʱ��) = "����ʱ��"

        .ColWidth(mCol.�����) = 0
        .TextMatrix(0, mCol.�����) = "�����"

        .ColWidth(mCol.���ʱ��) = 0
        .TextMatrix(0, mCol.���ʱ��) = "���ʱ��"
        
        .ColWidth(mCol.״̬) = 5000
        .TextMatrix(0, mCol.״̬) = "״̬"
        
        .ColWidth(mCol.˵��) = 8500
        .TextMatrix(0, mCol.˵��) = "˵��"
        
        .ColWidth(mCol.��ӡ����) = 0
        .TextMatrix(0, mCol.��ӡ����) = "��ӡ����"
        
        .ColWidth(mCol.����id) = 0
        .TextMatrix(0, mCol.����id) = "����ID"
        
        .ColWidth(mCol.ҽ��id) = 0
        .TextMatrix(0, mCol.ҽ��id) = "ҽ��ID"
        
        .ColWidth(mCol.�걾ID) = 0
        .TextMatrix(0, mCol.�걾ID) = "�걾ID"
    End With
    
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.hwnd)
    Call mobjICCard.SetParent(Me.hwnd)
    
    mInputType = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPrinterSetup", "���ҷ�ʽ", 0)
    mstrSource = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPrinterSetup", "������Դ", "0,0,0")
    mstrNO = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPrinterSetup", "���Ƶ���", "")
    mintClear = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPrinterSetup", "��ӡ��������", 0)
    mintBack = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPrinterSetup", "��ӡ����󷵻���ҳ", 0)
    tmrReturn.Enabled = False
    Select Case mInputType
    
        Case 0              '����
            Me.lblID.Caption = "��ɨ�����룺"
        Case 1              'סԺ
            Me.lblID.Caption = "��ɨ�����룺"
        Case 2              '����
            Me.lblID.Caption = "��ɨ�����룺"
        Case 3              '���￨
            Me.lblID.Caption = "��ˢ���￨��"
        Case 4              'IC��
            Me.lblID.Caption = "��ˢIC����"
        Case 5              '����ID
            Me.lblID.Caption = "��ɨ�����룺"
    
    End Select
    Call Form_Resize
    Call InitSysPar
End Sub

Private Sub Form_Resize()
    With Me.fraMain
        .Width = Me.Width - 100
    End With
    
    With Me.fratak
        .Width = fraMain.Width - 100
    End With

    With Me.TxtID
        .Width = fraMain.Width - .Left - 100
    End With
    
    With Me.msfMain
        .Width = Me.Width - 100
        .Height = Me.Height - .Top - 1000
    End With
    
    With Me.ctlClear
        .Left = Me.Width - .Width - 300
        .Top = Me.Height - .Height - 250
    End With
    
'    With Me.ctlPrinter
'        .Left = Me.ctlClear.Left - .Width - 300
'        .Top = Me.ctlClear.Top
'    End With
    
    With Me.ctlReturn
        .Left = Me.Left + 900
        .Top = Me.ctlClear.Top
    End With
    
    With Me.lbl��ʾ
        .Left = 10920
    End With
    
    picBack.Enabled = mlngHelpPage > 0
    picBack1.Enabled = mlngHelpPage > 0
    hsb.Enabled = mlngHelpPage > 0
    vsb.Enabled = mlngHelpPage > 0

    shp.Visible = mlngHelpPage > 0
    picBack.Visible = mlngHelpPage > 0
    picBack1.Visible = mlngHelpPage > 0
    hsb.Visible = mlngHelpPage > 0
    vsb.Visible = mlngHelpPage > 0
    
    If mlngHelpPage > 0 Then
        With Me.msfMain
            .Width = Me.Width - 100
            .Height = Me.Height - .Top - 4000
        End With
        
        
        QueryItem.Width = Screen.Width - 2010 - 45
        Call ResizeControl(shp, 15, Me.msfMain.Top + Me.msfMain.Height + 30, Me.ScaleWidth - 30, Me.ScaleHeight - (Me.msfMain.Top + Me.msfMain.Height + 30) - (Me.ScaleHeight - ctlClear.Top) - 100)
        
        Call ResizeControl(picBack, 45, Me.shp.Top + 30, Me.shp.Width - 60, Me.shp.Height - 60)
        Call ResizeControl(QueryItem, picBack.Left, 30, QueryItem.Width, QueryItem.Height)
        
        mvarLeftStart = QueryItem.Left
        
        Call ResizeControl(vsb, picBack.ScaleWidth - vsb.Width + 60, 0, vsb.Width, picBack.ScaleHeight - hsb.Height + 60)
        Call ResizeControl(hsb, 0, picBack.ScaleHeight - hsb.Height + 60, picBack.ScaleWidth - vsb.Width + 60, hsb.Height)
        picBack1.Left = vsb.Left
        picBack1.Top = hsb.Top
        
        Call CalcVsb
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNO As String)
    If Not TxtID.Locked And TxtID.Text = "" And Me.ActiveControl Is TxtID Then
        TxtID.Text = strCardNO
        Call TxtID_KeyPress(vbKeyReturn)
    End If
End Sub

Private Sub msfMain_Click()
    Me.TxtID.SetFocus
End Sub

Private Sub msfMain_SelChange()
    With Me.msfMain
        If .TextMatrix(.Row, mCol.״̬) = "���ɴ�ӡ" Then
            .ForeColorSel = &HC0&
        Else
            .ForeColorSel = &HFF0000
        End If
    End With
End Sub

Private Sub tmrReturn_Timer()
    ''��ʱ���
    If mintReturn - 1 < 0 Then
        Call ctlClear_CommandClick
        If mintBack = 1 Then
            ctlReturn_CommandClick
        End If
    Else
        mintReturn = mintReturn - 1
        lbl��ʾ.Caption = mintReturn & "���,�����������Ϣ!"
        If mintBack = 1 Then
        
             lbl��ʾ.Caption = lbl��ʾ.Caption & "��������ҳ"
        End If
    End If
End Sub

Private Sub TxtID_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (TxtID.Text = "" And Me.ActiveControl Is TxtID)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (TxtID.Text = "" And Me.ActiveControl Is TxtID)
End Sub

Private Sub TxtID_GotFocus()
    Me.TxtID.SelStart = 0
    Me.TxtID.SelLength = Len(Me.TxtID)
    If Not mobjIDCard Is Nothing And TxtID.Text = "" And Not TxtID.Locked Then mobjIDCard.SetEnabled (True)
    If Not mobjICCard Is Nothing And TxtID.Text = "" And Not TxtID.Locked Then mobjICCard.SetEnabled (True)
End Sub

Private Sub TxtID_KeyPress(KeyAscii As Integer)
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim intRow As Integer, intCol As Integer
    Dim blnPrinter As Boolean           '�Ƿ��п��Դ�ӡ��
    Dim blnCard As Boolean
    Dim strSource As String
    Dim strAdvice As Long               'ҽ��id
    Dim strTimeHorizon As String
    Dim intDays As Integer
    Dim intPrintend As Integer           '����ɴ�ӡ
    Dim intNotPrint As Integer           'δ��ӡ
    Dim intPrinting As Integer           '��ӡ��
    
    
    If CheckIsInclude(UCase(Chr(KeyAscii)), "'����;��:��?��|,��.��""") = True Then KeyAscii = 0
    Call zlCommFun.InputIsCard(TxtID, KeyAscii, glngSys)
    
    
    TxtID.Text = ReplaseSpecial(TxtID.Text)
    '�Ƿ�ˢ�����
    
    blnCard = KeyAscii <> 8 And Len(TxtID.Text) = gbytCardNOLen - 1 And TxtID.SelLength <> Len(TxtID.Text)
    If KeyAscii = 13 Then blnCard = True
    If mInputType = 3 Then
        '���￨ֻҪ�����λ������ִ��
        If blnCard = False Then Exit Sub
        If KeyAscii <> 13 Then
            Me.TxtID = Me.TxtID & Chr(KeyAscii)
        End If
        KeyAscii = 0
    Else
        If KeyAscii <> 13 Then Exit Sub
    End If
    
    strTimeHorizon = GetPara("�������ڷ�Χ", "0")
    If Split(strTimeHorizon, "-")(0) = "1" Then
        intDays = Val(Split(strTimeHorizon, "-")(1))
    Else
        intDays = 30
    End If
    
    If Me.TxtID = "" Then Exit Sub
    
    blnPrinter = False
    
    On Error GoTo errH
    
    strSQL = "Select /*+ rule */" & vbNewLine & _
            " A.ҽ������ As ������Ŀ, " & vbNewLine & _
            "         Decode(B.ҽ��id, Null, '1-δ����', Decode(B.������, Null, '2-δ����', Decode(B.������, Null, '3-�Ѳ���', '4-�ѽ���'))) As ״̬," & vbNewLine & _
            "         '' As ������, '' As ����ʱ��, '' As �����, '' As ���ʱ��, '' As ��ӡ����, " & vbNewLine & _
            "         e.����,e.�Ա�,e.����,a.���id as ҽ��ID,a.����ID,null as �걾ID " & vbNewLine & _
            "From ����ҽ����¼ A, ����ҽ������ B, ���ű� D, ������Ϣ E,��������Ӧ�� F,�����ļ��б� G " & vbNewLine & _
            "Where A.Id = B.ҽ��id And A.��������id = D.Id And A.����id = E.����id And nvl(B.ִ��״̬,0) = 0 And A.������� = 'C' " & vbNewLine & _
            " And a.������Ŀid = f.������ĿID and f.�����ļ�id = g.id  And decode(a.������Դ,3,1,a.������Դ) = f.Ӧ�ó��� " & vbNewLine & _
            " and a.������Դ in (Select * From Table(Cast(f_Num2list([2]) As zlTools.t_Numlist))) " & vbNewLine & _
            " And g.��� In (Select * From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist))) " & vbNewLine & _
            " And a.����ʱ�� + 0 between [4] and [5] " & vbNewLine & _
            "��ѯ����1" & vbNewLine

            strSQL = strSQL & " Union All" & vbNewLine & _
            "Select /*+ rule */" & vbNewLine & _
            "Distinct A.������Ŀ, Decode(����״̬, 1, '5-�Ѻ���', Decode(Sign(Nvl(��ӡ����, 0)), 1, '7-�Ѵ�ӡ', '6-�����')) As ״̬, A.������," & vbNewLine & _
            "         To_Char(A.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��, A.�����, To_Char(A.���ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ���ʱ��," & vbNewLine & _
            "         Decode(Nvl(A.��ӡ����, 0), 0, '', '��') As ��ӡ����,e.����,e.�Ա�,e.����, " & vbNewLine & _
            "         a.ҽ��ID,a.����ID,g.�걾ID " & vbNewLine & _
            "From ����걾��¼ A, ����ҽ����¼ B, ����ҽ������ D, ���ű� F, ������Ϣ E, ������Ŀ�ֲ� G,��������Ӧ�� H,�����ļ��б� I " & vbNewLine & _
            "Where A.Id = G.�걾id And G.ҽ��id = B.���id And B.���id = D.ҽ��id And B.��������id = F.Id And A.����id = E.����id" & vbNewLine & _
            " And b.������Ŀid = h.������ĿID And H.�����ļ�id = I.id And decode(b.������Դ,3,1,b.������Դ) = H.Ӧ�ó��� " & vbNewLine & _
            " And b.������Դ in (Select * From Table(Cast(f_Num2list([2]) As zlTools.t_Numlist))) " & vbNewLine & _
            " And  I.��� In (Select * From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist))) " & vbNewLine & _
            " And a.����ʱ�� + 0 between [4] and [5] " & vbNewLine & _
            "��ѯ����2 " & vbNewLine & _
            " order by ״̬ "


    '��������:0=����;1=סԺ;2=����;3=���￨;4=IC��;5=����ID
    Select Case mInputType
    
        Case 0              '����
            strSQL = Replace$(strSQL, "��ѯ����1", " And b.�������� = [1] ")
            strSQL = Replace$(strSQL, "��ѯ����2", " And d.�������� = [1] ")
        Case 1              'סԺ
            strSQL = Replace$(strSQL, "��ѯ����1", " And e.סԺ�� = [1] ")
            strSQL = Replace$(strSQL, "��ѯ����2", " And e.סԺ�� = [1] ")
        Case 2              '����
            strSQL = Replace$(strSQL, "��ѯ����1", " And e.����� = [1] ")
            strSQL = Replace$(strSQL, "��ѯ����2", " And e.����� = [1] ")
        Case 3              '���￨
            strSQL = Replace$(strSQL, "��ѯ����1", " And e.���￨�� = [1] ")
            strSQL = Replace$(strSQL, "��ѯ����2", " And e.���￨�� = [1] ")
        Case 4              'IC��
            strSQL = Replace$(strSQL, "��ѯ����1", " And e.IC���� = [1] ")
            strSQL = Replace$(strSQL, "��ѯ����2", " And e.IC���� = [1] ")
        Case 5              '����ID
            strSQL = Replace$(strSQL, "��ѯ����1", " And e.����ID = [1] ")
            strSQL = Replace$(strSQL, "��ѯ����2", " And e.����ID = [1] ")
    End Select
    strSource = IIf(Split(mstrSource, ",")(0) = 1, "1,3", 0)
    strSource = strSource & "," & IIf(Split(mstrSource, ",")(1) = 1, "2", 0)
    strSource = strSource & "," & IIf(Split(mstrSource, ",")(2) = 1, "4", 0)
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Me.TxtID, strSource, mstrNO, CDate(Format(Now - intDays, "yyyy-mm-dd 00:00:00")), CDate(Format(Now, "yyyy-mm-dd 23:59:59")))
    
    '��ռ�¼
    Call ctlClear_CommandClick
    
    'д���¼
    If rsTmp.RecordCount > Me.msfMain.Rows - 1 Then
        Me.msfMain.Rows = rsTmp.RecordCount + 1
    End If
    intRow = 0
    If rsTmp.RecordCount > 0 Then
        Me.lbl���� = "����:" & Nvl(rsTmp("����"))
        Me.lbl���� = "����:" & Nvl(rsTmp("����"))
        Me.lbl�Ա� = "�Ա�:" & Nvl(rsTmp("�Ա�"))
    End If
    Do While Not rsTmp.EOF
        If strAdvice <> Nvl(rsTmp("ҽ��id")) And Nvl(rsTmp("ҽ��id")) <> "" Then
            intRow = intRow + 1
            Me.msfMain.TextMatrix(intRow, mCol.������Ŀ) = Nvl(rsTmp("������Ŀ"), "")
            
            Me.msfMain.TextMatrix(intRow, mCol.������) = Nvl(rsTmp("������"), "")
            Me.msfMain.TextMatrix(intRow, mCol.����ʱ��) = Nvl(rsTmp("����ʱ��"), "")
            Me.msfMain.TextMatrix(intRow, mCol.�����) = Nvl(rsTmp("�����"), "")
            Me.msfMain.TextMatrix(intRow, mCol.���ʱ��) = Nvl(rsTmp("���ʱ��"), "")
            With Me.msfMain
                Select Case Nvl(rsTmp("״̬"), "")
                    Case "1-δ����", "2-δ����", "3-�Ѳ���", "4-�ѽ���", "5-�Ѻ���"
                        Me.msfMain.TextMatrix(intRow, mCol.״̬) = "���ɴ�ӡ"
                        .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = vbRed
                        .Cell(flexcpFontBold, intRow, 0, intRow, .Cols - 1) = True
                        intNotPrint = intNotPrint + 1
                    Case "7-�Ѵ�ӡ"
                        Me.msfMain.TextMatrix(intRow, mCol.״̬) = "���ɴ�ӡ"
                        .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = vbRed
                        .Cell(flexcpFontBold, intRow, 0, intRow, .Cols - 1) = True
                    Case "6-�����"
                        Me.msfMain.TextMatrix(intRow, mCol.״̬) = "���Դ�ӡ"
                        .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = vbBlack
                        .Cell(flexcpFontBold, intRow, 0, intRow, .Cols - 1) = False
                        intPrinting = intPrinting + 1
                End Select
            End With
            Select Case Nvl(rsTmp("״̬"), "")
                Case "1-δ����"
                    Me.msfMain.TextMatrix(intRow, mCol.˵��) = "ҽ��δ���ͣ�"
                Case "2-δ����"
                    Me.msfMain.TextMatrix(intRow, mCol.˵��) = "�걾δ������"
                Case "3-�Ѳ���"
                    Me.msfMain.TextMatrix(intRow, mCol.˵��) = "�걾�ȴ�����..."
                Case "4-�ѽ���"
                    Me.msfMain.TextMatrix(intRow, mCol.˵��) = "�걾���ڼ��飬���Ժ�������"
                Case "5-�Ѻ���"
                    Me.msfMain.TextMatrix(intRow, mCol.˵��) = "�걾���ڼ��飬���Ժ�������"
                Case "6-�����"
                    
                Case "7-�Ѵ�ӡ"
                    Me.msfMain.TextMatrix(intRow, mCol.˵��) = "�Ѵ�ӡ�����ٴ�ӡ��"
            End Select
            Me.msfMain.TextMatrix(intRow, mCol.��ӡ����) = Nvl(rsTmp("��ӡ����"), "")
            If Nvl(rsTmp("��ӡ����"), "") = "" Then
                blnPrinter = True
            End If
            Me.msfMain.TextMatrix(intRow, mCol.ҽ��id) = Nvl(rsTmp("ҽ��ID"), "")
            Me.msfMain.TextMatrix(intRow, mCol.����id) = Nvl(rsTmp("����ID"), "")
            Me.msfMain.TextMatrix(intRow, mCol.�걾ID) = Nvl(rsTmp("�걾ID"), "")
        End If
        strAdvice = Nvl(rsTmp("ҽ��id"))
        
        rsTmp.MoveNext
    Loop
    Me.lbl��ʾ.Caption = "���ڴ�ӡ����" & intPrinting & "�ţ������б���" & intNotPrint & "��δ��ӡ��"
    Call msfMain_SelChange
    Me.TxtID.Text = ""
    Me.TxtID.SetFocus
    
    '��ӡ����
    If blnPrinter Then
        With Me.msfMain
            For intRow = 1 To .Rows - 1
                If .TextMatrix(intRow, mCol.�����) <> "" And .TextMatrix(intRow, mCol.��ӡ����) = "" Then
                    .TextMatrix(intRow, mCol.״̬) = "���ڴ�ӡ��"
                    '����˲�û�д�ӡ��ʱ�Ž��д�ӡ
                    If ReportPrint(Val(.TextMatrix(intRow, mCol.ҽ��id)), Val(.TextMatrix(intRow, mCol.�걾ID)), Val(.TextMatrix(intRow, mCol.����id)), True) = True Then
                        .TextMatrix(intRow, mCol.��ӡ����) = Val(.TextMatrix(intRow, mCol.��ӡ����)) + 1
                        .TextMatrix(intRow, mCol.״̬) = "�Ѵ�ӡ"
                        .TextMatrix(intRow, mCol.˵��) = "�Ѵ�ӡ�����ٴ�ӡ��"
                        .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = vbRed
                        .Cell(flexcpFontBold, intRow, 0, intRow, .Cols - 1) = True
'                        intPrinting = intPrinting - 1
                        Me.lbl��ʾ.Caption = "���ι���ӡ����" & intPrinting & "�ţ������б���" & intNotPrint & "��δ��ӡ��"
'                        Me.lbl��ʾ.Caption = "���ι���ӡ����" & intPrinting & "�š�"
                    Else
                        .TextMatrix(intRow, mCol.״̬) = "��ӡʧ��"
                    End If
                    If mintPrintDelayed > 0 Then
                        Call Sleep(mintPrintDelayed * 1000)
                    End If
                End If
            Next
            Me.lbl��ʾ.Caption = "��ӡ��ɣ���ע��ȡ�����б��棡"
            Call Sleep(2 * 1000)
            Me.lbl��ʾ.Caption = ""
        End With
        Call msfMain_SelChange
        Me.TxtID.SetFocus
    End If
    If mintClear > 0 Then
        If rsTmp.RecordCount > 0 Then
            tmrReturn.Interval = 1000
            mintReturn = mintClear
            tmrReturn.Enabled = True
        Else
            If mintBack = 1 Then
                tmrReturn.Interval = 1000
                mintReturn = 5
                tmrReturn.Enabled = True
            End If
        End If
    Else
        If mintBack = 1 Then
            tmrReturn.Interval = 1000
            mintReturn = 5
            tmrReturn.Enabled = True
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Function ReadImageData(lngKeyID As Long) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim DrawIndex As Integer
    Dim StrTime As Date
    On Error GoTo errH
    StrTime = Now
    gstrSQL = "select id ,�걾ID,ͼ������ from ����ͼ���� where �걾id = [1] order by ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKeyID)

    
    Do Until rsTmp.EOF
        If Dir(App.Path & "\" & rsTmp("ID") & ".cht") = "" Then
             Call LoadImageData(App.Path, rsTmp("ID"))
        End If
        DrawIndex = DrawIndex + 1
        rsTmp.MoveNext
    Loop
    ReadImageData = True
    Exit Function
    
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function ReportPrint(ByVal lngҽ��ID As Long, ByVal lngKey As Long, ByVal lng����ID As Long, ByVal blnPrint As Boolean) As Boolean
    '���������ӡ
    
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim blnCurrMoved As Boolean
    Dim strSQL As String
    Dim strChart(1 To 9) As String
    Dim intLoop As Integer
    
    ReportPrint = False
    Me.MousePointer = 11
    zlCommFun.ShowFlash "���ڴ�ӡ��ȴ�...", Me
    
    '����ͼ�ι��Զ��屨�����
    ReadImageData lngKey
    strSQL = "select id from ����ͼ���� where �걾id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
    intLoop = 1
    Do Until rsTmp.EOF
        strChart(intLoop) = App.Path & "\" & rsTmp("ID") & ".cht"
        intLoop = intLoop + 1
        rsTmp.MoveNext
    Loop
    
    
    If GetReportCode(lngҽ��ID, lng���ͺ�, strReportCode, strReportParaNo, bytReportParaMode, blnCurrMoved) Then
        Call ReportOpen(gcnOracle, glngSys, strReportCode, Me, "NO=" & strReportParaNo, "����=" & bytReportParaMode, "ҽ��ID=" & lngҽ��ID, _
                        "����ID=" & lng����ID, "�걾ID=" & lngKey, "���ҽ��=" & lngҽ��ID, "����걾=" & lngKey, _
                        "ͼ��1=" & strChart(1), "ͼ��2=" & strChart(2), "ͼ��3=" & strChart(3), "ͼ��4=" & strChart(4), _
                        "ͼ��5=" & strChart(5), "ͼ��6=" & strChart(6), "ͼ��7=" & strChart(7), "ͼ��8=" & strChart(8), _
                        "ͼ��9=" & strChart(9), IIf(blnPrint, 2, 1))
    End If
    
    
    On Error GoTo errH

    gstrSQL = " select id from ����걾��¼ where ҽ��id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngҽ��ID)
    Do Until rsTmp.EOF
        strSQL = "ZL_����걾��¼_�걾�ʿ�(" & rsTmp("ID") & ",'',1)"
        zlDatabase.ExecuteProcedure strSQL, gstrSysName
        rsTmp.MoveNext
    Loop
    
    Me.MousePointer = 0
    zlCommFun.StopFlash
    ReportPrint = True
    On Error Resume Next
    'ɾ��ͼ���ļ�
    For intLoop = 1 To 9
        Kill strChart(intLoop)
    Next
    
    Exit Function
errH:
    Me.MousePointer = 0
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetReportCode(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, ByRef strCode As String, ByRef strNo As String, ByRef bytMode As Byte, Optional ByVal DataMoved As Boolean = False) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '����;
    '--------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If lngҽ��ID = 0 And lng���ͺ� = 0 Then Exit Function
    
'    strSQL = "SELECT DISTINCT 'ZLCISBILL'||Trim(To_Char(C.���,'00000'))||'-2' AS ������," & _
                       "A.NO," & _
                       "A.��¼���� " & _
                "FROM ����ҽ������ A,�����ļ��б� C,����ҽ����¼ D,��������Ӧ�� E " & _
                "Where E.�����ļ�id = C.ID " & _
                        "AND D.������ĿID=E.������ĿID " & _
                      "AND A.ҽ��ID=D.ID AND E.Ӧ�ó���=Decode(D.������Դ,2,2,4,4,1) " & _
                      " AND D.���id= [1] "
                      
    strSQL = "Select Distinct 'ZLCISBILL' || Trim(To_Char(C.���, '00000')) || '-2' As ������, A.NO, A.��¼����, F.ID, F.����" & vbNewLine & _
            "From ����ҽ������ A, �����ļ��б� C, ����ҽ����¼ D, ��������Ӧ�� E, ������ĿĿ¼ F" & vbNewLine & _
            "Where E.�����ļ�id = C.ID And D.������Ŀid = E.������Ŀid And D.������Ŀid = F.ID And A.ҽ��id = D.ID And" & vbNewLine & _
            "      E.Ӧ�ó��� = Decode(D.������Դ, 2, 2, 4, 4, 1) And D.���id = [1] " & vbNewLine & _
            "Order By F.���� "
                          
    If DataMoved Then
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If

'    strSQL = "SELECT DISTINCT 'ZLCISBILL'||Trim(To_Char(C.���,'00000'))||'-2' AS ������," & _
'                       "A.NO," & _
'                       "A.��¼���� " & _
'                "FROM ��������Ӧ�� A,�����ļ�Ŀ¼ C,����ҽ����¼ D,����ҽ������ B " & _
'                "Where A.�����ļ�id = C.ID " & _
'                      "AND A.������Ŀid=D.������ĿID " & _
'                      "AND B.����ID=D.����ID " & _
'                      "AND NVL(B.��ҳID,0)=NVL(D.��ҳID,0) " & _
'                      "AND B.�ļ�id=C.ID " & _
'                      "AND D.���id=" & lngҽ��id & " " & _
'                      "AND A.���ͺ�=" & lng���ͺ�

    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlLISWork", lngҽ��ID, lng���ͺ�)
                      
    
    If rs.BOF = False Then
        strCode = zlCommFun.Nvl(rs("������"))
        strNo = zlCommFun.Nvl(rs("NO"))
        bytMode = zlCommFun.Nvl(rs("��¼����"), 1)
    End If
    
    GetReportCode = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    ' 2007-08-17 ����һ��֧ͨ��
    Dim lngPreIDKind As Long
    If Not TxtID.Locked And TxtID.Text = "" And Me.ActiveControl Is TxtID Then
        TxtID.Text = strID
        Call TxtID_KeyPress(vbKeyReturn)
    End If
End Sub



Private Sub LoadPageItemList(ByVal PageNo As Long)
'����:����ҳ���ÿһ��ѯ��Ŀ
'����:PageNo            ҳ�����
'˵��:���ǲ�ѯ������ʾ�����岿��,��ʾ��ѯ����
    Dim FileName As String
    Dim W As Single
    Dim H As Single
    Dim vFont As New StdFont
    Dim i As Long
    Dim j As Long
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
    Dim vNextY As Single
    Dim vNextX As Single
    Dim objDraw As ctlQueryItem
    Dim vWidth As Single
    Dim vHeight As Single
    Dim vTmp As Single
    Dim vTmp1 As Single
    Dim vMaxWidth As Single
    Dim vVisible As Boolean
    Dim strText As String
    
    On Error GoTo errHand
    i = 1
    vNextY = 60 + (i - 1) * 600
    vNextX = 120
    vMaxWidth = 120
            
    ShowFlatFlash "���Ժ���������ҳ��...", Me
    DoEvents
    
    Set objDraw = QueryItem
    objDraw.ClientVisible = False
    Call objDraw.ClearAllPageItem
    
    '��ȡҳ��ı������������
'    Set gRs = OpenRecord(gRs, "select B.����,B.���� from ��ѯҳ��Ŀ¼ A,��ѯͼƬԪ�� B where A.��������=B.��� and A.ҳ�����=" & PageNo)
'    If gRs.BOF = False Then FrameDefault.AdviceMovie = IIf(IsNull(gRs!����), "", App.Path & "\ͼ��\" & gRs!���� & IIf(gRs!���� <> 2, ".pic", ".swf"))
                    
    '��ʼ�����Զ����ѯҳ��
    gstrSQL = "select ҳ�����,�������,�����ı�,����ͼ��,��������,����λ��,��������,����ҳ��,��������,��������,������,���λ��,��ͼ���,��ͼλ�� from ��ѯ����Ŀ¼ where ҳ�����=[1] order by �������"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo)
    If gRs.BOF = False Then
        While Not gRs.EOF
            strTmp = IIf(IsNull(gRs!��������), "����;12;0;0;0", gRs!��������)
            vFont.Name = Split(strTmp, ";")(0)
            vFont.Size = Val(Split(strTmp, ";")(1))
            vFont.Bold = Val(Split(strTmp, ";")(2))
            vFont.Italic = Val(Split(strTmp, ";")(3))
                                    
            FileName = ""
            '1.���ر������ݼ�����ͼ��
            vVisible = IIf(IsNull(gRs!��������), 1, gRs!��������)
            
            gstrSQL = "select ���� from ��ѯͼƬԪ�� where ���=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(IIf(IsNull(gRs!����ͼ��), 0, gRs!����ͼ��)))
            If rs.BOF = False Then
                FileName = GetFileName(IIf(IsNull(gRs!����ͼ��), 0, gRs!����ͼ��), W, H)
            End If
            Call objDraw.AddPageItemTitle(i, vNextY, IIf(IsNull(gRs!�����ı�), "", gRs!�����ı�), Val(Split(strTmp, ";")(4)), vFont, FileName, PageNo, IIf(IsNull(gRs!�������), 0, gRs!�������), vWidth, vHeight, Not vVisible, IIf(IsNull(gRs!����λ��), 0, gRs!����λ��))
                                                                                    
            If Not vVisible = True Then vNextY = vNextY + vHeight + 150
            
            
            Select Case zlCommFun.Nvl(gRs("��������").Value, 0)
            '----------------------------------------------------------------------------------------------------------
            Case 0          '���ı�����
                strTmp = IIf(IsNull(gRs!��������), "����;12;0;0;0", gRs!��������)
                j = objDraw.NextTxtIndex
                
                vWidth = QueryItem.Width - 330
                
                strText = Sys.ReadLob(glngSys, 29, PageNo & "," & gRs("�������").Value, "", 1)
                
                Call objDraw.AddPageItemTxt(j, vNextX, vNextY, strText & Chr(13) & Chr(10), strTmp, vWidth, vHeight)
                vNextY = vNextY + vHeight + 150
            '----------------------------------------------------------------------------------------------------------
            Case 1          '���������
                vHeight = 0
                Call InsertGrid(objDraw, IIf(IsNull(gRs!������), 0, gRs!������), vNextX, vNextY, vWidth, vHeight)
                If vHeight > 0 Then vNextY = vNextY + vHeight + 150
            '----------------------------------------------------------------------------------------------------------
            Case 2          '��ͼ������
                FileName = GetFileName(IIf(IsNull(gRs!��ͼ���), 0, gRs!��ͼ���), W, H)
                Call objDraw.AddPageItemPic(objDraw.NextPicIndex, 0, vNextY, FileName, vWidth, vHeight, W, H)
                vNextY = vNextY + vHeight + 150
            '----------------------------------------------------------------------------------------------------------
            Case 3          '����������
                gstrSQL = "select C.ҳ������||decode(B.�����ı�,NULL,'','��'||B.�����ı�) as �����ı�,A.����ҳ��,A.ҳ�ڶκ� from ��ѯ�������� A,��ѯ����Ŀ¼ B,��ѯҳ��Ŀ¼ C Where A.����ҳ��=C.ҳ����� and A.����ҳ��=B.ҳ�����(+) and A.ҳ�ڶκ�=B.�������(+) and A.ҳ����� = [1] And A.������� = [2]"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo, Val(IIf(IsNull(gRs!�������), 0, gRs!�������)))
                If rs.BOF = False Then
                    While Not rs.EOF
                        Call objDraw.AddPageItemConnect(objDraw.NextConnectIndex, vNextX + 150, vNextY, IIf(IsNull(rs!�����ı�), "", rs!�����ı�), IIf(IsNull(rs!����ҳ��), 0, rs!����ҳ��), IIf(IsNull(rs!ҳ�ڶκ�), 0, rs!ҳ�ڶκ�), vWidth, vHeight)
                        vNextY = vNextY + 300
                        rs.MoveNext
                    Wend
                    vNextY = vNextY + 150
                Else
                    '����Ƿ����ӵ�ZLHIS����Ա
                    gstrSQL = "select B.����,A.����ҳ��,A.ҳ�ڶκ� from ��ѯ�������� A,��Ա�� B Where A.ҳ�ڶκ�=B.id And (b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.����ʱ�� Is Null) and A.ҳ����� = [1] And A.������� = [2]"
                    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo, Val(IIf(IsNull(gRs!�������), 0, gRs!�������)))
                    If rs.BOF = False Then
                        While Not rs.EOF
                            Call objDraw.AddPageItemConnect(objDraw.NextConnectIndex, vNextX, vNextY, IIf(IsNull(rs!����), "", rs!����), IIf(IsNull(rs!����ҳ��), 0, rs!����ҳ��), IIf(IsNull(rs!ҳ�ڶκ�), 0, rs!ҳ�ڶκ�), vWidth, vHeight)
                            vNextY = vNextY + 300
                            rs.MoveNext
                        Wend
                        vNextY = vNextY + 150
                    End If
                End If

            '----------------------------------------------------------------------------------------------------------
            Case 4          '�ı��ͱ��
                strTmp = IIf(IsNull(gRs!��������), "����;12;0;0;0", gRs!��������)
                j = objDraw.NextTxtIndex
                strText = Sys.ReadLob(glngSys, 29, PageNo & "," & gRs("�������").Value, "", 1)
                Select Case IIf(IsNull(gRs!���λ��), 0, gRs!���λ��)
                Case 0
                    vHeight = 0
                    Call InsertGrid(objDraw, IIf(IsNull(gRs!������), 0, gRs!������), 0, vNextY, vTmp1, vTmp)
                    vWidth = QueryItem.Width - vTmp1 - 120 - 120
                    Call objDraw.AddPageItemTxt(j, vNextX + vTmp1 + 60, vNextY, strText & Chr(13) & Chr(10), strTmp, vWidth, vHeight)
                Case 1
                    Call objDraw.AddPageItemTxt(j, vNextX, vNextY, strText & Chr(13) & Chr(10), strTmp, vTmp1, vHeight)
                    Call InsertGrid(objDraw, IIf(IsNull(gRs!������), 0, gRs!������), 1, vNextY, vWidth, vTmp)
                End Select
                vNextY = vNextY + IIf(vTmp > vHeight, vTmp, vHeight) + 150
            '----------------------------------------------------------------------------------------------------------
            Case 5          '�ı���ͼ��
                FileName = GetFileName(IIf(IsNull(gRs!��ͼ���), 0, gRs!��ͼ���), W, H)
                strText = Sys.ReadLob(glngSys, 29, PageNo & "," & gRs("�������").Value, "", 1)
                strTmp = IIf(IsNull(gRs!��������), "����;12;0;0;0", gRs!��������)
                j = objDraw.NextTxtIndex
                Select Case IIf(IsNull(gRs!��ͼλ��), 0, gRs!��ͼλ��)
                Case 0
                    Call objDraw.AddPageItemPic(objDraw.NextPicIndex, 0, vNextY, FileName, vTmp1, vTmp, W, H)
                    vWidth = QueryItem.Width - vTmp1 - 120 - 120
                    Call objDraw.AddPageItemTxt(j, vNextX + vTmp1 + 60, vNextY, strText & Chr(13) & Chr(10), strTmp, vWidth, vHeight)
                Case 1
                    Call objDraw.AddPageItemPic(objDraw.NextPicIndex, 1, vNextY, FileName, vWidth, vTmp, W, H)
                    vTmp1 = QueryItem.Width - vWidth - 60 - 90
                    Call objDraw.AddPageItemTxt(j, vNextX, vNextY, strText & Chr(13) & Chr(10), strTmp, vTmp1, vHeight)
                End Select
                vNextY = vNextY + IIf(vTmp > vHeight, vTmp, vHeight) + 150
            End Select
                        
            '8.���÷���ҳ�ױ�־
            If IIf(IsNull(gRs!����ҳ��), 0, gRs!����ҳ��) = 1 Then
                vHeight = 0
                Call objDraw.AddReturnFlag(vNextX, vNextY, vHeight)
                If vHeight > 0 Then vNextY = vNextY + vHeight + 150
            End If
            
            i = i + 1
            gRs.MoveNext
        Wend
    End If
        
    Call objDraw.ResizePage(QueryItem.Width, vNextY)
    QueryItem.Height = vNextY
    'Call FrameDefault.InitNavigator(FrameDefault.ClientWidth, vNextY)
    
    '��ȡ����������ҳ�汳��
    gstrSQL = "select B.����,B.����,B.���,B.�߶� from ��ѯҳ��Ŀ¼ A,��ѯͼƬԪ�� B where A.ҳ�汳��=B.��� and A.ҳ�����=[1]"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo)
    If gRs.BOF = False Then
        Call objDraw.BackPicture(IIf(IsNull(gRs!����), "", App.Path & "\ͼ��\" & gRs!���� & IIf(gRs!���� <> 2, ".pic", ".swf")), IIf(IsNull(gRs!���), 0, gRs!���) * Screen.TwipsPerPixelX, IIf(IsNull(gRs!�߶�), 0, gRs!�߶�) * Screen.TwipsPerPixelY)
    End If
            
    Call objDraw.InitLoad
    objDraw.ClientVisible = True
    
    StopFlatFlash
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub hsb_Change()
    QueryItem.Left = mvarLeftStart - hsb.Value * 360
    If QueryItem.Left + QueryItem.Width < picBack.Left + picBack.Width - vsb.Width Then
        QueryItem.Left = picBack.Left + picBack.Width - QueryItem.Width - vsb.Width
    End If
    If QueryItem.Left > 0 Then QueryItem.Left = picBack.Width - vsb.Width - QueryItem.Width
End Sub

Private Sub hsb_KeyDown(KeyCode As Integer, Shift As Integer)
    Call picBack_KeyDown(KeyCode, Shift)
End Sub

Private Sub picBack_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If vsb.Enabled Then vsb.Value = IIf(vsb.Value < vsb.Max, vsb.Value + 1, vsb.Max)
    End If

    If KeyCode = vbKeyUp Then
        If vsb.Enabled Then vsb.Value = IIf(vsb.Value > 0, vsb.Value - 1, 0)
    End If

    If KeyCode = vbKeyRight Then
        If hsb.Enabled Then hsb.Value = IIf(hsb.Value < hsb.Max, hsb.Value + 1, hsb.Max)
    End If

    If KeyCode = vbKeyLeft Then
        If hsb.Enabled Then hsb.Value = IIf(hsb.Value > 0, hsb.Value - 1, 0)
    End If
End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call InitInternal
End Sub

Private Sub picBack_Paint()
    Call RaisEffect(picBack, -1)
End Sub

Private Sub picBack1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call picBack_KeyDown(KeyCode, Shift)
End Sub

Private Sub picBack1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call InitInternal
End Sub

Private Sub QueryItem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call InitInternal
End Sub

Private Sub vsb_Change()
    QueryItem.Top = 0 - vsb.Value * 360
    If QueryItem.Top + QueryItem.Height < picBack.Height - hsb.Height Then
        QueryItem.Top = picBack.Top + picBack.Height - hsb.Height - QueryItem.Height
    End If
    If QueryItem.Top > 30 Then QueryItem.Top = picBack.Height - hsb.Height - QueryItem.Height
    
End Sub

Private Sub CalcVsb()
    vsb.Max = 0 - Int(0 - (QueryItem.Height - picBack.ScaleHeight + hsb.Height + 45) / 360)
    If vsb.Max > 0 Then
        vsb.Enabled = True
        vsb.Visible = True
        vsb.SmallChange = 1
        vsb.LargeChange = 1
        vsb.Value = 0
        hsb.Width = picBack.Width - hsb.Width
    Else
        vsb.Enabled = False
        vsb.Visible = False
        hsb.Width = picBack.Width
    End If
    
    hsb.Max = 0 - Int(0 - (QueryItem.Width - picBack.ScaleWidth + vsb.Width + 45) / 360)
    If hsb.Max > 0 Then
        hsb.Enabled = True
        hsb.Visible = True
        hsb.SmallChange = 1
        hsb.LargeChange = 1
        hsb.Value = 0
        vsb.Height = picBack.Height - hsb.Height
    Else
        hsb.Enabled = False
        hsb.Visible = False
        vsb.Height = picBack.Height
    End If
End Sub

Private Sub vsb_KeyDown(KeyCode As Integer, Shift As Integer)
    Call picBack_KeyDown(KeyCode, Shift)
End Sub

