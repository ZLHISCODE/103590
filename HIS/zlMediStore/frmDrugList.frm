VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDrugList 
   BackColor       =   &H8000000C&
   Caption         =   "ҩƷ��ϸ��"
   ClientHeight    =   5445
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin VB.Frame shpback 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "ҩƷ"
      Height          =   4065
      Left            =   540
      TabIndex        =   2
      Top             =   885
      Width           =   5880
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgdData 
         Height          =   2565
         Left            =   495
         TabIndex        =   3
         Top             =   1170
         Width           =   5070
         _ExtentX        =   8943
         _ExtentY        =   4524
         _Version        =   393216
         BackColor       =   16777215
         Rows            =   10
         FixedCols       =   0
         BackColorFixed  =   16777215
         BackColorBkg    =   16777215
         SelectionMode   =   1
         AllowUserResizing=   3
         Appearance      =   0
         BandDisplay     =   1
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label lbl��λ 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "���۵�λ:"
         Height          =   180
         Left            =   3105
         TabIndex        =   10
         Top             =   945
         Width           =   2520
      End
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���:"
         Height          =   180
         Left            =   2430
         TabIndex        =   9
         Top             =   930
         Width           =   1800
      End
      Begin VB.Label lblҩƷ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ:"
         Height          =   180
         Left            =   510
         TabIndex        =   8
         Top             =   945
         Width           =   1365
      End
      Begin VB.Label lbl�ⷿ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ⷿ:"
         Height          =   180
         Left            =   510
         TabIndex        =   7
         Top             =   735
         Width           =   450
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ��ϸ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2205
         TabIndex        =   5
         Top             =   210
         Width           =   1905
      End
      Begin VB.Label lbl�ڼ� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ڼ�:   "
         Height          =   180
         Left            =   4620
         TabIndex        =   4
         Top             =   690
         Width           =   720
      End
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   1270
      BandCount       =   1
      _CBWidth        =   7785
      _CBHeight       =   720
      _Version        =   "6.7.8988"
      Child1          =   "tbrThis"
      MinHeight1      =   660
      Width1          =   5370
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgTbrStard"
         HotImageList    =   "imgTbrHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "��ӡ"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "ͼ��"
               Key             =   "ͼ��"
               Object.ToolTipText     =   "ͼ�η���"
               Object.Tag             =   "ͼ��"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   7
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgTbrHot 
      Left            =   1425
      Top             =   795
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":005E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":00BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":011A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":0178
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":01D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":0234
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTbrStard 
      Left            =   690
      Top             =   810
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":0292
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":02F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":034E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":03AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":040A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":0468
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":04C6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   5085
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7964
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
            Object.ToolTipText     =   "��ǰ���ּ�״̬"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
            Object.ToolTipText     =   "��ǰ��д��״̬"
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
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePrintView 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuEXCEL 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "��������(&J)"
      End
      Begin VB.Menu mnuViewLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolbarStand 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolbarText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFont 
         Caption         =   "����(&F)"
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "С����"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "������"
            Index           =   1
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "������"
            Index           =   2
         End
      End
      Begin VB.Menu mnuViewForeColor 
         Caption         =   "ǰ��ɫ(&C)"
      End
      Begin VB.Menu mnuViewBackColor 
         Caption         =   "����ɫ(&B)"
      End
      Begin VB.Menu mnuViewLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileReFresh 
         Caption         =   "����(&V) "
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web�ϵ�����"
         Begin VB.Menu mnuHelpZlWeb 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebSend 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelpWebL 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)��"
      End
   End
End
Attribute VB_Name = "frmDrugList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------------------------------------------------------
Public InDrugId As Long            'ҩƷid
Public inDeptId As Long            '�ⷿid
Public InDeptName  As String              '�ⷿ����
Public InDrugName  As String       'ҩƷ����
Public InDrugStAndard As String      'ҩƷ���
Public InDrugUnit As String          'ҩƷ��λ

Dim dtpStartDate As String        '��ֹ����
Dim dtpEndDate As String        '��ֹ����
Dim DataRecordSet As ADODB.Recordset
Dim RecTmpList As ADODB.Recordset
Dim blnFirst As Boolean              'ȷ���Ƿ��һ��ʹ�ñ�ϵͳ

Private mlngLevel As Integer        '��λ������1:�޼۵�λ;2:���ﵥλ��3���ⷿ��λ�� 4��סԺ��λ



Private Sub fgdData_DblClick()
    If Me.fgdData.RowData(fgdData.Row) = 0 Then Exit Sub
    If Me.fgdData.TextMatrix(fgdData.Row, 1) = "" Then Exit Sub
        
    Dim rsTemp As New ADODB.Recordset
    Dim strsql As String
    Dim int��¼״̬ As Integer
    
    On Error GoTo errHandle
    With rsTemp
        strsql = "Select id,����,NO,nvl(�۸�id,0) as �۸�id From ҩƷ�շ���¼ Where id=[1]"
        If .State = adStateOpen Then .Close
        Set rsTemp = zldatabase.OpenSQLRecord(strsql, "fgdData_DblClick", Me.fgdData.RowData(fgdData.Row))
        If .EOF Or .BOF Then Exit Sub
   '1-�⹺��⣻2-������⣻3-Эҩ��⣻4-������⣻5-��۵�����6-�ⷿ�Ƴ���7-�������ã�8-�շѴ�����9-���ʵ�������10-���ʱ�����11-�������⣻12-�̵㣻13-���۱䶯
        int��¼״̬ = Me.fgdData.TextMatrix(fgdData.Row, 13)
        Select Case !����
        Case 1
            frmPurchaseCard.ShowCard Me, !No, 4, int��¼״̬
        Case 2
            frmSelfMakeCard.ShowCard Me, !No, 4, int��¼״̬
        Case 3
            frmAccordDrugCard.ShowCard Me, !No, 4, int��¼״̬
        Case 4
            frmOtherInputCard.ShowCard Me, !No, 4, int��¼״̬
        Case 5
            frmDiffPriceAdjustCard.ShowCard Me, !No, 4, int��¼״̬
        Case 6
            frmTransferCard.ShowCard Me, !No, 4, int��¼״̬
        Case 7
            frmDrawCard.ShowCard Me, !No, 4, int��¼״̬
        Case 11
            frmOtherOutputCard.ShowCard Me, !No, 4, int��¼״̬
        Case 12
            frmCheckCard.ShowCard Me, !No, 4, int��¼״̬
        Case 13
            gstrUserName = UserInfo.�û�����
            With frmAdjust
                .lngBillId = rsTemp!�۸�id
                .lngMediId = 1
                .Show 1, Me
            End With


        Case Else
            Frm����See.byt���� = !����
            Frm����See.strNo = !No
            Frm����See.Show 1, Me
        End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If Not blnFirst Then Exit Sub
    
    Select Case frmDrugQuery.intChoose����
        Case 1
            lbl��λ.Caption = "�ۼ۵�λ��"
        Case 2
            lbl��λ.Caption = "���ﵥλ��"
        Case 3
            lbl��λ.Caption = "ҩ�ⵥλ��"
        Case 4
            lbl��λ.Caption = "סԺ��λ��"
    End Select
    
    
    Lbl���.Caption = "���" & InDrugStAndard
    lbl�ⷿ.Caption = "�ⷿ��" & InDeptName
    lblҩƷ.Caption = "ҩƷ��" & InDrugName
    lbl�ڼ�.Caption = "�ڼ�:" & dtpStartDate & "  ��  " & dtpEndDate
    ReFreshStru
    blnFirst = False
    If Not RefreshData Then Unload Me: Exit Sub
End Sub

Private Sub Form_Load()
    mlngLevel = GetSetting("ZLSOFT", "����\ҩƷ����ѯ", "���-��λ����", 1)
    blnFirst = True
    dtpStartDate = Format(DateAdd("m", -1, Currentdate()), "yyyy-MM-DD hh:mm:ss")
    dtpEndDate = Format(Currentdate(), "yyyy-MM-DD hh:mm:ss")
    RestoreWinState Me
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.MousePointer = 0
End Sub

Private Sub Form_Resize()
    Dim lngCbrHeight As Long, lngStbHeight As Long
    
    If Me.WindowState = 1 Then Exit Sub
    On Error Resume Next
    
    lngCbrHeight = IIf(Me.cbrThis.Visible, Me.cbrThis.Height, 0)
    lngStbHeight = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    Me.shpback.Left = Me.ScaleLeft + 50
    Me.shpback.Width = Me.ScaleWidth - 100
    Me.shpback.Top = Me.ScaleTop + lngCbrHeight + 50
    Me.shpback.Height = Me.ScaleHeight - (lngCbrHeight + lngStbHeight + 100)
    
    Me.LblTitle.Top = 150
    Me.LblTitle.Left = 0
    Me.LblTitle.Width = Me.shpback.Width
    
    With lbl�ⷿ
        .Top = Me.LblTitle.Top + 500
        .Left = 200
    End With
    
    With lblҩƷ
        .Top = Me.lbl�ⷿ.Top + Me.lbl�ⷿ.Height + 45
        .Left = 200
    End With
    With Me.fgdData
        .Left = 200
        .Width = Me.shpback.Width - 400
        .Top = Me.lblҩƷ.Top + Me.lblҩƷ.Height + 45
        .Height = Me.shpback.Height - Me.fgdData.Top - 400
    End With
    
    With Lbl���
        .Top = lblҩƷ.Top
        .Left = 200 + Abs((fgdData.Width - .Width)) / 2
    End With
    
    With lbl��λ
        .Top = lblҩƷ.Top
        .Left = 200 + fgdData.Width - .Width
    End With

    Me.lbl�ڼ�.Top = Me.LblTitle.Top + 500
    Me.lbl�ڼ�.Left = Me.fgdData.Width + Me.fgdData.Left - Me.lbl�ڼ�.Width
    If Me.shpback.Width < Me.LblTitle.Width Then
        Me.LblTitle.Visible = False
        Me.fgdData.Visible = False
        Lbl���.Visible = False
        lbl��λ.Visible = False
        Me.lbl�ڼ�.Visible = False
    Else
        Me.LblTitle.Visible = True
        Lbl���.Visible = True
        lbl��λ.Visible = True
        Me.fgdData.Visible = True
        Me.lbl�ڼ�.Visible = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Unload frmDrugListAsk
    SaveWinState Me
End Sub

Private Sub mnuEXCEL_Click()
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    objPrint.Title.Text = Me.LblTitle.Caption
     
     Set objRow = New zlTabAppRow
     objRow.Add Me.lbl�ⷿ.Caption
     objRow.Add Me.lbl�ڼ�.Caption
     objPrint.UnderAppRows.Add objRow
     
     Set objPrint.Body = fgdData
     
      Set objRow = New zlTabAppRow
     With objRow
        .Add "��ӡ��:" & UserInfo.�û�����
        .Add "��ӡʱ��:" & Format(Currentdate, "yyyy��MM��DD��")
     End With
     
     objPrint.BelowAppRows.Add objRow
    
     objPrint.PageFooter = 2
     
     zlPrintOrView1Grd objPrint, 3
     Set objPrint = Nothing

End Sub

Private Sub mnufileexit_Click()
    Unload Me
End Sub

Private Sub mnuFileOpen_Click()
    RefreshData
End Sub

Private Sub mnuFilePrint_Click()
    grdPrint True
End Sub

Private Sub mnuFilePrintSet_Click()
     zlPrintSet
End Sub

Private Sub mnuFilePrintView_Click()
  grdPrint False
End Sub
Private Sub grdPrint(blnIsPreview As Boolean)
    '---------------------------------------------------
    '���ܣ�    ������Ļ��֯���ϸ�����Ŀ����ӡԤ��
    '������blnIsPreview false��ʾԤ��
    '���أ�
    '---------------------------------------------------
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    objPrint.Title.Text = Me.LblTitle.Caption
     
     Set objRow = New zlTabAppRow
     objRow.Add Me.lblҩƷ.Caption
     objPrint.UnderAppRows.Add objRow
     
     Set objRow = New zlTabAppRow
     objRow.Add Me.Lbl���.Caption
     objPrint.UnderAppRows.Add objRow
     
     Set objRow = New zlTabAppRow
     objRow.Add Me.lbl��λ.Caption
     objRow.Add Me.lbl�ڼ�.Caption
     objRow.Add Me.lbl�ⷿ.Caption
     objPrint.UnderAppRows.Add objRow
     
     Set objRow = New zlTabAppRow
     objRow.Add "��ӡ��:" & UserInfo.�û�����
     objRow.Add "��ӡʱ��:" & Format(Currentdate, "yyyy��MM��DD�� HH:MM")
     objPrint.BelowAppRows.Add objRow
     Set objPrint.Body = fgdData
     
     objPrint.PageFooter = 2
     
    If Not blnIsPreview Then
             zlPrintOrView1Grd objPrint, 2
        Else
        Select Case zlPrintAsk(objPrint)
            Case 1
                 zlPrintOrView1Grd objPrint, 1
            Case 2
                zlPrintOrView1Grd objPrint, 2
            Case 3
                zlPrintOrView1Grd objPrint, 3
        End Select
    End If
    Set objPrint = Nothing
End Sub

Private Sub mnuFileReFresh_Click()
    fgdData_DblClick
End Sub

Private Sub mnuHelpAbout_Click()
   ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpHelp_Click()
         Shell "hh.exe " & App.Path & "\zlMediBill.chm::/ҩ��������/ҩƷ����ѯ.htm", vbNormalFocus
End Sub
Private Sub mnuHelpWebSend_Click()
    zlMailTo Me.hWnd
End Sub

Private Sub mnuHelpZlWeb_Click()
    zlHomePage Me.hWnd
End Sub



Private Sub mnuViewFontSize_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 2
        Me.mnuViewFontSize(i).Checked = False
    Next
    Me.mnuViewFontSize(Index).Checked = True
    
    Select Case Index
    Case 0
        Me.LblTitle.FontSize = 22
        Me.lbl�ⷿ.FontSize = 9
        Me.lbl�ڼ�.FontSize = 9
        Me.lbl��λ.FontSize = 9
        Me.Lbl���.FontSize = 9
        Me.lblҩƷ.FontSize = 9
        Me.fgdData.Font.Size = 9
        Me.fgdData.FontFixed.Size = 9

     Case 1
        Me.LblTitle.FontSize = 24
        Me.lbl�ⷿ.FontSize = 11
        Me.lbl�ڼ�.FontSize = 11
        Me.lbl��λ.FontSize = 11
        Me.Lbl���.FontSize = 11
        Me.lblҩƷ.FontSize = 11
        Me.fgdData.Font.Size = 11
        Me.fgdData.FontFixed.Size = 11

    Case 2
        Me.LblTitle.FontSize = 28
        Me.lbl�ⷿ.FontSize = 15
        Me.lbl�ڼ�.FontSize = 15
        Me.lbl��λ.FontSize = 15
        Me.Lbl���.FontSize = 15
        Me.lblҩƷ.FontSize = 15
        Me.fgdData.Font.Size = 15
        Me.fgdData.FontFixed.Size = 15
    End Select
    Form_Resize
    Me.Refresh
End Sub

Private Sub mnuViewForeColor_Click()
    Dim lngForeColor As Long
    lngForeColor = zlGetColor(Me.LblTitle.ForeColor)
    Me.LblTitle.ForeColor = lngForeColor
    Me.lbl�ⷿ.ForeColor = lngForeColor
    Me.lblҩƷ.ForeColor = lngForeColor
    Me.Lbl���.ForeColor = lngForeColor
    Me.lbl��λ.ForeColor = lngForeColor
    Me.lbl�ڼ�.ForeColor = lngForeColor
    Me.fgdData.ForeColor = lngForeColor
    Me.fgdData.ForeColorFixed = lngForeColor
End Sub
Private Sub mnuViewBackColor_Click()
    Dim lngBackColor As Long
    lngBackColor = zlGetColor(Me.fgdData.BackColor)
    Me.shpback.BackColor = lngBackColor
    Me.fgdData.BackColor = lngBackColor
    Me.fgdData.BackColorBkg = lngBackColor
    Me.fgdData.BackColorFixed = lngBackColor
End Sub


Private Sub mnuViewStatus_Click()
    Me.mnuViewStatus.Checked = Not Me.mnuViewStatus.Checked
    Me.stbThis.Visible = Me.mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolbarStAnd_Click()
    Dim intCount As Integer
    Me.mnuViewToolbarStand.Checked = Not Me.mnuViewToolbarStand.Checked
    Me.mnuViewToolbarText.Enabled = Me.mnuViewToolbarStand.Checked
    Me.cbrThis.Visible = Me.mnuViewToolbarStand.Checked
    If Me.mnuViewToolbarText.Checked Then
        For intCount = 1 To Me.tbrThis.Buttons.count
            Me.tbrThis.Buttons(intCount).Caption = Me.tbrThis.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To Me.tbrThis.Buttons.count
            Me.tbrThis.Buttons(intCount).Caption = ""
        Next
    End If
    Me.cbrThis.Bands(1).MinHeight = Me.tbrThis.Height
    Me.cbrThis.Refresh
    Form_Resize

End Sub
Private Sub mnuViewToolbarText_Click()
    Dim intCount As Integer
    Me.mnuViewToolbarText.Checked = Not Me.mnuViewToolbarText.Checked
    If Me.mnuViewToolbarText.Checked Then
        For intCount = 1 To Me.tbrThis.Buttons.count
            Me.tbrThis.Buttons(intCount).Caption = Me.tbrThis.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To Me.tbrThis.Buttons.count
            Me.tbrThis.Buttons(intCount).Caption = ""
        Next
    End If
    Me.cbrThis.Bands(1).MinHeight = Me.tbrThis.Height
    Me.cbrThis.Refresh
    Form_Resize

End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    With Button
        Select Case .Key
        Case "Ԥ��"
            mnuFilePrintView_Click
        Case "��ӡ"
            grdPrint True
        Case "����"
            mnuFileOpen_Click
'        Case "ͼ��"
'            mnuViewchart_Click
        Case "����"
             PopupMenu mnuViewFont
        Case "ǰ��ɫ"
            mnuViewForeColor_Click
        Case "����ɫ" '
            mnuViewBackColor_Click
        Case "����"
            mnuHelpHelp_Click
        Case "�˳�"
           mnufileexit_Click
        End Select
    End With
End Sub

Private Function RefreshData() As Boolean
    '-------------------------------------------------------------------------
    '--���ܣ�ˢ������
    '--����:                                                                --
    '--����:                                                                --
    '-------------------------------------------------------------------------
    Dim strsql As String
    Dim lngRow As Long
    Dim dblCurrNum As Double      '��ǰ�������
    Dim dblCurrMny As Double      '��ǰ�����
    Dim dblCurrDf As Double      '��ǰ�����
    Dim dblStartNum As Double   '��ʼʱ����
    Dim dblStartMny As Double   '��ʼʱ���
    Dim dblStartDf As Double   '��ʼʱ���
    Dim dblinNum As Double     '�������
    Dim dblInMny As Double     '�����
    Dim dblinDf As Double      '�����
    Dim dblOutNum As Double     '��������
    Dim dblOutMny As Double     '������
    Dim dblOutDf As Double      '������
    Dim intLevel As Integer     '��λ����
        
    On Error GoTo errHandle
    dblCurrNum = 0: dblCurrMny = 0: dblCurrDf = 0
    dblinNum = 0: dblInMny = 0: dblinDf = 0
    dblOutNum = 0: dblOutMny = 0: dblOutDf = 0
    Load frmDrugListAsk
    With frmDrugListAsk
        .dtpStartDate.Value = CDate(dtpStartDate)
        .dtpEndDate.Value = CDate(dtpEndDate)
        .dtpEndDate.MaxDate = Currentdate()
        .dtpStartDate.MaxDate = .dtpEndDate.MaxDate
        .inDeptId = inDeptId
        
        .InDrugId = InDrugId
        .InDrugName = InDrugName
        .InDrugStAndard = InDrugStAndard
        .InDrugUnit = InDrugUnit
        .Show 1, Me
        RefreshData = False
    End With
    If frmDrugListAsk.blnAskOk = False Then
        Exit Function
    End If
    
    With frmDrugListAsk
        dtpStartDate = Format(.dtpStartDate.Value, "yyyy-MM-DD hh:mm:ss")
        dtpEndDate = Format(.dtpEndDate.Value, "yyyy-MM-DD hh:mm:ss")
        InDrugId = .InDrugId
        inDeptId = .cob�ⷿ.ItemData(.cob�ⷿ.ListIndex)
        InDeptName = .cob�ⷿ.Text
        InDrugName = .InDrugName
        InDrugStAndard = .InDrugStAndard
        InDrugUnit = .InDrugUnit
        intLevel = frmDrugQuery.intChoose����
        
        
    End With
    
    '��ȡ��ǰ����
    ShowFlash "����װ�����ݣ����Ժ�", Me
    DoEvents
    On Error GoTo Err:
    
    Set RecTmpList = New ADODB.Recordset
    strsql = " Select Sum(ʵ������)" & IIf(Val(Me.Tag) = 0, "/1", "/" & Me.Tag) & " as ��ǰ����,Sum(ʵ�ʽ��) as ��ǰ���,Sum(ʵ�ʲ��) as ��ǰ���" & _
             " From ҩƷ��� " & _
             " Where  ����=1 And ҩƷid=" & InDrugId & IIf(inDeptId = 0, "", "  And �ⷿid=" & inDeptId)
    With RecTmpList
    Set RecTmpList = zldatabase.OpenSQLRecord(strsql, "RefreshData")

    If Not .EOF Then
        dblCurrNum = IIf(IsNull(.Fields("��ǰ����").Value), 0, .Fields("��ǰ����").Value)
        dblCurrMny = IIf(IsNull(.Fields("��ǰ���").Value), 0, .Fields("��ǰ���").Value)
        dblCurrDf = IIf(IsNull(.Fields("��ǰ���").Value), 0, .Fields("��ǰ���").Value)
    End If
     .Close
     
          
    '��ȡ��ʼ����
     strsql = " Select sum(�������)" & IIf(Val(Me.Tag) = 0, "/1", "/" & Me.Tag) & " as �������,sum(�����) as �����,sum(�����) as �����, " & _
            "        sum(��������)" & IIf(Val(Me.Tag) = 0, "/1", "/" & Me.Tag) & " as ��������,sum(������) as ������,sum(������) as ������ " & _
            "  From ( " & _
            "        Select 'Aid' as RiD,id, " & _
            "          Decode(���ϵ��,1,1,0)*ʵ������*���� as �������, " & _
            "          Decode(���ϵ��,1,1,0)*���۽�� as �����, " & _
            "          Decode(���ϵ��,1,1,0)*��� as �����, " & _
            "          Decode(���ϵ��,-1,1,0)*ʵ������*���� as ��������, " & _
            "          Decode(���ϵ��,-1,1,0)*���۽�� as  ������, " & _
            "          Decode(���ϵ��,-1,1,0)*��� as  ������ " & _
            "      From ҩƷ�շ���¼ " & _
            "      Where ����� Is Not Null " & IIf(inDeptId = 0, " ", " And �ⷿid=" & inDeptId) & _
            "          And ҩƷid=" & InDrugId & "And ������� >= " & " To_date('" & Format(dtpStartDate, "yyyy-MM-DD HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'))"
            
      Set RecTmpList = zldatabase.OpenSQLRecord(strsql, "RefreshData")
       
      
            dblinNum = IIf(IsNull(.Fields("�������").Value), 0, .Fields("�������").Value)
            dblInMny = IIf(IsNull(.Fields("�����").Value), 0, .Fields("�����").Value)
            dblinDf = IIf(IsNull(.Fields("�����").Value), 0, .Fields("�����").Value)
            dblOutNum = IIf(IsNull(.Fields("��������").Value), 0, .Fields("��������").Value)
            dblOutMny = IIf(IsNull(.Fields("������").Value), 0, .Fields("������").Value)
            dblOutDf = IIf(IsNull(.Fields("������").Value), 0, .Fields("������").Value)

            dblStartNum = dblCurrNum - dblinNum + dblOutNum
            dblStartMny = dblCurrMny - dblInMny + dblOutMny
            dblStartDf = dblCurrDf - dblinDf + dblOutDf

    End With
    
    '��ȡ��ϸ��¼
    '1-�⹺��⣻2-������⣻3-Эҩ��⣻4-������⣻5-��۵�����6-�ⷿ�Ƴ���7-�������ã�8-�շѴ�����9-���ʵ�������10-���ʱ�����11-�������⣻12-�̵㣻13-���۱䶯
    
    strsql = "Select max(a.id) as id, Decode(A.����,1,'�⹺',2,'����',3,'Э��',4,'���',5,'���',6,'�ƿ�',7,'����',8,'����',9,'����',10,'��ҩ',11,'����',12,'�̵�',13,'����')||A.No as no,A.����,A.�������,decode(a.��¼״̬,2,'��������',rtrim(A.ժҪ)|| Decode(B.��Ʊ��,null,' ',' ��Ʊ��:')||��Ʊ��) as ժҪ,A.����,  " & _
            "       sum(Decode(���ϵ��,1,1,0)*A.ʵ������*A.����" & IIf(Val(Me.Tag) = 0, "/1", "/" & Me.Tag) & ") as �������,  " & _
            "       sum(Decode(���ϵ��,1,1,0)*A.���۽��) as �����,  " & _
            "       sum(Decode(���ϵ��,1,1,0)*A.���) as �����,  " & _
            "       sum(Decode(���ϵ��,1,0,1)*A.ʵ������*A.����" & IIf(Val(Me.Tag) = 0, "/1", "/" & Me.Tag) & ") as ��������,  " & _
            "       sum(Decode(���ϵ��,1,0,1)*A.���۽��) as  ������,  " & _
            "       sum(Decode(���ϵ��,1,0,1)*A.���) as  ������, A.��¼״̬  " & _
            " From    ҩƷ�շ���¼ A,ҩƷӦ����¼ B  " & _
            " Where A.����� Is Not Null  And A.id=B.�շ�id(+) " & _
            "      And A.ҩƷid= " & InDrugId & IIf(inDeptId = 0, "", " And A.�ⷿid=" & inDeptId) & _
            "      And  A.������� between To_date('" & Format(dtpStartDate, "yyyy-MM-DD HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss') And To_date('" & Format(dtpEndDate, "yyyy-MM-DD HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')  " & _
            "    GROUP BY a.no, a.����, a.�������, a.��¼״̬, a.ժҪ, b.��Ʊ��, a.����,a.���� " & _
            " order by A.������� "
         'And Mod(A.��¼״̬,3)=1
    Set DataRecordSet = New ADODB.Recordset
    With DataRecordSet
        If .State = 1 Then .Close
        Set RecTmpList = zldatabase.OpenSQLRecord(strsql, "RefreshData")
'        If .RecordCount <> 0 Then
            ReFreshStru
'        End If
        
    
        
        lngRow = 2

'        If .RecordCount = 0 Then
'            StopFlash
'            MsgBox "ҩƷ�ڴ��ڼ����κ���ϸ!", vbInformation, gstrSysName
'            Exit Function
'        End If
        
        ReFreshStru
        Me.fgdData.rows = Me.fgdData.rows + 1
        Me.fgdData.TextMatrix(lngRow, 0) = Format(dtpStartDate, "yyyy��MM��DD��")
        Me.fgdData.TextMatrix(lngRow, 1) = ""
        Me.fgdData.TextMatrix(lngRow, 2) = "�ڳ����"
        Me.fgdData.TextMatrix(lngRow, 3) = ""
        Me.fgdData.TextMatrix(lngRow, 4) = ""
        Me.fgdData.TextMatrix(lngRow, 5) = ""
        Me.fgdData.TextMatrix(lngRow, 6) = ""
        Me.fgdData.TextMatrix(lngRow, 7) = ""
        Me.fgdData.TextMatrix(lngRow, 8) = ""
        Me.fgdData.TextMatrix(lngRow, 9) = ""
        Me.fgdData.TextMatrix(lngRow, 10) = Format(dblStartNum, "###0.000;-###0.000; ; ")
        Me.fgdData.TextMatrix(lngRow, 11) = Format(dblStartMny, "###0.00;-###0.00; ; ")
        Me.fgdData.TextMatrix(lngRow, 12) = Format(dblStartDf, "###0.00;-###0.00; ; ")
        Call RefreshGridColWidth(Me.fgdData, lngRow)
        Me.fgdData.RowData(lngRow) = "0"
        lngRow = lngRow + 1
       
        Select Case intLevel
            Case 1
                lbl��λ.Caption = "�ۼ۵�λ��" & InDrugUnit
            Case 2
                lbl��λ.Caption = "���ﵥλ��" & InDrugUnit
            Case 3
                lbl��λ.Caption = "ҩ�ⵥλ��" & InDrugUnit
            Case 4
                lbl��λ.Caption = "סԺ��λ��" & InDrugUnit
        End Select

        Lbl���.Caption = "���" & InDrugStAndard
        lbl�ⷿ.Caption = "�ⷿ��" & InDeptName
        lblҩƷ.Caption = "ҩƷ��" & InDrugName
        lbl�ڼ�.Caption = "�ڼ�:" & dtpStartDate & "  ��  " & dtpEndDate
'        lbl�ڼ�.Caption = "�ڼ�:" & dtpStartDate & "  ��  " & dtpEndDate
       
         If .RecordCount <> 0 Then
                Me.fgdData.rows = Me.fgdData.rows + .RecordCount
         End If
         
            dblinNum = 0
            dblInMny = 0
            dblinDf = 0
            dblOutNum = 0
            dblOutMny = 0
            dblOutDf = 0
         
         Do While Not .EOF
            dblStartNum = dblStartNum + IIf(IsNull(.Fields("�������").Value), 0, .Fields("�������").Value) - IIf(IsNull(.Fields("��������").Value), 0, .Fields("��������").Value)
            dblStartMny = dblStartMny + IIf(IsNull(.Fields("�����").Value), 0, .Fields("�����").Value) - IIf(IsNull(.Fields("������").Value), 0, .Fields("������").Value)
            dblStartDf = dblStartDf + IIf(IsNull(.Fields("�����").Value), 0, .Fields("�����").Value) - IIf(IsNull(.Fields("������").Value), 0, .Fields("������").Value)
            
            dblinNum = dblinNum + IIf(IsNull(.Fields("�������").Value), 0, .Fields("�������").Value)
            dblInMny = dblInMny + IIf(IsNull(.Fields("�����").Value), 0, .Fields("�����").Value)
            dblinDf = dblinDf + IIf(IsNull(.Fields("�����").Value), 0, .Fields("�����").Value)
            dblOutNum = dblOutNum + IIf(IsNull(.Fields("��������").Value), 0, .Fields("��������").Value)
            dblOutMny = dblOutMny + IIf(IsNull(.Fields("������").Value), 0, .Fields("������").Value)
            dblOutDf = dblOutDf + IIf(IsNull(.Fields("������").Value), 0, .Fields("������").Value)
            
            Me.fgdData.TextMatrix(lngRow, 0) = Format(.Fields("�������").Value, "yyyy��MM��DD��") & IIf(lngRow Mod 2 = 0, "", " ")
            Me.fgdData.TextMatrix(lngRow, 1) = IIf(IsNull(.Fields("no").Value), "", .Fields("no").Value) & IIf(lngRow Mod 2 = 0, "", " ")
            Me.fgdData.TextMatrix(lngRow, 2) = IIf(IsNull(.Fields("ժҪ").Value), "", .Fields("ժҪ").Value) & IIf(lngRow Mod 2 = 0, "", " ")
            Me.fgdData.TextMatrix(lngRow, 3) = IIf(IsNull(.Fields("����").Value), "", .Fields("����").Value) & IIf(lngRow Mod 2 = 0, "", " ")
            Me.fgdData.TextMatrix(lngRow, 4) = Format(.Fields("�������").Value, "###0.000;-###0.000; ; ")
            Me.fgdData.TextMatrix(lngRow, 5) = Format(.Fields("�����").Value, "###0.00;-###0.00; ; ")
            Me.fgdData.TextMatrix(lngRow, 6) = Format(.Fields("�����").Value, "###0.00;-###0.00; ; ")
            Me.fgdData.TextMatrix(lngRow, 7) = Format(.Fields("��������").Value, "###0.000;-###0.000; ; ")
            Me.fgdData.TextMatrix(lngRow, 8) = Format(.Fields("������").Value, "###0.00;-###0.00; ; ")
            Me.fgdData.TextMatrix(lngRow, 9) = Format(.Fields("������").Value, "###0.00;-###0.00; ; ")
            Me.fgdData.TextMatrix(lngRow, 10) = Format(dblStartNum, "###0.000;-###0.000; ; ")
            Me.fgdData.TextMatrix(lngRow, 11) = Format(dblStartMny, "###0.00;-###0.00; ; ")
            Me.fgdData.TextMatrix(lngRow, 12) = Format(dblStartDf, "###0.00;-###0.00; ; ")
            Me.fgdData.TextMatrix(lngRow, 13) = .Fields("��¼״̬")
            Call RefreshGridColWidth(Me.fgdData, lngRow)
            Me.fgdData.RowData(lngRow) = .Fields("ID").Value
            lngRow = lngRow + 1
            .MoveNext
        Loop
    End With
'    dblCurrNum = dblStartNum
'    dblCurrMny = dblStartMny
'    dblCurrDf = dblStartDf
    
    If dblCurrNum <> 0 Or dblCurrMny <> 0 Or dblCurrDf <> 0 Or _
        dblinNum <> 0 Or dblInMny <> 0 Or dblinDf <> 0 Or _
        dblOutNum <> 0 Or dblOutMny <> 0 Or dblOutDf <> 0 Then
        Me.fgdData.TextMatrix(lngRow, 0) = Format(dtpEndDate, "yyyy��MM��DD��") & Space(lngRow Mod 2)
        Me.fgdData.TextMatrix(lngRow, 1) = ""
        Me.fgdData.TextMatrix(lngRow, 2) = "��ĩ���"
        Me.fgdData.TextMatrix(lngRow, 3) = ""
        Me.fgdData.TextMatrix(lngRow, 4) = Format(dblinNum, "###0.000;-###0.000; ; ")
        Me.fgdData.TextMatrix(lngRow, 5) = Format(dblInMny, "###0.00;-###0.00; ; ")
        Me.fgdData.TextMatrix(lngRow, 6) = Format(dblinDf, "###0.00;-###0.00; ; ")
        Me.fgdData.TextMatrix(lngRow, 7) = Format(dblOutNum, "###0.000;-###0.000; ; ")
        Me.fgdData.TextMatrix(lngRow, 8) = Format(dblOutMny, "###0.00;-###0.00; ; ")
        Me.fgdData.TextMatrix(lngRow, 9) = Format(dblOutDf, "###0.00;-###0.00; ; ")
        Me.fgdData.TextMatrix(lngRow, 10) = Format(dblStartNum, "###0.000;-###0.000; ; ")
        Me.fgdData.TextMatrix(lngRow, 11) = Format(dblStartMny, "###0.00;-###0.00; ; ")
        Me.fgdData.TextMatrix(lngRow, 12) = Format(dblStartDf, "###0.00;-###0.00; ; ")
        Call RefreshGridColWidth(Me.fgdData, lngRow)
        Me.fgdData.RowData(lngRow) = "0"
        lngRow = lngRow + 1
    End If
    Me.fgdData.ColWidth(13) = 0
    RefreshData = True
    StopFlash
Exit Function
Err:
   StopFlash
   RefreshData = False
   MsgBox "�ڻ�ȡ��ϸ������ʱ,�����˲���Ԥ֪�Ĵ���!", vbInformation, gstrSysName
   Unload Me
   Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ReFreshStru()
    '-------------------------------------------------------------------------
    '--����:���»�ı�ͷ�ṹ
    '--����:
    '--����:
    '-------------------------------------------------------------------------
    Dim IntCol As Long
    Me.Caption = "ҩƷ��ϸ��"
    Me.LblTitle.Caption = GetUnitName & "ҩƷ��ϸ��"
    With Me.fgdData
            .Redraw = False
            For IntCol = 0 To .rows - 1
                .MergeRow(IntCol) = False
            Next
             .Clear
             .Cols = 14
             .rows = 3
             .FixedRows = 2
             .MergeCells = flexMergeRestrictRows
             .SelectionMode = flexSelectionByRow
            For IntCol = 0 To .Cols - 1
                .FixedAlignment(IntCol) = 4
                If IntCol = 0 Then
                    .ColWidth(IntCol) = 1350
                ElseIf IntCol = 1 Then
                    .ColWidth(IntCol) = 800
                ElseIf IntCol = 2 Then
                    .ColWidth(IntCol) = 1200
                ElseIf IntCol = 3 Then
                    .ColWidth(IntCol) = 800
                Else
                    .ColWidth(IntCol) = 800
                End If
                If IntCol <= 3 Then
                    .ColAlignment(IntCol) = 1
                    .MergeCol(IntCol) = True
                Else
                    .MergeCol(IntCol) = False
                    .ColAlignment(IntCol) = 7
                End If
            Next
            .ColWidth(13) = 0
            .MergeCells = 1
            .TextMatrix(0, 0) = "����"
            .TextMatrix(1, 0) = "����"
            .TextMatrix(0, 1) = "���ݺ�"
            .TextMatrix(1, 1) = "���ݺ�"
            .TextMatrix(0, 2) = "ժҪ"
            .TextMatrix(1, 2) = "ժҪ"
            .TextMatrix(0, 3) = "����"
            .TextMatrix(1, 3) = "����"
            .TextMatrix(0, 4) = "���"
            .TextMatrix(0, 5) = "���"
            .TextMatrix(0, 6) = "���"
            .TextMatrix(1, 4) = "����"
            .TextMatrix(1, 5) = "���"
            .TextMatrix(1, 6) = "���"
            .MergeRow(0) = True
            .MergeRow(1) = True
            .TextMatrix(0, 7) = "����"
            .TextMatrix(0, 8) = "����"
            .TextMatrix(0, 9) = "����"
            .TextMatrix(1, 7) = "����"
            .TextMatrix(1, 8) = "���"
            .TextMatrix(1, 9) = "���"
        
            .TextMatrix(0, 10) = "���"
            .TextMatrix(0, 11) = "���"
            .TextMatrix(0, 12) = "���"
            .TextMatrix(1, 10) = "����"
            .TextMatrix(1, 11) = "���"
            .TextMatrix(1, 12) = "���"
'            Call RefreshGridColWidth(Me.fgdData, 0)
            .Redraw = True
    End With

End Sub

Private Function GetLevel(ByVal lng����id As Long) As Integer
    '�жϸò���ֻ��ҩ�������ҩ��
    Dim rsTemp As New ADODB.Recordset
    Dim intChoose���� As Integer
    Dim strsql As String
    
    On Error GoTo errHandle
    strsql = "Select * From ��������˵�� " & _
        " Where ����id=[1] And �������� IN ('��ҩ��','��ҩ��','��ҩ��','�Ƽ���','��ҩ��','��ҩ��','��ҩ��') "
    
    Set rsTemp = zldatabase.OpenSQLRecord(strsql, "GetLevel", lng����id)
    If Not rsTemp.EOF Then
        Select Case rsTemp!�������
            Case 0
                intChoose���� = 3
            Case 1, 3
                intChoose���� = 2
            Case 2
                intChoose���� = 4
            Case Else
                intChoose���� = 1
        End Select
    Else
        intChoose���� = 1
    End If
   
    rsTemp.Close
    
    GetLevel = intChoose����
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

