VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm�󴦷���� 
   BackColor       =   &H8000000A&
   Caption         =   "�󴦷����"
   ClientHeight    =   6690
   ClientLeft      =   165
   ClientTop       =   3750
   ClientWidth     =   11760
   Icon            =   "frm�󴦷����.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSum 
      Height          =   2325
      Left            =   540
      TabIndex        =   5
      Top             =   1200
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   4101
      _Version        =   393216
      FixedCols       =   0
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      GridColor       =   8421504
      GridColorFixed  =   8421504
      GridColorUnpopulated=   8421504
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox picV 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1995
      Left            =   3180
      MousePointer    =   9  'Size W E
      ScaleHeight     =   1995
      ScaleWidth      =   45
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2550
      Width           =   45
   End
   Begin MSComctlLib.ImageList ilsColor 
      Left            =   4800
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�󴦷����.frx":030A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�󴦷����.frx":0526
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�󴦷����.frx":0742
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�󴦷����.frx":095C
            Key             =   "Font"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�󴦷����.frx":0B78
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�󴦷����.frx":0D94
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMono 
      Left            =   4080
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�󴦷����.frx":0FB0
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�󴦷����.frx":11CC
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�󴦷����.frx":13E8
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�󴦷����.frx":1602
            Key             =   "Font"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�󴦷����.frx":181E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�󴦷����.frx":1A3A
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   11760
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   5370
      Key1            =   "one"
      NewRow1         =   0   'False
      Caption2        =   "����"
      Child2          =   "cmbDept"
      MinHeight2      =   300
      Width2          =   765
      Key2            =   "two"
      NewRow2         =   0   'False
      Begin VB.ComboBox cmbDept 
         Height          =   300
         Left            =   5985
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   5685
      End
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMono"
         HotImageList    =   "ilsColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Open"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageKey        =   "Open"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6330
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   635
      SimpleText      =   $"frm�󴦷����.frx":1C56
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm�󴦷����.frx":1C9D
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15663
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   1215
      Left            =   6510
      TabIndex        =   6
      Top             =   2490
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   2143
      _Version        =   393216
      FixedCols       =   0
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      GridColor       =   8421504
      GridColorFixed  =   8421504
      GridColorUnpopulated=   8421504
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblDetail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "������ϸ"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6120
      TabIndex        =   7
      Top             =   1380
      Width           =   3015
   End
   Begin VB.Label lblSum 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�����б�"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   930
      Width           =   4440
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "����(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuReportItem 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolSplit 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOpen 
         Caption         =   "��������(&J)"
      End
      Begin VB.Menu mnuViewSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R) "
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
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
Attribute VB_Name = "frm�󴦷����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnLoad As Boolean
Dim mdatBegin As Date, mdatEnd As Date            '��ѯ��ʱ�䷶Χ

'����ģ�����
Dim mdblMax As Double                               '����׼ֵ
Dim mintUnit As Integer                             'ҩƷ��λ��0���ۼ۵�λ��1��ҩ����λ

Dim mstrType As String                            '��������(��ѡ)
Dim msngStartX As Single, msngStartY As Single    '�ƶ�ǰ����λ��
Dim mlngID As Long          'ǰһ�����ŵ�ID
Dim mstrNo As String        'ǰһ�Ŵ�����NO
Dim mstr���� As String      'ǰһ�Ŵ���������
Dim mlng����ID As Long      'ǰһ�Ŵ����Ĳ���ID
Dim mlngRow As Long         'ǰһ������ʱ��������
Private mlngMode As Long
Private mstrPrivs As String             '��ǰ�û����еĵ�ǰģ��Ĺ���


Private Sub cbrThis_HeightChanged(ByVal NewHeight As Single)
    Call Form_Resize
End Sub

Private Sub cmbDept_Click()
    If mblnLoad = False Then
        Call FillSum
    End If
End Sub

Private Sub Form_Activate()
    If mblnLoad = True Then
        FillDept
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    mblnLoad = True
    mlngMode = glngModul
    mstrPrivs = gstrprivs
    
    RestoreWinState Me, App.ProductName
    Call zldatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    '�õ���ѯ��ʱ�䷶Χ
    mdatEnd = CDate(Format(Sys.Currentdate, "yyyy-MM-dd"))
    mdatBegin = DateAdd("d", -10, mdatEnd) + 1
    
    mdblMax = Val(zldatabase.GetPara("����׼", glngSys, 1347))
    mintUnit = Val(zldatabase.GetPara("ҩƷ��λ", glngSys, 1347))

    Call InitSum
End Sub

Private Sub InitSum()
'��ʼ�����ܱ����ʽ
    With mshSum
        ClearGrid mshSum, 8
        .TextMatrix(0, 0) = "ҽ��"
        .TextMatrix(0, 1) = "����"
        .TextMatrix(0, 2) = "����"
        .TextMatrix(0, 3) = "������"
        .TextMatrix(0, 4) = "���"
        .TextMatrix(0, 5) = "��������"
        .TextMatrix(0, 6) = "����/סԺ��"
        .TextMatrix(0, 7) = "����ID"
        
        .ColWidth(0) = 690
        .ColWidth(1) = 1020
        .ColWidth(2) = 540
        .ColWidth(3) = 840
        .ColWidth(4) = 945
        .ColWidth(5) = 900
        .ColWidth(6) = 1500
        .ColWidth(7) = 0
        
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 7
        .ColAlignment(5) = 1
        .ColAlignment(6) = 1
        
        .MergeCells = flexMergeRestrictRows
        .MergeCol(0) = True
        .MergeCol(1) = True
    End With
    
    With mshDetail
        ClearGrid mshDetail, 7
        .TextMatrix(0, 0) = "ҩƷ����"
        .TextMatrix(0, 1) = "���"
        .TextMatrix(0, 2) = "���� "
        .TextMatrix(0, 3) = "��λ"
        .TextMatrix(0, 4) = "����"
        .TextMatrix(0, 5) = "����"
        .TextMatrix(0, 6) = "���"
        
        .ColWidth(0) = 2500
        .ColWidth(1) = 1200
        .ColWidth(2) = 1300
        .ColWidth(3) = 450
        .ColWidth(4) = 825
        .ColWidth(5) = 840
        .ColWidth(6) = 1140
        
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 4
        .ColAlignment(4) = 7
        .ColAlignment(5) = 7
        .ColAlignment(6) = 7
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlngID = -1
    
    zldatabase.SetPara "ҩƷ��λ", mintUnit, glngSys, 1347

    SaveWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    sngTop = IIf(cbrThis.Visible, cbrThis.Top + cbrThis.Height, 0)
    sngBottom = ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    lblSum.Top = sngTop
    lblSum.Left = 0
    mshSum.Left = lblSum.Left
    mshSum.Width = lblSum.Width
    mshSum.Top = lblSum.Top + lblSum.Height
    If sngBottom - mshSum.Top > 0 Then mshSum.Height = sngBottom - mshSum.Top
    
    picV.Top = lblSum.Top
    picV.Left = lblSum.Left + lblSum.Width
    picV.Height = sngBottom - picV.Top
    
    
    lblDetail.Left = picV.Left + picV.Width
    mshDetail.Left = lblDetail.Left
    lblDetail.Top = lblSum.Top
    mshDetail.Top = mshSum.Top
    
    lblDetail.Width = IIf(ScaleWidth - lblDetail.Left > 0, ScaleWidth - lblDetail.Left, 0)
    mshDetail.Width = lblDetail.Width
    mshDetail.Height = mshSum.Height
    
    Refresh
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    'Ĭ�ϲ�������ʼʱ��=�Ǽ�ʱ�俪ʼ������ʱ��=�Ǽ�ʱ�������NO=����NO,���˿���=���˿���ID,����ID=����ID
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "���˿���=" & IIf(Val(cmbDept.ItemData(cmbDept.ListIndex)) = 0, "", Val(cmbDept.ItemData(cmbDept.ListIndex))), _
        "��ʼʱ��=" & Format(mdatBegin, "yyyy-mm-dd"), _
        "����ʱ��=" & Format(mdatEnd, "yyyy-mm-dd"), _
        "NO=" & mstrNo, _
        "����ID=" & IIf(mlng����ID = 0, "", mlng����ID))
End Sub
Private Sub mnuViewOpen_Click()
    If frm�������.GetCondition(mdatBegin, mdatEnd, mintUnit, mstrType, mstrPrivs, Me) = True Then
        mlngID = -1
        Call FillSum
    End If
End Sub

Private Sub mnuViewRefresh_Click()
    FillDept
End Sub

Private Sub mshSum_EnterCell()
    SetColor False, mlngRow
    mlngRow = mshSum.Row
    SetColor True, mlngRow
    
    If mstrNo = mshSum.TextMatrix(mshSum.Row, 3) And mstr���� = mshSum.TextMatrix(mshSum.Row, 2) Then Exit Sub
    
    mstrNo = mshSum.TextMatrix(mshSum.Row, 3)
    mstr���� = mshSum.TextMatrix(mshSum.Row, 2)
    mlng����ID = Val(mshSum.TextMatrix(mshSum.Row, 7))
    Call FillDetail
End Sub

Private Sub SetColor(ByVal blnChange As Boolean, ByVal lngRow As Long)
'����:blnChange  Ϊ��ı������ɫ��Ϊ�ٱ�ʾ��ԭ
    Dim lngTemp As Long
    Dim i As Long

    With mshSum
        If lngRow < 0 Or lngRow > .rows - 1 Then Exit Sub
        .Redraw = False
        lngTemp = .Row
        .Row = lngRow
        If blnChange = True Then
            For i = 2 To .Cols - 1
                .Col = i
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
            Next
        Else
            For i = 2 To .Cols - 1
                .Col = i
                .CellBackColor = &H80000005
                .CellForeColor = &H80000008
            Next
        End If
        .Row = lngTemp
        .Redraw = True
    End With
End Sub

Private Sub mshSum_GotFocus()
    Call MenuSet
End Sub

Private Sub mshSum_LostFocus()
    Call MenuSet
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool, 2
    End If
End Sub

Private Sub picV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartX = X
    End If
End Sub

Private Sub picV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picV.Left + X - msngStartX
        If sngTemp > lblSum.Left + 600 And ScaleWidth - (sngTemp + picV.Width) > 1600 Then
            picV.Left = sngTemp
            lblSum.Width = picV.Left - lblSum.Left
            Form_Resize
        End If
    End If
End Sub

Private Sub mnufileexit_Click()
    Unload Me
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFilePreView_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub


Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Open"
            mnuViewOpen_Click
        Case "Quit"
            mnufileexit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnuFilePreView_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For Each buttTemp In tbrThis.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    cbrThis.Bands("one").MinHeight = tbrThis.Height
    Form_Resize
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(hWnd)
End Sub

Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    If mshSum Is ActiveControl Then
        Set objPrint.Body = mshSum
        objPrint.Title.Text = "�󴦷��б�"
        objRow.Add " "
        objRow.Add "��ѯʱ�䣺" & Format(mdatBegin, "yyyy-MM-dd") & " �� " & Format(mdatEnd, "yyyy-MM-dd")
        objPrint.UnderAppRows.Add objRow
        
        Set objRow = New zlTabAppRow
        objRow.Add "��ӡ�ˣ�" & gstrUserName
        objRow.Add "��ӡʱ�䣺" & Format(Sys.Currentdate, "yyyy-MM-dd")
        objPrint.BelowAppRows.Add objRow
    Else
        Set objPrint.Body = mshDetail
        objPrint.Title.Text = "������ϸ"
        objRow.Add "�����ţ�" & mshSum.TextMatrix(mshSum.Row, 3)
        objRow.Add "��ѯʱ�䣺" & Format(mdatBegin, "yyyy-MM-dd") & " �� " & Format(mdatEnd, "yyyy-MM-dd")
        objPrint.UnderAppRows.Add objRow
        
        Set objRow = New zlTabAppRow
        objRow.Add "��ӡ�ˣ�" & gstrUserName
        objRow.Add "��ӡʱ�䣺" & Format(Sys.Currentdate, "yyyy-MM-dd")
        objPrint.BelowAppRows.Add objRow
    End If
    If mshSum Is ActiveControl Then
        mshSum.Redraw = False
        SetColor False, mshSum.Row
    End If
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    If mshSum Is ActiveControl Then
        mshSum.Redraw = True
        SetColor True, mshSum.Row
    End If
End Sub

Private Function FillDept() As Boolean
'����:װ��ҩƷ��Ӧ��
    
    Dim rstemp As New ADODB.Recordset
    Dim strTemp As String
    Dim LngID As Long
    
    mlngID = -1     'ȫ��ˢ��ʱ���൱���û�û����κνڵ�
    If cmbDept.ListIndex > 0 Then
        LngID = cmbDept.ItemData(cmbDept.ListIndex)
    End If
    
    On Error GoTo errHandle
    rstemp.CursorLocation = adUseClient
    gstrSQL = "select id,���� from ���ű� A,��������˵�� B  where (A.վ�� = '" & gstrNodeNo & "' Or A.վ�� is Null) And (A.����ʱ��=to_date('3000-01-01','yyyy-mm-dd') or A.����ʱ�� is null) " & _
         " and A.ID=B.����ID and B.��������='�ٴ�' and B.�������<>0 order by A.����"
    Call zldatabase.OpenRecordset(rstemp, gstrSQL, Me.Caption)
    
    If rstemp.RecordCount = 0 Then
        MsgBox "���������ҡ�����Ϣ��ȫ���޷����в�ѯ��", vbExclamation, gstrSysName
        FillDept = False
        Exit Function
    End If
    
    
    With cmbDept
        .Clear
        .AddItem "���п���"
        Do Until rstemp.EOF
            .AddItem rstemp("����")
            .ItemData(.NewIndex) = rstemp("ID")
            If rstemp("ID") = LngID Then
                .ListIndex = .NewIndex
            End If
            rstemp.MoveNext
        Loop
        If .ListIndex < 0 Then .ListIndex = 0
    End With
    
    FillDept = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub FillSum()
'����:װ�����ͳ������
    Dim rstemp As New ADODB.Recordset
    Dim strBegin As String, strEnd As String
    Dim lngRow As Long
    Dim str�������� As String
    Dim str��������1 As String
    
    On Error GoTo errHandle
    If cmbDept.ListIndex < 0 Then Exit Sub
    If mlngID = cmbDept.ItemData(cmbDept.ListIndex) Then Exit Sub
    mlngID = cmbDept.ItemData(cmbDept.ListIndex)
    '��ʼ��ѯ
    
    strBegin = Format(mdatBegin, "yyyy-MM-dd")
    strEnd = Format(mdatEnd + 1, "yyyy-MM-dd")
    
    If InStr(1, mstrType, "012") > 0 Then
        str�������� = ""
    ElseIf InStr(1, mstrType, "01") > 0 Then
        str�������� = " And (A.�����־=1 Or A.�����־=4) "
    ElseIf InStr(1, mstrType, "02") > 0 Then
        str�������� = " And (A.�����־=1 Or A.�����־=4) And A.���ʷ���=0 "
        str��������1 = " And A.�����־=2 And A.���ʷ���=1 "
    ElseIf InStr(1, mstrType, "12") > 0 Then
        str�������� = " And A.���ʷ���=1 "
    ElseIf InStr(1, mstrType, "0") > 0 Then
         str�������� = " And (A.�����־=1 Or A.�����־=4) And A.���ʷ���=0 "
    ElseIf InStr(1, mstrType, "1") > 0 Then
         str�������� = " And (A.�����־=1 Or A.�����־=4) And A.���ʷ���=1 "
    ElseIf InStr(1, mstrType, "2") > 0 Then
         str�������� = " And A.�����־=2 And A.���ʷ���=1 "
    End If
    
    MousePointer = 11
    
    '�ٵõ�������SQL���
    If mlngID = 0 Then
        '���п��ҵĲ���Ա
        gstrSQL = "select A.������,to_char(A.�Ǽ�ʱ��,'yyyy-mm-dd') as ����,A.NO,decode(A.��¼����,1,'�շ�','����') as ��������,sum(A.ʵ�ս��) as ���," & _
                   " A.����,decode(A.�����־,1,'(����)',4,'(����)',decode(A.�����־,2,'(סԺ)','(����)'))||A.��ʶ�� ��ʶ��,A.����ID " & _
                   " from ������ü�¼ A,�շ���ĿĿ¼ C, " & _
                   " (Select Distinct ����,No, ����id From ҩƷ�շ���¼ Where ���� In (8, 9) And Mod(��¼״̬, 3) = 1 And ��������>=[2] And ��������<[3]) B " & _
                   " where A.�Ǽ�ʱ��>=[2] and A.�Ǽ�ʱ��<[3] " & _
                   "       and (A.��¼����=1 or A.��¼����=2)  and ��¼״̬=1 and A.������ is not null " & _
                   " And A.�շ�ϸĿid=C.Id And C.��� In('5','6','7') And A.No = B.No And A.Id = B.����id" & str�������� & _
                   " group by A.������,A.�Ǽ�ʱ��,A.NO ,A.��¼����,A.����,Decode(A.�����־, 1, '(����)',4,'(����)', Decode(A.�����־, 2, '(סԺ)', '(����)')) || A.��ʶ��,A.����ID " & _
                   " Having Sum(A.ʵ�ս��) >= [1] "
                   
    Else
        gstrSQL = "select A.������,to_char(A.�Ǽ�ʱ��,'yyyy-mm-dd') as ����,A.NO,decode(A.��¼����,1,'�շ�','����') as ��������,sum(A.ʵ�ս��) as ���, " & _
                   " ����,decode(A.�����־,1,'(����)',4,'(����)',decode(A.�����־,2,'(סԺ)','(����)'))||A.��ʶ�� ��ʶ��,A.����ID  " & _
                   " from ������ü�¼ A,�շ���ĿĿ¼ C, " & _
                   " (Select Distinct ����,No, ����id From ҩƷ�շ���¼ Where ���� In (8, 9) And Mod(��¼״̬, 3) = 1 And ��������>=[2] And ��������<[3]) B " & _
                   " where A.�Ǽ�ʱ��>=[2] and A.�Ǽ�ʱ��<[3] " & _
                   "       and (A.��¼����=1 or A.��¼����=2)  and A.��¼״̬=1  and A.������ is not null and A.��������ID+0=[4] " & _
                   " And A.�շ�ϸĿid=C.Id And C.��� In('5','6','7') And A.No = B.No And A.Id = B.����id" & str�������� & _
                   " group by A.������,A.�Ǽ�ʱ��,A.NO ,A.��¼����,A.����,Decode(A.�����־, 1, '(����)',4,'(����)', Decode(A.�����־, 2, '(סԺ)', '(����)')) || A.��ʶ��,A.����ID  " & _
                   " Having Sum(A.ʵ�ս��) >= [1] "
    End If
    
    If str��������1 <> "" Then
        If mlngID = 0 Then
            gstrSQL = gstrSQL & _
                   " UNION ALL" & _
                   " select A.������,to_char(A.�Ǽ�ʱ��,'yyyy-mm-dd') as ����,A.NO,decode(A.��¼����,1,'�շ�','����') as ��������,sum(A.ʵ�ս��) as ���," & _
                   " A.����,decode(A.�����־,1,'(����)',4,'(����)',decode(A.�����־,2,'(סԺ)','(����)'))||A.��ʶ�� ��ʶ��,A.����ID " & _
                   " from ������ü�¼ A,�շ���ĿĿ¼ C, " & _
                   " (Select Distinct ����,No, ����id From ҩƷ�շ���¼ Where ���� In (8, 9) And Mod(��¼״̬, 3) = 1 And ��������>=[2] And ��������<[3]) B " & _
                   " where A.�Ǽ�ʱ��>=[2] and A.�Ǽ�ʱ��<[3] " & _
                   "       and (A.��¼����=1 or A.��¼����=2)  and ��¼״̬=1 and A.������ is not null " & _
                   " And A.�շ�ϸĿid=C.Id And C.��� In('5','6','7') And A.No = B.No And A.Id = B.����id" & str��������1 & _
                   " group by A.������,A.�Ǽ�ʱ��,A.NO ,A.��¼����,A.����,Decode(A.�����־, 1, '(����)', 4,'(����)',Decode(A.�����־, 2, '(סԺ)', '(����)')) || A.��ʶ��,A.����ID " & _
                   " Having Sum(A.ʵ�ս��) >= [1] "
        Else
            gstrSQL = gstrSQL & _
                    " UNION ALL" & _
                    "select A.������,to_char(A.�Ǽ�ʱ��,'yyyy-mm-dd') as ����,A.NO,decode(A.��¼����,1,'�շ�','����') as ��������,sum(A.ʵ�ս��) as ���, " & _
                    " ����,decode(A.�����־,1,'(����)',4,'(����)',decode(A.�����־,2,'(סԺ)','(����)'))||A.��ʶ�� ��ʶ��,A.����ID  " & _
                    " from ������ü�¼ A,�շ���ĿĿ¼ C, " & _
                    " (Select Distinct ����,No, ����id From ҩƷ�շ���¼ Where ���� In (8, 9) And Mod(��¼״̬, 3) = 1 And ��������>=[2] And ��������<[3]) B " & _
                    " where A.�Ǽ�ʱ��>=[2] and A.�Ǽ�ʱ��<[3] " & _
                    "       and (A.��¼����=1 or A.��¼����=2)  and A.��¼״̬=1  and A.������ is not null and A.��������ID+0=[4] " & _
                    " And A.�շ�ϸĿid=C.Id And C.��� In('5','6','7') And A.No = B.No And A.Id = B.����id" & str��������1 & _
                    " group by A.������,A.�Ǽ�ʱ��,A.NO ,A.��¼����,A.����,Decode(A.�����־, 1, '(����)',4,'(����)', Decode(A.�����־, 2, '(סԺ)', '(����)')) || A.��ʶ��,A.����ID  " & _
                    " Having Sum(A.ʵ�ս��) >= [1] "
        End If
    End If
    
    If mstrType = "0" Or mstrType = "1" Or mstrType = "01" Then
    ElseIf mstrType = "2" Then
        gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
    Else
        gstrSQL = gstrSQL & " Union All " & Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
    End If
    gstrSQL = gstrSQL & " order by ������,����,NO"
    
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mdblMax, CDate(strBegin), CDate(strEnd), mlngID)
    
    mshSum.Redraw = False
    If rstemp.RecordCount = 0 Then
        ClearGrid mshSum
    Else
        mshSum.rows = rstemp.RecordCount + 1
    End If
    lngRow = 1
    With mshSum
        Do Until rstemp.EOF
            .TextMatrix(lngRow, 0) = rstemp("������")
            .TextMatrix(lngRow, 1) = IIf(IsNull(rstemp("����")), "", rstemp("����"))
            .TextMatrix(lngRow, 2) = IIf(IsNull(rstemp("��������")), "", rstemp("��������"))
            .TextMatrix(lngRow, 3) = IIf(IsNull(rstemp("NO")), "", rstemp("NO"))
            .TextMatrix(lngRow, 4) = Format(rstemp("���"), "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 5) = IIf(IsNull(rstemp("����")), "", rstemp("����"))
            .TextMatrix(lngRow, 6) = IIf(IsNull(rstemp("��ʶ��")), "", rstemp("��ʶ��"))
            .TextMatrix(lngRow, 7) = IIf(IsNull(rstemp("����ID")), "", rstemp("����ID"))
            lngRow = lngRow + 1
            rstemp.MoveNext
        Loop
    End With
    mshSum.Redraw = True
    
    stbThis.Panels(2).Text = "ʱ�䷶Χ��" & Format(mdatBegin, "yyyy-MM-dd") & " �� " & Format(mdatEnd, "yyyy-MM-dd") & _
                            "�� ��������" & rstemp.RecordCount
    MousePointer = 0
    Call FillDetail
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FillDetail()
'����:װ����ϸ����
    Dim rstemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim strNo As String, lng���� As Long
    Dim int��ʶ As String
    Dim strUnitQuantity As String
    Dim strFormat As String
    
    On Error GoTo errHandle
    MousePointer = 11
    
    mshDetail.Redraw = False
    '��ʼ�����
    ClearGrid mshDetail
    
    '�õ���ѯ����
    strNo = mshSum.TextMatrix(mshSum.Row, 3)
    lng���� = IIf(mshSum.TextMatrix(mshSum.Row, 2) = "�շ�", 8, 9)
    If InStrB(1, mshSum.TextMatrix(mshSum.Row, 6), "(����)") > 0 Then
        int��ʶ = 1
    Else
        int��ʶ = 2
    End If
    
    If strNo = "" Then
        mshDetail.Redraw = True
        MousePointer = 0
        
        Call MenuSet
        Exit Sub
    End If
    
    Select Case mintUnit
        Case 1      'ҩ����λ
            If int��ʶ = 1 Then
                '���ﵥλ
                strUnitQuantity = ",B.���ﵥλ AS ��λ,(A.ʵ������ / B.�����װ) AS ʵ������,a.���ۼ�*B.�����װ as ���ۼ�"
            Else
                'סԺ��λ
                strUnitQuantity = ",B.סԺ��λ AS ��λ,(A.ʵ������ / B.סԺ��װ) AS ʵ������,a.���ۼ�*B.סԺ��װ as ���ۼ�"
            End If
        Case Else   '�ۼ۵�λ
            strUnitQuantity = ",C.���㵥λ AS ��λ, a.ʵ������,a.�ɱ���,a.���ۼ�"
    End Select
    
    
    '��ʼ��ѯ
    gstrSQL = " SELECT DISTINCT C.���� ͨ������,C.���,A.����||DECODE(NVL(A.����,0),0,'','('||A.����||')') ����," & _
           " A.���۽��,A.��� " & strUnitQuantity & _
           " FROM ҩƷ�շ���¼ A,ҩƷ��� B,�շ���ĿĿ¼ C " & _
           " WHERE A.ҩƷID +0 = B.ҩƷID AND B.ҩƷID=C.ID " & _
           "       AND A.NO=[1] AND A.����=[2] AND MOD(A.��¼״̬,3)=1 " & _
           " ORDER BY A.���"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, lng����)
    
    If rstemp.RecordCount > 0 Then
        mshDetail.rows = rstemp.RecordCount + 1
        lngRow = 1
        With mshDetail
            Do Until rstemp.EOF
                .TextMatrix(lngRow, 0) = rstemp("ͨ������")
                .TextMatrix(lngRow, 1) = IIf(IsNull(rstemp("���")), "", rstemp("���"))
                .TextMatrix(lngRow, 2) = IIf(IsNull(rstemp("����")), "", rstemp("����"))
                .TextMatrix(lngRow, 3) = IIf(IsNull(rstemp("��λ")), "", rstemp("��λ"))
                .TextMatrix(lngRow, 4) = Format(rstemp("ʵ������"), "###########0.000;-###########0.000; ; ")
                .TextMatrix(lngRow, 5) = Format(rstemp("���ۼ�"), "###########0.000;-###########0.000; ; ")
                .TextMatrix(lngRow, 6) = Format(rstemp("���۽��"), "###########0.00;-###########0.00; ; ")
                    
                lngRow = lngRow + 1
                rstemp.MoveNext
            Loop
            
        End With
    End If
    mshDetail.Redraw = True
    MousePointer = 0
    
    Call MenuSet
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ClearGrid(objGrid As MSHFlexGrid, Optional lngCols As Long = 0)
'���ܣ�������,����ɲ��ֳ�ʼ��
    Dim i As Long
    
    With objGrid
        If lngCols > 0 Then
            '������������������Ǿͳ�ʼ����
            .Cols = lngCols
            .AllowBigSelection = True
            .FillStyle = flexFillRepeat
            .Col = 0
            .Row = 0
            .ColSel = .Cols - 1
            .RowSel = 0
            .CellAlignment = 4
            .FillStyle = flexFillSingle
            .AllowBigSelection = False
        End If
        
        .rows = 2
        .Row = 1
        .RowData(1) = 0
        For i = 0 To objGrid.Cols - 1
            objGrid.TextMatrix(1, i) = ""
            If objGrid Is mshSum And i > 1 Then
                mlngRow = 1
                .Col = i
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
            End If
        Next
    End With
End Sub

Private Sub MenuSet()
'����:��ʾ�˵��͹�������״̬(��ӡ)
    Dim blnPrint As Boolean
    
    If ActiveControl Is mshSum Then
        blnPrint = Not (mshSum.rows = 2 And mshSum.TextMatrix(1, 0) = "")
    Else
        blnPrint = Not (mshDetail.rows = 2 And mshDetail.TextMatrix(1, 0) = "")
    End If

    mnuFilePreview.Enabled = blnPrint
    mnuFilePrint.Enabled = blnPrint
    mnuFileExcel.Enabled = blnPrint
    tbrThis.Buttons("Preview").Enabled = blnPrint
    tbrThis.Buttons("Print").Enabled = blnPrint
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

