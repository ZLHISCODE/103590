VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDrugListRepot 
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
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   5880
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgdData 
         Height          =   2565
         Left            =   495
         TabIndex        =   3
         Top             =   945
         Width           =   5070
         _ExtentX        =   8943
         _ExtentY        =   4524
         _Version        =   393216
         BackColor       =   16777215
         Rows            =   10
         FixedCols       =   0
         BackColorFixed  =   16777215
         BackColorBkg    =   16777215
         AllowUserResizing=   3
         Appearance      =   0
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
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
            Picture         =   "frmDrugListRepot.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":005E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":00BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":011A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":0178
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":01D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":0234
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
            Picture         =   "frmDrugListRepot.frx":0292
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":02F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":034E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":03AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":040A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":0468
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":04C6
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
Attribute VB_Name = "frmDrugListRepot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------------------------------------------------------
Public inDeptId As Long            '�ⷿid
Public InDeptName  As String              '�ⷿ����
Public inDrugType As Long          'ҩƷ����id
Public inDrugTypeName As String        'ҩƷ��������

Dim dtpStartDate As String        '��ֹ����
Dim dtpEndDate As String        '��ֹ����
Dim strStartDate As String        '��ֹ����
Dim strEndDate As String        '��ֹ����
Dim DataRecordSet As ADODB.Recordset
Dim blnFirst As Boolean              'ȷ���Ƿ��һ��ʹ�ñ�ϵͳ
Dim Bln����ҩ As Boolean '��ʾ�Ƿ���в�ѯ����ҩ��Ȩ��
Dim Bln�г�ҩ As Boolean '��ʾ�Ƿ���в�ѯ�г�ҩ��Ȩ��
Dim Bln�в�ҩ As Boolean '��ʾ�Ƿ���в�ѯ�в�ҩ��Ȩ��
Dim Str���� As String


Private Sub Form_Activate()
    If Not blnFirst Then Exit Sub
    lbl�ⷿ.Caption = "�ⷿ��" & InDeptName
    lbl�ڼ�.Caption = "�ڼ�:" & dtpStartDate & "  ��  " & dtpEndDate
    ReFreshStru
    blnFirst = False
    If Not RefreshData Then Unload Me: Exit Sub
End Sub

Private Sub Form_Load()
    blnFirst = True
    dtpStartDate = Format(DateAdd("m", -1, Currentdate()), "yyyy-MM-DD HH:mm:ss")
    dtpEndDate = Format(Currentdate(), "yyyy-MM-DD HH:mm:ss")
    RestoreWinState Me
    If InStr(gstrStockSearchPrivs, "����ҩ") <> 0 Then
        Bln����ҩ = True
    Else
        Bln����ҩ = False
    End If
    
    If InStr(gstrStockSearchPrivs, "�г�ҩ") <> 0 Then
        Bln�г�ҩ = True
    Else
        Bln�г�ҩ = False
    End If
    
    If InStr(gstrStockSearchPrivs, "�в�ҩ") <> 0 Then
        Bln�в�ҩ = True
    Else
        Bln�в�ҩ = False
    End If

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
    
    Me.lblTitle.Top = 150
    Me.lblTitle.Left = 0
    Me.lblTitle.Width = Me.shpback.Width
    
    With lbl�ⷿ
        .Top = Me.lblTitle.Top + 500
        .Left = 200
    End With
    
    With Me.fgdData
        .Left = 200
        .Width = Me.shpback.Width - 400
        .Top = Me.lbl�ⷿ.Top + Me.lbl�ⷿ.Height + 45
        .Height = Me.shpback.Height - Me.fgdData.Top - 400
    End With
    Me.lbl�ڼ�.Top = Me.lblTitle.Top + 500
    Me.lbl�ڼ�.Left = Me.fgdData.Width + Me.fgdData.Left - Me.lbl�ڼ�.Width
    
    If Me.shpback.Width < Me.lblTitle.Width Then
        Me.lblTitle.Visible = False
        Me.fgdData.Visible = False
        Me.lbl�ڼ�.Visible = False
    Else
        Me.lblTitle.Visible = True
        Me.fgdData.Visible = True
        Me.lbl�ڼ�.Visible = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me
    Unload frmDrugListRepotAsk
End Sub

Private Sub mnuEXCEL_Click()
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    objPrint.Title.Text = Me.lblTitle.Caption
     
     Set objRow = New zlTabAppRow
     objRow.Add Me.lbl�ⷿ.Caption
     objRow.Add Me.lbl�ڼ�.Caption
     objPrint.UnderAppRows.Add objRow
     
     Set objPrint.Body = fgdData
     
      Set objRow = New zlTabAppRow
     With objRow
        .Add "��ӡ��:"
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
    
    objPrint.Title.Text = Me.lblTitle.Caption
     
     Set objRow = New zlTabAppRow
     objRow.Add Me.lbl�ⷿ.Caption
     objRow.Add Me.lbl�ڼ�.Caption
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
        Me.lblTitle.FontSize = 22
        Me.lbl�ⷿ.FontSize = 9
        Me.lbl�ڼ�.FontSize = 9
        Me.fgdData.Font.Size = 9
        Me.fgdData.FontFixed.Size = 9

     Case 1
        Me.lblTitle.FontSize = 24
        Me.lbl�ⷿ.FontSize = 11
        Me.lbl�ڼ�.FontSize = 11
        Me.fgdData.Font.Size = 11
        Me.fgdData.FontFixed.Size = 11

    Case 2
        Me.lblTitle.FontSize = 28
        Me.lbl�ⷿ.FontSize = 15
        Me.lbl�ڼ�.FontSize = 15
        Me.fgdData.Font.Size = 15
        Me.fgdData.FontFixed.Size = 15
    End Select
    Form_Resize
    Me.Refresh
End Sub

Private Sub mnuViewForeColor_Click()
    Dim lngForeColor As Long
    lngForeColor = zlGetColor(Me.lblTitle.ForeColor)
    Me.lblTitle.ForeColor = lngForeColor
    Me.lbl�ⷿ.ForeColor = lngForeColor
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
        Case "ͼ��"
            
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
    Dim Str��λ As String
    Dim Strϵ�� As String
    
    Dim lngRow As Long
    Dim i As Long
    Dim frmNewAsk As New frmDrugListRepotAsk
    Dim lng��ǰ��� As Double
    Dim lng��ǰ���� As Double
    Dim lng��ǰ��� As Double
    Dim lng�ڳ���� As Double
    Dim lng�ڳ���� As Double
    Dim lng�ڳ����� As Double
    Dim lng������� As Double
    Dim lng����� As Double
    Dim lng����� As Double
    Dim lng������ As Double
    Dim lng������ As Double
    Dim lng�������� As Double
    Dim lng��ĩ��� As Double
    Dim lng��ĩ��� As Double
    Dim lng��ĩ���� As Double
    Dim lng���۽�� As Double
    Dim lng���۲�� As Double
    
    Dim dbl��ǰ��� As Double
    Dim dbl��ǰ���� As Double
    Dim dbl��ǰ��� As Double
    Dim dbl�ڳ����� As Double
    Dim Dbl�ڳ���� As Double
    Dim Dbl�ڳ���� As Double
    Dim dbl������� As Double
    Dim dbl����� As Double
    Dim dbl����� As Double
    Dim dbl�������� As Double
    Dim dbl������ As Double
    Dim dbl������ As Double
    Dim dbl��ĩ��� As Double
    Dim dbl���۽�� As Double
    Dim dbl���۲�� As Double
    Dim dbl��ĩ��� As Double
    Dim dbl��ĩ���� As Double
    
    Dim str��; As String
    
    On Error GoTo errHandle
    RefreshData = False
    Load frmDrugListRepotAsk
    With frmDrugListRepotAsk
        .inDeptId = inDeptId
        .Show 1, Me
        If Not .blnAskOk Then Exit Function
        inDeptId = .cbo�ⷿ.ItemData(.cbo�ⷿ.ListIndex)
        InDeptName = .cbo�ⷿ.Text
        strStartDate = Format(.dtpStartDate.Value, "yyyyMMDDHHmmss")
        strEndDate = Format(.dtpEndDate.Value, "yyyyMMDDHHmmss")
        dtpStartDate = Format(.dtpStartDate.Value, "yyyy-MM-DD HH:mm:ss")
        dtpEndDate = Format(.dtpEndDate.Value, "yyyy-MM-DD HH:mm:ss")
                
    End With
    
    Str���� = "''"
    If Bln����ҩ Then Str���� = "'����ҩ'"
    If Bln�г�ҩ Then
        If Bln����ҩ Then
            Str���� = Str���� & ",'�г�ҩ'"
        Else
            Str���� = "'�г�ҩ'"
        End If
    End If
    If Bln�в�ҩ Then
        If Bln�г�ҩ Or Bln����ҩ Then
            Str���� = Str���� & ",'�в�ҩ'"
        Else
            Str���� = "'�в�ҩ'"
        End If
    End If
    
    Select Case frmDrugQuery.intChoose����
            Case 1
                Str��λ = "B.�ۼ۵�λ As ��λ,"
                Strϵ�� = "1"
            Case 2
                Str��λ = "B.���ﵥλ As ��λ,"
                Strϵ�� = "B.�����װ"
            Case 3
                Str��λ = "B.ҩ�ⵥλ As ��λ,"
                Strϵ�� = "B.ҩ���װ"
            Case 4
                Str��λ = "B.סԺ��λ As ��λ,"
                Strϵ�� = "B.סԺ��װ"
    End Select
    
    
    str��; = frmDrugQuery.tvwSection_S.SelectedItem.Key
    
'    StrSql = "Select Distinct A.��ǰ����/Decode(" & Strϵ�� & ",0,1," & Strϵ�� & ") As ��ǰ����,A.��ǰ��� As ��ǰ���,A.��ǰ��� As ��ǰ���," & _
             " C.����ǰ��������/Decode(" & Strϵ�� & ",0,1," & Strϵ�� & ") As ����ǰ��������,C.����ǰ������� As ����ǰ�������,C.����ǰ������� As ����ǰ�������," & _
             " C.����ĩ�������/Decode(" & Strϵ�� & ",0,1," & Strϵ�� & ") As ����ĩ�������,C.����ĩ����� As ����ĩ�����,C.����ĩ����� As ����ĩ�����," & _
             " C.����ĩ��������/Decode(" & Strϵ�� & ",0,1," & Strϵ�� & ") As ����ĩ��������,C.����ĩ������ As ����ĩ������,C.����ĩ������ As ����ĩ������," & _
              Str��λ & "B.ҩƷid,B.���� As ����,X.ͨ������ As ����,B.��� As ���" & _
            " From (Select ҩƷid,Sum(ʵ������) As ��ǰ����,Sum(ʵ�ʽ��) As ��ǰ���,Sum(ʵ�ʲ��) As ��ǰ��� From ҩƷ��� Where ����=1 " & IIf(inDeptId = 0, "", "And �ⷿid=" & inDeptId) & "Group by ҩƷid ) A," & _
            " (Select ҩƷid,Sum(����ǰ��������) As ����ǰ��������,Sum(����ǰ�������) As ����ǰ�������,Sum(����ǰ�������) As ����ǰ�������," & _
            "        Sum(����ĩ�������) As ����ĩ�������,Sum(����ĩ�����) As ����ĩ�����,Sum(����ĩ�����) As ����ĩ�����, " & _
            "        Sum(����ĩ��������) As ����ĩ��������,Sum(����ĩ������) As ����ĩ������,Sum(����ĩ������) As ����ĩ������" & _
            "  From (Select E.ҩƷid,Sum(����) As ����ǰ��������,Sum(E.���) As ����ǰ�������,Sum(���) As ����ǰ�������," & _
            "       Decode(Sign(To_number(To_Char(E.����,'yyyymmdd'))-" & strEndDate & "),-1,0,decode(F.ϵ��,-1,0,1)*Sum(E.����)) as ����ĩ�������, " & _
            "       Decode(Sign(To_number(To_Char(E.����,'yyyymmdd'))-" & strEndDate & "),-1,0,decode(F.ϵ��,-1,-1,0)*Sum(E.����)) as ����ĩ��������, " & _
            "       Decode(Sign(To_number(To_Char(E.����,'yyyymmdd'))-" & strEndDate & "),-1,0,decode(F.ϵ��,-1,0,1)*Sum(E.���)) as ����ĩ�����, " & _
            "       Decode(Sign(To_number(To_Char(E.����,'yyyymmdd'))-" & strEndDate & "),-1,0,decode(F.ϵ��,-1,-1,0)*Sum(E.���)) as ����ĩ������, " & _
            "       Decode(Sign(To_number(To_Char(E.����,'yyyymmdd'))-" & strEndDate & "),-1,0,decode(F.ϵ��,-1,0,1)*Sum(E.���)) as ����ĩ�����, " & _
            "       Decode(Sign(To_number(To_Char(E.����,'yyyymmdd'))-" & strEndDate & "),-1,0,decode(F.ϵ��,-1,-1,0)*Sum(E.���)) as ����ĩ������  " & _
            "        From ҩƷ�շ����� E,ҩƷ������ F " & _
            "        Where To_number(To_Char(E.����,'yyyymmdd'))>= " & strStartDate & IIf(inDeptId = 0, "", "And �ⷿid=" & inDeptId) & " And E.���id=F.id " & _
            "        Group By E.ҩƷid,E.����,F.ϵ��)" & _
            "  Group By ҩƷid)C," & _
            " ҩƷĿ¼ B,ҩƷ��Ϣ X" & _
            " Where B.ҩƷid=A.ҩƷid(+) And B.ҩƷid=C.ҩƷid(+) And B.ҩ��id=X.ҩ��id and (B.����ʱ�� IS NULL OR TO_CHAR (B.����ʱ��, 'yyyy-MM-dd') = '3000-01-01') " _
            & IIf(Left(str��;, 1) = "R", " and x.���ʷ��� In ('" & Mid(str��;, 2) & "')", " And x.��;����id in ( Select id From ҩƷ��;���� Q start with Q.id= " & Mid(str��;, 2) & " connect by prior id=�ϼ�id)") _
            & " Order By B.ҩƷid "
    
    strsql = "Select Distinct A.��ǰ����/Decode(" & Strϵ�� & ",0,1," & Strϵ�� & ") As ��ǰ����,A.��ǰ��� As ��ǰ���,A.��ǰ��� As ��ǰ���," & _
             " C.����ǰ��������/Decode(" & Strϵ�� & ",0,1," & Strϵ�� & ") As ����ǰ��������,C.����ǰ������� As ����ǰ�������,C.����ǰ������� As ����ǰ�������," & _
             " C.����ĩ�������/Decode(" & Strϵ�� & ",0,1," & Strϵ�� & ") As ����ĩ�������,C.����ĩ����� As ����ĩ�����,C.����ĩ����� As ����ĩ�����," & _
             " C.����ĩ��������/Decode(" & Strϵ�� & ",0,1," & Strϵ�� & ") As ����ĩ��������,C.����ĩ������ As ����ĩ������,C.����ĩ������ As ����ĩ������," & _
              Str��λ & "B.ҩƷid,B.���� As ����,X.ͨ������ As ����,B.��� As ���" & _
            " From (Select ҩƷid,Sum(ʵ������) As ��ǰ����,Sum(ʵ�ʽ��) As ��ǰ���,Sum(ʵ�ʲ��) As ��ǰ��� From ҩƷ��� Where ����=1 " & IIf(inDeptId = 0, "", "And �ⷿid=" & inDeptId) & "Group by ҩƷid ) A," & _
            " (SELECT ҩƷid," _
                & "(sum(DECODE(���ϵ��,-1,0,1)*ʵ������)- sum(DECODE(���ϵ��,-1,1,0)*ʵ������)) as ����ǰ��������," _
                & "(sum(DECODE(���ϵ��,-1,0,1)*���۽��)-sum(DECODE(���ϵ��,-1,1,0)*���۽��)) as ����ǰ�������," _
                & "(SUM(DECODE(���ϵ��,-1,0,1)*���)-SUM(DECODE(���ϵ��,-1,1,0)*���)) as  ����ǰ�������," _
                & "sum(Decode(Sign(To_number(To_Char(�������,'yyyymmddhh24miss'))-" & strEndDate & "),1,0,DECODE(���ϵ��,-1,0,1)*ʵ������)) AS ����ĩ�������," _
                & "sum(Decode(Sign(To_number(To_Char(�������,'yyyymmddhh24miss'))-" & strEndDate & "),1,0,DECODE(���ϵ��,-1,0,1)*���۽��)) AS ����ĩ�����," _
                & "sum(Decode(Sign(To_number(To_Char(�������,'yyyymmddhh24miss'))-" & strEndDate & "),1,0,DECODE(���ϵ��,-1,0,1)*���)) AS ����ĩ�����," _
                & "sum(Decode(Sign(To_number(To_Char(�������,'yyyymmddhh24miss'))-" & strEndDate & "),1,0,DECODE(���ϵ��,-1,1,0)*ʵ������)) AS ����ĩ��������," _
                & "sum(Decode(Sign(To_number(To_Char(�������,'yyyymmddhh24miss'))-" & strEndDate & "),1,0,DECODE(���ϵ��,-1,1,0)*���۽��)) AS ����ĩ������," _
                & "sum(Decode(Sign(To_number(To_Char(�������,'yyyymmddhh24miss'))-" & strEndDate & "),1,0,DECODE(���ϵ��, -1, 1, 0) * ���)) As ����ĩ������ " _
                & " From ҩƷ�շ���¼ " _
               & " WHERE �������>=to_date('" & strStartDate & "','yyyy-mm-dd hh24:mi:ss') " _
               & IIf(inDeptId = 0, "", "And �ⷿid=" & inDeptId) & _
            "  Group By ҩƷid)C," & _
            " ҩƷĿ¼ B,ҩƷ��Ϣ X" & _
            " Where B.ҩƷid=A.ҩƷid(+) And B.ҩƷid=C.ҩƷid(+) And B.ҩ��id=X.ҩ��id and (B.����ʱ�� IS NULL OR TO_CHAR (B.����ʱ��, 'yyyy-MM-dd') = '3000-01-01') " _
            & IIf(Left(str��;, 1) = "R", " and x.���ʷ��� In ('" & Mid(str��;, 2) & "')", " And x.��;����id in ( Select id From ҩƷ��;���� Q start with Q.id= " & Mid(str��;, 2) & " connect by prior id=�ϼ�id)") _
            & " Order By B.ҩƷid "
    Set DataRecordSet = New ADODB.Recordset
    ShowFlash "����װ�����ݣ����Ժ�", Me
    DoEvents
   
    With DataRecordSet
        Call SQLTest(App.Title, Me.Caption, strsql)
        Set DataRecordSet = zldatabase.OpenSQLRecord(strsql, "RefreshData")
        Call SQLTest
        If .RecordCount = 0 Then
            StopFlash
            MsgBox "�ڴ�������Ȩ�޷�Χ��,���κ���ϸ���¼��", vbInformation, gstrSysName
            Exit Function
        End If
        lng�ڳ���� = 0: lng�ڳ���� = 0: lng����� = 0: lng����� = 0: lng������ = 0: lng������ = 0: lng��ĩ��� = 0:        lng��ĩ��� = 0
        lng���۽�� = 0: lng���۲�� = 0: lng�ڳ����� = 0
        Dbl�ڳ���� = 0: Dbl�ڳ���� = 0: dbl����� = 0: dbl����� = 0: dbl������ = 0: dbl������ = 0: dbl��ĩ��� = 0:        dbl��ĩ��� = 0
        dbl���۽�� = 0: dbl���۲�� = 0: dbl�ڳ����� = 0
        
        fgdData.rows = IIf(.RecordCount = 0, 1, .RecordCount) + 2
        Call RefreshGridColWidth(Me.fgdData, 0)
         i = 2
        Do While Not .EOF
            
            Dbl�ڳ���� = IIf(IsNull(.Fields("��ǰ���").Value), 0, .Fields("��ǰ���").Value) - IIf(IsNull(.Fields("����ǰ�������").Value), 0, .Fields("����ǰ�������").Value)
            Dbl�ڳ���� = IIf(IsNull(.Fields("��ǰ���").Value), 0, .Fields("��ǰ���").Value) - IIf(IsNull(.Fields("����ǰ�������").Value), 0, .Fields("����ǰ�������").Value)
            dbl�ڳ����� = IIf(IsNull(.Fields("��ǰ����").Value), 0, .Fields("��ǰ����").Value) - IIf(IsNull(.Fields("����ǰ��������").Value), 0, .Fields("����ǰ��������").Value)
            
            dbl����� = IIf(IsNull(.Fields("����ĩ�����").Value), 0, .Fields("����ĩ�����").Value)
            dbl����� = IIf(IsNull(.Fields("����ĩ�����").Value), 0, .Fields("����ĩ�����").Value)
            dbl������� = IIf(IsNull(.Fields("����ĩ�������").Value), 0, .Fields("����ĩ�������").Value)
            dbl������ = IIf(IsNull(.Fields("����ĩ������").Value), 0, .Fields("����ĩ������").Value)
            dbl������ = IIf(IsNull(.Fields("����ĩ������").Value), 0, .Fields("����ĩ������").Value)
            dbl�������� = IIf(IsNull(.Fields("����ĩ��������").Value), 0, .Fields("����ĩ��������").Value)
                        
            dbl��ĩ��� = Dbl�ڳ���� + dbl����� - dbl������
            dbl��ĩ��� = Dbl�ڳ���� + dbl����� - dbl������
            dbl��ĩ���� = dbl�ڳ����� + dbl������� - dbl��������
            
            
            lng�ڳ���� = lng�ڳ���� + Dbl�ڳ����
            lng�ڳ���� = lng�ڳ���� + Dbl�ڳ����
            lng����� = lng����� + dbl�����
            lng����� = lng����� + dbl�����
            lng������ = lng������ + dbl������
            lng������ = lng������ + dbl������
            lng��ĩ��� = lng��ĩ��� + dbl��ĩ���
            lng��ĩ��� = lng��ĩ��� + dbl��ĩ���
            
            fgdData.TextMatrix(i, 0) = IIf(IsNull(.Fields("����").Value), "", .Fields("����").Value)
            fgdData.TextMatrix(i, 1) = IIf(IsNull(.Fields("����").Value), "", .Fields("����").Value)
            fgdData.TextMatrix(i, 2) = IIf(IsNull(.Fields("���").Value), "", .Fields("���").Value)
            fgdData.TextMatrix(i, 3) = IIf(IsNull(.Fields("��λ").Value), "", .Fields("��λ").Value)
            
            fgdData.TextMatrix(i, 4) = " " & Format(dbl�ڳ�����, "##,###0.000;-##,###0.000; ; ")
            fgdData.TextMatrix(i, 5) = " " & Format(Dbl�ڳ����, "##,###0.00;-##,###0.00; ; ")
            fgdData.TextMatrix(i, 6) = Format(Dbl�ڳ����, "##,###0.00;-##,###0.00; ; ")
            
            fgdData.TextMatrix(i, 7) = " " & Format(dbl�������, "##,###0.000;-##,###0.000; ; ")
            fgdData.TextMatrix(i, 8) = " " & Format(dbl�����, "##,###0.00;-##,###0.00; ; ")
            fgdData.TextMatrix(i, 9) = Format(dbl�����, "##,###0.00;-##,###0.00; ; ")
            
            fgdData.TextMatrix(i, 10) = " " & Format(dbl��������, "##,###0.000;-##,###0.000; ; ")
            fgdData.TextMatrix(i, 11) = " " & Format(dbl������, "##,###0.00;-##,###0.00; ; ")
            fgdData.TextMatrix(i, 12) = Format(dbl������, "##,###0.00;-##,###0.00; ; ")
'            fgdData.TextMatrix(i, 13) = " "
'            fgdData.TextMatrix(i, 14) = " "
            fgdData.TextMatrix(i, 13) = " " & Format(dbl��ĩ����, "##,###0.000;-##,###0.000; ; ")
            fgdData.TextMatrix(i, 14) = " " & Format(dbl��ĩ���, "##,###0.00;-##,###0.00; ; ")
            fgdData.TextMatrix(i, 15) = Format(dbl��ĩ���, "##,###0.00;-##,###0.00; ; ")
            Call RefreshGridColWidth(Me.fgdData, i)
            .MoveNext
            i = i + 1
        Loop
        If lng�ڳ���� <> 0 Or lng�ڳ���� <> 0 Or lng����� <> 0 Or lng����� <> 0 Or lng������ <> 0 Or lng������ <> 0 Or _
            lng��ĩ��� <> 0 Or lng��ĩ��� <> 0 Then
            fgdData.rows = fgdData.rows + 1
            fgdData.MergeRow(i) = True
            fgdData.TextMatrix(i, 0) = "�ϼ�"
            fgdData.TextMatrix(i, 1) = "�ϼ�"
            fgdData.TextMatrix(i, 2) = "�ϼ�"
            fgdData.TextMatrix(i, 3) = "�ϼ�"
            fgdData.TextMatrix(i, 4) = ""
            fgdData.TextMatrix(i, 5) = " " & Format(lng�ڳ����, "##,###0.00;-##,###0.00; ; ")
            fgdData.TextMatrix(i, 6) = Format(lng�ڳ����, "##,###0.00;-##,###0.00; ; ")
            fgdData.TextMatrix(i, 7) = "   "
            fgdData.TextMatrix(i, 8) = Format(lng�����, "##,###0.00;-##,###0.00; ; ")
            fgdData.TextMatrix(i, 9) = "  " & Format(lng�����, "##,###0.00;-##,###0.00; ; ")
            fgdData.TextMatrix(i, 10) = ""
            fgdData.TextMatrix(i, 11) = "  " & Format(lng������, "##,###0.00;-##,###0.00; ; ")
            fgdData.TextMatrix(i, 12) = Format(lng������, "##,###0.00;-##,###0.00; ; ")
'            fgdData.TextMatrix(i, 13) = " "
'            fgdData.TextMatrix(i, 14) = " "
            fgdData.TextMatrix(i, 13) = "  "
            fgdData.TextMatrix(i, 14) = Format(lng��ĩ���, "##,###0.00;-##,###0.00; ; ")
            fgdData.TextMatrix(i, 15) = "  " & Format(lng��ĩ���, "##,###0.00;-##,###0.00; ; ")
            Call RefreshGridColWidth(Me.fgdData, i)
        End If
        fgdData.Redraw = True
    End With
    lbl�ⷿ.Caption = "�ⷿ��" & InDeptName & Space(6) & "ҩƷ��;:" & inDrugTypeName
    lbl�ڼ�.Caption = "�ڼ�:" & dtpStartDate & "  ��  " & dtpEndDate
    StopFlash
    RefreshData = True
Exit Function
Err:
    StopFlash
    RefreshData = False
    Me.fgdData.Redraw = True
    MsgBox "�ڻ�ȡҩƷ��ϸ��ʱ,�����˲���Ԥ֪�Ĵ���!", vbInformation, gstrSysName
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
    Me.lblTitle.Caption = GetUnitName & "ҩƷ��ϸ��"

     With fgdData
            .Cols = 16
            .Redraw = False
            .rows = 6
            .FixedRows = 2
            .FixedCols = 0
            .MergeCells = flexMergeRestrictRows
            For IntCol = 0 To .Cols - 1
                .ColAlignmentFixed(IntCol) = 4
                If IntCol <= 3 Then
                    .ColAlignment(IntCol) = 1
                Else
                    .ColAlignment(IntCol) = 7
                End If
                If IntCol <= 3 Then
                    .ColWidth(IntCol) = IIf(IntCol <> 1, IIf(IntCol = 2, 1200, IIf(IntCol = 0, 600, 400)), 1400)
                Else
                    .ColWidth(IntCol) = 1000
                End If
            Next
            .MergeRow(0) = True
            .MergeCol(0) = True
            .MergeCol(1) = True
            .MergeCol(2) = True
            .MergeCol(3) = True
            
            .TextMatrix(0, 0) = "����"
            .TextMatrix(1, 0) = "����"
            .TextMatrix(0, 1) = "����"
            .TextMatrix(1, 1) = "����"
            .TextMatrix(0, 2) = "���"
            .TextMatrix(1, 2) = "���"
            
            Select Case frmDrugQuery.intChoose����
                Case 1
                    .TextMatrix(0, 3) = "�ۼ۵�λ"
                    .TextMatrix(1, 3) = "�ۼ۵�λ"
                Case 2
                    .TextMatrix(0, 3) = "���ﵥλ"
                    .TextMatrix(1, 3) = "���ﵥλ"
                Case 3
                    .TextMatrix(0, 3) = "�ⷿ��λ"
                    .TextMatrix(1, 3) = "�ⷿ��λ"
                Case 4
                    .TextMatrix(0, 3) = "סԺ��λ"
                    .TextMatrix(1, 3) = "סԺ��λ"
            End Select

            
            .TextMatrix(0, 4) = "�ڳ�"
            .TextMatrix(0, 5) = "�ڳ�"
            .TextMatrix(0, 6) = "�ڳ�"
            .TextMatrix(1, 4) = "����"
            .TextMatrix(1, 5) = "���"
            .TextMatrix(1, 6) = "���"
            
            .TextMatrix(0, 7) = "�������"
            .TextMatrix(0, 8) = "�������"
            .TextMatrix(0, 9) = "�������"
            .TextMatrix(1, 7) = "����"
            .TextMatrix(1, 8) = "���"
            .TextMatrix(1, 9) = "���"
            
            .TextMatrix(0, 10) = "���ڳ���"
            .TextMatrix(0, 11) = "���ڳ���"
            .TextMatrix(0, 12) = "���ڳ���"
            .TextMatrix(1, 10) = "����"
            .TextMatrix(1, 11) = "���"
            .TextMatrix(1, 12) = "���"
            
            
            .TextMatrix(0, 13) = "��ĩ"
            .TextMatrix(0, 14) = "��ĩ"
            .TextMatrix(0, 15) = "��ĩ"
            .TextMatrix(1, 13) = "����"
            .TextMatrix(1, 14) = "���"
            .TextMatrix(1, 15) = "���"
             Call RefreshGridColWidth(Me.fgdData, 0)
            .Redraw = True
        End With
End Sub
