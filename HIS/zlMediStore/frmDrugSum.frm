VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDrugSum 
   BackColor       =   &H8000000C&
   Caption         =   "ҩƷ����"
   ClientHeight    =   5445
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7785
   Icon            =   "frmDrugSum.frx":0000
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
         HighLight       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
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
         Caption         =   "ҩƷ����"
         BeginProperty Font 
            Name            =   "����_GB2312"
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
      _Version        =   "6.0.8169"
      Child1          =   "tbrThis"
      MinHeight1      =   660
      Width1          =   5370
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   660
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   1164
         ButtonWidth     =   820
         ButtonHeight    =   1164
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
            Picture         =   "frmDrugSum.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSum.frx":0526
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSum.frx":0742
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSum.frx":095C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSum.frx":0B78
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSum.frx":0D94
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSum.frx":0FB0
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
            Picture         =   "frmDrugSum.frx":11CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSum.frx":13E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSum.frx":1604
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSum.frx":181E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSum.frx":1A3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSum.frx":1C56
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSum.frx":1E72
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
            Picture         =   "frmDrugSum.frx":208E
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7990
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
      Begin VB.Menu mnuViewBlc 
         Caption         =   "��ʾ���(&Z)"
         Checked         =   -1  'True
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
Attribute VB_Name = "frmDrugSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------------------------------------------------------
Public inDeptId As Long            '�ⷿid
Public InDeptName  As String              '�ⷿ����
Public Bln��� As Boolean        '�Ƿ�Ķ����ѡ��
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



Private Sub fgdData_DblClick()

    If Me.fgdData.RowData(fgdData.Row) = 999999 Then Exit Sub
    If Me.fgdData.TextMatrix(fgdData.Row, 1) = "" Then Exit Sub
    
    Dim rsTemp As New ADODB.Recordset
    Dim StrSQL As String
    With rsTemp
        StrSQL = "Select id,����,NO,nvl(�۸�id,0) as �۸�id" & _
                " From ҩƷ�շ���¼" & _
                " Where No='" & Mid(Trim(Me.fgdData.TextMatrix(fgdData.Row, 1)), 3) & "'" & _
                "       And ����=" & Me.fgdData.RowData(fgdData.Row)
        If .State = adStateOpen Then .Close
        .Open StrSQL, gcnOracle, adOpenKeyset
        If .EOF Or .BOF Then Exit Sub
        
  '1-�⹺��⣻2-������⣻3-Эҩ��⣻4-������⣻5-��۵�����6-�ⷿ�Ƴ���7-�������ã�8-�շѴ�����9-���ʵ�������10-���ʱ�����11-�������⣻12-�̵㣻13-���۱䶯
        
        Select Case !����
        Case 1
            frmPurchaseCard.ShowCard Me, !No, 4
        Case 2
            frmSelfMakeCard.ShowCard Me, !No, 4
        Case 3
            frmAccordDrugCard.ShowCard Me, !No, 4
        Case 4
            frmOtherInputCard.ShowCard Me, !No, 4
        Case 5
            frmDiffPriceAdjustCard.ShowCard Me, !No, 4
        Case 6
            frmTransferCard.ShowCard Me, !No, 4
        Case 7
            frmDrawCard.ShowCard Me, !No, 4
        Case 11
            frmOtherOutputCard.ShowCard Me, !No, 4
        Case 12
            frmCheckCard.ShowCard Me, !No, 4
        Case 13
            gstrUserName = UserInfo.�û�����
            With frmAdjust
                .lngBillId = rsTemp!�۸�id
                .lngMediId = 1
                .Show 1, Me
            End With
        Case Else
            Frm����See.byt���� = !����
            Frm����See.StrNo = !No
            Frm����See.Show 1, Me
        End Select
    End With

End Sub

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
    Bln��� = False
    dtpStartDate = Format(DateAdd("m", -1, Currentdate()), "yyyy-MM-DD")
    dtpEndDate = Format(Currentdate(), "yyyy-MM-DD")
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
    On Error Resume Next
    Unload frmDrugSumAsk
    SaveWinState Me
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
    zlMailTo Me.hwnd
End Sub

Private Sub mnuHelpZlWeb_Click()
    zlHomePage Me.hwnd
End Sub


Private Sub mnuViewBlc_Click()
    mnuViewBlc.Checked = Not mnuViewBlc.Checked
    Bln��� = True
    Call ReFreshStru
    Call RefreshData
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
        For intCount = 1 To Me.tbrThis.Buttons.Count
            Me.tbrThis.Buttons(intCount).Caption = Me.tbrThis.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To Me.tbrThis.Buttons.Count
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
        For intCount = 1 To Me.tbrThis.Buttons.Count
            Me.tbrThis.Buttons(intCount).Caption = Me.tbrThis.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To Me.tbrThis.Buttons.Count
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
       
        Case "����"
             PopupMenu mnuViewFont
        Case "ǰ��ɫ"
            mnuViewForeColor_Click
        Case "����ɫ" '
            mnuViewBackColor_Click
'        Case "����"
'            mnuHelpHelp_Click
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
    Dim StrSQL As String
    Dim lngRow As Long
    Dim rsRecord As ADODB.Recordset
'    Dim frmNewAsk As New frmDrugSumAsk
    Dim dblCurr��� As Double
    Dim dblCurrMny As Double
    Dim dblOutMny As Double
    Dim dblInMny As Double
    Dim DblCgMny As Double
    
    Dim Dbl�ڳ���� As Double
    Dim Dbl�ڳ���� As Double
    Dim DBlʵ�ʽ�� As Double
    Dim DBlʵ�ʲ�� As Double
    RefreshData = False
    
    If Bln��� = False Then
        Load frmDrugSumAsk
        With frmDrugSumAsk
            .inDeptId = inDeptId
            .Show 1, Me
            If Not .blnAskOk Then Exit Function
            inDeptId = .cbo�ⷿ.ItemData(.cbo�ⷿ.ListIndex)
            InDeptName = .cbo�ⷿ.Text
            strStartDate = Format(.dtpStartDate.Value, "yyyyMMDD")
            dtpStartDate = Format(.dtpStartDate.Value, "yyyy-MM-DD")
            strEndDate = Format(.dtpEndDate.Value, "yyyyMMDD")
            dtpEndDate = Format(.dtpEndDate.Value, "yyyy-MM-DD")
        End With
    
    Else: Bln��� = False
    
    End If
    
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

    
    ShowFlash "����װ�����ݣ����Ժ�", Me
    DoEvents
    
    '�����ڵ�ʵ�ʽ�ʼ,���ڳ����
     Set rsRecord = New ADODB.Recordset
     StrSQL = " Select Sum(ʵ�ʽ��) as ʵ�ʽ��,Sum(ʵ�ʲ��) as ʵ�ʲ�� " & _
            "From ҩƷ��� Where ����=1 " & IIf(inDeptId = 0, "", " And �ⷿid =" & inDeptId) & _
            "And ҩƷid In (Select A.ҩƷid From ҩƷĿ¼ A,ҩƷ��Ϣ B Where A.ҩ��id=B.ҩ��id And B.���ʷ��� In (" & Str���� & ")) "
    Call SQLTest(App.Title, Me.Caption, StrSQL)
    rsRecord.Open StrSQL, gcnOracle
    Call SQLTest


          
     DBlʵ�ʽ�� = IIf(IsNull(rsRecord!ʵ�ʽ��), 0, rsRecord!ʵ�ʽ��)
     DBlʵ�ʲ�� = IIf(IsNull(rsRecord!ʵ�ʲ��), 0, rsRecord!ʵ�ʲ��)
     rsRecord.Close
    
'    StrSql = " Select  sum(A.���) as ���,sum(A.���) as ���" & _
            "  From ( Select  " & _
            "         B.���,B.��� " & _
            "         From ҩƷ�շ����� B,ҩƷ������ C " & _
            "         Where B.���� >=To_Date('" & strStartDate & "','yyyymmdd') And " & IIf(inDeptId = 0, "", "B.�ⷿID = " & inDeptId & " And ") & " B.���id=C.id" & _
            "               And B.ҩƷid In (Select X.ҩƷid From ҩƷĿ¼ X,ҩƷ��Ϣ Y Where X.ҩ��id=Y.ҩ��id And Y.���ʷ��� In (" & Str���� & "))) A"
    
    StrSQL = " Select  sum(���) as ���,sum(���) as ���" & _
            "    From ҩƷ�շ����� B " & _
            "   Where B.���� >=To_Date('" & strStartDate & "','yyyymmdd') " _
              & IIf(inDeptId = 0, "", " and B.�ⷿID = " & inDeptId) _
              & " And B.ҩƷid In " _
              & " (Select X.ҩƷid From ҩƷĿ¼ X,ҩƷ��Ϣ Y Where X.ҩ��id=Y.ҩ��id And Y.���ʷ��� In (" & Str���� & ")) "
    
    Set DataRecordSet = New ADODB.Recordset
    Call SQLTest(App.Title, Me.Caption, StrSQL)
    DataRecordSet.Open StrSQL, gcnOracle
    Call SQLTest
    
    
    Dbl�ڳ���� = DBlʵ�ʽ��
    Dbl�ڳ���� = DBlʵ�ʲ��
    With DataRecordSet
        
            Dbl�ڳ���� = Dbl�ڳ���� - IIf(IsNull(!���), 0, !���)
            Dbl�ڳ���� = Dbl�ڳ���� - IIf(IsNull(!���), 0, !���)
       
    End With
    
    
    
    '���ڼ䷢����
    '1-�⹺��⣻2-������⣻3-Эҩ��⣻4-������⣻5-��۵�����6-�ⷿ�Ƴ���7-�������ã�8-�շѴ�����9-���ʵ�������10-���ʱ�����11-�������⣻12-�̵㣻13-���۱䶯
    
        StrSQL = " Select A.�������,A.NO,����,ltrim(C.����) as ժҪ, " & _
            "        abs(sum(A.�ɹ����)) as �ɹ����,sum(A.�����) as �����,sum(A.������) As ������ ,sum(A.���) As ���" & _
            " From ( " & _
            "     Select 'A'||id as RiD," & _
            "         ���� as ����,  " & _
            "         Decode(����,1,'�⹺',2,'����',3,'Э��',4,'���',5,'���',6,'�ƿ�',7,'����',8,'����',9,'����',10,'��ҩ',11,'����',12,'�̵�',13,'����')||No as no,  " & _
            "         ������� ,  " & _
            "         ��ҩ��λID as ��ҩ��λID  ," & _
            "         �ⷿid as �ⷿid ,���ϵ��*Decode(���,null,0,���) As ���," & _
            "         (���ϵ��* Decode(����,5,0,13,0,1)*�ɱ����) as �ɹ����, " & _
            "         Decode(���ϵ��,-1,0,1)*���۽�� as �����, " & _
            "         Decode(���ϵ��,1,0,1)*���۽�� as  ������ " & _
            "     From ҩƷ�շ���¼ A, ҩƷĿ¼ X, ҩƷ��Ϣ Y " & _
            "     Where A.ҩƷid = X.ҩƷid AND X.ҩ��id = Y.ҩ��id And Y.���ʷ��� In (" & Str���� & ") " & _
            "         And ������� <=to_date('" & strEndDate & "','yyyymmdd')+1" & "And ������� >=to_date('" & strStartDate & "','yyyymmdd') " & _
            "         And ����� Is Not Null " & IIf(inDeptId = 0, "", " And �ⷿID = " & inDeptId) & " And Mod(��¼״̬,3)=1 " & _
            "     ) A,���ű� B,ҩƷ��Ӧ�� C " & _
            " Where A.��ҩ��λid=C.id(+) And A.�ⷿid=B.id(+) " & _
            " Group by A.�������,A.����,A.NO,C.���� " & _
            " having sum(A.�ɹ����)<>0 or sum(A.�����) <>0 or sum(A.������) <>0 order by A.�������"
            
            'Decode(����,5,0,13,0,1)*�ɱ���� as �ɹ����
    DataRecordSet.Close
    With DataRecordSet
        Call SQLTest(App.Title, Me.Caption, StrSQL)
        .Open StrSQL, gcnOracle
        Call SQLTest
'        If .RecordCount = 0 Then
'            StopFlash
'            MsgBox "�ڴ�������Ȩ�޷�Χ��,���κ����ʼ�¼��", vbInformation, gstrSysName
'            Exit Function
'        End If
        dblInMny = 0
        DblCgMny = 0
        dblOutMny = 0
        dblCurrMny = Dbl�ڳ����
        dblCurr��� = Dbl�ڳ����
        Dim colWidth As Long
        Me.fgdData.Rows = .RecordCount + 3
        lngRow = 2
        colWidth = 0
        Me.fgdData.Redraw = False
        
        Call RefreshGridColWidth(Me.fgdData, 0)
        If .RecordCount <> 0 Then
            .MoveFirst
            Do While Not .EOF
             dblInMny = dblInMny + IIf(IsNull(.Fields("�����").Value), 0, .Fields("�����").Value)
             DblCgMny = DblCgMny + IIf(IsNull(.Fields("�ɹ����").Value), 0, .Fields("�ɹ����").Value)
             dblOutMny = dblOutMny + IIf(IsNull(.Fields("������").Value), 0, .Fields("������").Value)
             dblCurrMny = dblCurrMny + IIf(IsNull(.Fields("�����").Value), 0, .Fields("�����").Value) - IIf(IsNull(.Fields("������").Value), 0, .Fields("������").Value)
             dblCurr��� = dblCurr��� + IIf(IsNull(.Fields("���").Value), 0, .Fields("���").Value)
             Me.fgdData.TextMatrix(lngRow, 0) = IIf(Format(.Fields("�������").Value, "yyyy-mm-dd") = "1932-09-09", dtpStartDate, Format(.Fields("�������").Value, "yyyy-mm-dd")) & IIf(lngRow Mod 2 = 0, "", " ")
             Me.fgdData.TextMatrix(lngRow, 1) = IIf(IsNull(.Fields("no").Value), "", IIf(Format(.Fields("�������").Value, "yyyy-mm-dd") = "1932-09-09", "", .Fields("no").Value)) & IIf(lngRow Mod 2 = 0, "", " ")
             Me.fgdData.TextMatrix(lngRow, 2) = IIf(Format(.Fields("�������").Value, "yyyy-mm-dd") = "1932-09-09", "�ڳ�������", IIf(IsNull(.Fields("ժҪ").Value), "", .Fields("ժҪ").Value)) & IIf(lngRow Mod 2 = 0, "", " ")
             Me.fgdData.TextMatrix(lngRow, 3) = " " & Format(.Fields("�ɹ����").Value, "##,###0.00;-##,###0.00; ; ")
             Me.fgdData.TextMatrix(lngRow, 4) = Format(.Fields("�����").Value, "##,###0.00;-##,###0.00; ; ")
             Me.fgdData.TextMatrix(lngRow, 5) = " " & Format(.Fields("������").Value, "##,###0.00;-##,###0.00; ; ")
             Me.fgdData.TextMatrix(lngRow, 6) = Format(dblCurrMny, "##,###0.00;-##,###0.00; ; ")
             
             If mnuViewBlc.Checked Then Me.fgdData.TextMatrix(lngRow, 7) = Format(dblCurr���, "##,###0.00;-##,###0.00; ; ")
             
             Me.fgdData.RowData(lngRow) = IIf(IsNull(.Fields("����").Value), 999999, .Fields("����").Value)
             Call RefreshGridColWidth(Me.fgdData, lngRow)
             lngRow = lngRow + 1
             .MoveNext
           Loop
        End If
            Me.fgdData.MergeRow(1) = True
            Me.fgdData.TextMatrix(1, 0) = "�ڳ�"
            Me.fgdData.TextMatrix(1, 1) = "�ڳ�"
            Me.fgdData.TextMatrix(1, 2) = "�ڳ�"
            Me.fgdData.TextMatrix(1, 3) = ""
            Me.fgdData.TextMatrix(1, 4) = " "
            Me.fgdData.TextMatrix(1, 5) = ""
            Me.fgdData.TextMatrix(1, 6) = Format(Dbl�ڳ����, "##,###0.00;-##,###0.00; ; ")
            If mnuViewBlc.Checked Then Me.fgdData.TextMatrix(1, 7) = Format(Dbl�ڳ����, "##,###0.00;-##,###0.00; ; ")
         
            
'        If dblInMny <> 0 Or DblCgMny <> 0 Or dblOutMny <> 0 Then
            Me.fgdData.MergeRow(Me.fgdData.Rows - 1) = True
            Me.fgdData.RowData(lngRow) = 999999
            Me.fgdData.TextMatrix(Me.fgdData.Rows - 1, 0) = "�ϼ�"
            Me.fgdData.TextMatrix(Me.fgdData.Rows - 1, 1) = "�ϼ�"
            Me.fgdData.TextMatrix(Me.fgdData.Rows - 1, 2) = "�ϼ�"
            Me.fgdData.TextMatrix(Me.fgdData.Rows - 1, 3) = " " & Format(DblCgMny, "##,###0.00;-##,###0.00; ; ")
            Me.fgdData.TextMatrix(Me.fgdData.Rows - 1, 4) = Format(dblInMny, "##,###0.00;-##,###0.00; ; ")
            Me.fgdData.TextMatrix(Me.fgdData.Rows - 1, 5) = " " & Format(dblOutMny, "##,###0.00;-##,###0.00; ; ")
            Me.fgdData.TextMatrix(Me.fgdData.Rows - 1, 6) = Format(dblCurrMny, "##,###0.00;-##,###0.00; ; ")
            If mnuViewBlc.Checked Then Me.fgdData.TextMatrix(Me.fgdData.Rows - 1, 7) = Format(dblCurr���, "##,###0.00;-##,###0.00; ; ")
            Call RefreshGridColWidth(Me.fgdData, lngRow)
'        End If
        
        Me.fgdData.Redraw = True
    End With
    lbl�ⷿ.Caption = "�ⷿ��" & InDeptName
    lbl�ڼ�.Caption = "�ڼ�:" & dtpStartDate & "  ��  " & dtpEndDate
    StopFlash
    RefreshData = True
Exit Function
Err:
    StopFlash
    RefreshData = False
    Me.fgdData.Redraw = True
    MsgBox "�ڻ�ȡҩƷ����ʱ,�����˲���Ԥ֪�Ĵ���!", vbInformation, gstrSysName
End Function

Private Sub ReFreshStru()
    '-------------------------------------------------------------------------
    '--����:���»�ı�ͷ�ṹ
    '--����:
    '--����:
    '-------------------------------------------------------------------------
    Dim IntCol As Long
    Me.Caption = "ҩƷ����"
    Me.lblTitle.Caption = GetUnitName & "ҩƷ����"
    With Me.fgdData
            .Redraw = False
             .Clear
             .Cols = 7
             If mnuViewBlc.Checked Then .Cols = 8
             .Rows = 3
             .MergeCells = flexMergeRestrictRows
             For IntCol = 0 To .Rows - 1
                .MergeRow(IntCol) = False
                .CellAlignment = 1
             Next
            For IntCol = 0 To .Cols - 1
                .FixedAlignment(IntCol) = 4
            Next
            .ColAlignment(0) = 1
            .ColAlignment(1) = 1
            .ColAlignment(2) = 1
            .ColAlignment(3) = 7
            .ColAlignment(4) = 7
            .ColAlignment(5) = 7
            .ColAlignment(6) = 7
            If mnuViewBlc.Checked Then .ColAlignment(7) = 7
            
            .colWidth(0) = 400
            .colWidth(1) = 600
            .colWidth(2) = 400
            .colWidth(3) = 800
            .colWidth(4) = 800
            .colWidth(5) = 800
            .colWidth(6) = 800
            If mnuViewBlc.Checked Then .colWidth(7) = 800
            
            .TextMatrix(0, 0) = "����"
            .TextMatrix(0, 1) = "���ݺ�"
            .TextMatrix(0, 2) = "ժҪ"
            .TextMatrix(0, 3) = "�ɹ����"
            .TextMatrix(0, 4) = "�����"
            .TextMatrix(0, 5) = "������"
            .TextMatrix(0, 6) = "�����"
            If mnuViewBlc.Checked Then .TextMatrix(0, 7) = "���"
            .Redraw = True
    End With
End Sub
