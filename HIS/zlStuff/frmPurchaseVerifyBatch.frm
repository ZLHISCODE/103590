VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmPurchaseVerifyBatch 
   BackColor       =   &H8000000A&
   Caption         =   "���������������"
   ClientHeight    =   6120
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11760
   Icon            =   "frmPurchaseVerifyBatch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   11760
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picSelPatient 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   4440
      ScaleHeight     =   2055
      ScaleWidth      =   4995
      TabIndex        =   14
      Top             =   960
      Width           =   4995
      Begin VB.PictureBox picTitlePatient 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   120
         ScaleHeight     =   405
         ScaleWidth      =   3015
         TabIndex        =   18
         Top             =   0
         Width           =   3015
         Begin VB.Label lblSelPatient 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            Caption         =   "��ѡ��Ĳ����б�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   120
            TabIndex        =   19
            Top             =   100
            Width           =   1560
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfPatient 
         Height          =   945
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   4545
         _cx             =   8017
         _cy             =   1667
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
         BackColor       =   -2147483628
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   14
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483628
         GridColor       =   12632256
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483628
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPurchaseVerifyBatch.frx":014A
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
         ExplorerBar     =   7
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   0   'False
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
      Begin VB.Label lblCostAmount 
         AutoSize        =   -1  'True
         Caption         =   "�ϼƳɱ���"
         Height          =   180
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   1260
      End
      Begin VB.Label lblIVAmount 
         AutoSize        =   -1  'True
         Caption         =   "�ϼƷ�Ʊ��"
         Height          =   180
         Left            =   2760
         TabIndex        =   17
         Top             =   1680
         Width           =   1260
      End
   End
   Begin VB.PictureBox picSelMaterial 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   4560
      ScaleHeight     =   2055
      ScaleWidth      =   4995
      TabIndex        =   10
      Top             =   3600
      Width           =   4995
      Begin VB.PictureBox picTitleMaterial 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   120
         ScaleHeight     =   405
         ScaleWidth      =   4815
         TabIndex        =   20
         Top             =   0
         Width           =   4815
         Begin VB.TextBox txtFindMaterial 
            Height          =   270
            Left            =   2880
            TabIndex        =   21
            Top             =   70
            Width           =   1815
         End
         Begin VB.Label lblFindMaterial 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            Caption         =   "���Ҳ���(&M)"
            Height          =   180
            Left            =   1800
            TabIndex        =   23
            Top             =   100
            Width           =   990
         End
         Begin VB.Label lblMaterial 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            Caption         =   "����ʹ�ò�����ϸ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   120
            TabIndex        =   22
            Top             =   100
            Width           =   1560
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMaterial 
         Height          =   945
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   4545
         _cx             =   8017
         _cy             =   1667
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
         BackColor       =   -2147483628
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483628
         GridColor       =   12632256
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483628
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPurchaseVerifyBatch.frx":021F
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
         ExplorerBar     =   7
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   0   'False
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
      Begin VB.Label lblIV 
         AutoSize        =   -1  'True
         Caption         =   "С�Ʒ�Ʊ��"
         Height          =   180
         Left            =   2760
         TabIndex        =   13
         Top             =   1680
         Width           =   1260
      End
      Begin VB.Label lblCost 
         AutoSize        =   -1  'True
         Caption         =   "С�Ƴɱ���"
         Height          =   180
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   1260
      End
   End
   Begin VB.PictureBox pic������ 
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   360
      ScaleHeight     =   4455
      ScaleWidth      =   3615
      TabIndex        =   9
      Top             =   600
      Width           =   3615
      Begin VB.CheckBox chkȫѡ 
         Caption         =   "ȫѡ"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtPatientInfo 
         Height          =   270
         Left            =   1200
         TabIndex        =   8
         Top             =   3960
         Width           =   2295
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "ѡ����������"
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   615
         Width           =   690
      End
      Begin MSComctlLib.ListView lvwPatient 
         Height          =   2175
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox txtProvider 
         Height          =   300
         Left            =   885
         TabIndex        =   1
         Top             =   240
         Width           =   2200
      End
      Begin VB.CommandButton cmdProvider 
         Caption         =   "��"
         Height          =   300
         Left            =   3120
         TabIndex        =   2
         Top             =   240
         Width           =   300
      End
      Begin MSComCtl2.DTPicker dtpDateBegin 
         Height          =   315
         Left            =   885
         TabIndex        =   4
         Top             =   600
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   50987011
         CurrentDate     =   36263
      End
      Begin MSComCtl2.DTPicker dtpDateEnd 
         Height          =   315
         Left            =   885
         TabIndex        =   5
         Top             =   960
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   50987011
         CurrentDate     =   36263
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "���Ҳ���(&F)"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   3960
         Width           =   990
      End
      Begin VB.Label lblProvider 
         AutoSize        =   -1  'True
         Caption         =   "��Ӧ��"
         Height          =   180
         Left            =   120
         TabIndex        =   0
         Top             =   270
         Width           =   540
      End
   End
   Begin XtremeCommandBars.CommandBars cmbMain 
      Left            =   480
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPurchaseVerifyBatch.frx":02F4
      Left            =   120
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPurchaseVerifyBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MCON_ģ��� = 1712

Private Enum enm_CommandBarID
    GetData = 3052
    Verify = 8044
    Cancel = 2613
End Enum

Private mlngStockID As Long
Private mFMT As g_FmtString
Private mintUnit As Integer              '0��ɢװ��λ�� 1����װ��λ
Private mbln��Ҫ�˲� As Boolean
Private mstrPrivs As String
Private mintMode As Integer              '0���޺˲黷�ڵ���ˣ� 1���к˲黷�ڵĺ˲飻 2���к˲黷�ڵ����

Private mdblIVAmountOld As Double
Private mdblIVAmountNew As Double

Private mdatBegin As Date
Private mdatEnd As Date
Private mblnUpdate As Boolean                   '�Ƿ��µ��ۼ۸��µ��ݣ���Ҫ��������ʱ���۸��µ����
Public Sub ShowMe(ByVal frmMain As Form, ByVal strPrivs As String, ByVal intMode As Integer, ByVal lngStockID As Long)
    mlngStockID = lngStockID
    mstrPrivs = strPrivs
    mintMode = intMode
    If mintMode = 1 Then
        Caption = "�������������˲�"
    Else
        Caption = "���������������"
    End If
    Show vbModal, frmMain
End Sub

Private Sub InitDKPMain()
'��ʼ��dkpMain
    Dim pneParameter As Pane, pneInvoice As Pane, pneMaterial As Pane
    With dkpMain
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        
        Set pneParameter = .CreatePane(1, ScaleHeight, 250, DockLeftOf)
        pneParameter.Options = PaneNoCloseable + PaneNoHideable + PaneNoFloatable
        pneParameter.Title = "��������"
        pneParameter.MinTrackSize.Width = 150
        pneParameter.MaxTrackSize.Width = 350
        
        Set pneInvoice = .CreatePane(2, 1000, 10, DockRightOf)
        pneInvoice.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        pneInvoice.Title = "ѡ����A"
        pneInvoice.MinTrackSize.Height = 120
        
        Set pneMaterial = .CreatePane(3, 100, 10, DockBottomOf, pneInvoice)
        pneMaterial.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        pneMaterial.Title = "ѡ����B"
        pneMaterial.MinTrackSize.Height = 120
        
        If Not cmbMain Is Nothing Then Call .SetCommandBars(cmbMain)
    End With
End Sub

Private Sub InitCommandBar()
    Dim cbcControl As CommandBarControl
    Dim cbrToolBar As CommandBar

    cmbMain.VisualTheme = xtpThemeOffice2003
    With cmbMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True                 '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize False, 24, 24
        .SetIconSize True, 16, 16
    End With
    cmbMain.EnableCustomization False
'    cmbMain.Icons = frmPubIcons.imgPublic.Icons
    Set cmbMain.Icons = zlCommFun.GetPubIcons

    Set cbrToolBar = cmbMain.Add("������", xtpBarTop)
    'cbrToolBar.ShowTextBelowIcons = True
    cbrToolBar.EnableDocking xtpFlagAlignTop
    With cbrToolBar.Controls
        Set cbcControl = .Add(xtpControlButton, enm_CommandBarID.GetData, "��ȡ")
        If mintMode = 1 Then
            Set cbcControl = .Add(xtpControlButton, enm_CommandBarID.Verify, "�˲�")
        Else
            Set cbcControl = .Add(xtpControlButton, enm_CommandBarID.Verify, "���")
        End If
        Set cbcControl = .Add(xtpControlButton, enm_CommandBarID.Cancel, "�ر�")
        cbcControl.BeginGroup = True
    End With
    For Each cbcControl In cbrToolBar.Controls
        If cbcControl.Type = xtpControlButton Then
            cbcControl.Style = xtpButtonIconAndCaption
        End If
    Next
End Sub

Private Sub chkDate_Click()
    dtpDateBegin.Enabled = chkDate.Value = 1
    dtpDateEnd.Enabled = chkDate.Value = 1
End Sub

Private Sub chkDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub chkȫѡ_Click()
    Dim i As Integer
    
    vsfPatient.Rows = 1
    vsfMaterial.Rows = 1
    With lvwPatient
        For i = 1 To .ListItems.Count
            .ListItems.Item(i).Checked = chkȫѡ.Value
            If chkȫѡ.Value = 1 Then
                Call SelPatient(.ListItems(i))
            Else
                Call CalPatient(.ListItems(i))
            End If
        Next
    End With
End Sub

Private Sub cmbMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case enm_CommandBarID.GetData
            Call GetCheckData
        Case enm_CommandBarID.Verify
            Control.Enabled = False
            Call VerifyBatch
            Control.Enabled = True
        Case enm_CommandBarID.Cancel
            Unload Me
    End Select
End Sub

Private Sub GetCheckData()
    If chkDate.Value = 1 And DateDiff("M", dtpDateBegin.Value, dtpDateEnd.Value) > 3 Then
        MsgBox "ѡ���������ڷ�Χ���ܳ��������£�", vbInformation, gstrSysName
        Exit Sub
    End If
    If vsfPatient.Rows > 1 Or vsfMaterial.Rows > 1 Then
        If MsgBox("����ѡ��������δ����Ҫ��������", vbInformation + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    End If
    If Val(txtProvider.Tag) = 0 Then
        MsgBox "��¼�롰��Ӧ�̡���Ϣ��", vbInformation, gstrSysName
        Exit Sub
    End If
    Dim cbcControl As CommandBarControl
    
    Set cbcControl = Me.cmbMain.FindControl(, enm_CommandBarID.GetData)
    If cbcControl Is Nothing Then Exit Sub
    
    cbcControl.Enabled = False
    MousePointer = vbHourglass
    Call FillLVWPatient
    lvwPatient.SetFocus
    MousePointer = vbDefault
    cbcControl.Enabled = True
End Sub

Private Sub cmbMain_Resize()
    On Error Resume Next
    With txtProvider
        .Width = pic������.Width - .Left - cmdProvider.Width - 100
    End With
    With cmdProvider
        .Left = pic������.Width - .Width - 100
    End With
    With dtpDateBegin
        .Width = txtProvider.Width
    End With
    With dtpDateEnd
        .Width = txtProvider.Width
    End With
    With chkȫѡ
        .Left = chkDate.Left
        .Top = dtpDateEnd.Top + dtpDateEnd.Height + 50
    End With
    With lvwPatient
        .Top = chkȫѡ.Top + chkȫѡ.Height + 50
        .Left = 0
        .Width = pic������.Width
        .Height = pic������.Height - dtpDateEnd.Top - dtpDateEnd.Height - txtPatientInfo.Height - chkȫѡ.Height - 200
    End With
    With lblFind
        .Top = pic������.Height - txtPatientInfo.Height - 70
        .Left = lblProvider.Left
    End With
    With txtPatientInfo
        .Top = pic������.Height - txtPatientInfo.Height - 100
        .Left = lblFind.Left + lblFind.Width + 50
        .Width = pic������.Width - lblFind.Left - lblFind.Width - 50
    End With

    With picTitlePatient
        .Top = 0
        .Left = 0
        .Width = picSelPatient.Width
        .Height = 400
    End With
    With vsfPatient
        .Top = picTitlePatient.Height
        .Left = 0
        .Width = picSelPatient.Width
        .Height = picSelPatient.Height - .Top - lblCostAmount.Height - 100 * 2
    End With
    With lblCostAmount
        .Top = picSelPatient.Height - .Height - 100
        .Left = lblSelPatient.Left
    End With
    With lblIVAmount
        .Top = lblCostAmount.Top
        .Left = picSelPatient.Width / 2
    End With

    With picTitleMaterial
        .Top = 0
        .Left = 0
        .Width = picSelPatient.Width
        .Height = 400
    End With
    With lblFindMaterial
        .Left = picSelMaterial.Width - txtFindMaterial.Width - .Width - 100
    End With
    With txtFindMaterial
        .Left = lblFindMaterial.Left + lblFindMaterial.Width + 50
    End With
    With vsfMaterial
        .Top = picTitleMaterial.Height
        .Left = 0
        .Width = picSelMaterial.Width
        .Height = picSelMaterial.Height - .Top - lblCost.Height - 200
    End With
    With lblCost
        .Top = picSelMaterial.Height - .Height - 100
        .Left = lblMaterial.Left
    End With
    With lblIV
        .Top = lblCost.Top
        .Left = picSelMaterial.Width / 2
    End With
    err.Clear: On Error GoTo 0
End Sub

Private Sub cmdProvider_Click()
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim vRect As RECT

    vRect = zlControl.GetControlRect(txtProvider.hwnd)
    
    gstrSQL = "" & _
        "   Select id,�ϼ�ID,����,����,����,ĩ�� " & _
        "   From ��Ӧ�� " & _
        "   Where  (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null) " & _
        "       And (substr(����,5,1)=1 And (վ��=[1] or վ�� is null)  Or Nvl(ĩ��,0)=0) " & _
        "   Start with �ϼ�ID is null connect by prior ID =�ϼ�ID " & _
        "   Order by level,ID "
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 2, "��Ӧ��", True, "", "" _
              , True, True, False, vRect.Left - 15, vRect.Top, txtProvider.Height, blnCancel, False, False, gstrNodeNo)
    If blnCancel = False Then
        If Not rsTmp Is Nothing Then
            txtProvider.Text = zlStr.Nvl(rsTmp!����)
            txtProvider.Tag = zlStr.Nvl(rsTmp!Id)
        Else
            txtProvider.Text = ""
            txtProvider.Tag = "0"
        End If
    End If
    txtProvider.SetFocus
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
        Case 1: Item.Handle = pic������.hwnd
        Case 2: Item.Handle = picSelPatient.hwnd
        Case 3: Item.Handle = picSelMaterial.hwnd
    End Select
End Sub

Private Sub dtpDateBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub dtpDateEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    mblnUpdate = False
    
    dtpDateBegin.Value = DateAdd("M", -1, sys.Currentdate)
    dtpDateEnd.Value = sys.Currentdate
    Call chkDate_Click
    
    mbln��Ҫ�˲� = Val(zlDatabase.GetPara("�����⹺��Ҫ�˲�", glngSys, "0")) = 1
    
    mintUnit = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, MCON_ģ���, "0"))
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
        .FM_ɢװ���ۼ� = GetFmtString(0, g_�ۼ�)
    End With
    
    cmbMain.ActiveMenuBar.Visible = False
    Call InitDKPMain
    Call InitCommandBar
    Call InitLVWPatient
    Call InitVSFPatient
    Call InitVSFMaterial
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����", 0, 0)) = 1 Then
        RestoreWinState Me, App.ProductName, Me.Caption
    End If
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    If Width < 10000 Then Width = 10000
    If Height < 6000 Then Height = 6000
End Sub

Private Sub InitLVWPatient()
    With lvwPatient
        .Checkboxes = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .FullRowSelect = True
        .Sorted = True
        
        .ColumnHeaders.Add , "����", "����", 1500
        .ColumnHeaders.Add , "����", "����", 800
        .ColumnHeaders.Add , "�Ա�", "�Ա�", 600
        .ColumnHeaders.Add , "����", "����", 600
        .ColumnHeaders.Add , "סԺ��", "סԺ��", 800
    End With
End Sub

Private Sub InitVSFPatient()
Const conHead = "H_����ID,,|����,,1000|H_����ID,,|����,,2000|��Ʊ��,,1000|��Ʊ����,,1000|��Ʊ����,,1000,d|��Ʊ���,,1000,N|�ɱ����,,1000,n"
    
    With vsfPatient
        .Rows = 1
        SetVSFHead vsfPatient, conHead
    End With
End Sub

Private Sub InitVSFMaterial()
Const conHead = "H_����ID,,|H_����ID,,|��������,,2000|���,,1000|��Ʊ���,,1000,N|��λ,,500|����,,1000,N|�ɱ���,,1000,N|�ɱ����,,1000,N" & _
                "|H_NO,,|H_���,,|H_�ⷿID,,|H_��ҩ��λID,,|H_����,,|H_����,,|H_��������,,|H_Ч��,,|H_�������,,|H_���Ч��,," & _
                "|H_����,,|H_���ۼ�,,|H_���۽��,,|H_���,,|H_��۽��,,|H_ժҪ,,|H_�ڲ�����,,|H_����ϵ��,," & _
                "|H_ע��֤��,,|H_������,,|H_��������,,|H_��Ʊ��,,|H_��Ʊ����,,|H_��Ʊ���,,|H_�˲���,,|H_�˲�����,,|H_����,," & _
                "|H_��ֵ����,,|H_��Ʒ����,,|H_����ID,,|H_Verify,,|H_����ID,,"
                
    With vsfMaterial
        .Rows = 1
        .ExplorerBar = flexExMove
        SetVSFHead vsfMaterial, conHead
    End With
End Sub

Private Sub FillLVWPatient()
    Dim rsTmp As ADODB.Recordset
    Dim lsItem As ListItem
    Dim i As Long
    Dim lngColor As Long
    
    On Error GoTo ErrHandle
    MousePointer = vbHourglass
    '���
    If mbln��Ҫ�˲� Then
        If mintMode = 1 Then
            '�˲�
            gstrSQL = "Select Distinct a.����ID, a.����, a.�Ա�, b.����, a.סԺ��, b.���˿���ID, d.���� ���� " & _
                      "From ������Ϣ A, סԺ���ü�¼ B, ҩƷ�շ���¼ C, ���ű� D " & _
                      "Where a.����id = b.����id And b.Id = c.����id And b.���˿���id = d.Id " & _
                      "    And c.����ID > 0 And c.��ҩ���� is null And c.��ҩ��λID + 0 = [1] And c.�ⷿid = [2] " & _
                      "    And c.���� = 15 "
        Else
            '���
            gstrSQL = "Select Distinct a.����ID, a.����, a.�Ա�, b.����, a.סԺ��, b.���˿���ID, d.���� ���� " & _
                      "From ������Ϣ A, סԺ���ü�¼ B, ҩƷ�շ���¼ C, ���ű� D " & _
                      "Where a.����id = b.����id And b.Id = c.����id And b.���˿���id = d.Id " & _
                      "    And c.����ID > 0 And c.��ҩ���� is not null And c.������� is null " & _
                      "    And c.��ҩ��λID + 0 = [1] And c.�ⷿid = [2] And c.���� = 15 "
        End If
    Else
        'ֱ�����
        gstrSQL = "Select Distinct a.����ID, a.����, a.�Ա�, b.����, a.סԺ��, b.���˿���ID, d.���� ���� " & _
                  "From ������Ϣ A, סԺ���ü�¼ B, ҩƷ�շ���¼ C, ���ű� D " & _
                  "Where a.����id = b.����id And b.Id = c.����id And b.���˿���id = d.Id " & _
                  "    And c.����ID > 0 And c.������� is null And c.��ҩ��λID + 0 = [1] And c.�ⷿid = [2] " & _
                  "    And c.���� = 15 "
    End If
    
    If chkDate.Value = 1 Then
        mdatBegin = dtpDateBegin.Value
        mdatEnd = dtpDateEnd.Value
    Else
        mdatBegin = DateAdd("M", -1, sys.Currentdate)
        mdatEnd = sys.Currentdate
    End If
    gstrSQL = gstrSQL & " And c.�������� between to_date('" & Format(mdatBegin, "yyyy-mm-dd 00:00:00") & "', 'yyyy-mm-dd hh24:mi:ss') " & _
              " And to_date('" & Format(mdatEnd, "yyyy-mm-dd 23:59:59") & "', 'yyyy-mm-dd hh24:mi:ss') " & vbNewLine
    '�ϲ���������
    gstrSQL = gstrSQL & " Union All " & Replace(gstrSQL, "סԺ���ü�¼", "������ü�¼") & " Order By ����, ���� "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Caption & "-������Ϣ", Val(txtProvider.Tag), mlngStockID)
    
    With vsfPatient
        For i = .Rows - 1 To 1 Step -1
            .RemoveItem i
        Next
    End With
    With vsfMaterial
        For i = .Rows - 1 To 1 Step -1
            .RemoveItem i
        Next
    End With
    
    MousePointer = vbDefault
    
    With lvwPatient
        .ListItems.Clear
        .Tag = txtProvider.Tag
        Do While Not rsTmp.EOF
            Set lsItem = .ListItems.Add(, "_" & zlStr.Nvl(rsTmp!����ID) & "_" & zlStr.Nvl(rsTmp!���˿���ID), zlStr.Nvl(rsTmp!����))
            lsItem.SubItems(1) = zlStr.Nvl(rsTmp!����)
            lsItem.SubItems(2) = zlStr.Nvl(rsTmp!�Ա�)
            lsItem.SubItems(3) = zlStr.Nvl(rsTmp!����)
            lsItem.SubItems(4) = zlStr.Nvl(rsTmp!סԺ��)
            rsTmp.MoveNext
        Loop
        If rsTmp.RecordCount <= 0 Then
            MsgBox "�޿ɡ�" & IIf(mintMode = 1, "�˲�", "���") & "�������ݣ�", vbInformation, gstrSysName
        End If
    End With
    
'    If mintMode <> 1 Then
        lngColor = vbBlue
'    Else
'        lngColor = vbBlack
'    End If
    With vsfPatient
        .Cell(flexcpForeColor, 0, .ColIndex("��Ʊ��"), .Rows - 1, .ColIndex("��Ʊ��")) = lngColor
        .Cell(flexcpForeColor, 0, .ColIndex("��Ʊ����"), .Rows - 1, .ColIndex("��Ʊ����")) = lngColor
        .Cell(flexcpForeColor, 0, .ColIndex("��Ʊ����"), .Rows - 1, .ColIndex("��Ʊ����")) = lngColor
        .Cell(flexcpForeColor, 0, .ColIndex("��Ʊ���"), .Rows - 1, .ColIndex("��Ʊ���")) = lngColor
    End With
    With vsfMaterial
        .Cell(flexcpForeColor, 0, .ColIndex("��Ʊ���"), .Rows - 1, .ColIndex("��Ʊ���")) = lngColor
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����", 0, 0)) = 1 Then
        SaveWinState Me, App.ProductName, Me.Caption
    End If
End Sub

Private Sub lvwPatient_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ColumnHeader.Index >= 1 And ColumnHeader.Index <= 2 Or ColumnHeader.Index = 5 Then
        lvwPatient.SortKey = ColumnHeader.Index - 1
    End If
End Sub

Private Sub lvwPatient_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        'ȷ��ѡ��
        SelPatient Item
    Else
        'ȡ��ѡ��
        If CalPatient(Item) = False Then Item.Checked = True
    End If
    Call ShowAmount
End Sub

Private Sub txtFindMaterial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
        If Trim(txtFindMaterial.Text) = "" Then Exit Sub
        Dim i As Long, lngStart As Long
        With vsfMaterial
            lngStart = IIf(KeyCode = vbKeyF3, IIf(.Rows - 1 > .Row, .Row + 1, 1), 1)
            For i = lngStart To .Rows - 1
                If InStr(.TextMatrix(i, .ColIndex("��������")), Trim(txtFindMaterial.Text)) > 0 Then
                    .Row = i
                    Exit For
                End If
            Next
        End With
    End If
End Sub

Private Sub txtPatientInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
        If Trim(txtPatientInfo.Text) = "" Then Exit Sub
        Dim i As Long, lngStart As Long
        With lvwPatient
            lngStart = IIf(KeyCode = vbKeyF3, IIf(.ListItems.Count > .SelectedItem.Index, .SelectedItem.Index + 1, 1), 1)
            For i = lngStart To .ListItems.Count
                If InStr(.ListItems(i).SubItems(1), Trim(txtPatientInfo.Text)) > 0 Then
                    .ListItems.Item(i).Selected = True
                    Exit For
                End If
            Next
        End With
    End If
End Sub

Private Sub txtProvider_Change()
    With txtProvider
        .Text = UCase(.Text)
        .SelStart = Len(.Text)
    End With
End Sub

Private Sub txtProvider_GotFocus()
    txtProvider.SelStart = 0
    txtProvider.SelLength = Len(txtProvider.Text)
End Sub

Private Sub txtProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strProviderText As String
    Dim rsProvider As Recordset
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    vRect = zlControl.GetControlRect(txtProvider.hwnd)
    
    With txtProvider
        If Trim(.Text) = "" Then Exit Sub
        strProviderText = GetMatchingSting(UCase(.Text))
        
        gstrSQL = "" & _
            "   Select id,����,����,���� " & _
            "   From ��Ӧ�� " & _
            "   Where (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null) " & _
            "       And (վ��=[2] or վ�� is null) And ĩ��=1 And (substr(����,5,1) = 1 ) " & _
            "       And (���� like [1] Or ���� like [1] or upper(����) like [1]) "
        Set rsProvider = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "��Ӧ��", False, "", "", False, False, True, _
                            vRect.Left, vRect.Top, txtProvider.Height, blnCancel, False, False, _
                            strProviderText, gstrNodeNo)
        If Not rsProvider Is Nothing Then
            txtProvider.Text = zlStr.Nvl(rsProvider!����)
            txtProvider.Tag = zlStr.Nvl(rsProvider!Id)
            chkDate.SetFocus
        Else
            txtProvider.Text = ""
            txtProvider.Tag = "0"
        End If
    End With
End Sub

Private Sub txtProvider_LostFocus()
    If txtProvider.Text = "" Then
        txtProvider.Tag = "0"
        Exit Sub
    End If
End Sub

Private Sub txtProvider_Validate(Cancel As Boolean)
    If txtProvider.Text = "" Then
        txtProvider.Tag = "0"
        Exit Sub
    End If
    
    If CheckQualifications(MCON_ģ���, 2, Val(txtProvider.Tag)) = False Then
        txtProvider.Text = ""
        txtProvider.Tag = "0"
        Exit Sub
    End If
    
End Sub

Private Sub SetVSFHead(ByVal vsfObject As VSFlexGrid, ByVal strHead As String)
    Dim arrCols As Variant, arrRows As Variant
    Dim i As Integer
    
    arrRows = Split(strHead, "|")
    With vsfObject
        If .Rows = 0 Then .Rows = 1
        .Cols = UBound(arrRows) + 1
        For i = LBound(arrRows) To UBound(arrRows)
            If arrRows(i) = "" Then
                .TextMatrix(0, i) = ""
            Else
                arrCols = Split(arrRows(i), ",")
                '��1Ԫ�أ���ʾֵ
                .TextMatrix(0, i) = arrCols(0)
                '��2Ԫ�أ�Keyֵ
                If arrCols(1) = "" Then
                    If Left(arrCols(0), 2) = "H_" Then
                        .ColKey(i) = Mid(arrCols(0), 3, Len(arrCols(0)))
                    Else
                        .ColKey(i) = arrCols(0)
                    End If
                Else
                    .ColKey(i) = arrCols(1)
                End If
                '��3Ԫ�أ����
                .ColWidth(i) = Val(arrCols(2))
                'H_Ϊ������
                If Left(arrCols(0), 2) = "H_" Then
                    .ColHidden(i) = True
                Else
                    .ColHidden(i) = False
                    '��4Ԫ�أ���ʾ��ʽ
                    If UBound(arrCols) > 2 Then
                        If UCase(arrCols(3)) = "D" Then
                            .ColFormat(i) = "yyyy-mm-dd"
                            .ColAlignment(i) = flexAlignCenterCenter
                        ElseIf UCase(arrCols(3)) = "T" Then
                            .ColFormat(i) = "hh:mi:ss"
                            .ColAlignment(i) = flexAlignCenterCenter
                        ElseIf UCase(arrCols(3)) = "DT" Then
                            .ColFormat(i) = "yyyy-mm-dd hh:mi:ss"
                            .ColAlignment(i) = flexAlignCenterCenter
                        ElseIf UCase(arrCols(3)) = "N" Then
                            .ColAlignment(i) = flexAlignRightCenter
                        Else
                            .ColAlignment(i) = flexAlignLeftCenter
                        End If
                    Else
                        .ColAlignment(i) = flexAlignLeftCenter
                    End If
                End If
            End If
            .Cell(flexcpAlignment, 0, i) = flexAlignCenterCenter
        Next
    End With
End Sub

Private Sub SelPatient(ByVal lsItem As ListItem)
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, lngCurRow As Long
    Dim dblCostAmount As Double
    Dim dbl�ɱ���� As Double
    Dim lng���˿���ID As Long
    Dim blnInfo As Boolean
    Dim strIVNO As String, strIVCode As String, strIVDate As String
        
    lng���˿���ID = Val(Mid(lsItem.Key, InStr(2, lsItem.Key, "_") + 1))
    On Error GoTo ErrHandle
    
    '��� vsfMaterial
    gstrSQL = _
        "   SELECT distinct a.ҩƷid ����id, A.NO, A.���, ('[' || D.���� || ']' || D.����) AS ������Ϣ, D.���, D.���� as ԭ����, A.����," & _
        "          A.����, Nvl(A.����,0) ����, to_char(A.��������,'yyyy-mm-dd') ��������, A.Ч��, A.�������, A.���Ч��, " & _
        IIf(mintUnit = 1, "ltrim(rtrim(to_char(A.�ɱ��� * c.����ϵ��, " & gOraFmt_Max.FM_�ɱ��� & "))) as �ɱ���, ", "A.�ɱ���, ") & _
        IIf(mintUnit = 1, "ltrim(rtrim(to_char(A.ʵ������ / c.����ϵ��, " & gOraFmt_Max.FM_���� & "))) as ʵ������, ", "A.ʵ������, ") & _
        "         decode(nvl(A.��ҩ��ʽ,0),1,-1,1) * A.�ɱ���� AS ������, Nvl(A.��ҩ��ʽ,0) �˻�, " & _
        "         DECODE(A.����, NULL, 0, A.����) AS ����, " & _
        IIf(mintUnit = 1, "ltrim(rtrim(to_char(A.���ۼ� * c.����ϵ��, " & gOraFmt_Max.FM_���ۼ� & "))) as ���ۼ�, ", "A.���ۼ�, ") & _
        "         decode(nvl(A.��ҩ��ʽ,0),1,-1,1)*A.���۽�� as ���۽��, decode(nvl(A.��ҩ��ʽ,0),1,-1,1)* A.��� ���, " & _
        "         decode(nvl(A.��ҩ��ʽ,0),1,-1,1)*to_number(A.�÷�," & gOraFmt_Max.FM_��� & " )  as ���۲��, " & _
        "         a.��ҩ��λid,a.ע��֤��,a.��Ʒ����, a.ժҪ,A.������,A.��������,A.��ҩ�� as �˲���,A.��ҩ���� as �˲�����," & _
        "         a.�ⷿid, a.�ڲ�����, a.����id, b.���˿���ID, " & _
        IIf(mintMode = 1, "'' ��Ʊ��, '' ��Ʊ����, '' ��Ʊ����, '' ��Ʊ���, ", "f.��Ʊ��,f.��Ʊ����, f.��Ʊ����, f.��Ʊ���, ") & _
        IIf(mintUnit = 1, " c.����ϵ�� as ����ϵ��, ", "1 as ����ϵ��,") & _
        IIf(mintUnit = 1, " c.��װ��λ as ��λ, ", " d.���㵥λ as ��λ, ") & _
        "         decode(E.�շ�ID, null, '', E.���� || ',' || nvl(E.��������,'') || ',' || nvl(E.סԺ��,'') || ',' || nvl(E.����,'') ) AS ��ֵ���� " & _
        "       FROM ҩƷ�շ���¼ A, סԺ���ü�¼ B, �������� C, �շ���ĿĿ¼ D, �շ���¼������Ϣ E, Ӧ����¼ F " & _
        "       Where A.����ID = B.ID And a.ҩƷID = c.����ID And a.ҩƷid = d.ID And A.ID = E.�շ�ID(+) And a.id=f.�շ�id(+) And A.����ID > 0 And " & _
        "         a.��ҩ��λid + 0 = [1] and a.�ⷿid = [2] AND A.��¼״̬ = 1 And A.���� = 15 AND B.����ID + 0 = [3] And B.���˿���ID = [4] And " & _
        "         a.�������� between to_date('" & Format(mdatBegin, "yyyy-mm-dd 00:00:00") & "', 'yyyy-mm-dd hh24:mi:ss') And " & _
        "           to_date('" & Format(mdatEnd, "yyyy-mm-dd 23:59:59") & "', 'yyyy-mm-dd hh24:mi:ss') "
    If mbln��Ҫ�˲� Then
        If mintMode = 1 Then
            gstrSQL = gstrSQL & " And A.������� is null And A.��ҩ���� is null "
        Else
            gstrSQL = gstrSQL & " And A.������� is null And A.��ҩ���� is not null "
        End If
    Else
        gstrSQL = gstrSQL & " And A.������� is null "
    End If
    gstrSQL = gstrSQL & " Union All " & Replace(gstrSQL, "סԺ���ü�¼", "������ü�¼") & " ORDER BY NO, ��� "
    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "���˲�����ϸ", Val(lvwPatient.Tag), mlngStockID, Val(Mid(lsItem.Key, 2)), lng���˿���ID)
    With vsfMaterial
        blnInfo = rsTmp.RecordCount > 0
        dblCostAmount = 0
        Do While Not rsTmp.EOF
            .Rows = .Rows + 1
            lngCurRow = .Rows - 1
            .TextMatrix(lngCurRow, .ColIndex("����ID")) = Val(Mid(lsItem.Key, 2))
            .TextMatrix(lngCurRow, .ColIndex("����ID")) = lng���˿���ID
            .TextMatrix(lngCurRow, .ColIndex("����ID")) = zlStr.Nvl(rsTmp!����ID)
            .TextMatrix(lngCurRow, .ColIndex("��������")) = zlStr.Nvl(rsTmp!������Ϣ)
            .TextMatrix(lngCurRow, .ColIndex("���")) = zlStr.Nvl(rsTmp!���)
            .TextMatrix(lngCurRow, .ColIndex("����")) = zlStr.Nvl(rsTmp!ʵ������)
            .TextMatrix(lngCurRow, .ColIndex("�ɱ���")) = zlStr.Nvl(rsTmp!�ɱ���)
            .TextMatrix(lngCurRow, .ColIndex("�ɱ����")) = zlStr.Nvl(rsTmp!������)
            .TextMatrix(lngCurRow, .ColIndex("��Ʊ���")) = IIf(mintMode = 1, zlStr.Nvl(rsTmp!������), zlStr.Nvl(rsTmp!��Ʊ���))
            .TextMatrix(lngCurRow, .ColIndex("NO")) = zlStr.Nvl(rsTmp!NO)
            .TextMatrix(lngCurRow, .ColIndex("���")) = zlStr.Nvl(rsTmp!���)
            .TextMatrix(lngCurRow, .ColIndex("�ⷿID")) = zlStr.Nvl(rsTmp!�ⷿID)
            .TextMatrix(lngCurRow, .ColIndex("��ҩ��λID")) = zlStr.Nvl(rsTmp!��ҩ��λID)
            .TextMatrix(lngCurRow, .ColIndex("����")) = zlStr.Nvl(rsTmp!����)
            .TextMatrix(lngCurRow, .ColIndex("����")) = zlStr.Nvl(rsTmp!����)
            .TextMatrix(lngCurRow, .ColIndex("����ϵ��")) = zlStr.Nvl(rsTmp!����ϵ��)
            .TextMatrix(lngCurRow, .ColIndex("��λ")) = zlStr.Nvl(rsTmp!��λ)
            .TextMatrix(lngCurRow, .ColIndex("��������")) = zlStr.Nvl(rsTmp!��������)
            .TextMatrix(lngCurRow, .ColIndex("Ч��")) = zlStr.Nvl(rsTmp!Ч��)
            .TextMatrix(lngCurRow, .ColIndex("�������")) = zlStr.Nvl(rsTmp!�������)
            .TextMatrix(lngCurRow, .ColIndex("���Ч��")) = zlStr.Nvl(rsTmp!���Ч��)
            .TextMatrix(lngCurRow, .ColIndex("����")) = zlStr.Nvl(rsTmp!����)
            .TextMatrix(lngCurRow, .ColIndex("���ۼ�")) = zlStr.Nvl(rsTmp!���ۼ�)
            .TextMatrix(lngCurRow, .ColIndex("���۽��")) = zlStr.Nvl(rsTmp!���۽��)
            .TextMatrix(lngCurRow, .ColIndex("���")) = zlStr.Nvl(rsTmp!���)
            .TextMatrix(lngCurRow, .ColIndex("��۽��")) = zlStr.Nvl(rsTmp!���۲��)
            .TextMatrix(lngCurRow, .ColIndex("ժҪ")) = zlStr.Nvl(rsTmp!ժҪ)
            .TextMatrix(lngCurRow, .ColIndex("ע��֤��")) = zlStr.Nvl(rsTmp!ע��֤��)
            .TextMatrix(lngCurRow, .ColIndex("������")) = zlStr.Nvl(rsTmp!������)
            .TextMatrix(lngCurRow, .ColIndex("��������")) = zlStr.Nvl(rsTmp!��������)
            .TextMatrix(lngCurRow, .ColIndex("�˲���")) = zlStr.Nvl(rsTmp!�˲���)
            .TextMatrix(lngCurRow, .ColIndex("�˲�����")) = zlStr.Nvl(rsTmp!�˲�����)
            .TextMatrix(lngCurRow, .ColIndex("����")) = zlStr.Nvl(rsTmp!����)
            .TextMatrix(lngCurRow, .ColIndex("��ֵ����")) = zlStr.Nvl(rsTmp!��ֵ����)
            .TextMatrix(lngCurRow, .ColIndex("��Ʒ����")) = zlStr.Nvl(rsTmp!��Ʒ����)
            .TextMatrix(lngCurRow, .ColIndex("�ڲ�����")) = zlStr.Nvl(rsTmp!�ڲ�����)
            .TextMatrix(lngCurRow, .ColIndex("����ID")) = zlStr.Nvl(rsTmp!����ID)
            If mintMode = 1 Then
                dblCostAmount = dblCostAmount + zlStr.Nvl(rsTmp!������, 0)
            Else
                dblCostAmount = dblCostAmount + zlStr.Nvl(rsTmp!��Ʊ���, 0)
            End If
            dbl�ɱ���� = dbl�ɱ���� + Nvl(rsTmp!������)
            If strIVNO = "" Then strIVNO = zlStr.Nvl(rsTmp!��Ʊ��)
            If strIVCode = "" Then strIVCode = zlStr.Nvl(rsTmp!��Ʊ����)
            If strIVDate = "" Then strIVDate = IIf(IsNull(rsTmp!��Ʊ����), "", Format(rsTmp!��Ʊ����, "yyyy-mm-dd"))
            rsTmp.MoveNext
        Loop
        If .Rows > 1 Then
            .Row = 1
        End If
    End With
    rsTmp.Close

    If blnInfo Then
        '��� vsfPatient
        With vsfPatient
            .Rows = .Rows + 1
            lngCurRow = .Rows - 1
            .TextMatrix(lngCurRow, .ColIndex("����ID")) = Val(Mid(lsItem.Key, 2))
            .TextMatrix(lngCurRow, .ColIndex("����ID")) = lng���˿���ID
            .TextMatrix(lngCurRow, .ColIndex("����")) = lsItem.Text
            .TextMatrix(lngCurRow, .ColIndex("����")) = lsItem.SubItems(1)
            .TextMatrix(lngCurRow, .ColIndex("�ɱ����")) = dbl�ɱ����
            .TextMatrix(lngCurRow, .ColIndex("��Ʊ���")) = dblCostAmount
            .TextMatrix(lngCurRow, .ColIndex("��Ʊ��")) = strIVNO
            .TextMatrix(lngCurRow, .ColIndex("��Ʊ����")) = strIVCode
            .TextMatrix(lngCurRow, .ColIndex("��Ʊ����")) = strIVDate
            .Row = lngCurRow
        End With
    End If
    
    zl_VsGridLOSTFOCUS vsfPatient
    zl_VsGridLOSTFOCUS vsfMaterial
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CalPatient(ByVal lsItem As ListItem) As Boolean
    Dim i As Long
    Dim lngPatientID As Long, lngPatientDrugID As Long
    
    '����ID
    lngPatientID = Val(Mid(lsItem.Key, 2))
    '���˿���ID
    lngPatientDrugID = Val(Mid(lsItem.Key, InStr(2, lsItem.Key, "_") + 1))
    
    '����Ƿ�¼�����ݣ��о�ѯ����ʾ
    With vsfPatient
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("����ID"))) = lngPatientID And Val(.TextMatrix(i, .ColIndex("����ID"))) = lngPatientDrugID Then
                If Trim(.TextMatrix(i, .ColIndex("��Ʊ��"))) <> "" Or _
                   Trim(.TextMatrix(i, .ColIndex("��Ʊ����"))) <> "" Or Trim(.TextMatrix(i, .ColIndex("��Ʊ����"))) <> "" Then
                    If MsgBox("�������Ѿ���¼�뷢Ʊ��Ϣ�������Զ������Ҫ������", vbInformation + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then
                        CalPatient = False
                        Exit Function
                    End If
                End If
            End If
        Next
    End With
    
    '����ѡ��������
    With vsfPatient
        For i = .Rows - 1 To 1 Step -1
            If Val(.TextMatrix(i, .ColIndex("����ID"))) = lngPatientID And Val(.TextMatrix(i, .ColIndex("����ID"))) = lngPatientDrugID Then
                .RemoveItem i
                Exit For
            End If
        Next
    End With
    With vsfMaterial
        For i = .Rows - 1 To 1 Step -1
            If Val(.TextMatrix(i, .ColIndex("����ID"))) = lngPatientID And Val(.TextMatrix(i, .ColIndex("����ID"))) = lngPatientDrugID Then
                .RemoveItem i
            End If
        Next
    End With
    
    CalPatient = True
    
End Function

Private Sub vsfMaterial_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    Dim dblTmp As Double
    
    With vsfMaterial
        If .ColIndex("��Ʊ���") = Col Then
            For i = 1 To .Rows - 1
                If .RowHidden(i) = False Then
                    dblTmp = dblTmp + Val(.TextMatrix(i, .ColIndex("��Ʊ���")))
                End If
            Next
            vsfPatient.TextMatrix(vsfPatient.Row, vsfPatient.ColIndex("��Ʊ���")) = Format(dblTmp, mFMT.FM_���)
            Call ShowAmount
        End If
    End With
End Sub

Private Sub vsfMaterial_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If vsfMaterial.Rows > 1 Then Call zl_VsGridRowChange(vsfMaterial, OldRow, NewRow, OldCol, NewCol)
    vsfMaterial.Cell(flexcpBackColor, 0, 0, 0, vsfMaterial.Cols - 1) = &H8000000F
End Sub

Private Sub vsfMaterial_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfMaterial
        If .ColIndex("��Ʊ���") = Col Then
'            If mintMode <> 1 Then
                Cancel = False
'            Else
'                Cancel = True
'            End If
        Else
            Cancel = True
        End If
    End With
End Sub

Private Sub vsfMaterial_GotFocus()
    If vsfMaterial.Rows > 1 Then zl_VsGridGotFocus vsfMaterial
End Sub

Private Sub vsfMaterial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlVsMoveGridCell vsfMaterial
    End If
End Sub

Private Sub vsfMaterial_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = vsfMaterial.ColIndex("��Ʊ���") Then
        VsFlxGridCheckKeyPress vsfMaterial, Row, Col, KeyAscii, m���ʽ
    End If
End Sub

Private Sub vsfMaterial_LostFocus()
    zl_VsGridLOSTFOCUS vsfMaterial
End Sub

Private Sub vsfPatient_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    Dim blnEdit As Boolean
    
    With vsfPatient
        If .ColIndex("��Ʊ��") = Col Then
            .TextMatrix(Row, Col) = UCase(.TextMatrix(Row, Col))
            blnEdit = (.TextMatrix(Row, Col) <> "")
        ElseIf .ColIndex("��Ʊ����") = Col Then
            .TextMatrix(Row, Col) = UCase(.TextMatrix(Row, Col))
            blnEdit = (.TextMatrix(Row, Col) <> "")
        ElseIf .ColIndex("��Ʊ����") = Col Then
            If Not IsDate(Trim(.Text)) And Trim(.Text) <> "" Then
                .Text = ""
                MsgBox "���������ڡ�¼���ʽ����", vbInformation, gstrSysName
                Exit Sub
            End If
            blnEdit = True
        ElseIf .ColIndex("��Ʊ���") = Col Then
            mdblIVAmountNew = Val(.Text)
            ApportionInvoiceAmount mdblIVAmountOld, mdblIVAmountNew
            Call ShowAmount
            blnEdit = False
        End If
        
        If blnEdit Then
            For i = 1 To .Rows - 1
                If .RowHidden(i) = False And Trim(.TextMatrix(i, Col)) = "" And i <> Row Then
                    .TextMatrix(i, Col) = .TextMatrix(Row, Col)
                End If
            Next
        End If
    End With
End Sub

Private Sub vsfPatient_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfPatient
        If .Rows > 1 Then
            Call zl_VsGridRowChange(vsfPatient, OldRow, NewRow, OldCol, NewCol)
            Call FillMaterial(Val(.TextMatrix(NewRow, .ColIndex("����ID"))), Val(.TextMatrix(NewRow, .ColIndex("����ID"))))
            Call ShowAmount
        End If
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = &H8000000F
    End With
    
End Sub

Private Sub vsfPatient_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfPatient
        If .ColIndex("��Ʊ��") = Col Or .ColIndex("��Ʊ����") = Col Or .ColIndex("��Ʊ����") = Col Or .ColIndex("��Ʊ���") = Col Then
'            If mintMode <> 1 Then
                Cancel = False
                mdblIVAmountOld = 0
                If .ColIndex("��Ʊ���") = Col Then
                    mdblIVAmountOld = Val(.Text)
                End If
'            Else
'                Cancel = True
'            End If
        Else
            Cancel = True
        End If
        
        
    End With
End Sub

Private Sub vsfPatient_GotFocus()
    If vsfMaterial.Rows > 1 Then zl_VsGridGotFocus vsfPatient
End Sub

Private Sub vsfPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, lngPatient As Long, lngPatientDrug As Long
        
    If KeyCode = vbKeyReturn Then
        zlVsMoveGridCell vsfPatient
    ElseIf KeyCode = vbKeyDelete Then
        lngPatient = Val(vsfPatient.TextMatrix(vsfPatient.Row, vsfPatient.ColIndex("����ID")))
        lngPatientDrug = Val(vsfPatient.TextMatrix(vsfPatient.Row, vsfPatient.ColIndex("����ID")))
        
        With lvwPatient
            For i = 1 To .ListItems.Count
                If Val(Mid(.ListItems(i).Key, 2)) = lngPatient And Val(Mid(.ListItems(i).Key, InStr(2, .ListItems(i).Key, "_") + 1)) = lngPatientDrug Then
                    .ListItems(i).Checked = False
                    lvwPatient_ItemCheck .ListItems(i)
                End If
            Next
        End With
        With vsfPatient
            If .Rows > 1 Then
                Call FillMaterial(Val(.TextMatrix(.Row, .ColIndex("����ID"))), Val(.TextMatrix(.Row, .ColIndex("����ID"))))
                Call ShowAmount
            End If
        End With
    End If
End Sub

Private Sub vsfPatient_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = vsfPatient.ColIndex("��Ʊ���") Then
        VsFlxGridCheckKeyPress vsfPatient, Row, Col, KeyAscii, m���ʽ
    End If
    
    
    If KeyAscii <> 13 Then
        If Col = vsfPatient.ColIndex("��Ʊ����") Then
            If InStr("0123456789", Chr(KeyAscii)) > 0 Or KeyAscii = 8 Then
                If Len(vsfPatient.EditText) <= 19 Or KeyAscii = 8 Then
                    KeyAscii = KeyAscii
                Else
                    KeyAscii = 0
                End If
            Else
                KeyAscii = 0
            End If
        End If
    End If
    
End Sub

Private Sub vsfPatient_LostFocus()
    zl_VsGridLOSTFOCUS vsfPatient
End Sub

Private Sub ShowAmount()
    Dim dblCostAmount As Double, dblIVAmount As Double
    Dim dblCost As Double, dblIV As Double
    Dim i As Long, crFore1 As Long, crFore2 As Long
    
    With vsfPatient
        For i = 1 To .Rows - 1
            If .RowHidden(i) = False Then
                dblCostAmount = dblCostAmount + Val(.TextMatrix(i, .ColIndex("�ɱ����")))
                dblIVAmount = dblIVAmount + Val(.TextMatrix(i, .ColIndex("��Ʊ���")))
            End If
        Next
    End With
    With vsfMaterial
        For i = 1 To .Rows - 1
            If .RowHidden(i) = False Then
                dblCost = dblCost + Val(.TextMatrix(i, .ColIndex("�ɱ����")))
                dblIV = dblIV + Val(.TextMatrix(i, .ColIndex("��Ʊ���")))
            End If
        Next
    End With
    lblCostAmount.Caption = "�ϼƳɱ���" & Format(dblCostAmount, "###,###,###,##0.000")
    lblIVAmount.Caption = "�ϼƷ�Ʊ��" & Format(dblIVAmount, "###,###,###,##0.000")
    lblCost.Caption = "С�Ƴɱ���" & Format(dblCost, "###,###,###,##0.000")
    lblIV.Caption = "С�Ʒ�Ʊ��" & Format(dblIV, "###,###,###,##0.000")
    If dblCostAmount <> dblIVAmount Then
        crFore1 = vbRed
    Else
        crFore1 = vbBlack
    End If
    If dblCost <> dblIV Then
        crFore2 = vbRed
    Else
        crFore2 = vbBlack
    End If
    lblCostAmount.ForeColor = crFore1
    lblIVAmount.ForeColor = crFore1
    lblCost.ForeColor = crFore2
    lblIV.ForeColor = crFore2
End Sub

Private Sub ApportionInvoiceAmount(ByVal dblAmountOld As Double, ByVal dblAmountNew As Double)
'��̯��Ʊ����ϸ�ķ�Ʊ���
    Dim i As Long, LngLastRow As Long
    Dim dblTmp As Double
    
    With vsfMaterial
        For i = 1 To .Rows - 1
            If .RowHidden(i) = False Then
                LngLastRow = i
                 .TextMatrix(i, .ColIndex("��Ʊ���")) = Format(Val(.TextMatrix(i, .ColIndex("��Ʊ���"))) / IIf(dblAmountOld = 0, 1, dblAmountOld) * dblAmountNew, mFMT.FM_���)
            End If
        Next
        For i = 1 To .Rows - 1
            If .RowHidden(i) = False Then
                If LngLastRow = i Then
                    dblTmp = dblAmountNew - dblTmp
                    .TextMatrix(i, .ColIndex("��Ʊ���")) = Format(dblTmp, mFMT.FM_���)
                    Exit For
                Else
                    dblTmp = dblTmp + Val(.TextMatrix(i, .ColIndex("��Ʊ���")))
                End If
            End If
        Next
    End With
End Sub

Private Sub FillMaterial(ByVal lngPatientID As Long, ByVal lngPatientDrugID As Long)
'��䵱ǰ����ʹ�õĲ�����Ϣ
    Dim i As Long, lngTop As Long
    With vsfMaterial
        For i = 1 To .Rows - 1
            If lngPatientID = Val(.TextMatrix(i, .ColIndex("����ID"))) And lngPatientDrugID = Val(.TextMatrix(i, .ColIndex("����ID"))) Then
                .RowHidden(i) = False
                If lngTop = 0 Then lngTop = i
            Else
                .RowHidden(i) = True
            End If
        Next
        If lngTop > 0 Then .Row = lngTop
    End With
End Sub

Private Sub VerifyBatch()
'�������
    Dim i As Long, j As Long
    Dim strNo As String
    Dim strTime_Start As String, strTime_End As String
    Dim lngPatientID As Long, lngPatientDrugID As Long
    Dim strNewPirce As String
    Dim intCount As Integer
    
    If vsfPatient.Rows <= 1 Or vsfMaterial.Rows <= 1 Then Exit Sub
    
    If MsgBox("��ȷ��������" & IIf(mintMode = 1, "�˲�", "���") & "����", vbInformation + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    
    'If mintMode <> 1 Then
        With vsfPatient
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("��Ʊ��"))) = "" Then
                    MsgBox "����Ʊ�š�δ¼�룡", vbInformation, gstrSysName
                    .Col = .ColIndex("��Ʊ��")
                    .Row = i
                    .SetFocus
                    Exit Sub
                End If
                If Trim(.TextMatrix(i, .ColIndex("��Ʊ����"))) = "" Then
                    MsgBox "����Ʊ���롱δ¼�룡", vbInformation, gstrSysName
                    .Col = .ColIndex("��Ʊ����")
                    .Row = i
                    .SetFocus
                    Exit Sub
                End If
                If Trim(.TextMatrix(i, .ColIndex("��Ʊ����"))) = "" Then
                    MsgBox "����Ʊ���ڡ�δ¼�룡", vbInformation, gstrSysName
                    .Col = .ColIndex("��Ʊ����")
                    .Row = i
                    .SetFocus
                    Exit Sub
                End If
                If Val(.TextMatrix(i, .ColIndex("��Ʊ���"))) < 0 Then
                    MsgBox "����Ʊ��δ¼�룡", vbInformation, gstrSysName
                    .Col = .ColIndex("��Ʊ���")
                    .Row = i
                    .SetFocus
                    Exit Sub
                End If
            Next
        End With
    'End If
    
    With vsfMaterial
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("��Ʊ���"))) < 0 Then
                MsgBox "����Ʊ��δ¼�룡", vbInformation, gstrSysName
                .Col = .ColIndex("��Ʊ���")
                .Row = i
                .SetFocus
                Exit Sub
            End If
        Next
    End With
    
    '���۸�䶯������۸�䶯����½�������
    If mblnUpdate = False Then
        strNo = ""
        For i = 1 To vsfMaterial.Rows - 1
            strNo = vsfMaterial.TextMatrix(i, vsfMaterial.ColIndex("NO"))
            If Not CheckValuePrice(15, strNo) = True Then
                intCount = intCount + 1
                If intCount <= 5 Then
                    strNewPirce = IIf(strNewPirce = "", "", strNewPirce & vbCrLf) & vsfMaterial.TextMatrix(i, vsfMaterial.ColIndex("��������"))
                End If
            End If
        Next
        
        If strNewPirce <> "" Then
            ShowMsgBox "��ֵ������ⵥ�м۸��ѵ��ۣ��������Զ���ɸ��£��ۼۡ��ۼ۽����ɱ��ۡ��ɱ�����ۣ�,���飡" & vbCrLf & strNewPirce
            mblnUpdate = True
            Call ShowAmount
            Exit Sub
        End If
    End If
    
    strNo = ""
    For i = 1 To vsfMaterial.Rows - 1
        If strNo = vsfMaterial.TextMatrix(i, vsfMaterial.ColIndex("NO")) Then GoTo Continue
        
        strNo = vsfMaterial.TextMatrix(i, vsfMaterial.ColIndex("NO"))
                
        gcnOracle.BeginTrans
        
        '���浥��
        If SaveCard(strNo) = False Then
            gcnOracle.RollbackTrans
            Exit Sub
        End If
        
        If mbln��Ҫ�˲� Then
            If mintMode = 1 Then
                strTime_Start = GetBillInfo(15, strNo, False, True)
                strTime_End = GetBillInfo(15, strNo, False, True)
                If strTime_Start = "" Then strTime_Start = GetBillInfo(15, strNo)
                If strTime_End = "" Then strTime_End = GetBillInfo(15, strNo)
            Else
                strTime_Start = GetBillInfo(15, strNo)
                strTime_End = GetBillInfo(15, strNo)
            End If
        Else
            strTime_Start = GetBillInfo(15, strNo)
            strTime_End = GetBillInfo(15, strNo)
        End If
        If strTime_End = "" Then
            gcnOracle.RollbackTrans
            MsgBox "ע��:" & vbCrLf & "  �õ����Ѿ�����������Աɾ��,���ܼ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mintMode <> 1 Then
            '��˵���
            If SaveCheck(strNo) = True Then
                gcnOracle.CommitTrans
                '��������ɵĵ���
                For j = 1 To vsfMaterial.Rows - 1
                    If vsfMaterial.TextMatrix(j, vsfMaterial.ColIndex("NO")) = strNo Then
                        vsfMaterial.TextMatrix(j, vsfMaterial.ColIndex("Verify")) = "1"
                    End If
                Next
            Else
                gcnOracle.RollbackTrans
            End If
        Else
            gcnOracle.CommitTrans
            '�����ɵĵ���
            For j = 1 To vsfMaterial.Rows - 1
                If vsfMaterial.TextMatrix(j, vsfMaterial.ColIndex("NO")) = strNo Then
                    vsfMaterial.TextMatrix(j, vsfMaterial.ColIndex("Verify")) = "1"
                End If
            Next
        End If
        
Continue:
    Next
    
    '�����������
    With vsfMaterial
        lngPatientID = 0
        '������ѡ���Ĳ��˲���
        For i = .Rows - 1 To 1 Step -1
            If .TextMatrix(i, .ColIndex("Verify")) = "1" Then
                'If lngPatientID <> Val(.TextMatrix(i, .ColIndex("����ID"))) Then
                    lngPatientID = Val(.TextMatrix(i, .ColIndex("����ID")))
                    lngPatientDrugID = Val(.TextMatrix(i, .ColIndex("����ID")))
                    '������ѡ���Ĳ���
                    For j = vsfPatient.Rows - 1 To 1 Step -1
                        If Val(vsfPatient.TextMatrix(j, vsfPatient.ColIndex("����ID"))) = lngPatientID And Val(vsfPatient.TextMatrix(j, vsfPatient.ColIndex("����ID"))) = lngPatientDrugID Then
                            vsfPatient.RemoveItem j
                            Exit For
                        End If
                    Next
                    '�������б�
                    For j = 1 To lvwPatient.ListItems.Count
                        If Val(Mid(lvwPatient.ListItems(j).Key, 2)) = lngPatientID And Val(Mid(lvwPatient.ListItems(j).Key, InStr(2, lvwPatient.ListItems(j).Key, "_") + 1)) = lngPatientDrugID Then
                            lvwPatient.ListItems.Remove j
                            Exit For
                        End If
                    Next
                'End If
                .RemoveItem i
            End If
        Next
    End With
End Sub

Private Function CheckValuePrice(ByVal int���� As Integer, ByVal strNo As String) As Boolean
    '����ֵ������������������ⵥ�ļ۸��м۸�䶯ʱ���½���۸񣬽��
    '��������ⵥ�������ں��Ƿ����ͬ���εĵ��ۼ�¼������е��ۼ�¼��������ĵ��ۼ�¼�͵�ǰ��ⵥ�ļ۸���бȽ�
    'ֻ���ʱ�����ĵ��ۼۺͳɱ���
    '���أ�true-���ͨ��,false-�м۸�䶯
    Dim rsData As ADODB.Recordset
    Dim rsprice As ADODB.Recordset
    Dim lng����ID As Long
    Dim lng���� As Long
    Dim str�������� As String
    Dim dblԭ�� As Double
    Dim dbl���ۼ� As Double
    Dim dbl�ֳɱ��� As Double
    Dim strAdjustList As String '��Ҫ�䶯���嵥������id,����,���ۼ�(Ϊ0��ʾ�۸��ޱ仯),�ֳɱ���(Ϊ0��ʾ�۸��ޱ仯)
    Dim lngRow As Long
    Dim lngRows As Long
    Dim dbl���� As Double
    Dim dbl�ɱ���� As Double
    Dim dbl���۽�� As Double
    Dim dbl��� As Double
    Dim lng����id As Long
    Dim lng����id As Long
    Dim bln�ɱ��۱䶯 As Boolean
    Dim dbl�ɱ����ϼ� As Double
    Dim blnUpdate As Boolean
    
    gstrSQL = " Select '�����ۼ�' As ����, a.ҩƷid As ����id, Nvl(a.����, 0) As ����, a.���ۼ� As ԭ��, a.�������� " & _
            " From ҩƷ�շ���¼ A," & _
                 " (Select �շ�ϸĿid, Nvl(�ּ�, 0) �ּ�, ִ������" & _
                   " From �շѼ�Ŀ" & _
                   " Where (��ֹ���� Is Null Or Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd')))" & _
                   GetPriceClassString("") & ") B, �շ���ĿĿ¼ C" & _
            " Where a.���� = [1] And a.No = [2] And a.ҩƷid = b.�շ�ϸĿid And c.Id = b.�շ�ϸĿid And Round(a.���ۼ�,2) <> Round(b.�ּ�, 2) And" & _
              "    NVL(c.�Ƿ���, 0) = 0 " & _
        " Union All" & vbNewLine & _
        "Select 'ʱ���ۼ�' As ����, a.ҩƷid As ����id, Nvl(a.����, 0) As ����, a.���ۼ� As ԭ��, a.�������� " & vbNewLine & _
        " From ҩƷ�շ���¼ A, �շ���ĿĿ¼ C" & vbNewLine & _
        " Where a.���� = [1] And a.No = [2] And c.Id = a.ҩƷid And Nvl(c.�Ƿ���, 0) = 1 And a.����id > 0 And Exists" & vbNewLine & _
        " (Select 1" & vbNewLine & _
        "       From ҩƷ�շ���¼ B" & vbNewLine & _
        "       Where a.ҩƷid = b.ҩƷid And a.���� = b.���� And b.���� = 13 And b.������� > a.�������� And b.ժҪ = '���ĵ���')" & vbNewLine & _
        " Union All" & vbNewLine & _
        " Select '�ɱ���' As ����, a.ҩƷid As ����id, Nvl(a.����, 0) As ����, a.�ɱ��� As ԭ��, a.�������� " & vbNewLine & _
        " From ҩƷ�շ���¼ A" & vbNewLine & _
        " Where a.���� = [1] And a.No = [2] And a.����id > 0 And Exists" & vbNewLine & _
        " (Select 1" & vbNewLine & _
        "       From ҩƷ�շ���¼ B" & vbNewLine & _
        "       Where a.ҩƷid = b.ҩƷid And a.���� = b.���� And b.���� = 18 And b.������� > a.�������� And b.ժҪ = '�������ϳɱ��۵���') "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "CheckValuePrice", int����, strNo)
        
    If rsData.RecordCount = 0 Then
        CheckValuePrice = True
        Exit Function
    End If
    
    '��鵽�е��ۼ�¼��Ƚϼ۸����ڵ��ۼ�¼�����ж�����ȡ���һ���۸����Ƚ�
    Do While Not rsData.EOF
        lng����ID = rsData!����ID
        lng���� = rsData!����
        str�������� = Format(rsData!��������, "yyyy-mm-dd hh:mm:ss")
        dblԭ�� = rsData!ԭ��
        
        dbl���ۼ� = 0
        dbl�ֳɱ��� = 0
        bln�ɱ��۱䶯 = False
        
        If rsData!���� = "�����ۼ�" Then
            gstrSQL = " Select nvl(�ּ�,0) �ּ� From �շѼ�Ŀ " & _
            " Where �շ�ϸĿid=[1] and (��ֹ���� Is NULL Or sysdate Between ִ������ And nvl(��ֹ����,to_date('3000-01-01','yyyy-MM-dd')))" & _
            GetPriceClassString("")
            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, "CheckValuePrice", lng����ID)
            
            If rsprice.RecordCount > 0 Then
                If Round(rsprice!�ּ�, 2) <> Round(dblԭ��, 2) Then
                    dbl���ۼ� = rsprice!�ּ�
                    blnUpdate = True
                End If
            End If
        End If
        
        If rsData!���� = "ʱ���ۼ�" Then
            gstrSQL = "Select ���ۼ� As �ּ� " & _
                " From ҩƷ�շ���¼ " & _
                " Where ID = (Select Max(ID) " & _
                " From ҩƷ�շ���¼ B " & _
                " Where b.ҩƷid = [1] And b.���� = [2] And b.���� = 13 And b.������� > [3] And b.ժҪ = '���ĵ���') "
            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, "CheckValuePrice", lng����ID, lng����, CDate(str��������))
            
            If rsprice.RecordCount > 0 Then
                If Round(rsprice!�ּ�, 2) <> Round(dblԭ��, 2) Then
                    dbl���ۼ� = rsprice!�ּ�
                    blnUpdate = True
                End If
            End If
        End If
        
        If rsData!���� = "�ɱ���" Then
            gstrSQL = "Select ���� As �ּ� " & _
                " From ҩƷ�շ���¼ " & _
                " Where ID = (Select Max(ID) " & _
                " From ҩƷ�շ���¼ B " & _
                " Where b.ҩƷid = [1] And b.���� = [2] And b.���� = 18 And b.������� > [3] And b.ժҪ = '�������ϳɱ��۵���') "
            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, "CheckValuePrice", lng����ID, lng����, CDate(str��������))
            
            If rsprice.RecordCount > 0 Then
                If Round(rsprice!�ּ�, 2) <> Round(dblԭ��, 2) Then
                    dbl�ֳɱ��� = rsprice!�ּ�
                    bln�ɱ��۱䶯 = True
                    blnUpdate = True
                End If
            End If
        End If
        
        '�Ե�ǰ���¼۸����µ���������ݣ����ۡ����۽���ۣ�
        With vsfMaterial
            lngRows = vsfMaterial.Rows - 1
            For lngRow = 1 To lngRows
                If strNo = .TextMatrix(lngRow, .ColIndex("NO")) And lng����ID = Val(.TextMatrix(lngRow, .ColIndex("����ID"))) And (dbl���ۼ� <> 0 Or dbl�ֳɱ��� <> 0) Then
                    dbl���� = Val(.TextMatrix(lngRow, .ColIndex("����")))
                    lng����id = Val(.TextMatrix(lngRow, .ColIndex("����ID")))
                    lng����id = Val(.TextMatrix(lngRow, .ColIndex("����ID")))
                    If dbl���ۼ� <> 0 Then
                        dbl���ۼ� = Val(Format(dbl���ۼ� * Val(.TextMatrix(lngRow, .ColIndex("����ϵ��"))), mFMT.FM_���ۼ�))
                        dbl���۽�� = dbl���ۼ� * dbl����
                    Else
                        dbl���ۼ� = Val(.TextMatrix(lngRow, .ColIndex("���ۼ�")))
                        dbl���۽�� = Val(.TextMatrix(lngRow, .ColIndex("���۽��")))
                    End If
                    
                    If dbl�ֳɱ��� <> 0 Then
                        dbl�ֳɱ��� = Val(Format(dbl�ֳɱ��� * Val(.TextMatrix(lngRow, .ColIndex("����ϵ��"))), mFMT.FM_�ɱ���))
                        dbl�ɱ���� = dbl�ֳɱ��� * dbl����
                    Else
                        dbl�ֳɱ��� = Val(.TextMatrix(lngRow, .ColIndex("�ɱ���")))
                        dbl�ɱ���� = Val(.TextMatrix(lngRow, .ColIndex("�ɱ����")))
                    End If
                    
                    dbl��� = dbl���۽�� - dbl�ɱ����
                    
                    .TextMatrix(lngRow, .ColIndex("�ɱ���")) = Format(dbl�ֳɱ���, mFMT.FM_�ɱ���)
                    .TextMatrix(lngRow, .ColIndex("�ɱ����")) = Format(dbl�ɱ����, mFMT.FM_���)
                    .TextMatrix(lngRow, .ColIndex("���ۼ�")) = Format(dbl���ۼ�, mFMT.FM_���ۼ�)
                    .TextMatrix(lngRow, .ColIndex("���۽��")) = Format(dbl���۽��, mFMT.FM_���)
                    .TextMatrix(lngRow, .ColIndex("���")) = Format(dbl���, mFMT.FM_���)

                End If
            Next
                    
            dbl�ɱ����ϼ� = 0
            If bln�ɱ��۱䶯 = True Then
                For lngRow = 1 To lngRows
                    If lng����id = Val(.TextMatrix(lngRow, .ColIndex("����ID"))) And lng����id = Val(.TextMatrix(lngRow, .ColIndex("����ID"))) Then
                        dbl�ɱ����ϼ� = dbl�ɱ����ϼ� + .TextMatrix(lngRow, .ColIndex("�ɱ����"))
                    End If
                Next
                
                '���²����б�ɱ����
                lngRows = vsfPatient.Rows - 1
                For lngRow = 1 To lngRows
                    If lng����id = Val(vsfPatient.TextMatrix(lngRow, vsfPatient.ColIndex("����ID"))) And lng����id = Val(vsfPatient.TextMatrix(lngRow, vsfPatient.ColIndex("����ID"))) Then
                        vsfPatient.TextMatrix(lngRow, vsfPatient.ColIndex("�ɱ����")) = dbl�ɱ����ϼ�
                        vsfPatient.Row = lngRow
                        vsfPatient.TopRow = lngRow
                    End If
                Next
            End If
        End With
        
        rsData.MoveNext
    Loop
    
    CheckValuePrice = Not blnUpdate
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function SaveCheck(Optional ByVal strNo As String = "") As Boolean
   
    gstrSQL = "zl_�����⹺_Verify('" & strNo & "','" & UserInfo.�û��� & "')"
    
    On Error GoTo ErrHandle
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-���")
    
    SaveCheck = True
    Exit Function

ErrHandle:
    Call ErrCenter
    Call SaveErrLog
End Function

Private Function SaveCard(ByVal strNo As String) As Boolean
    Dim lng��� As Long
    Dim lngStockID As Long
    Dim lng������λid As Long
    Dim lng����ID As Long
    Dim str���� As String
    Dim str���� As String
    Dim strЧ�� As String
    Dim dblʵ������ As Double
    Dim dbl�ɱ��� As Double
    Dim dbl�ɱ���� As Double
    Dim dbl���� As Double
    Dim dbl���ۼ� As Double
    Dim dbl���۽�� As Double
    Dim dbl��� As Double
    Dim dbl���۲�� As Double
    Dim strժҪ As String
    Dim str������ As String
    Dim str�������� As String
    Dim str����� As String
    Dim datAssessDate As String
    Dim str��Ʊ�� As String
    Dim str��Ʊ���� As String
    Dim str��Ʊ���� As String
    Dim str������� As String
    Dim str���ʧЧ�� As String
    Dim dbl��Ʊ��� As Double
    Dim str��������  As String
    Dim str�˲��� As String
    Dim str�˲����� As String
    Dim strע��֤�� As String
    Dim intUnit As Integer
    Dim strUnit As String
    Dim strָ�������� As String
    Dim str������� As String
    Dim str��Ʒ���� As String
    Dim str�ڲ����� As String
    Dim lng����ID As Long
    Dim intRow As Integer
    Dim str��ֵ���� As String
    Dim str���� As String
    Dim strTmp As String
    
    
    SaveCard = False
    
    With vsfMaterial
        
        lngStockID = mlngStockID
        lng������λid = lvwPatient.Tag
        
        On Error GoTo ErrHandle
        
        gstrSQL = "zl_�����⹺_Delete('" & strNo & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, .ColIndex("NO")) <> strNo Then GoTo Continue
            
            lng��� = lng��� + 1
            strժҪ = Trim(.TextMatrix(intRow, .ColIndex("ժҪ")))
            str������ = Trim(.TextMatrix(intRow, .ColIndex("������")))
            str����� = UserInfo.����
            str�������� = Trim(.TextMatrix(intRow, .ColIndex("��������")))
            
            If mbln��Ҫ�˲� Then
                If mintMode = 1 Then
                    str�˲��� = UserInfo.����
                    str�˲����� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                Else
                    str�˲��� = Trim(.TextMatrix(intRow, .ColIndex("�˲���")))
                    str�˲����� = Trim(.TextMatrix(intRow, .ColIndex("�˲�����")))
                    If str�˲��� = "" Then
                        str�˲��� = UserInfo.����
                    End If
                    If str�˲����� = "" Then
                        str�˲����� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                    End If
                End If
            Else
                str�˲��� = Trim(.TextMatrix(intRow, .ColIndex("�˲���")))
                str�˲����� = Trim(.TextMatrix(intRow, .ColIndex("�˲�����")))
                If str�˲��� = "" Then
                    str�˲��� = UserInfo.����
                End If
                If str�˲����� = "" Then
                    str�˲����� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                End If
            End If
            
            lng����ID = .TextMatrix(intRow, .ColIndex("����ID"))
            str���� = .TextMatrix(intRow, .ColIndex("����"))
            str���� = .TextMatrix(intRow, .ColIndex("����"))
            strЧ�� = .TextMatrix(intRow, .ColIndex("Ч��"))
                
            strTmp = Val(.TextMatrix(intRow, .ColIndex("����"))) * Val(.TextMatrix(intRow, .ColIndex("����ϵ��")))
            dblʵ������ = Format(strTmp, mFMT.FM_����)
            
            strTmp = Val(.TextMatrix(intRow, .ColIndex("�ɱ���"))) / Val(.TextMatrix(intRow, .ColIndex("����ϵ��")))
            dbl�ɱ��� = Round(Val(strTmp), g_С��λ��.obj_ɢװС��.�ɱ���С��)
            
            strTmp = Val(.TextMatrix(intRow, .ColIndex("���ۼ�"))) / Val(.TextMatrix(intRow, .ColIndex("����ϵ��")))
            dbl���ۼ� = Round(Val(strTmp), g_С��λ��.obj_ɢװС��.���ۼ�С��)
            
            dbl���� = Val(.TextMatrix(intRow, .ColIndex("����")))
            dbl�ɱ���� = Val(.TextMatrix(intRow, .ColIndex("�ɱ����")))
            dbl���۽�� = Val(.TextMatrix(intRow, .ColIndex("���۽��")))
            dbl��� = Val(.TextMatrix(intRow, .ColIndex("���")))
            dbl���۲�� = Val(.TextMatrix(intRow, .ColIndex("��۽��")))
            str������� = ""
                
            If GetInvoiceInfo(.TextMatrix(intRow, .ColIndex("����ID")), str��Ʊ��, str��Ʊ����, str��Ʊ����) = False Then
                Exit Function
            End If
            
            dbl��Ʊ��� = Val(.TextMatrix(intRow, .ColIndex("��Ʊ���")))
            str������� = Trim(IIf(.TextMatrix(intRow, .ColIndex("�������")) = "", "", .TextMatrix(intRow, .ColIndex("�������"))))
            str���ʧЧ�� = Trim(IIf(.TextMatrix(intRow, .ColIndex("���Ч��")) = "", "", .TextMatrix(intRow, .ColIndex("���Ч��"))))
            str�������� = Trim(IIf(.TextMatrix(intRow, .ColIndex("��������")) = "", "", .TextMatrix(intRow, .ColIndex("��������"))))
            strע��֤�� = Trim(.TextMatrix(intRow, .ColIndex("ע��֤��")))
            str�ڲ����� = Trim(.TextMatrix(intRow, .ColIndex("�ڲ�����")))
            str��Ʒ���� = Trim(.TextMatrix(intRow, .ColIndex("��Ʒ����")))
            lng����ID = Val(.TextMatrix(intRow, .ColIndex("����ID")))
            str��ֵ���� = Trim(.TextMatrix(intRow, .ColIndex("��ֵ����")))
            str���� = Trim(.TextMatrix(intRow, .ColIndex("����")))
                
            ' Zl_�����⹺_Insert
            gstrSQL = "zl_�����⹺_INSERT("
            '  No_In         In ҩƷ�շ���¼.NO%Type,
            gstrSQL = gstrSQL & "'" & strNo & "',"
            '  ���_In       In ҩƷ�շ���¼.���%Type,
            gstrSQL = gstrSQL & "" & lng��� & ","
            '  �ⷿid_In     In ҩƷ�շ���¼.�ⷿid%Type,
            gstrSQL = gstrSQL & "" & lngStockID & ","
            '  ��ҩ��λid_In In ҩƷ�շ���¼.��ҩ��λid%Type,
            gstrSQL = gstrSQL & "" & lng������λid & ","
            '  ����id_In     In ҩƷ�շ���¼.ҩƷid%Type,
            gstrSQL = gstrSQL & "" & lng����ID & ","
            '  ����_In       In ҩƷ�շ���¼.����%Type := Null,
            gstrSQL = gstrSQL & "'" & str���� & "',"
            '  ����_In       In ҩƷ�շ���¼.����%Type := Null,
            gstrSQL = gstrSQL & "'" & str���� & "',"
            '  ��������_In   In ҩƷ�շ���¼.��������%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(str�������� = "", "Null", "to_date('" & Format(str��������, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
            '  Ч��_In       In ҩƷ�շ���¼.Ч��%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(strЧ�� = "", "Null", "to_date('" & Format(strЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
            '  �������_In   In ҩƷ�շ���¼.�������%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(str������� = "", "Null", "to_date('" & Format(str�������, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
            '  ���Ч��_In   In ҩƷ�շ���¼.���Ч��%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(str���ʧЧ�� = "", "Null", "to_date('" & Format(str���ʧЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
            '  ʵ������_In   In ҩƷ�շ���¼.ʵ������%Type := Null,
            gstrSQL = gstrSQL & "" & dblʵ������ & ","
            '  �ɱ���_In     In ҩƷ�շ���¼.�ɱ���%Type := Null,
            gstrSQL = gstrSQL & "" & dbl�ɱ��� & ","
            '  �ɱ����_In   In ҩƷ�շ���¼.�ɱ����%Type := Null,
            gstrSQL = gstrSQL & "" & dbl�ɱ���� & ","
            '  ����_In       In ҩƷ�շ���¼.����%Type := Null,
            gstrSQL = gstrSQL & "" & dbl���� & ","
            '  ���ۼ�_In     In ҩƷ�շ���¼.���ۼ�%Type := Null,
            gstrSQL = gstrSQL & "" & dbl���ۼ� & ","
            '  ���۽��_In   In ҩƷ�շ���¼.���۽��%Type := Null,
            gstrSQL = gstrSQL & "" & dbl���۽�� & ","
            '  ���_In       In ҩƷ�շ���¼.���%Type := Null,
            gstrSQL = gstrSQL & "" & dbl��� & ","
            '  ���۲��_In   In ҩƷ�շ���¼.���%Type := Null,Ŀǰ������÷��ֶ�
            gstrSQL = gstrSQL & "" & dbl���۲�� & ","
            '  ժҪ_In       In ҩƷ�շ���¼.ժҪ%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(strժҪ = "", "NULL", "'" & strժҪ & "'") & ","
            '  ע��֤��_In   In ҩƷ�շ���¼.ע��֤��%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(strע��֤�� = "", "NULL", "'" & strע��֤�� & "'") & ","
            '  ������_In     In ҩƷ�շ���¼.������%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(str������ = "", "NULL", "'" & str������ & "'") & ","
            '  �������_In   In Ӧ����¼.�������%Type := Null
            gstrSQL = gstrSQL & "" & IIf(str������� = "", "NULL", "'" & str������� & "'") & ","
            '  ��Ʊ��_In     In Ӧ����¼.��Ʊ��%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(str��Ʊ�� = "", "NULL", "'" & str��Ʊ�� & "'") & ","
            '  ��Ʊ����_In   In Ӧ����¼.��Ʊ����%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(str��Ʊ���� = "", "Null", "to_date('" & Format(str��Ʊ����, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
            '  ��Ʊ���_In   In Ӧ����¼.��Ʊ���%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(dbl��Ʊ��� = 0, "Null", dbl��Ʊ���) & ","
            '  ��������_In   In ҩƷ�շ���¼.��������%Type := Null,
            gstrSQL = gstrSQL & "to_date('" & str�������� & "','yyyy-mm-dd HH24:MI:SS'),"
            '  �˲���_In     In ҩƷ�շ���¼.��ҩ��%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(str�˲��� = "", "NULL", "'" & str�˲��� & "'") & ","
            '  �˲�����_In   In ҩƷ�շ���¼.��ҩ����%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(str�˲����� = "", "Null", "to_date('" & str�˲����� & "','yyyy-mm-dd hh24:mi:ss')") & ","
            '  ����_In       In ҩƷ�շ���¼.����%Type := 0,
            gstrSQL = gstrSQL & "" & IIf(str���� = "", "Null", "'" & str���� & "'") & ","
            '  �˻�_In       In Number := 1
            gstrSQL = gstrSQL & "1,"
            '  ��ֵ����_In   In varchar2(250)
            gstrSQL = gstrSQL & "" & IIf(str��ֵ���� = "", "Null", "'" & str��ֵ���� & "'") & ","
            '  ��Ʒ����_In   In ҩƷ�շ���¼.��Ʒ����%Type :=Null
            gstrSQL = gstrSQL & "" & IIf(str��Ʒ���� = "", "NULL", "'" & str��Ʒ���� & "'") & ","
            '  �ڲ�����
            gstrSQL = gstrSQL & IIf(str�ڲ����� = "", "Null", "'" & str�ڲ����� & "'") & ","
            '  ����ID
            gstrSQL = gstrSQL & IIf(lng����ID = 0, "Null", lng����ID) & ","
            '  ��Ʊ����
            gstrSQL = gstrSQL & "" & IIf(str��Ʊ���� = "", "NULL", "'" & str��Ʊ���� & "'")
            gstrSQL = gstrSQL & ")"
            
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
Continue:
        Next
        
    End With
    SaveCard = True
    Exit Function
    
ErrHandle:
    Call ErrCenter
    Call SaveErrLog
End Function

Private Function GetInvoiceInfo(ByVal lngPatientID, ByRef strIVNO As String, ByRef strIVCode As String, ByRef strIVDate As String) As Boolean
'��ȡ��Ʊ��
    Dim i As Long
    
    With vsfPatient
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("����ID")) = lngPatientID Then
                strIVNO = .TextMatrix(i, .ColIndex("��Ʊ��"))
                strIVCode = .TextMatrix(i, .ColIndex("��Ʊ����"))
                strIVDate = .TextMatrix(i, .ColIndex("��Ʊ����"))
                'dblIVAmount = .TextMatrix(i, .ColIndex("��Ʊ���"))
                GetInvoiceInfo = True
                Exit Function
            End If
        Next
        strIVNO = ""
        strIVDate = ""
        'dblIVAmount = 0
    End With
End Function


