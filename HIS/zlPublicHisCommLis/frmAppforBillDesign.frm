VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppforBillDesign 
   Caption         =   "���뵥���"
   ClientHeight    =   9075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15105
   Icon            =   "frmAppforBillDesign.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9075
   ScaleWidth      =   15105
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboDeptSel 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4170
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   870
      Width           =   2025
   End
   Begin VB.PictureBox picRight 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6345
      Left            =   6480
      ScaleHeight     =   6345
      ScaleWidth      =   3225
      TabIndex        =   1
      Top             =   450
      Width           =   3225
      Begin VSFlex8Ctl.VSFlexGrid VSFListDept 
         Height          =   2835
         Left            =   60
         TabIndex        =   3
         Top             =   3150
         Width           =   2895
         _cx             =   5106
         _cy             =   5001
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
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
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483635
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   2
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   0
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
      Begin VSFlex8Ctl.VSFlexGrid VSFList 
         Height          =   1995
         Left            =   150
         TabIndex        =   8
         Top             =   480
         Width           =   2895
         _cx             =   5106
         _cy             =   3519
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
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
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483635
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   2
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   0
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
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1111111"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   930
            TabIndex        =   10
            Top             =   1500
            Visible         =   0   'False
            Width           =   840
         End
      End
      Begin XtremeSuiteControls.ShortcutCaption ShortCaptionDept 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   2760
         Width           =   2745
         _Version        =   589884
         _ExtentX        =   4842
         _ExtentY        =   556
         _StockProps     =   6
         Caption         =   "���뵥ִ��С��"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         GradientColorLight=   14737632
         GradientColorDark=   14737632
      End
      Begin XtremeSuiteControls.ShortcutCaption ShortCaptionItem 
         Height          =   315
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2745
         _Version        =   589884
         _ExtentX        =   4842
         _ExtentY        =   556
         _StockProps     =   6
         Caption         =   "���뵥��Ŀ(���""����˳��""֮����϶��ı�˳��)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         GradientColorLight=   14737632
         GradientColorDark=   14737632
      End
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5235
      Left            =   60
      ScaleHeight     =   5235
      ScaleWidth      =   3225
      TabIndex        =   0
      Top             =   840
      Width           =   3225
      Begin VSFlex8Ctl.VSFlexGrid VSFType 
         Height          =   1995
         Left            =   30
         TabIndex        =   9
         Top             =   510
         Width           =   2895
         _cx             =   5106
         _cy             =   3519
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
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
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483635
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   2
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   0
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
      Begin XtremeSuiteControls.ShortcutCaption ShortCaptionType 
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   2745
         _Version        =   589884
         _ExtentX        =   4842
         _ExtentY        =   661
         _StockProps     =   6
         Caption         =   "���뵥"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         GradientColorLight=   14737632
         GradientColorDark=   14737632
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   8700
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   23945
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   390
      Top             =   60
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmAppforBillDesign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnfrmIfShow As Boolean                                        '�����Ƿ���ʾ���
Private mlngkeyID As Long                                               '����ID
Private mblnAllSite As Boolean                                          '�鿴����վ��
Private mblnItemSort As Boolean                                         '�Ƿ���˳�����״̬

'ʵ���϶�Ч����Ҫ�ı���
Private mlngMouseRow As Long            '��������
Private mlngMouseDownRow As Long        '��갴�µ���

Public Sub ShowMe(objFrm As Object, ByVal blnAddSite As Boolean)
    mblnAllSite = blnAddSite
    Me.Show 1, objFrm
End Sub

Private Sub cboDeptSel_Click()
    Call ReadTypeData
End Sub

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case ConMenu_Appfro_AddBill                     '�������뵥
            frmAppforBillDesignEditBill.ShowMe Me, 0, Me.cboDeptSel.ItemData(Me.cboDeptSel.ListIndex), "", "", 0, False
            Call ReadTypeData
        Case ConMenu_Appfro_ModifyBill                  '�޸ķ���
            Call ModifyType
        Case ConMenu_Appfro_DelBill                     'ɾ�����뵥
            Call DelType
        Case ConMenu_Appfro_ModifyItem                  '�޸�������Ŀ
            Call ModifyItem
        Case ConMenu_Appfro_Group                       'ѡ�����
            Call SelectGroup
        Case ConMenu_Appfro_ModifyDept                  'ִ��С��
            Call ModifyGroup
        Case ConMenu_Appfor_ItemSort                    '����˳��
            If Control.Caption = "����" Then
                Call SaveItemSort
                Control.Caption = "����˳��"
                cbrthis.RecalcLayout
            Else
                mblnItemSort = True
                Control.Caption = "����"
                cbrthis.RecalcLayout
            End If
        Case ConMenu_Appfro_Refresh                     'ˢ��
            Call ReadTypeData
        Case ConMenu_Appfro_Exit                        '�˳�
            Unload Me
    End Select
End Sub

Private Sub cbrthis_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    With Me.picLeft
        .Top = Top
        .Left = Left + 10
        .Width = (Right - Left) / 3
        .Height = Bottom - Top - stbThis.Height + 25
    End With
    With Me.picRight
        .Top = Top + 6
        .Left = picLeft.Left + picLeft.Width + 25
        .Width = (Right - Left) - .Left - 25
        .Height = Me.picLeft.Height
    End With
End Sub

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case ConMenu_Appfro_AddBill                     '�������뵥
            Control.Enabled = Not mblnItemSort
        Case ConMenu_Appfro_ModifyBill                  '�޸ķ���
            Control.Enabled = Not mblnItemSort
        Case ConMenu_Appfro_DelBill                     'ɾ�����뵥
            Control.Enabled = Not mblnItemSort
        Case ConMenu_Appfro_ModifyItem                  '�޸�������Ŀ
            Control.Enabled = Not mblnItemSort
        Case ConMenu_Appfro_Group                       'ѡ�����
            Control.Enabled = Not mblnItemSort
        Case ConMenu_Appfro_ModifyDept                  'ִ��С��
            Control.Enabled = Not mblnItemSort
        Case ConMenu_Appfro_Refresh                     'ˢ��
            Control.Enabled = Not mblnItemSort
    End Select
    picLeft.Enabled = Not mblnItemSort
End Sub

Private Sub Form_Activate()
    If mblnfrmIfShow = False Then
        Call InitVSF
        Call LoadDept
        
        mblnfrmIfShow = True
    End If
End Sub

Private Sub Form_Load()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrthis.VisualTheme = xtpThemeOffice2003
    Me.cbrthis.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbrthis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrthis.EnableCustomization False

    '-----------------------------------------------------
    '�˵�����
    Me.cbrthis.ActiveMenuBar.Title = "�˵�"
    Me.cbrthis.ActiveMenuBar.Visible = False
    Set cbrToolBar = Me.cbrthis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfro_AddBill, "����")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfro_ModifyBill, "�޸�")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfro_DelBill, "ɾ��")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfro_ModifyItem, "�޸���Ŀ")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfro_Group, "ѡ�����")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfro_ModifyDept, "ִ��С��")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfor_ItemSort, "����˳��")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfro_Refresh, "ˢ��")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfro_Exit, "�˳�")
        
        
    End With
    
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlLabel, 0, "Ӧ�ÿ���")
    cbrControl.Flags = xtpFlagRightAlign

    Set cbrCustom = cbrToolBar.Controls.Add(xtpControlCustom, ConMenu_Appfro_DeptSel, "Ӧ�ÿ���")
    cbrCustom.ShortcutText = "Ӧ�ÿ���"
    cbrCustom.Handle = Me.cboDeptSel.hWnd
    cbrCustom.Flags = xtpFlagRightAlign
    cbrCustom.Style = xtpButtonIconAndCaption
    
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type = xtpControlButton Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnfrmIfShow = False
    mblnAllSite = False
    mblnItemSort = False
    mlngMouseRow = 0
    mlngMouseDownRow = 0
End Sub

Private Sub picLeft_Resize()
    On Error Resume Next
    With ShortCaptionType
        .Top = 6
        .Left = 6
        .Width = Me.picLeft.ScaleWidth
    End With
    With VSFType
        .Top = ShortCaptionType.Height + 12
        .Left = 6
        .Width = picLeft.ScaleWidth - .Left * 2
        .Height = picLeft.ScaleHeight - .Top
    End With
End Sub

Private Sub picRight_Resize()
    On Error Resume Next
    With ShortCaptionItem
        .Top = 6
        .Left = 6
        .Width = Me.picRight.ScaleWidth
    End With
    With VSFList
        .Top = ShortCaptionItem.Top + ShortCaptionItem.Height + 6
        .Left = 6
        .Width = picRight.ScaleWidth - .Left * 2
        .Height = picRight.ScaleHeight - VSFListDept.Height - ShortCaptionDept.Height - 48
    End With
    With ShortCaptionDept
        .Top = VSFList.Top + VSFList.Height + 6
        .Left = 6
        .Width = picRight.ScaleWidth
    End With
    With VSFListDept
        .Top = ShortCaptionDept.Top + ShortCaptionDept.Height + 6
        .Left = 6
        .Width = picRight.ScaleWidth
        .Height = Me.picRight.ScaleHeight - .Top
    End With
    
End Sub
Private Sub InitVSF()
      '��ʼ���б�
1         On Error GoTo InitVSF_Error

2         With Me.VSFList
3             .Rows = 2
4             .Cols = 6
5             .FixedRows = 1
6             .ColKey(0) = "����": .ColWidth(.ColIndex("����")) = 2000: .ColAlignment(.ColIndex("����")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("����")) = "����"
7             .Cell(flexcpAlignment, 0, .ColIndex("����"), 0, .ColIndex("����")) = flexAlignCenterCenter
8             .ColKey(1) = "�����Ŀ": .ColWidth(.ColIndex("�����Ŀ")) = 3000: .ColAlignment(.ColIndex("�����Ŀ")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("�����Ŀ")) = "�����Ŀ"
9             .Cell(flexcpAlignment, 0, .ColIndex("�����Ŀ"), 0, .ColIndex("�����Ŀ")) = flexAlignCenterCenter
10            .ColKey(2) = "����": .ColWidth(.ColIndex("����")) = 2000: .ColAlignment(.ColIndex("����")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("����")) = "����"
11            .Cell(flexcpAlignment, 0, .ColIndex("����"), 0, .ColIndex("����")) = flexAlignCenterCenter
12            .ColKey(3) = "����˳��": .ColAlignment(.ColIndex("����˳��")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("����˳��")) = "����˳��"
13            .Cell(flexcpAlignment, 0, .ColIndex("����˳��"), 0, .ColIndex("����˳��")) = flexAlignCenterCenter: .ColHidden(.ColIndex("����˳��")) = True
14            .ColKey(4) = "ID": .ColAlignment(.ColIndex("ID")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("ID")) = "ID"
15            .Cell(flexcpAlignment, 0, .ColIndex("ID"), 0, .ColIndex("ID")) = flexAlignCenterCenter: .ColHidden(.ColIndex("ID")) = True
16            .ColKey(5) = "����": .ColWidth(.ColIndex("����")) = 2000: .ColAlignment(.ColIndex("����")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("����")) = "����"
17            .Cell(flexcpAlignment, 0, .ColIndex("����"), 0, .ColIndex("����")) = flexAlignCenterCenter
18        End With

19        With Me.VSFType
20            .Rows = 2
21            .Cols = 5
22            .FixedRows = 1
23            .ColKey(0) = "ID": .ColWidth(.ColIndex("ID")) = 1000: .ColAlignment(.ColIndex("ID")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("ID")) = "ID": .ColHidden(.ColIndex("ID")) = True
24            .ColKey(1) = "����": .ColWidth(.ColIndex("����")) = 1000: .ColAlignment(.ColIndex("����")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("����")) = "����"
25            .Cell(flexcpAlignment, 0, .ColIndex("����"), 0, .ColIndex("����")) = flexAlignCenterCenter
26            .ColKey(2) = "����": .ColWidth(.ColIndex("����")) = 2000: .ColAlignment(.ColIndex("����")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("����")) = "����"
27            .Cell(flexcpAlignment, 0, .ColIndex("����"), 0, .ColIndex("����")) = flexAlignCenterCenter
28            .ColKey(3) = "��ɫ": .ColWidth(.ColIndex("��ɫ")) = 700: .ColAlignment(.ColIndex("��ɫ")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("��ɫ")) = "��ɫ"
29            .Cell(flexcpAlignment, 0, .ColIndex("��ɫ"), 0, .ColIndex("��ɫ")) = flexAlignCenterCenter
30            .ColKey(4) = "��������": .ColWidth(.ColIndex("��������")) = 700: .ColAlignment(.ColIndex("��������")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("��������")) = "��������"
31            .Cell(flexcpAlignment, 0, .ColIndex("��������"), 0, .ColIndex("��������")) = flexAlignCenterCenter
32        End With

33        With Me.VSFListDept
34            .Rows = 2
35            .Cols = 4
36            .FixedRows = 1
37            .ColKey(0) = "����": .ColWidth(.ColIndex("����")) = 2000: .ColAlignment(.ColIndex("����")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("����")) = "����"
38            .Cell(flexcpAlignment, 0, .ColIndex("����"), 0, .ColIndex("����")) = flexAlignCenterCenter
39            .ColKey(1) = "С������": .ColWidth(.ColIndex("С������")) = 2500: .ColAlignment(.ColIndex("С������")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("С������")) = "С������"
40            .Cell(flexcpAlignment, 0, .ColIndex("С������"), 0, .ColIndex("С������")) = flexAlignCenterCenter
41            .ColKey(2) = "HIS���ű���": .ColWidth(.ColIndex("HIS���ű���")) = 2500: .ColAlignment(.ColIndex("HIS���ű���")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("HIS���ű���")) = "HIS���ű���"
42            .Cell(flexcpAlignment, 0, .ColIndex("HIS���ű���"), 0, .ColIndex("HIS���ű���")) = flexAlignCenterCenter
43            .ColKey(3) = "Ĭ��": .ColWidth(.ColIndex("Ĭ��")) = 700: .ColAlignment(.ColIndex("Ĭ��")) = flexAlignCenterCenter: .TextMatrix(0, .ColIndex("Ĭ��")) = "Ĭ��"
44        End With


45        Exit Sub
InitVSF_Error:
46        Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesign", "ִ��(InitVSF)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
47        Err.Clear
End Sub
Private Sub ReadTypeData()
          '����   ����������ݵ��б���
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim intloop As Integer
          
1         On Error GoTo ReadTypeData_Error

2         VSFList.Rows = 1: VSFList.Rows = 2
3         VSFListDept.Rows = 1: VSFListDept.Rows = 2
          
4         strSQL = " select ID,����,����,��ɫ,�Ƿ��������뵥 from �������뵥 where nvl(����ID,0) = [1] order by ���� "
5         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�������뵥", Val(Me.cboDeptSel.ItemData(Me.cboDeptSel.ListIndex)))
6         With Me.VSFType
7             .Rows = 1
8             Do Until rsTmp.EOF
9                 .Rows = .Rows + 1
10                .TextMatrix(.Rows - 1, .ColIndex("ID")) = rsTmp("ID") & ""
11                .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTmp("����") & ""
12                .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTmp("����") & ""
13                .Cell(flexcpBackColor, .Rows - 1, .ColIndex("��ɫ"), .Rows - 1, .ColIndex("��ɫ")) = Val(rsTmp("��ɫ") & "")
14                .TextMatrix(.Rows - 1, .ColIndex("��������")) = IIf(Val(rsTmp("�Ƿ��������뵥") & "") = 1, "��", "")
15                rsTmp.MoveNext
16            Loop
17            If .Rows = 1 Then
18                .Rows = .Rows + 1
19                .Row = 1
20                mlngkeyID = 0
21                Exit Sub
22            End If
23            If mlngkeyID = 0 Then
24                .Row = 1
25                Exit Sub
26            End If
27            For intloop = 1 To .Rows - 1
28                If .TextMatrix(intloop, .ColIndex("ID")) = mlngkeyID Then
29                    .Row = intloop
30                    Exit For
31                End If
32            Next
'33            mlngkeyID = 0
34        End With


35        Exit Sub
ReadTypeData_Error:
36        Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesign", "ִ��(ReadTypeData)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
37        Err.Clear
End Sub

Private Sub ReadItemData()
    '����       ���������ϸ
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strItem As String
    Dim strDefault As String
    
    On Error GoTo ReadItemData_Error
    
    '��Ŀ��Ϣ
    strSQL = "Select a.id, b.����, b.����, b.����,a.����˳��,d.���� ����" & vbNewLine & _
             " From �������뵥��ϸ A, ���������Ŀ B,�������뵥 c,�������뵥���� d Where a.���뵥id =c.id and A.���id = B.Id and a.����id=d.id(+) and b.ͣ������ is null and a.���뵥ID = [1] order by a.����ID,a.����˳��, b.���� "
    Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���������ϸ", mlngkeyID)
    With Me.VSFList
        .Rows = 1
        
        Do Until rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("id")) = rsTmp("id") & ""
            .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTmp("����") & ""
            .TextMatrix(.Rows - 1, .ColIndex("�����Ŀ")) = rsTmp("����") & ""
            .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTmp("����") & ""
            .TextMatrix(.Rows - 1, .ColIndex("����˳��")) = rsTmp("����˳��") & ""
            .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTmp("����") & ""
            
            rsTmp.MoveNext
        Loop
        If .Rows = 1 Then .Rows = 2
    End With
    
    'ִ��С��
    strSQL = "Select ִ��С��,Ĭ��ִ��С�� From �������뵥 Where ID = [1]"
    Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���뵥ִ��С��", mlngkeyID)
    If Not rsTmp.EOF Then
        strItem = rsTmp("ִ��С��") & ""
        strDefault = rsTmp("Ĭ��ִ��С��") & ""
    End If
    
    If gUserInfo.NodeNo = "-" Or mblnAllSite Then
        strSQL = "select ����,���� С������,HIS���ű��� from ����С���¼" & vbNewLine & _
                "where ���� in (Select * From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist)))"
    Else
        strSQL = "select ����,���� С������,HIS���ű��� from ����С���¼" & vbNewLine & _
                "where ���� in (Select * From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist))) and (վ��=[2] or վ�� is null)"
    End If
    Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�������С��", strItem, gUserInfo.NodeNo)
    With Me.VSFListDept
        .Rows = 1
        Do Until rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTmp("����") & ""
            .TextMatrix(.Rows - 1, .ColIndex("С������")) = rsTmp("С������") & ""
            .TextMatrix(.Rows - 1, .ColIndex("HIS���ű���")) = rsTmp("HIS���ű���") & ""
            
            .Cell(flexcpChecked, .Rows - 1, .ColIndex("Ĭ��"), .Rows - 1, .ColIndex("Ĭ��")) = 2
            .Cell(flexcpPictureAlignment, .Rows - 1, .ColIndex("Ĭ��"), .Rows - 1, .ColIndex("Ĭ��")) = flexAlignCenterCenter
            If InStr("," & strDefault & ",", "," & rsTmp("����") & ",") > 0 Then
                .Cell(flexcpChecked, .Rows - 1, .ColIndex("Ĭ��"), .Rows - 1, .ColIndex("Ĭ��")) = 1
            End If
            rsTmp.MoveNext
        Loop
    End With


    Exit Sub
ReadItemData_Error:
    Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesign", "ִ��(ReadItemData)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
    Err.Clear

End Sub

Private Sub DelType()
          '����   ɾ������
          Dim strSQL As String
          
1         On Error GoTo DelType_Error

2         If mlngkeyID = 0 Then
3             MsgBox "��ѡ��һ������!", vbInformation, "ɾ������"
4             Exit Sub
5         End If
6         With Me.VSFType
7             If MsgBox("��ȷ��Ҫɾ��<" & .TextMatrix(.Row, .ColIndex("����")) & ">����?", vbYesNo + vbInformation + vbDefaultButton2) = vbNo Then
8                 Exit Sub
9             End If
              
10            strSQL = "Zl_�������뵥_Edit('" & 3 & "','" & mlngkeyID & "','','')"
11            Call ComExecuteProc(Sel_Lis_DB, strSQL, "ɾ������")
12            SaveDBLog 18, 6, 0, "ɾ��", "ɾ�����뵥:" & .TextMatrix(.Row, .ColIndex("����")), 1012, "���뵥����"
13            mlngkeyID = 0
14            Call ReadTypeData
15        End With
          


16        Exit Sub
DelType_Error:
17        Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesign", "ִ��(DelType)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
18        Err.Clear
         
End Sub

Private Sub ModifyType()
          '����   �޸ķ���
1         On Error GoTo ModifyType_Error

2         With Me.VSFType
3             If mlngkeyID = 0 Then
4                 MsgBox "��ѡ��һ������!", vbInformation, "�޸ķ���"
5                 Exit Sub
6             End If
7             frmAppforBillDesignEditBill.ShowMe Me, mlngkeyID, Me.cboDeptSel.ItemData(Me.cboDeptSel.ListIndex), _
                              .TextMatrix(.Row, .ColIndex("����")), .TextMatrix(.Row, .ColIndex("����")), _
                              Val(.Cell(flexcpBackColor, .Row, .ColIndex("��ɫ"), .Row, .ColIndex("��ɫ"))), _
                              IIf(Trim(.TextMatrix(.Row, .ColIndex("��������"))) = "��", True, False)
8             Call ReadTypeData
              
9         End With


10        Exit Sub
ModifyType_Error:
11        Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesign", "ִ��(ModifyType)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
12        Err.Clear
End Sub

Private Sub ModifyItem()
    '����   �޸���Ŀ
    If mlngkeyID = 0 Then
        MsgBox "��ѡ��һ������!", vbInformation, "�޸���Ŀ"
        Exit Sub
    End If
    frmAppforBillDesignEditItem.ShowMe Me, mlngkeyID, mblnAllSite, IIf(VSFType.TextMatrix(VSFType.Row, VSFType.ColIndex("��������")) <> "", True, False)
    Call ReadTypeData
End Sub




Private Sub VSFType_RowColChange()
    With Me.VSFType
        If .Rows = 0 Then Exit Sub
        If .ColIndex("ID") = -1 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("ID"))) <= 0 Then Exit Sub
        
        'If Val(.TextMatrix(.Row, .ColIndex("ID"))) <> mlngkeyID Then
            mlngkeyID = Val(.TextMatrix(.Row, .ColIndex("ID")))
            Call ReadItemData
        'End If
    End With
End Sub

Private Sub LoadDept()
          '����   �������
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
              
1         On Error GoTo LoadDept_Error

2         strSQL = "Select distinct ID, ����, ���� From ���ű� A, ��������˵�� B Where A.Id = B.����id And B.�������� In ('����', '�ٴ�')"
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "���벿�ű�")
4         With cboDeptSel
5             .Clear
6             .AddItem "���п���"
7             .ItemData(.NewIndex) = 0
8             Do Until rsTmp.EOF
9                 .AddItem rsTmp("����") & "-" & rsTmp("����")
10                .ItemData(.NewIndex) = rsTmp("ID")
11                rsTmp.MoveNext
12            Loop
13            .ListIndex = 0
14        End With


15        Exit Sub
LoadDept_Error:
16        Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesign", "ִ��(LoadDept)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
17        Err.Clear
          
End Sub
Private Sub ModifyGroup()
    '����   �޸ĵ���ִ��С��
    With Me.VSFType
        If .Rows = 1 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("ID"))) <> 0 Then
            frmAppforBillDesignDept.ShowMe Me, .TextMatrix(.Row, .ColIndex("ID"))
            ReadItemData
        End If
    End With
End Sub

Private Sub SelectGroup()
    '����   �޸���Ŀ
    If mlngkeyID = 0 Then
        MsgBox "��ѡ��һ������!", vbInformation, "ѡ�����"
        Exit Sub
    End If
    frmAppforBillGroup.ShowMe Me, mlngkeyID, VSFType.TextMatrix(VSFType.Row, VSFType.ColIndex("����"))
    Call ReadTypeData
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/6/7
'��    ��:��������
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub SaveItemSort()
          Dim lngRow As Long
          Dim strSQL As String
          Dim lngCount As Long
          
1         On Error GoTo SaveItemSort_Error

2         ReDim strArrSQL(0)
3         With Me.VSFList
4             For lngRow = 1 To .Rows - 1
5                 If .RowHidden(lngRow) = False Then
6                     lngCount = lngCount + 1
7                     strSQL = "Zl_���뵥��ϸ_Sort(" & Val(.TextMatrix(lngRow, .ColIndex("ID"))) & "," & lngCount & ")"
8                     Call ComExecuteProc(Sel_Lis_DB, strSQL, "���뵥����")
9                 End If
10            Next
11        End With
12        mblnItemSort = False


13        Exit Sub
SaveItemSort_Error:
14        Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesign", "ִ��(SaveItemSort)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
15        Err.Clear
End Sub


'==========���´��빦��Ϊ:���Ҳ�VSF�б��е������϶������VSF��=============
'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/6/14
'��    ��:ģ���϶�������б�ʱ������ǩ��λ�������λ�ã������������ƶ�
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub vsfList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1         On Error GoTo VSFList_MouseDown_Error

2         If Button <> 1 Then Exit Sub
3         If Not mblnItemSort Then Exit Sub
          
4         With Me.VSFList
5             If .MouseRow <= 0 Or .MouseCol < 0 Then Exit Sub
6             Me.lblShow.Caption = .TextMatrix(.MouseRow, .ColIndex("�����Ŀ"))
7             Me.lblShow.Tag = .TextMatrix(.MouseRow, .ColIndex("id")) & "|" & .TextMatrix(.MouseRow, .ColIndex("����")) & "|" & .TextMatrix(.MouseRow, .ColIndex("����")) & "|" & .TextMatrix(.MouseRow, .ColIndex("�����Ŀ")) & "|" & .TextMatrix(.MouseRow, .ColIndex("����"))
8             mlngMouseDownRow = .MouseRow
9         End With


10        Exit Sub
VSFList_MouseDown_Error:
11        MsgBox "ִ��(VSFList_MouseDown)ʱ����,��������:" & Err.Description & " �����:" & Err.Number & " ������:" & Erl, vbInformation, "��ʾ"
12        Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/6/14
'��    ��:��ǩ��������ƶ�
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub VSFList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
          Dim lngRow As Long
          Dim lngCol As Long
          
1         On Error GoTo VSFList_MouseMove_Error

2         If Button <> 1 Then Exit Sub
3         If Not mblnItemSort Then Exit Sub
          
4         If Me.lblShow.Caption = "" Then Exit Sub
5         With Me.lblShow
6             If .Visible = False Then .Visible = True
7             .Left = X - (.Width / 2)
8             .Top = Y - (.Height / 2)
9         End With
          
          '�����Ҳ��б����϶����ʱ��Ч��
10        With Me.VSFList
11            lngRow = .MouseRow
12            lngCol = .MouseCol
13            If lngRow > -1 And lngCol > -1 Then
14                If mlngMouseRow <> lngRow And mlngMouseRow > 0 And lngRow > 0 Then
                      '�ƶ���ĳһ����֮������һ������
15                    If mlngMouseRow <= .Rows - 1 Then
16                        If Trim(.TextMatrix(mlngMouseRow, .ColIndex("�����Ŀ"))) = "" Then .RemoveItem mlngMouseRow    '���Ƴ�֮ǰ�Ŀ���
17                    End If
18                    Debug.Print 1
19                    .AddItem "", lngRow
20                    mlngMouseRow = lngRow
21                    .Row = mlngMouseRow
22                ElseIf mlngMouseRow = 0 And lngRow > 0 Then
23                    Debug.Print 2
24                    .AddItem "", lngRow
25                    mlngMouseRow = lngRow
26                ElseIf lngRow = .Rows - 1 And Trim(.TextMatrix(.Rows - 1, .ColIndex("�����Ŀ"))) <> "" Then
                      '����ƶ������һ��,�����������һ��
27                    Debug.Print 3
28                    .AddItem "", .Rows
29                    mlngMouseRow = .Rows
30                End If
31            ElseIf lngRow = -1 And .Rows < 2 Then
32                Debug.Print 4
33                .Rows = .Rows + 1
34                mlngMouseRow = .Rows - 1
35            ElseIf lngRow = -1 And lngCol = -1 And mlngMouseRow <= .Rows - 1 Then
36                If Trim(.TextMatrix(mlngMouseRow, .ColIndex("�����Ŀ"))) = "" Then
37                    .RemoveItem mlngMouseRow
38                End If
39            End If
40        End With
          

41        Exit Sub
VSFList_MouseMove_Error:
42        MsgBox "ִ��(VSFList_MouseMove)ʱ����,��������:" & Err.Description & " �����:" & Err.Number & " ������:" & Erl, vbInformation, "��ʾ"
43        Err.Clear

End Sub


'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/6/14
'��    ��:�ɿ����ʱ,���϶���ֵ���Ƶ��ұߵ�VSF��
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub VSFList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
          
1         On Error GoTo VSFList_MouseUp_Error

2         If Button <> 1 Then Exit Sub
3         If Not mblnItemSort Then Exit Sub
4         With Me.VSFList
              '�����Ҳ��б����϶�����ʱ
5             If .MouseCol > -1 And mlngMouseRow > 0 And mlngMouseRow <= .Rows - 1 Then
6                 .TextMatrix(mlngMouseRow, .ColIndex("id")) = Split(Me.lblShow.Tag, "|")(0): .ColAlignment(.ColIndex("id")) = flexAlignLeftCenter
7                 .TextMatrix(mlngMouseRow, .ColIndex("����")) = Split(Me.lblShow.Tag, "|")(1): .ColAlignment(.ColIndex("����")) = flexAlignLeftCenter
8                 .TextMatrix(mlngMouseRow, .ColIndex("����")) = Split(Me.lblShow.Tag, "|")(2): .ColAlignment(.ColIndex("����")) = flexAlignLeftCenter
9                 .TextMatrix(mlngMouseRow, .ColIndex("�����Ŀ")) = Split(Me.lblShow.Tag, "|")(3): .ColAlignment(.ColIndex("�����Ŀ")) = flexAlignLeftCenter
10                .TextMatrix(mlngMouseRow, .ColIndex("����")) = Split(Me.lblShow.Tag, "|")(4): .ColAlignment(.ColIndex("����")) = flexAlignLeftCenter
11                If mlngMouseDownRow > 0 And Me.lblShow.Visible = True Then
12                    If mlngMouseRow > mlngMouseDownRow Then
13                        .RemoveItem mlngMouseDownRow
14                    ElseIf mlngMouseDownRow + 1 <= .Rows - 1 Then
15                        If .MouseCol > -1 Then .RemoveItem mlngMouseDownRow + 1
16                    End If
17                End If
18            End If
              
19        End With
20        mlngMouseRow = 0
21        Me.lblShow.Caption = ""
22        If Me.lblShow.Visible = True Then Me.lblShow.Visible = False

23        Exit Sub
VSFList_MouseUp_Error:
24        MsgBox "ִ��(VSFList_MouseUp)ʱ����,��������:" & Err.Description & " �����:" & Err.Number & " ������:" & Erl, vbInformation, "��ʾ"
          'WriteLog "ִ��(VSFList_MouseUp)ʱ����,��������:" & Err.Description & " �����:" & Err.Number & " ������:" & Erl
25        Err.Clear
End Sub
'=============================================================
