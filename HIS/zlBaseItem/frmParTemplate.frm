VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmParTemplate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "临床参数设置"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12555
   Icon            =   "frmParTemplate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   12555
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7440
      Index           =   0
      Left            =   2400
      ScaleHeight     =   7410
      ScaleWidth      =   10065
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   10095
      Begin VB.TextBox txtUD 
         Alignment       =   2  'Center
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1590
         MaxLength       =   4
         TabIndex        =   39
         Text            =   "30"
         Top             =   120
         Width           =   450
      End
      Begin VB.Frame fra抗菌目的 
         BorderStyle     =   0  'None
         Height          =   1935
         Index           =   2
         Left            =   120
         TabIndex        =   29
         Top             =   3360
         Width           =   4380
         Begin VB.Frame fra抗菌目的 
            Caption         =   "住院"
            Height          =   680
            Index           =   1
            Left            =   0
            TabIndex        =   34
            Top             =   1200
            Width           =   4305
            Begin VB.OptionButton opt抗菌目的住院 
               Caption         =   "治疗"
               Height          =   180
               Index           =   2
               Left            =   3000
               TabIndex        =   37
               Top             =   300
               Value           =   -1  'True
               Width           =   680
            End
            Begin VB.OptionButton opt抗菌目的住院 
               Caption         =   "预防"
               Height          =   180
               Index           =   1
               Left            =   1920
               TabIndex        =   36
               Top             =   300
               Width           =   680
            End
            Begin VB.OptionButton opt抗菌目的住院 
               Caption         =   "下达时确定"
               Height          =   180
               Index           =   0
               Left            =   255
               TabIndex        =   35
               Top             =   300
               Width           =   1275
            End
         End
         Begin VB.Frame fra抗菌目的 
            Caption         =   "门诊"
            Height          =   680
            Index           =   0
            Left            =   0
            TabIndex        =   30
            Top             =   330
            Width           =   4305
            Begin VB.OptionButton opt抗菌目的门诊 
               Caption         =   "下达时确定"
               Height          =   180
               Index           =   0
               Left            =   255
               TabIndex        =   33
               Top             =   300
               Width           =   1275
            End
            Begin VB.OptionButton opt抗菌目的门诊 
               Caption         =   "预防"
               Height          =   180
               Index           =   1
               Left            =   1920
               TabIndex        =   32
               Top             =   300
               Width           =   680
            End
            Begin VB.OptionButton opt抗菌目的门诊 
               Caption         =   "治疗"
               Height          =   180
               Index           =   2
               Left            =   3000
               TabIndex        =   31
               Top             =   300
               Value           =   -1  'True
               Width           =   680
            End
         End
         Begin VB.Label lbl抗菌目的 
            Caption         =   "抗菌药物缺省用药目的"
            Height          =   255
            Left            =   0
            TabIndex        =   38
            Top             =   60
            Width           =   1935
         End
      End
      Begin VB.CommandButton cmdAdvice 
         Caption         =   "医嘱内容定义(&F)"
         Height          =   405
         Left            =   120
         TabIndex        =   28
         Top             =   2760
         Width           =   1680
      End
      Begin VB.Frame fra入院诊断 
         Caption         =   "住院下达以下类别医嘱时检查是否填写入院诊断"
         Height          =   1365
         Left            =   120
         TabIndex        =   23
         Top             =   5880
         Width           =   4320
         Begin VB.CommandButton cmd住院检查入院诊断 
            Caption         =   "全清"
            Height          =   300
            Index           =   1
            Left            =   3120
            TabIndex        =   26
            Top             =   720
            Width           =   900
         End
         Begin VB.CommandButton cmd住院检查入院诊断 
            Caption         =   "全选"
            Height          =   300
            Index           =   0
            Left            =   3120
            TabIndex        =   25
            Top             =   360
            Width           =   900
         End
         Begin VB.ListBox lst 
            Columns         =   3
            ForeColor       =   &H80000012&
            Height          =   900
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   24
            Top             =   375
            Width           =   2940
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "过敏登记有效天数"
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   570
         Width           =   1740
      End
      Begin VB.TextBox txtUD 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   1890
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "1"
         Top             =   555
         Width           =   495
      End
      Begin VB.CheckBox chk 
         Caption         =   "下达医嘱时显示产地"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1980
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   960
         Width           =   1455
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   0
         Left            =   2400
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   555
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtUD(0)"
         BuddyDispid     =   196625
         BuddyIndex      =   0
         OrigLeft        =   2400
         OrigTop         =   1380
         OrigRight       =   2655
         OrigBottom      =   1680
         Max             =   365
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   1
         Left            =   2040
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   120
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   30
         BuddyControl    =   "txtUD(1)"
         BuddyDispid     =   196625
         BuddyIndex      =   1
         OrigLeft        =   2085
         OrigTop         =   120
         OrigRight       =   2340
         OrigBottom      =   390
         Max             =   9999
         Min             =   10
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "补录医嘱识别间隔         分钟"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   180
         Width           =   2610
      End
      Begin VB.Label lbl 
         Caption         =   "中药配方每行"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   990
         Width           =   1095
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7455
      Index           =   1
      Left            =   2400
      ScaleHeight     =   7425
      ScaleWidth      =   10185
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   10215
      Begin VB.TextBox txt 
         Height          =   735
         Index           =   1
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   46
         Top             =   6480
         Width           =   5295
      End
      Begin VB.TextBox txt 
         Height          =   855
         Index           =   0
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   45
         Top             =   5160
         Width           =   5295
      End
      Begin VB.Frame fraCLKS 
         Height          =   4650
         Left            =   120
         TabIndex        =   42
         Top             =   0
         Width           =   5295
         Begin VSFlex8Ctl.VSFlexGrid vsUnWriteDept 
            Height          =   3885
            Left            =   120
            TabIndex        =   43
            Top             =   525
            Width           =   5040
            _cx             =   8890
            _cy             =   6853
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
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
            ForeColorSel    =   0
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   10
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParTemplate.frx":6852
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
            Editable        =   2
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
         Begin VB.Label lbl 
            Caption         =   "设置可不录入超量原因的科室，例如：精神科。"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   44
            Top             =   300
            Width           =   4095
         End
      End
      Begin VB.Label lblBloodPrompt 
         Caption         =   "住院输血申请注意事项"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   48
         Top             =   6240
         Width           =   2655
      End
      Begin VB.Label lblBloodPrompt 
         Caption         =   "门诊输血申请注意事项"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   47
         Top             =   4920
         Width           =   2655
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7455
      Index           =   2
      Left            =   2400
      ScaleHeight     =   7425
      ScaleWidth      =   10185
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   10215
   End
   Begin VB.PictureBox picFunc 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      FillColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   7440
      Left            =   0
      ScaleHeight     =   7440
      ScaleWidth      =   2415
      TabIndex        =   14
      Top             =   0
      Width           =   2415
      Begin VB.PictureBox picTPL 
         BorderStyle     =   0  'None
         Height          =   6135
         Left            =   0
         ScaleHeight     =   6135
         ScaleWidth      =   2250
         TabIndex        =   16
         Top             =   0
         Width           =   2250
         Begin XtremeSuiteControls.TaskPanel tplFunc 
            Height          =   5250
            Left            =   0
            TabIndex        =   17
            Top             =   720
            Width           =   2205
            _Version        =   589884
            _ExtentX        =   3889
            _ExtentY        =   9260
            _StockProps     =   64
            Behaviour       =   1
            ItemLayout      =   2
            HotTrackStyle   =   3
         End
         Begin XtremeCommandBars.ImageManager imgFunc 
            Left            =   1800
            Top             =   360
            _Version        =   589884
            _ExtentX        =   635
            _ExtentY        =   635
            _StockProps     =   0
            Icons           =   "frmParTemplate.frx":6908
         End
         Begin XtremeSuiteControls.ShortcutCaption sccFunc 
            Height          =   300
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   2200
            _Version        =   589884
            _ExtentX        =   3881
            _ExtentY        =   529
            _StockProps     =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            Alignment       =   1
         End
      End
      Begin VB.PictureBox picVbar 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         FillColor       =   &H8000000A&
         Height          =   5820
         Left            =   2280
         MousePointer    =   9  'Size W E
         ScaleHeight     =   5820
         ScaleWidth      =   45
         TabIndex        =   15
         Top             =   120
         Width           =   45
      End
      Begin XtremeSuiteControls.ShortcutBar scbFunc 
         Height          =   6765
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   2400
         _Version        =   589884
         _ExtentX        =   4233
         _ExtentY        =   11933
         _StockProps     =   64
      End
      Begin XtremeCommandBars.ImageManager imgType 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         Icons           =   "frmParTemplate.frx":A246
      End
   End
   Begin VB.PictureBox PicBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   590
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   12555
      TabIndex        =   1
      Top             =   7440
      Width           =   12555
      Begin VB.TextBox txtLocate 
         Height          =   300
         Index           =   1
         Left            =   4700
         TabIndex        =   21
         Top             =   120
         Width           =   1200
      End
      Begin VB.TextBox txtLocate 
         Height          =   300
         Index           =   0
         Left            =   2400
         TabIndex        =   7
         Top             =   120
         Width           =   1200
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   60
         TabIndex        =   4
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   11400
         TabIndex        =   3
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   10245
         TabIndex        =   2
         Top             =   120
         Width           =   1100
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   6000
         TabIndex        =   22
         Top             =   165
         Width           =   4095
      End
      Begin VB.Label lblLocate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "科室查找(&F)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   20
         Top             =   165
         Width           =   1095
      End
      Begin VB.Label lblLocate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "参数查找(&S)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   6
         Top             =   168
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmParTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsPar As ADODB.Recordset '参数与控件对应记录集（同一个参数可能对应一组多个控件）
Private marrFunc(2) As String
Private mlngPreFind As Long

Private Enum constTxtLocate
    txt_Par = 0
    txt_Dept = 1
End Enum

Private Enum constChk
    chk_下达医嘱时显示产地 = 0
    chk_过敏登记有效天数 = 1
End Enum

Private Enum constCbo
    cbo_中药配方 = 0
End Enum

Private Enum constUpDown
    ud_过敏登记有效天数 = 0
    ud_补录医嘱识别间隔 = 1
End Enum

Private Enum constTxt
    txt_门诊输血申请注意事项 = 0
    txt_住院输血申请注意事项 = 1
End Enum

Private Enum constListBox
    lst_住院检查入院诊断 = 0
End Enum

Private Sub cmdHelp_Click()
     ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_Activate()
    If Me.Tag = "初始成功" Then
        Call scbFunc_SelectedChanged(scbFunc.Selected)
        Me.Tag = ""
    End If
End Sub

Private Sub Form_Load()
    Dim strCategory As String
    
    strCategory = "参数设置,基础设置"
    
    '图标编号,TaskPanelItem的ID(同时也是参数容器Picture控件数组号),TaskPanelItem的标题;......
    marrFunc(0) = "100,0,医嘱下达选项;101,1,业务流程控制"
    marrFunc(1) = "102,2,科室药房设置"

    '1.初始化快捷面板的一级分类列表,缺省选中第一个
    Call InitSCBItem(scbFunc, strCategory, picTPL.hwnd)
    Call scbFunc.Icons.AddIcons(imgType.Icons)
      
    '2.初始化任务面板的二级分类列表,缺省选中第一个
    Call InitTPLItem(sccFunc, tplFunc, scbFunc.Selected.Caption, marrFunc(0))
    Call tplFunc.Icons.AddIcons(imgFunc.Icons)
    
    
    Call InitData
    
    Me.Tag = "初始成功"
End Sub

Private Sub tplFunc_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    Dim i As Long
    
    For i = 0 To picPar.UBound
        picPar(i).Visible = (i = Item.ID)
    Next
    
    lblLocate(txt_Dept).Visible = (Item.ID = GetFuncID("业务流程控制", marrFunc))
    txtLocate(txt_Dept).Visible = lblLocate(txt_Dept).Visible
    If txtLocate(txt_Dept).Visible Then
        lblPrompt.Left = txtLocate(txt_Dept).Left + txtLocate(txt_Dept).Width + 60
    Else
        lblPrompt.Left = txtLocate(txt_Par).Left + txtLocate(txt_Par).Width + 60
    End If
    lblPrompt.Width = cmdOK.Left - lblPrompt.Left - 120
    
    mlngPreFind = 1
    
    tplFunc.Tag = Item.ID   '用于获取当前选中的TaskPanelItem
End Sub

Private Sub Form_Resize()
    Dim i As Long
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If picVbar.Left < 1500 Then picVbar.Left = 1500
    If picVbar.Left > Me.ScaleWidth - 3000 Then picVbar.Left = Me.ScaleWidth - 3000
    picVbar.Top = 0
    
    picFunc.Width = picVbar.Left + picVbar.Width
    
    For i = 0 To picPar.UBound
        picPar(i).Top = Me.ScaleTop
        picPar(i).Left = picFunc.Left + picFunc.ScaleWidth
        picPar(i).Width = Me.ScaleWidth - picPar(i).Left
        picPar(i).Height = Me.ScaleHeight - PicBottom.ScaleHeight
    Next
End Sub


Private Sub scbFunc_ExpandButtonDown(CancelMenu As Boolean)
    CancelMenu = True
End Sub

Private Sub picBottom_Resize()
    cmdCancel.Left = PicBottom.ScaleWidth - cmdCancel.Width - 120
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 120
End Sub


Private Sub picFunc_Resize()
    scbFunc.Top = picFunc.ScaleTop
    scbFunc.Left = picFunc.ScaleLeft + 45
    scbFunc.Width = picFunc.ScaleWidth - picVbar.Width - 45
    scbFunc.Height = picFunc.ScaleHeight
    
    picVbar.Height = picFunc.ScaleHeight
End Sub

Private Sub picTPL_Resize()
    sccFunc.Left = picTPL.ScaleLeft
    sccFunc.Width = picTPL.ScaleWidth
    
    tplFunc.Left = picTPL.ScaleLeft
    tplFunc.Top = sccFunc.Top + sccFunc.Height
    tplFunc.Height = picTPL.ScaleHeight - sccFunc.Height
    tplFunc.Width = picTPL.ScaleWidth
End Sub


Private Sub picVbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picVbar.Left = IIF(picVbar.Left + X < 2000, 2000, picVbar.Left + X)
        Call Form_Resize
    End If
End Sub

Private Sub scbFunc_SelectedChanged(ByVal Item As XtremeSuiteControls.IShortcutBarItem)
    If Me.Visible Then
        Call InitTPLItem(sccFunc, tplFunc, Item.Caption, marrFunc(Item.ID - 1)) 'ID是从1开始的（因为同时为图标序号）,数组是从0开始
        Call tplFunc_ItemClick(tplFunc.Groups(1).Items(1))
    End If
End Sub


Public Sub LocateFuncItem(ByVal lngFunc As Long)
'功能：根据ID选中一级和二级分类
    Dim i As Long, j As Long, lngID As Long
    Dim arrTmp As Variant
    Dim n As Long
    
    For i = 0 To UBound(marrFunc)
        arrTmp = Split(marrFunc(i), ";")
        For j = 0 To UBound(arrTmp)
            lngID = Split(arrTmp(j), ",")(1)
            If lngFunc = lngID Then
                tplFunc.Tag = lngID
                Set scbFunc.Selected = scbFunc(i)
                
                For n = 1 To tplFunc.Groups(1).Items.Count
                    tplFunc.Groups(1).Items(n).Selected = tplFunc.Groups(1).Items(n).ID = lngID
                Next
            End If
        Next
    Next
End Sub


Private Sub Form_Unload(Cancel As Integer)
    mrsPar.Filter = "修改状态=1"
    If mrsPar.RecordCount > 0 Or cmdAdvice.Tag = "已修改" Then
        If MsgBox("你已修改部分参数，如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
    
    Set mrsPar = Nothing
End Sub

Private Sub InitData()
'功能：初始化界面控件,读取并加载数据
    
    '1.初始化变量
    mlngPreFind = 1
    
    Call InitSystemPara
    
    
    '2.初始化界面控件
    Call InitEnv
    
    
    '3.加载系统参数
    Call LoadPar
    
End Sub

Private Sub LoadPar()
'功能：读取并加载参数到界面控件
    Dim strValue As String, strTmp As String
    Dim i As Long
    Dim rsTmp As ADODB.Recordset
    Dim arrObj As Variant  '数组对象：模块1,参数号1,控件对象1,模块2,参数号2,控件对象2,......
    
    Set rsTmp = GetPar(mrsPar, p门诊医嘱下达 & "," & p住院医嘱下达 & "," & p住院医嘱发送)
        
     '1.设置CheckBox类参数
    strTmp = "0:162:" & chk_下达医嘱时显示产地 & _
            ",0:70:" & chk_过敏登记有效天数
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '2.设置ComboBox类参数
    strTmp = ""
    Call SetParToControl(strTmp, mrsPar, cbo)
            
    '3.设置UpDown类参数
    strTmp = "0:5:" & ud_补录医嘱识别间隔
    Call SetParToControl(strTmp, mrsPar, ud)     'mrsPar存储的控件名是txtUD
    
    '4.设置TextBox类参数
    strTmp = p门诊医嘱下达 & ":53:" & txt_门诊输血申请注意事项 & _
        "," & p住院医嘱下达 & ":56:" & txt_住院输血申请注意事项
    Call SetParToControl(strTmp, mrsPar, txt)
    
    
    '5.设置ListBox类参数
    strTmp = p住院医嘱下达 & ":4:" & lst_住院检查入院诊断
    Call SetParToControl(strTmp, mrsPar, lst)
        
    '6.设置OptionButton类参数
    arrObj = Array(p门诊医嘱下达, 45, opt抗菌目的门诊, _
                    p住院医嘱下达, 51, opt抗菌目的住院)
    Call SetParToControl("", mrsPar, arrObj)
    
    
    '7.其他系统参数
    rsTmp.Filter = "模块=0"
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数号
        Case 70
            ud(ud_过敏登记有效天数).Value = IIF(Val(strValue) = 0, 1, Val(strValue))
            
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "") '已有CheckBox控件，所以需要再产生一条记录
            Call SetParRelation(txtUD, ud_过敏登记有效天数, mrsPar)
            
        Case 233
            Call Load不写超量科室(strValue)
            Call SetParRelation(vsUnWriteDept, 0, mrsPar, rsTmp!参数号)
        End Select
        
        rsTmp.MoveNext
    Loop
    
    '8.其他模块参数设置
    rsTmp.Filter = "模块=" & p门诊医嘱下达
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数号
        
        End Select
        rsTmp.MoveNext
    Loop
    
End Sub

Private Sub InitEnv()
'功能：初始化界面控件，加载基础数据
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, strTmp As String
    Dim blnTmp As Boolean
    
    On Error GoTo ErrHandle

    vsUnWriteDept.ComboList = "..."
    vsUnWriteDept.RowHeightMin = 280
    
    cbo(cbo_中药配方).AddItem "0-三味中药"
    cbo(cbo_中药配方).AddItem "1-四味中药"
    cbo(cbo_中药配方).ListIndex = 0

    
    '读取医嘱发送为划价类别
    gstrSQL = "Select 编码,名称 From 诊疗项目类别 Where 编码 Not IN('5','6','7','8','9')" & _
        " Union All Select '5','药品' From Dual Order by 编码"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)

    If rsTmp.RecordCount > 0 Then rsTmp.Filter = "编码<>'4'"
    Do While Not rsTmp.EOF
        lst(lst_住院检查入院诊断).AddItem rsTmp!编码 & "-" & rsTmp!名称
        lst(lst_住院检查入院诊断).ItemData(lst(lst_住院检查入院诊断).NewIndex) = Asc(rsTmp!编码)
        rsTmp.MoveNext
    Loop
    lst(lst_住院检查入院诊断).ListIndex = 0
  
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    
    If ValidateData() = False Then Exit Sub
    
    Call Save医嘱内容
    
    If SavePar(mrsPar, Me) = False Then Exit Sub
    Call zlDatabase.ClearParaCache
    
    Unload Me
End Sub

Private Function ValidateData() As Boolean
'功能：验证数据的有效性
    
    ValidateData = True
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub txtLocate_Change(Index As Integer)
    If Index = txt_Dept Then
        mlngPreFind = 1
    ElseIf Index = txt_Par Then
        txtLocate(Index).Tag = ""
    End If
End Sub

Private Sub txtLocate_GotFocus(Index As Integer)
    txtLocate(Index).SelStart = 0
    txtLocate(Index).SelLength = Len(txtLocate(Index).Text)
End Sub

Private Sub txtLocate_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim strFind As String
        
        If Trim(txtLocate(Index).Text) = "" Then Exit Sub
        strFind = UCase(Trim(txtLocate(Index).Text))
        
        Select Case Index
        Case txt_Par
            Call LocatePar(txtLocate(Index), Me)
        Case txt_Dept
            If vsUnWriteDept.Visible Then
                Call LocateDept(strFind, vsUnWriteDept)
            End If
        End Select
    End If
End Sub

Private Sub LocateDept(ByVal strFind As String, ByRef objTmp As Object)
'功能：检查不写超量的科室
    Dim i As Long, j As Long
    Dim lngRows As Long, lngStart As Long
    Dim strCode As String, strName As String
    
    If TypeName(objTmp) = "ListBox" Then
        With objTmp
            lngRows = .ListCount - 1
            
            lngStart = IIF(mlngPreFind = 1, 0, mlngPreFind)
            For i = lngStart To .ListCount - 1
                strCode = Split(.List(i), "-")(0)
                strName = Split(.List(i), "-")(1)
                If strCode Like strFind & "*" Or strName Like IIF(gstrLike <> "", "*", "") & strFind & "*" Then
                    .ListIndex = i
                    .SetFocus
                    Exit For
                End If
            Next
        End With
        If i < lngRows Then
            mlngPreFind = i + 1
        Else
            If mlngPreFind = 1 Then
                MsgBox "没有找到匹配的，请检查输入的内容。", vbInformation, Me.Caption
                txtLocate(txt_Dept).SetFocus
            Else
                MsgBox "全部找完了，后面没有了。", vbInformation, Me.Caption
                mlngPreFind = 1
            End If
        End If
    Else
        '考虑到此功能的使用频率低，未支持连续查找
        With objTmp
            For i = 0 To .Rows - 1
                For j = 0 To .Cols - 1
                    If .ColHidden(j) = False Then
                        If .TextMatrix(i, j) Like IIF(gstrLike <> "", "*", "") & strFind & "*" Then
                            .Row = i: .Col = j
                            .ShowCell i, j
                            Exit Sub
                        End If
                    End If
                Next
            Next
            
            MsgBox "没有找到匹配的科室，请检查输入的内容。", vbInformation, Me.Caption
            txtLocate(txt_Dept).SetFocus
        End With
    End If
End Sub


Private Sub lst_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(lst, Index, mrsPar)
End Sub

Private Sub chk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(chk, Index, mrsPar)
End Sub

Private Sub opt抗菌目的门诊_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt抗菌目的门诊, Index, mrsPar)
 
End Sub

Private Sub opt抗菌目的门诊_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt抗菌目的门诊_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt抗菌目的门诊, Index, mrsPar)
End Sub

Private Sub opt抗菌目的住院_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt抗菌目的住院, Index, mrsPar)
End Sub

Private Sub opt抗菌目的住院_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt抗菌目的住院_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt抗菌目的住院, Index, mrsPar)
End Sub

Private Sub txt_Change(Index As Integer)
    If Me.Visible Then Call SetParChange(txt, Index, mrsPar)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii = Asc(gstrParSplit1) Or KeyAscii = Asc(gstrParSplit2) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, Index, mrsPar)
End Sub

Private Sub txtUD_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txtUD, Index, mrsPar)
End Sub

Private Sub cbo_GotFocus(Index As Integer)
    Call SetParTip(cbo, Index, mrsPar)
End Sub


Private Sub lst_ItemCheck(Index As Integer, Item As Integer)
    If Me.Visible Then Call SetParChange(lst, Index, mrsPar)
End Sub

Private Sub txtUD_Validate(Index As Integer, Cancel As Boolean)
    If Val(txtUD(Index).Text) > ud(Index).Max Or Val(txtUD(Index).Text) < ud(Index).Min Then
        txtUD(Index).Text = ud(Index).Value
    End If
End Sub

Private Sub txtUD_Change(Index As Integer)
    If Me.Visible Then Call SetParChange(txtUD, Index, mrsPar)
End Sub

Private Sub txtUD_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtUD(Index))
End Sub

Private Sub txtUD_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub cmdAdvice_Click()
    'If frmAdviceDefine.ShowMe(Me, mrsAdvice) Then
        '标记为已变化,需要保存
        cmdAdvice.Tag = "已修改"
    'End If
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub lst_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub



Private Sub cbo_Click(Index As Integer)
    Dim blnValue As Boolean, strValue As String
    
    Select Case Index
    Case cbo_中药配方
        blnValue = True
        strValue = IIF(cbo(cbo_中药配方).ListIndex = 1, 4, 3)
    End Select
    
    If Me.Visible Then
        Call SetParChange(cbo, Index, mrsPar, blnValue, strValue)
    End If
    
End Sub


Private Sub chk_Click(Index As Integer)
    Dim blnValue As Boolean, strValue As String
    
    Select Case Index
        Case chk_过敏登记有效天数
            txtUD(ud_过敏登记有效天数).Enabled = chk(Index).Value = 1
            txtUD(ud_过敏登记有效天数).BackColor = IIF(chk(Index).Value = 1, RGB(255, 255, 255), Me.BackColor)
            ud(ud_过敏登记有效天数).Enabled = txtUD(ud_过敏登记有效天数).Enabled
            strValue = IIF(chk(Index).Value = 1, ud(ud_过敏登记有效天数).Value, "0")
            blnValue = True
    End Select
    
    If Me.Visible Then
        Call SetParChange(chk, Index, mrsPar, blnValue, strValue)
    End If
End Sub

Private Sub cmd住院检查入院诊断_Click(Index As Integer)
    Call SetLstSelected(lst(lst_住院检查入院诊断), Index = 0)
End Sub

Private Sub vsUnWriteDept_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(vsUnWriteDept, 0, mrsPar)
End Sub


Private Sub vsUnWriteDept_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String
    
    If Me.Visible Then
        'strValue = Get不写超量科室
        Call SetParChange(vsUnWriteDept, 0, mrsPar, True, strValue)
    End If
End Sub

Private Sub Save医嘱内容()
'保存医嘱内容定义
    Dim blnTrans As Boolean

    On Error GoTo ErrHandle
    If cmdAdvice.Tag = "已修改" Then
        
        gcnOracle.BeginTrans: blnTrans = True
'        gstrSQL = "zl_医嘱内容定义_Delete"
'        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
'
'        mrsAdvice.Filter = 0
'        Do While Not mrsAdvice.EOF
'            If Not IsNull(mrsAdvice!医嘱内容) Then
'                gstrSQL = "zl_医嘱内容定义_Insert('" & mrsAdvice!诊疗类别 & "','" & Replace(mrsAdvice!医嘱内容, "'", "''") & "')"
'                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
'            End If
'            mrsAdvice.MoveNext
'        Loop
        gcnOracle.CommitTrans: blnTrans = False
        cmdAdvice.Tag = ""
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    If blnTrans Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub Load不写超量科室(ByVal strIn As String)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim lngRow As Long
    Dim lngCol As Long
    
    vsUnWriteDept.Clear
    If strIn = "" Then Exit Sub
    
    strIn = Replace(strIn, "|", ",")
    strSQL = "select id,名称 from 部门表 where id in (Select Column_Value From Table(f_Num2list([1]))) Order by 编码"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIn)
    If rsTmp.EOF Then Exit Sub
    
    With vsUnWriteDept
        lngRow = (rsTmp.RecordCount + 3) \ 4
        If lngRow > 5 Then .Rows = lngRow
        
        For i = 1 To rsTmp.RecordCount
            'Call mcol科室.Add(rsTmp!ID & "", "_" & rsTmp!ID)
            lngRow = (i - 1) \ 4
            lngCol = (i - 1) Mod 4
            
            .TextMatrix(lngRow, lngCol) = rsTmp!名称
            .Cell(flexcpData, lngRow, lngCol) = rsTmp!名称 & ""
            .TextMatrix(lngRow, lngCol + 4) = rsTmp!ID
            rsTmp.MoveNext
        Next
    End With
    
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
