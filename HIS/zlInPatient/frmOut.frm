VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "病人出院"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8520
   Icon            =   "frmOut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleMode       =   0  'User
   ScaleWidth      =   6712.303
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7185
      TabIndex        =   18
      Top             =   795
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7185
      TabIndex        =   17
      Top             =   360
      Width           =   1100
   End
   Begin VB.Frame fraInfo 
      Height          =   5535
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   7035
      Begin VB.ComboBox cbo出院情况 
         Height          =   300
         Left            =   5790
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1020
         Width           =   1110
      End
      Begin VB.CheckBox chk随诊 
         Alignment       =   1  'Right Justify
         Caption         =   "随诊"
         Height          =   195
         Left            =   2955
         TabIndex        =   13
         Top             =   5130
         Width           =   660
      End
      Begin VB.TextBox txt随诊 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4290
         MaxLength       =   3
         TabIndex        =   14
         Top             =   5070
         Width           =   525
      End
      Begin VB.CheckBox chk尸检 
         Alignment       =   1  'Right Justify
         Caption         =   "尸检"
         Enabled         =   0   'False
         Height          =   195
         Left            =   5910
         TabIndex        =   12
         Top             =   4740
         Width           =   660
      End
      Begin VB.TextBox txt床号 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   630
         Width           =   2055
      End
      Begin VB.TextBox txt住院号 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   630
         Width           =   2130
      End
      Begin VB.TextBox txt年龄 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5685
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txt性别 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox txt姓名 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   2130
      End
      Begin VB.TextBox txt出院诊断 
         Height          =   300
         Left            =   960
         MaxLength       =   200
         TabIndex        =   2
         Top             =   1020
         Width           =   3900
      End
      Begin VB.CheckBox chk疑诊 
         Alignment       =   1  'Right Justify
         Caption         =   "确诊"
         Height          =   195
         Left            =   2955
         TabIndex        =   9
         Top             =   4740
         Width           =   660
      End
      Begin VB.TextBox txt中医诊断 
         Height          =   300
         Left            =   960
         MaxLength       =   200
         TabIndex        =   5
         Top             =   2820
         Width           =   3900
      End
      Begin VB.ComboBox cbo中医出院情况 
         Height          =   300
         Left            =   5790
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2820
         Width           =   1110
      End
      Begin VB.ComboBox cbo出院方式 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   4680
         Width           =   1830
      End
      Begin VB.ComboBox cbo随诊 
         Height          =   300
         ItemData        =   "frmOut.frx":030A
         Left            =   5040
         List            =   "frmOut.frx":031D
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   5070
         Width           =   1095
      End
      Begin MSComCtl2.UpDown UD随诊 
         Height          =   300
         Left            =   4800
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   5070
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txt随诊"
         BuddyDispid     =   196614
         OrigLeft        =   3945
         OrigTop         =   645
         OrigRight       =   4185
         OrigBottom      =   930
         Max             =   99999
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   300
         Left            =   960
         TabIndex        =   11
         Top             =   5070
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   19
         Format          =   "yyyy-MM-dd hh:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VSFlex8Ctl.VSFlexGrid vfg西医 
         Height          =   1335
         Left            =   960
         TabIndex        =   4
         Top             =   1400
         Width           =   5895
         _cx             =   10398
         _cy             =   2355
         Appearance      =   1
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
         BackColorSel    =   16635590
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   280
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
      Begin VSFlex8Ctl.VSFlexGrid vfg中医 
         Height          =   1335
         Left            =   960
         TabIndex        =   7
         Top             =   3200
         Width           =   5895
         _cx             =   10398
         _cy             =   2355
         Appearance      =   1
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
         BackColorSel    =   16635590
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   280
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
      Begin MSMask.MaskEdBox txtOkDate 
         Height          =   300
         Left            =   3720
         TabIndex        =   10
         Top             =   4680
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   19
         Format          =   "yyyy-MM-dd hh:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl中医其它 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "其它诊断"
         Height          =   180
         Left            =   180
         TabIndex        =   37
         Top             =   3240
         Width           =   720
      End
      Begin VB.Label lbl西医其它 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "其它诊断"
         Height          =   180
         Left            =   180
         TabIndex        =   36
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院情况"
         Height          =   180
         Left            =   4995
         TabIndex        =   35
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院时间"
         Height          =   180
         Left            =   180
         TabIndex        =   34
         Top             =   5130
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "期限"
         Height          =   180
         Left            =   3900
         TabIndex        =   33
         Top             =   5130
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前床位"
         Height          =   180
         Left            =   3975
         TabIndex        =   32
         Top             =   690
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   180
         Left            =   5205
         TabIndex        =   31
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   180
         Left            =   3345
         TabIndex        =   30
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   525
         TabIndex        =   29
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "住院号"
         Height          =   180
         Left            =   360
         TabIndex        =   28
         Top             =   690
         Width           =   540
      End
      Begin VB.Label lbl出院诊断 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院诊断"
         Height          =   180
         Left            =   180
         TabIndex        =   27
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label lbl中医诊断 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "中医诊断"
         Height          =   180
         Left            =   180
         TabIndex        =   26
         Top             =   2880
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院情况"
         Height          =   180
         Left            =   4995
         TabIndex        =   25
         Top             =   2880
         Width           =   720
      End
      Begin VB.Label lbl出院方式 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院方式"
         Height          =   180
         Left            =   180
         TabIndex        =   24
         Top             =   4740
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   7185
      TabIndex        =   0
      Top             =   4950
      Width           =   1100
   End
End
Attribute VB_Name = "frmOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Public mstrPrivs As String
Public mlng病人ID As Long, mlng主页ID As Long

Private mrsPatiInfo As ADODB.Recordset
Private mint默认诊断 As Integer
Private mblnOutDeath As Boolean
Private mstrOldName As String
Private mdteDeathDate As Date
Private mintDeath As Integer
Private mfrmParent As Object

Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

Private Sub cbo出院方式_Click()
    Dim i As Integer
    If InStr(cbo出院方式.Text, "死亡") > 0 Then
        txt随诊.Text = ""
        chk随诊.Value = 0

        txt随诊.Enabled = (chk随诊.Value = 1)
        UD随诊.Enabled = txt随诊.Enabled

        chk随诊.Enabled = False
    
        chk尸检.Enabled = True
    Else
        chk随诊.Enabled = True
        
        chk尸检.Value = 0
        chk尸检.Enabled = False
        If mrsPatiInfo Is Nothing Then Exit Sub
        If mrsPatiInfo.RecordCount = 0 Then Exit Sub
         '49163,刘鹏飞,2012-09-07,增加随诊标志和随诊期限
        If Not IsNull(mrsPatiInfo!随诊期限) Then
            chk随诊.Value = 1
            txt随诊.Text = Nvl(mrsPatiInfo!随诊期限)
            i = cbo.FindIndex(cbo随诊, Decode(Val(Nvl(mrsPatiInfo!随诊标志)), 1, "月", 2, "年", 3, "周", 4, "天", 9, "终身", ""), True)
            If i <> -1 Then cbo随诊.ListIndex = i
            
            txt随诊.Enabled = False
            UD随诊.Enabled = False
            chk随诊.Enabled = False
            cbo随诊.Enabled = False
        End If
    End If
End Sub

Private Sub cbo出院情况_Click()
    Dim i As Integer
    If InStr(cbo出院情况.Text, "死亡") > 0 Then
        i = cbo.FindIndex(cbo出院方式, "死亡", True)
        If i <> -1 Then cbo出院方式.ListIndex = i
    End If
End Sub

Private Sub cbo随诊_Click()
    txt随诊.Enabled = (cbo随诊.ItemData(cbo随诊.ListIndex) <> 9)
    UD随诊.Enabled = txt随诊.Enabled
End Sub

Private Sub cbo中医出院情况_Click()
    Dim i As Integer
    If InStr(cbo中医出院情况.Text, "死亡") > 0 Then
        i = cbo.FindIndex(cbo出院方式, "死亡", True)
        If i <> -1 Then cbo出院方式.ListIndex = i
    End If
End Sub

Private Sub chk随诊_Click()
    txt随诊.Enabled = (chk随诊.Value = 1)
    UD随诊.Enabled = txt随诊.Enabled
    cbo随诊.Enabled = txt随诊.Enabled
    Call zlCommFun.PressKey(vbKeyTab)
End Sub
'问题28982 by lesfeng 2010-06-09
Private Sub chk疑诊_Click()
    If chk疑诊.Value = 1 Then
        txtOkDate.Enabled = True
    Else
        txtOkDate.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '问题28139 by lesfeng 2010-03-02 增加其它诊断处理
    If KeyCode = 13 Then
        If Not Me.ActiveControl Is txt出院诊断 _
            And Not Me.ActiveControl Is txt中医诊断 And Not Me.ActiveControl Is vfg西医 And Not Me.ActiveControl Is vfg中医 Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '问题28139 by lesfeng 2010-03-02 增加其它诊断处理
    If KeyAscii = Asc("'") And Not (Me.ActiveControl Is txt出院诊断 Or Me.ActiveControl Is txt中医诊断 Or Me.ActiveControl Is vfg西医 Or Me.ActiveControl Is vfg中医) Then KeyAscii = 0       '诊断内容中可能有'号
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, strSQL As String, str床号 As String, str房间号 As String
    Dim dMax As Date, int原因 As Integer, int险类 As Integer
    Dim rsDiagnosis As ADODB.Recordset
    Dim str出院情况 As String, str中医出院情况 As String, strTmp As String
    '问题28138 by lesfeng 2010-03-01
    mint默认诊断 = Val(zlDatabase.GetPara("默认诊断", glngSys, glngModul))
    '63706:刘鹏飞,2014-08-11
    mblnOutDeath = (Val(zlDatabase.GetPara("出院死亡", glngSys, glngModul)) = 1)
    '问题28612 by lesfeng 2010-07-05
    mintDeath = 0
    mdteDeathDate = GetdeathTime(mlng病人ID, mlng主页ID)

    gblnOK = False
    Set mrsPatiInfo = GetPatiInfo(mlng病人ID, mlng主页ID)
    int险类 = Val("" & mrsPatiInfo!险类)
    
    '医保病人出院检查
    If int险类 <> 0 Then '是否允许未结清出院
        If Not gclsInsure.GetCapability(support未结清出院, mlng病人ID, int险类) Then
            Set rsTmp = GetMoneyInfo(mlng病人ID, , , , , , , mlng主页ID)
            If Not rsTmp Is Nothing Then
                If Nvl(rsTmp!费用余额, 0) <> 0 Then
                    MsgBox "该保险病人的费用尚未结清,请先结帐后再出院！", vbInformation, gstrSysName
                    Unload Me: Exit Sub
                End If
            End If
        End If
    End If

    '黑名单提醒
    strTmp = inBlackList(mlng病人ID)
    If strTmp <> "" Then
        If MsgBox("该病人在特殊病人名单中。" & vbCrLf & vbCrLf & "原因：" & vbCrLf & vbCrLf & "　　" & strTmp & vbCrLf & vbCrLf & "要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Unload Me: Exit Sub
        End If
    End If
    
    If gbln医生允许才能出院 Then
        If Not Check医生下达出院医嘱(mlng病人ID, mlng主页ID) Then
            MsgBox "医生对病人下达出院(或转院、死亡)医嘱后才允许出院！", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
    End If
    
    '病人存在销帐申请时不允许办理出院
    If gbln在院病人不准出院结帐 Then
        strTmp = ""
        '56323:刘鹏飞,2013-02-18,加强销帐为审核单据的提示信息内容
        strTmp = Check销帐申请(mlng病人ID, mlng主页ID)
        If strTmp <> "" Then
            MsgBox "该病人存在以下未审核的销帐申请单据：" & vbCrLf & vbCrLf & strTmp & vbCrLf & vbCrLf & "不能办理出院！", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
    End If
    
    With mrsPatiInfo
        txt姓名.Text = !姓名
        txt性别.Text = "" & !性别
        txt年龄.Text = "" & !年龄
        txt住院号.Text = "" & !住院号
        
        txt中医诊断.Enabled = (InStr(1, "," & GetDepCharacter(Val("" & !出院科室id)) & ",", ",中医科,") > 0)
        txt中医诊断.ToolTipText = "只有当病人所在科室的性质为中医科时才允许输入中医诊断!"
        cbo中医出院情况.Enabled = txt中医诊断.Enabled
    End With
        
    Set rsTmp = GetPatiBeds(mlng病人ID)
    str房间号 = ""
    If rsTmp.RecordCount = 0 Then
        str床号 = "家庭病床"
    Else
        Do While Not rsTmp.EOF
            str床号 = str床号 & "," & rsTmp!床号
            If Nvl(rsTmp!床号) = Nvl(mrsPatiInfo!主要床号) And Nvl(rsTmp!科室ID) = Nvl(mrsPatiInfo!入住科室id) Then
                str房间号 = Nvl(rsTmp!房间号)
            End If
            rsTmp.MoveNext
        Loop
        str床号 = Mid(str床号, 2)
    End If
    txt床号.Text = str床号
    txt床号.Tag = str房间号
    
    txtDate.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    dMax = GetMaxDate(mlng病人ID, mlng主页ID, int原因)
    If int原因 = 10 Then
        '59094:刘鹏飞,2013-04-24,修改为只加1s,原来为1m
        txtDate.Text = Format(dMax + 1 / 24 / 60 / 60, "yyyy-MM-dd HH:mm:ss")
    Else
        If dMax > CDate(txtDate.Text) Then
            txtDate.Text = Format(dMax + 1 / 24 / 60, "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    '问题28612 by lesfeng 2010-07-05
    If mintDeath = 1 Then
        txtDate.Text = Format(mdteDeathDate, "yyyy-MM-dd HH:mm:ss")
    End If
    
    '75308,没有调整出院时间权限,则不允许修改出院时间
    If InStr(1, ";" & mstrPrivs & ";", ";调整出院时间;") = 0 Then
        txtDate.Enabled = False
    End If
    '显示病人诊断记录
    Set rsDiagnosis = GetDiagnosticInfo(mlng病人ID, mlng主页ID, "1,11,2,12,3,13", "2,3")
    If Not rsDiagnosis Is Nothing Then
        'a.西医诊断
        rsDiagnosis.Filter = "诊断类型=3 and 记录来源=3"            '先取首页整理的出院诊断
        If Not rsDiagnosis.EOF Then
            txt出院诊断.Text = Nvl(rsDiagnosis!诊断描述): txt出院诊断.Tag = Nvl(rsDiagnosis!疾病ID, rsDiagnosis!诊断ID & ";"): lbl出院诊断.Tag = txt出院诊断.Text
            str出院情况 = "" & rsDiagnosis!出院情况
            '问题28982 by lesfeng 2010-06-09
            chk疑诊.Value = IIf(Val("" & rsDiagnosis!是否疑诊) = 1, 0, 1)
        Else
            '问题28483 by lesfeng 2010-03-01
            rsDiagnosis.Filter = "诊断类型=3 and 记录来源=2"        '再取入院登记的出院诊断
            If Not rsDiagnosis.EOF Then
                txt出院诊断.Text = Nvl(rsDiagnosis!诊断描述): txt出院诊断.Tag = Nvl(rsDiagnosis!疾病ID, rsDiagnosis!诊断ID & ";"): lbl出院诊断.Tag = txt出院诊断.Text
                str出院情况 = "" & rsDiagnosis!出院情况
                '问题28982 by lesfeng 2010-06-09
                chk疑诊.Value = IIf(Val("" & rsDiagnosis!是否疑诊) = 1, 0, 1)
            Else
                '问题28138 by lesfeng 2010-03-01 增加默认诊断的判断 不获取门诊诊断及入院诊断
                If mint默认诊断 = 1 Then
                    rsDiagnosis.Filter = "诊断类型=2 and 记录来源=2"        '再取入院登记的入院诊断
                    If Not rsDiagnosis.EOF Then
                        txt出院诊断.Text = Nvl(rsDiagnosis!诊断描述): txt出院诊断.Tag = Nvl(rsDiagnosis!疾病ID, rsDiagnosis!诊断ID & ";"): lbl出院诊断.Tag = txt出院诊断.Text
                    Else
                        rsDiagnosis.Filter = "诊断类型=1 and 记录来源=2"    '最后取入院登记的门诊诊断
                        If Not rsDiagnosis.EOF Then
                            txt出院诊断.Text = Nvl(rsDiagnosis!诊断描述): txt出院诊断.Tag = Nvl(rsDiagnosis!疾病ID, rsDiagnosis!诊断ID & ";"): lbl出院诊断.Tag = txt出院诊断.Text
                        End If
                    End If
                End If
            End If
        End If
        
        'b.中医诊断
        If txt中医诊断.Enabled Then
            rsDiagnosis.Filter = "诊断类型=13 and 记录来源=3"            '先取首页整理的出院诊断
            If Not rsDiagnosis.EOF Then
                txt中医诊断.Text = Nvl(rsDiagnosis!诊断描述): txt中医诊断.Tag = Nvl(rsDiagnosis!疾病ID, rsDiagnosis!诊断ID & ";"): lbl中医诊断.Tag = txt中医诊断.Text
                str中医出院情况 = "" & rsDiagnosis!出院情况
            Else
                '问题28483 by lesfeng 2010-03-01
                rsDiagnosis.Filter = "诊断类型=13 and 记录来源=2"        '再取入院登记的出院诊断
                If Not rsDiagnosis.EOF Then
                    txt中医诊断.Text = Nvl(rsDiagnosis!诊断描述): txt中医诊断.Tag = Nvl(rsDiagnosis!疾病ID, rsDiagnosis!诊断ID & ";"): lbl中医诊断.Tag = txt中医诊断.Text
                    str中医出院情况 = "" & rsDiagnosis!出院情况
                Else
                    '问题28138 by lesfeng 2010-03-01 增加默认诊断的判断 不获取门诊诊断及入院诊断
                    If mint默认诊断 = 1 Then
                        rsDiagnosis.Filter = "诊断类型=12 and 记录来源=2"        '再取入院登记的入院诊断
                        If Not rsDiagnosis.EOF Then
                            txt中医诊断.Text = Nvl(rsDiagnosis!诊断描述): txt中医诊断.Tag = Nvl(rsDiagnosis!疾病ID, rsDiagnosis!诊断ID & ";"): lbl中医诊断.Tag = txt中医诊断.Text
                        Else
                            rsDiagnosis.Filter = "诊断类型=11 and 记录来源=2"    '最后取入院登记的门诊诊断
                            If Not rsDiagnosis.EOF Then
                                txt中医诊断.Text = Nvl(rsDiagnosis!诊断描述): txt中医诊断.Tag = Nvl(rsDiagnosis!疾病ID, rsDiagnosis!诊断ID & ";"): lbl中医诊断.Tag = txt中医诊断.Text
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    '问题28982 by lesfeng 2010-06-09
    If Not IsNull(mrsPatiInfo!确诊日期) Then
        txtOkDate.Text = Format(mrsPatiInfo!确诊日期, "yyyy-MM-dd HH:mm:ss")
        chk疑诊.Value = IIf(Val("" & mrsPatiInfo!是否确诊) = 1, 1, 0)
        If chk疑诊.Value = 0 Then chk疑诊.Value = 1
        chk疑诊.Enabled = False
        txtOkDate.Enabled = False
    End If
    
    '出院情况
    cbo出院情况.AddItem "": cbo出院情况.ListIndex = cbo出院情况.NewIndex
    If cbo中医出院情况.Enabled Then cbo中医出院情况.AddItem "": cbo中医出院情况.ListIndex = cbo中医出院情况.NewIndex
    
     On Error GoTo errH
    strSQL = "Select 编码,名称,Nvl(缺省标志,0) as 缺省 From 治疗结果 Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo出院情况.AddItem rsTmp!编码 & "-" & rsTmp!名称
            If rsTmp!缺省 = 1 Then
                If txt出院诊断.Text <> "" Then cbo出院情况.ListIndex = cbo出院情况.NewIndex
                cbo出院情况.ItemData(cbo出院情况.NewIndex) = 1
            End If
            
            If cbo中医出院情况.Enabled Then
                cbo中医出院情况.AddItem rsTmp!编码 & "-" & rsTmp!名称
                If rsTmp!缺省 = 1 Then
                    If txt中医诊断.Text <> "" Then cbo中医出院情况.ListIndex = cbo中医出院情况.NewIndex
                    cbo中医出院情况.ItemData(cbo中医出院情况.NewIndex) = 1
                End If
            End If
            rsTmp.MoveNext
        Next
    End If
    Call cbo.Locate(cbo出院情况, str出院情况)
    Call cbo.Locate(cbo中医出院情况, str中医出院情况)
    
    '出院方式
    strSQL = "Select 编码,名称,Nvl(缺省标志,0) as 缺省 From 出院方式 Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo出院方式.AddItem rsTmp!编码 & "-" & rsTmp!名称
            If rsTmp!缺省 = 1 Then cbo出院方式.ListIndex = cbo出院方式.NewIndex
            rsTmp.MoveNext
        Next
    End If
    If (Nvl(mrsPatiInfo!出院方式) <> "") Then
        Call cbo.Locate(cbo出院方式, Nvl(mrsPatiInfo!出院方式))
    End If
    
    '47955:刘鹏飞,2012-09-18,如果存在死亡医嘱，出院方式默认选择"死亡"
    If InStr(cbo出院情况.Text, "死亡") > 0 Or InStr(cbo中医出院情况.Text, "死亡") > 0 Or mintDeath = 1 Then
        i = cbo.FindIndex(cbo出院方式, "死亡", True)
        If i <> -1 Then cbo出院方式.ListIndex = i
    End If
    chk尸检.Value = IIf(Val(Nvl(mrsPatiInfo!尸检标志)) = 1, 1, 0)
    
    cbo随诊.ListIndex = 0
    '49163,刘鹏飞,2012-09-07,增加随诊标志和随诊期限
    If InStr(cbo出院方式.Text, "死亡") = 0 And Not IsNull(mrsPatiInfo!随诊期限) Then
        chk随诊.Value = 1
        txt随诊.Text = Nvl(mrsPatiInfo!随诊期限)
        i = cbo.FindIndex(cbo随诊, Decode(Val(Nvl(mrsPatiInfo!随诊标志)), 1, "月", 2, "年", 3, "周", 4, "天", 9, "终身", ""), True)
        If i <> -1 Then cbo随诊.ListIndex = i
        
        txt随诊.Enabled = False
        UD随诊.Enabled = False
        chk随诊.Enabled = False
        cbo随诊.Enabled = False
    End If
    '问题28139 by lesfeng 2010-03-02
    Call LoadVfgData(vfg西医, 1)
    Call LoadVfgData(vfg中医, 2)
    If chk随诊.Enabled = True Then Call chk随诊_Click
    If chk疑诊.Enabled Then Call chk疑诊_Click
    
    '创建消息对象
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, 1132, mstrPrivs, mfrmParent.hWnd)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim dMax As Date, Curdate As Date, i As Integer
    Dim strSQL As String, strInfo As String, blnTrans As Boolean
    Dim lng西医疾病ID As Long, lng中医疾病ID As Long
    Dim lng西医诊断ID As Long, lng中医诊断ID As Long
    Dim int随诊 As Integer, int险类 As Integer
    Dim strTmp As String, str其它诊断 As String, str其它情况 As String
    Dim int次数 As Integer, int诊断类型 As Integer
    Dim int诊断次序 As Integer
    Dim strICD编码 As String
    Dim str确诊日期  As String
    Dim str入院时间 As String
    Dim lngRow As Long, strRow其它诊断 As String, strRowICD编码 As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    '出院诊断
    If Not CheckLen(txt出院诊断, txt出院诊断.MaxLength) Then Exit Sub
    If Not CheckLen(txt中医诊断, txt中医诊断.MaxLength) Then Exit Sub
    
    int险类 = Val("" & mrsPatiInfo!险类)
    If int险类 <> 0 Then
        If gclsInsure.GetCapability(support必须录入入出诊断, mlng病人ID, int险类) Then
            If txt出院诊断.Text = "" Then
                MsgBox "请填写该病人的出院诊断！", vbInformation, gstrSysName
                txt出院诊断.SetFocus: Exit Sub
            End If
        End If
    End If
    If txt出院诊断.Text <> "" And cbo出院情况.Text = "" Then
        MsgBox "请选择出院诊断的出院情况。", vbInformation, gstrSysName
        cbo出院情况.SetFocus: Exit Sub
    End If
    If txt中医诊断.Text <> "" And cbo中医出院情况.Text = "" And cbo中医出院情况.Enabled Then
        MsgBox "请选择中医出院诊断的出院情况。", vbInformation, gstrSysName
        cbo中医出院情况.SetFocus: Exit Sub
    End If
    
    '问题28139 by lesfeng 2010-03-03 数据判断
    strTmp = Replace(txt出院诊断.Text, "'", "''")
    With vfg西医
        For int次数 = 1 To .Rows - 1
            str其它诊断 = Trim(.TextMatrix(int次数, .ColIndex("诊断描述")))
            str其它情况 = Trim(.TextMatrix(int次数, .ColIndex("出院情况")))
            strICD编码 = Trim(.TextMatrix(int次数, .ColIndex("ICD编码")))
            If str其它诊断 <> "" And strTmp = "" Then
                MsgBox "请填写该病人的出院诊断！才能填写其它诊断！", vbInformation, gstrSysName
                txt出院诊断.SetFocus: Exit Sub
            End If
            If str其它诊断 <> "" And str其它情况 = "" Then
                MsgBox "请选择其它出院诊断的出院情况", vbInformation, gstrSysName
                vfg西医.SetFocus
                .Select int次数, .ColIndex("出院情况")
                Exit Sub
            End If
            If strICD编码 <> "" Then
                str其它诊断 = "(" & strICD编码 & ")" & str其它诊断
            End If
            If str其它诊断 = strTmp And strTmp <> "" Then
                MsgBox "该病人的出院诊断与出院诊断相同！请调整再保存！", vbInformation, gstrSysName
                .Select int次数, .ColIndex("诊断描述")
                Exit Sub
            End If
            '50337:刘鹏飞,2012-09-18,检查其他诊断是否重复
            For lngRow = int次数 + 1 To .Rows - 1
                strRow其它诊断 = Trim(.TextMatrix(lngRow, .ColIndex("诊断描述")))
                strRowICD编码 = Trim(.TextMatrix(lngRow, .ColIndex("ICD编码")))
                If strRowICD编码 <> "" Then
                    strRow其它诊断 = "(" & strRowICD编码 & ")" & strRow其它诊断
                End If
                If str其它诊断 = strRow其它诊断 And str其它诊断 <> "" And strRow其它诊断 <> "" Then
                    MsgBox "该病人的出院其它诊断列表中第" & int次数 & "，" & lngRow & "列的诊断相同，请调整在保存！", vbInformation, gstrSysName
                    .Select lngRow, .ColIndex("诊断描述")
                    Exit Sub
                End If
            Next lngRow
        Next
    End With
    
    If cbo中医出院情况.Enabled Then
        strTmp = Replace(txt中医诊断.Text, "'", "''")
        With vfg中医
            For int次数 = 1 To .Rows - 1
                str其它诊断 = Trim(.TextMatrix(int次数, .ColIndex("诊断描述")))
                str其它情况 = Trim(.TextMatrix(int次数, .ColIndex("出院情况")))
                strICD编码 = Trim(.TextMatrix(int次数, .ColIndex("中医编码")))
                If str其它诊断 <> "" And strTmp = "" Then
                    MsgBox "请填写该病人的出院诊断！才能填写其它诊断！", vbInformation, gstrSysName
                    txt中医诊断.SetFocus: Exit Sub
                End If
                If str其它诊断 <> "" And str其它情况 = "" Then
                    MsgBox "请选择其它出院诊断的出院情况", vbInformation, gstrSysName
                    vfg中医.SetFocus
                    .Select int次数, .ColIndex("出院情况")
                    Exit Sub
                End If
                If strICD编码 <> "" Then
                    str其它诊断 = "(" & strICD编码 & ")" & str其它诊断
                End If
                If str其它诊断 = strTmp And strTmp <> "" Then
                    MsgBox "填写该病人的出院诊断与中医诊断相同！请调整再保存！", vbInformation, gstrSysName
                    .Select int次数, .ColIndex("诊断描述")
                    Exit Sub
                End If
                
                '50337:刘鹏飞,2012-09-18,检查其他诊断是否重复
                For lngRow = int次数 + 1 To .Rows - 1
                    strRow其它诊断 = Trim(.TextMatrix(lngRow, .ColIndex("诊断描述")))
                    strRowICD编码 = Trim(.TextMatrix(lngRow, .ColIndex("中医编码")))
                    If strRowICD编码 <> "" Then
                        strRow其它诊断 = "(" & strRowICD编码 & ")" & strRow其它诊断
                    End If
                    If str其它诊断 = strRow其它诊断 And str其它诊断 <> "" And strRow其它诊断 <> "" Then
                        MsgBox "该病人的中医其它诊断列表中第" & int次数 & "，" & lngRow & "列的诊断相同，请调整在保存！", vbInformation, gstrSysName
                        .Select lngRow, .ColIndex("诊断描述")
                        Exit Sub
                    End If
                Next lngRow
            Next
        End With
    End If
    
    If Not IsDate(txtDate.Text) Then
        MsgBox "请输入正确的病人出院时间！", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    '时间不能超过当前时间太长(一周)
    Curdate = zlDatabase.Currentdate
    If CDate(txtDate.Text) > Curdate Then
        If CDate(txtDate.Text) - Curdate > 7 Then
            MsgBox "出院时间比当前时间大得过多,请检查！", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        If MsgBox("出院时间大于了当前系统时间,确实要出院吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txtDate.SetFocus: Exit Sub
        End If
    End If
    
    '检查病人是否有未执行完成的诊疗项目及未发药品
    strInfo = ""
    If gbyt出院时检查未执行 <> 0 Then
        strInfo = ExistWaitExe(mlng病人ID, mlng主页ID)
        If strInfo <> "" Then
            If gbyt出院时检查未执行 = 1 Then
                If MsgBox("发现病人" & txt姓名.Text & "存在尚未执行完成的内容：" & _
                    vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "确定要出院吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            Else
                MsgBox "发现病人" & txt姓名.Text & "存在尚未执行完成的内容：" & vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "不允许出院.", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    '问题30208 by lesfeng 2010-08-02 撤分参数22及32 新增154、155
    strInfo = ""
    If gbyt出院时检查药品未执行 <> 0 Then
        strInfo = ExistWaitDrug(mlng病人ID, mlng主页ID)
        If strInfo <> "" Then
            If gbyt出院时检查药品未执行 = 1 Then
                If MsgBox("发现病人" & txt姓名.Text & strInfo & vbCrLf & vbCrLf & "确定要出院吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            Else
                MsgBox "发现病人" & txt姓名.Text & strInfo & vbCrLf & vbCrLf & "不允许出院。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        '30339:刘鹏飞,2012-09-14,检查是否发血
        strInfo = ExistWaitBool(mlng病人ID, mlng主页ID)
        If strInfo <> "" Then
            If gbyt出院时检查药品未执行 = 1 Then
                If MsgBox("发现病人" & txt姓名.Text & strInfo & vbCrLf & vbCrLf & "确定要出院吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            Else
                MsgBox "发现病人" & txt姓名.Text & strInfo & vbCrLf & vbCrLf & "不允许出院。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    If GetUnAuditReFee(mlng病人ID, mlng主页ID) Then
        If MsgBox("病人" & txt姓名.Text & "存在已申请退费但未审核的记录,确定要出院吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    dMax = GetMaxDate(mlng病人ID, mlng主页ID)
    If Format(txtDate.Text, "yyyyMMddHHmmss") <= Format(dMax, "yyyyMMddHHmmss") Then
        MsgBox "病人出院时间必须大于该病人上次变动时间 " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & " ！", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    dMax = GetLastAdviceTime(mlng病人ID, mlng主页ID)
    If Format(txtDate.Text, "yyyyMMddHHmmss") < Format(dMax, "yyyyMMddHHmmss") Then
        If MsgBox("出院时间小于该病人最后有效医嘱的时间 " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & ",确实要出院吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txtDate.SetFocus: Exit Sub
        End If
    End If
    '问题28612 by lesfeng 2010-07-05
    If InStr(cbo出院方式.Text, "死亡") = 0 And mintDeath = 1 Then
        If MsgBox("该病人存在有效临床死亡医嘱,其死亡医嘱的时间 " & Format(mdteDeathDate, "yyyy-MM-dd HH:mm:ss") & ",但出院方式不为死亡,确实要出院吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            cbo出院方式.SetFocus: Exit Sub
        End If
    End If
    
    If InStr(cbo出院方式.Text, "死亡") > 0 And mintDeath = 1 Then
        If Format(txtDate.Text, "yyyyMMddHHmmss") <> Format(mdteDeathDate, "yyyyMMddHHmmss") Then
            If MsgBox("出院时间不等于该病人有效临床死亡医嘱的时间 " & Format(mdteDeathDate, "yyyy-MM-dd HH:mm:ss") & ",确实要出院吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                txtDate.SetFocus: Exit Sub
            End If
        End If
    End If
    
    If InStr(cbo出院方式.Text, "死亡") > 0 And mintDeath = 0 And mblnOutDeath = True Then
        MsgBox "该病人出院方式为死亡,但不存在有效临床死亡医嘱,不允许出院!", vbInformation, gstrSysName
        cbo出院方式.SetFocus: Exit Sub
    End If
    
    '68953:刘鹏飞,2012-09-14
    strInfo = ""
    If gbyt出院时超期护理数据检查 <> 0 Then
        strInfo = ExistNurseData(mlng病人ID, mlng主页ID, CDate(Format(txtDate.Text, "YYYY-MM-DD HH:mm:ss")))
        If strInfo <> "" Then
            If strInfo = "OK" Then
                '老版
                If gbyt出院时超期护理数据检查 = 1 Then
                    If MsgBox("发现病人" & txt姓名.Text & "存在出院时间之后的护理数据，确定要出院吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                Else
                    MsgBox "发现病人" & txt姓名.Text & "存在出院时间之后的护理数据，不允许出院.", vbInformation, gstrSysName
                    Exit Sub
                End If
            Else
                '新版
                If gbyt出院时超期护理数据检查 = 1 Then
                    If MsgBox("发现病人" & txt姓名.Text & "存在出院时间之后的护理数据：" & _
                        vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "确定要出院吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                Else
                    MsgBox "发现病人" & txt姓名.Text & "存在出院时间之后的护理数据：" & vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "不允许出院.", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
    End If
    
    '问题28982 by lesfeng 2010-06-09
    str确诊日期 = ""
    If chk疑诊.Value = 1 Then
        str入院时间 = Format(mrsPatiInfo!入院时间, "yyyy-MM-dd HH:mm:ss")
        If Not IsDate(txtOkDate.Text) Then
            MsgBox "请输入正确的病人确诊时间！", vbInformation, gstrSysName
            If txtOkDate.Enabled Then txtOkDate.SetFocus: Exit Sub
        End If
        If Format(txtOkDate.Text, "yyyyMMddHHmmss") >= Format(txtDate.Text, "yyyyMMddHHmmss") Then
            MsgBox "确诊时间必须小于病人出院时间 " & Format(txtDate.Text, "yyyy-MM-dd HH:mm:ss") & " ！", vbInformation, gstrSysName
            If txtOkDate.Enabled Then txtOkDate.SetFocus: Exit Sub
        End If
        
        If Format(str入院时间, "yyyyMMddHHmmss") > Format(txtOkDate.Text, "yyyyMMddHHmmss") Then
            MsgBox "确诊时间必须大于等于病人入院时间 " & Format(str入院时间, "yyyy-MM-dd HH:mm:ss") & " ！", vbInformation, gstrSysName
            If txtOkDate.Enabled Then txtOkDate.SetFocus: Exit Sub
        End If
        
        str确诊日期 = Format(txtOkDate.Text, "yyyy-MM-dd HH:mm:ss")
    End If
   
    If cbo随诊.ListIndex <> -1 Then int随诊 = cbo随诊.ItemData(cbo随诊.ListIndex)
    
    If InStr(1, txt出院诊断.Tag, ";") <= 0 Then
        lng西医疾病ID = Val(txt出院诊断.Tag)
    Else
        lng西医诊断ID = Val(txt出院诊断.Tag)
    End If
    If InStr(1, txt中医诊断.Tag, ";") <= 0 Then
        lng中医疾病ID = Val(txt中医诊断.Tag)
    Else
        lng中医诊断ID = Val(txt中医诊断.Tag)
    End If
    '问题28982 by lesfeng 2010-06-09 增加确诊日期
    strSQL = "zl_病人变动记录_Out(" & mlng病人ID & "," & mlng主页ID & "," & _
        ZVal(lng西医疾病ID) & "," & ZVal(lng西医诊断ID) & ",'" & Replace(txt出院诊断.Text, "'", "''") & "','" & zlCommFun.GetNeedName(cbo出院情况.Text) & "'," & _
        ZVal(lng中医疾病ID) & "," & ZVal(lng中医诊断ID) & ",'" & Replace(txt中医诊断.Text, "'", "''") & "','" & zlCommFun.GetNeedName(cbo中医出院情况.Text) & "'," & _
        chk疑诊.Value & ",'" & zlCommFun.GetNeedName(cbo出院方式.Text) & "',To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS')," & _
        IIf(chk随诊.Value = 1, int随诊, 0) & "," & IIf(chk随诊.Value = 1 And int随诊 <> 9, Val(txt随诊.Text), "Null") & "," & IIf(chk尸检.Enabled, chk尸检.Value, "NULL") & "," & _
        "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & IIf(str确诊日期 = "", "NULL", "To_Date('" & str确诊日期 & "','YYYY-MM-DD HH24:MI:SS')") & ")"
    
    gcnOracle.BeginTrans
    blnTrans = True
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
     '问题28139 by lesfeng 2010-03-03 数据判断
    With vfg西医
        For int次数 = 1 To .Rows - 1
            str其它诊断 = Trim(.TextMatrix(int次数, .ColIndex("诊断描述")))
            str其它情况 = Trim(.TextMatrix(int次数, .ColIndex("出院情况")))
            
            lng西医疾病ID = Val(.TextMatrix(int次数, .ColIndex("疾病ID")))
            lng西医诊断ID = Val(.TextMatrix(int次数, .ColIndex("诊断ID")))
            int诊断类型 = 3
            int诊断次序 = int次数 + 1
            If str其它诊断 <> "" Then
                '病人id,主页id,诊断类型,诊断次序,疾病id,诊断id,出院情况,描述信息,是否疑诊
                strSQL = "Zl_病人诊断情况_Other(" & mlng病人ID & "," & mlng主页ID & "," & int诊断类型 & "," & int诊断次序 & _
                        "," & ZVal(lng西医疾病ID) & "," & ZVal(lng西医诊断ID) & _
                        ",'" & zlCommFun.GetNeedName(str其它情况) & "','" & Replace(str其它诊断, "'", "’") & _
                        "'" & IIf(.TextMatrix(int次数, .ColIndex("疑诊")) <> "", ",1", ",0") & ",'" & UserInfo.姓名 & "')"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
        Next
    End With
    If cbo中医出院情况.Enabled Then
        With vfg中医
            For int次数 = 1 To .Rows - 1
                str其它诊断 = Trim(.TextMatrix(int次数, .ColIndex("诊断描述")))
                str其它情况 = Trim(.TextMatrix(int次数, .ColIndex("出院情况")))
                lng中医疾病ID = Val(.TextMatrix(int次数, .ColIndex("疾病ID")))
                lng中医诊断ID = Val(.TextMatrix(int次数, .ColIndex("诊断ID")))
                int诊断类型 = 13
                int诊断次序 = int次数 + 1
                If str其它诊断 <> "" Then
                    strSQL = "Zl_病人诊断情况_Other(" & mlng病人ID & "," & mlng主页ID & "," & int诊断类型 & "," & int诊断次序 & _
                        "," & ZVal(lng中医疾病ID) & "," & ZVal(lng中医诊断ID) & _
                        ",'" & zlCommFun.GetNeedName(str其它情况) & "','" & Replace(str其它诊断, "'", "’") & "',0,'" & UserInfo.姓名 & "')"
                    zlDatabase.ExecuteProcedure strSQL, Me.Caption
                End If
            Next
        End With
    End If
    
    '医保改动
    If int险类 <> 0 Then
        If Not gclsInsure.LeaveSwap(mlng病人ID, mlng主页ID, int险类) Then
            gcnOracle.RollbackTrans: Exit Sub
        End If
    End If
    gcnOracle.CommitTrans
    blnTrans = False
    
    On Error Resume Next
    '出院成功后触发此消息
    If mclsMipModule.IsConnect = True Then
        mclsXML.ClearXmlText '清除缓存中的XML
        '--进行消息组装
        '病人信息
        mclsXML.AppendNode "in_patient"
        'patient_id      病人id  1   N
        mclsXML.appendData "patient_id", mlng病人ID, xsNumber  '病人ID
        'page_id     主页id  1   N
        mclsXML.appendData "page_id", mlng主页ID, xsNumber '主页ID
        'patient_name        姓名    1   S
        mclsXML.appendData "patient_name", txt姓名.Text, xsString  '姓名
        'patient_sex     性别    0..1    S
        mclsXML.appendData "patient_sex", txt性别.Text, xsString  '性别
        'in_number       住院号  1   S
        mclsXML.appendData "in_number", txt住院号.Text, xsString '住院号
        mclsXML.AppendNode "in_patient", True
        
        strSQL = "Select ID 变动id,开始时间 变动时间 From 病人变动记录 where 病人ID=[1] And 主页Id=[2] And 终止原因=1 And 终止时间 IS NOT NULL And NVL(附加床位,0)=0 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病人变动记录", mlng病人ID, mlng主页ID)
        
        'out_hospital        病人出院    1
        mclsXML.AppendNode "out_hospital"
        'change_id       出院变更id  1   N
        mclsXML.appendData "change_id", rsTmp!变动ID, xsNumber
        'out_date        变更时间    1   s
        mclsXML.appendData "out_date", Format(rsTmp!变动时间, "YYYY-MM-DD HH:mm:ss"), xsString
        'out_area_id     当前病区id  0..1    N
        mclsXML.appendData "out_area_id", Nvl(mrsPatiInfo!当前病区ID, 0), xsNumber
        'out_area_title      当前病区    0..1    S
        mclsXML.appendData "out_area_title", Nvl(mrsPatiInfo!当前病区), xsString
        'out_dept_id     当前科室id    1   N
        mclsXML.appendData "out_dept_id", Nvl(mrsPatiInfo!出院科室id, 0), xsNumber
        'out_dept_title      当前科室  1   S
        mclsXML.appendData "out_dept_title", Nvl(mrsPatiInfo!当前科室), xsString
        'out_room        当前病房    0..1    S
        mclsXML.appendData "out_room", txt床号.Tag, xsString
        'out_bed     当前病床    1   S
        mclsXML.appendData "out_bed", Nvl(mrsPatiInfo!主要床号), xsString
        'out_way     出院方式    1   S
        mclsXML.appendData "out_way", zlCommFun.GetNeedName(cbo出院方式.Text), xsString
        'treat_state     治愈情况    1   S
        mclsXML.appendData "treat_state", zlCommFun.GetNeedName(cbo出院情况.Text), xsString
        
        mclsXML.AppendNode "out_hospital", True
        mclsMipModule.CommitMessage "ZLHIS_PATIENT_010", mclsXML.XmlText
    End If
    If Err <> 0 Then Err.Clear
    
     '调用外挂接口
    If CreatePlugInOK(glngModul) Then
        On Error Resume Next
        Call gobjPlugIn.InPatiCheckOutAfter(mlng病人ID, mlng主页ID)
        Call zlPlugInErrH(Err, "InPatiCheckOutAfter")
        Err.Clear: On Error GoTo 0
    End If
    
    On Error GoTo errH
    
    Dim strOut As String
    Call zlExcuteUploadSwap(mlng病人ID, strOut) '发卡了调用宁波一卡通上传功能
    
    '出院后自动计算病人的床位费用和护理费用(如果放在出院前执行，则当使用半天模式时会多算半天费用)
    strSQL = "ZL1_AUTOCPTPATI(" & mlng病人ID & "," & mlng主页ID & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
       
    gblnOK = True
    
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '卸载消息对象
    If Not (mclsMipModule Is Nothing) Then
        Call mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    If Not (mclsXML Is Nothing) Then
        Set mclsXML = Nothing
    End If
    
    Call SaveHead(vfg西医, 1)
    Call SaveHead(vfg中医, 2)
End Sub

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_LostFocus()
    If Not IsDate(txtDate.Text) Then txtDate.SetFocus
End Sub

Private Sub txt随诊_GotFocus()
    zlControl.TxtSelAll txt随诊
End Sub

Private Sub txt随诊_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt出院诊断_GotFocus()
    zlControl.TxtSelAll txt出院诊断
End Sub

Private Sub txt中医诊断_GotFocus()
    zlControl.TxtSelAll txt中医诊断
End Sub

Private Sub txt出院诊断_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String, lngTxtHeight As Long, vPoint As POINTAPI
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not RequestCode Then
            '问题25785 by lesfeng 2009-10-20 处理允许自由录入规则
            '************************************************
            If gint住院诊断输入 = 1 Then
                strInput = UCase(txt出院诊断.Text)
                strSex = zlCommFun.GetNeedName(txt性别.Text)
                
                If zlCommFun.IsCharChinese(strInput) Then
                    strSQL = "名称 Like [2] or '('||编码||')'||名称 Like [2]" '输入汉字时只匹配名称
                Else
                    strSQL = "编码 Like [1] Or 名称 Like [2] Or " & IIf(gbytCode = 0, "简码", "五笔码") & " Like [2]"
                End If
                
                strSQL = _
                        " Select ID,ID as 项目ID,编码,附码,名称," & IIf(gbytCode = 0, "简码", "五笔码 as 简码") & ",说明" & _
                        " From 疾病编码目录 Where Instr([3],类别)>0 And (" & strSQL & ")" & _
                        IIf(strSex <> "", " And (性别限制=[4] Or 性别限制 is NULL)", "") & _
                        " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " Order by 编码"
                '问题27613 by lesfeng 2010-01-21
                '自由录入时有多个匹配(汉字)不进行选择,数字及字母则进行选择
                If zlCommFun.IsCharChinese(strInput) Then
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", gstrLike & strInput & "%", "D", strSex, gbytCode + 1)
                    If rsTmp.EOF Then
                        Set rsTmp = Nothing
                    ElseIf rsTmp.RecordCount > 1 Then
                        Set rsTmp = Nothing '自由录入时有多个匹配不进行选择
                    End If
                Else
                    vPoint = zlControl.GetCoordPos(fraInfo.hWnd, txt出院诊断.Left, txt出院诊断.Top)
                    strInput = UCase(txt出院诊断.Text)
                    strSex = zlCommFun.GetNeedName(txt性别.Text)
                    lngTxtHeight = txt出院诊断.Height
                    Set rsTmp = GetDiseaseCodeNew(Me, blnCancel, strInput, strSex, "D", vPoint.X, vPoint.Y, lngTxtHeight)
                    If Not rsTmp Is Nothing Then
                        If rsTmp.EOF Then
                            Set rsTmp = Nothing
                        End If
                    End If
                End If
                If Not rsTmp Is Nothing Then
                    '数据库中只有一个匹配项目，则以该匹配的项目为准
                    txt出院诊断.Tag = rsTmp!ID
                    txt出院诊断.Text = "(" & rsTmp!编码 & ")" & rsTmp!名称 '
                    lbl出院诊断.Tag = txt出院诊断.Text '用于恢复显示
                Else
                    '多项或者无匹配项目时才以输入的为准
                    txt出院诊断.Tag = ""
                    lbl出院诊断.Tag = txt出院诊断.Text '用于恢复显示
                End If
            End If
            '************************************************
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt出院诊断.Text = lbl出院诊断.Tag And txt出院诊断.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt出院诊断.Text = "" Then
            txt出院诊断.Tag = "": lbl出院诊断.Tag = ""
            Call zlCommFun.PressKey(vbKeyTab) '允许不输入
        Else
            vPoint = zlControl.GetCoordPos(fraInfo.hWnd, txt出院诊断.Left, txt出院诊断.Top)
            strInput = UCase(txt出院诊断.Text)
            strSex = zlCommFun.GetNeedName(txt性别.Text)
            lngTxtHeight = txt出院诊断.Height
            Set rsTmp = GetDiseaseCode(Me, blnCancel, strInput, strSex, "D", vPoint.X, vPoint.Y, lngTxtHeight)
            If Not rsTmp Is Nothing Then
                txt出院诊断.Tag = rsTmp!ID
                txt出院诊断.Text = "(" & rsTmp!编码 & ")" & rsTmp!名称
                lbl出院诊断.Tag = txt出院诊断.Text '用于恢复显示
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If Not blnCancel Then
                    MsgBox "没有找到匹配的疾病编码。", vbInformation, gstrSysName
                End If
                If lbl出院诊断.Tag <> "" Then txt出院诊断.Text = lbl出院诊断.Tag
                Call txt出院诊断_GotFocus
                txt出院诊断.SetFocus
            End If
        End If
    Else
        CheckInputLen txt出院诊断, KeyAscii
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt中医诊断_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String, lngTxtHeight As Long, vPoint As POINTAPI
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not RequestCode Then
            '问题25785 by lesfeng 2009-10-20 处理允许自由录入规则
            '************************************************
            If gint住院诊断输入 = 1 Then
                strInput = UCase(txt中医诊断.Text)
                strSex = zlCommFun.GetNeedName(txt性别.Text)
                
                If zlCommFun.IsCharChinese(strInput) Then
                    strSQL = "名称 Like [2] or '('||编码||')'||名称 Like [2]" '输入汉字时只匹配名称
                Else
                    strSQL = "编码 Like [1] Or 名称 Like [2] Or " & IIf(gbytCode = 0, "简码", "五笔码") & " Like [2]"
                End If
                
                strSQL = _
                        " Select ID,ID as 项目ID,编码,附码,名称," & IIf(gbytCode = 0, "简码", "五笔码 as 简码") & ",说明" & _
                        " From 疾病编码目录 Where Instr([3],类别)>0 And (" & strSQL & ")" & _
                        IIf(strSex <> "", " And (性别限制=[4] Or 性别限制 is NULL)", "") & _
                        " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " Order by 编码"
                '问题27613 by lesfeng 2010-01-21
                '自由录入时有多个匹配(汉字)不进行选择,数字及字母则进行选择
                If zlCommFun.IsCharChinese(strInput) Then
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", gstrLike & strInput & "%", "B", strSex, gbytCode + 1)
                    If rsTmp.EOF Then
                        Set rsTmp = Nothing
                    ElseIf rsTmp.RecordCount > 1 Then
                        Set rsTmp = Nothing '自由录入时有多个匹配不进行选择
                    End If
                Else
                    vPoint = zlControl.GetCoordPos(fraInfo.hWnd, txt中医诊断.Left, txt中医诊断.Top)
                    strInput = UCase(txt中医诊断.Text)
                    strSex = zlCommFun.GetNeedName(txt性别.Text)
                    lngTxtHeight = txt中医诊断.Height
                    Set rsTmp = GetDiseaseCodeNew(Me, blnCancel, strInput, strSex, "B", vPoint.X, vPoint.Y, lngTxtHeight)
                    If Not rsTmp Is Nothing Then
                        If rsTmp.EOF Then
                            Set rsTmp = Nothing
                        End If
                    End If
                End If
                If Not rsTmp Is Nothing Then
                    '数据库中只有一个匹配项目，则以该匹配的项目为准
                    txt中医诊断.Tag = rsTmp!ID
                    txt中医诊断.Text = "(" & rsTmp!编码 & ")" & rsTmp!名称 '
                    lbl中医诊断.Tag = txt中医诊断.Text '用于恢复显示
                Else
                    '多项或者无匹配项目时才以输入的为准
                    txt中医诊断.Tag = ""
                    lbl中医诊断.Tag = txt中医诊断.Text '用于恢复显示
                End If
            End If
            '************************************************
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt中医诊断.Text = lbl中医诊断.Tag And txt中医诊断.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt中医诊断.Text = "" Then
            txt中医诊断.Tag = "": lbl中医诊断.Tag = ""
            Call zlCommFun.PressKey(vbKeyTab) '允许不输入
        Else
            vPoint = zlControl.GetCoordPos(fraInfo.hWnd, txt中医诊断.Left, txt中医诊断.Top)
            strInput = UCase(txt中医诊断.Text)
            strSex = zlCommFun.GetNeedName(txt性别.Text)
            lngTxtHeight = txt中医诊断.Height
            Set rsTmp = GetDiseaseCode(Me, blnCancel, strInput, strSex, "B", vPoint.X, vPoint.Y, lngTxtHeight)
            
            If Not rsTmp Is Nothing Then
                txt中医诊断.Tag = rsTmp!ID
                txt中医诊断.Text = "(" & rsTmp!编码 & ")" & rsTmp!名称
                lbl中医诊断.Tag = txt中医诊断.Text '用于恢复显示
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If Not blnCancel Then
                    MsgBox "没有找到匹配的疾病编码。", vbInformation, gstrSysName
                End If
                If lbl中医诊断.Tag <> "" Then txt中医诊断.Text = lbl中医诊断.Tag
                Call txt中医诊断_GotFocus
                txt中医诊断.SetFocus
            End If
        End If
    Else
        CheckInputLen txt中医诊断, KeyAscii
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt出院诊断_Validate(Cancel As Boolean)
    If Val(txt出院诊断.Tag) > 0 And txt出院诊断.Text <> lbl出院诊断.Tag Then
        txt出院诊断.Text = lbl出院诊断.Tag
    ElseIf Val(txt出院诊断.Tag) = 0 And RequestCode Then
        txt出院诊断.Text = ""
    End If
    
    If txt出院诊断.Text <> "" And cbo出院情况.Text = "" Then
        cbo出院情况.ListIndex = cbo.FindIndex(cbo出院情况, 1)
        If cbo出院情况.ListIndex = -1 Then cbo出院情况.ListIndex = 0
    ElseIf txt出院诊断.Text = "" And cbo出院情况.Text <> "" Then
        cbo出院情况.ListIndex = 0
    End If
End Sub

Private Sub txt中医诊断_Validate(Cancel As Boolean)
    If Val(txt中医诊断.Tag) > 0 And txt中医诊断.Text <> lbl中医诊断.Tag Then
        txt中医诊断.Text = lbl中医诊断.Tag
    ElseIf Val(txt中医诊断.Tag) = 0 And RequestCode Then
        txt中医诊断.Text = ""
    End If
    
    If txt中医诊断.Text <> "" And cbo中医出院情况.Text = "" Then
        cbo中医出院情况.ListIndex = cbo.FindIndex(cbo中医出院情况, 1)
        If cbo中医出院情况.ListIndex = -1 Then cbo中医出院情况.ListIndex = 0
    ElseIf txt中医诊断.Text = "" And cbo中医出院情况.Text <> "" Then
        cbo中医出院情况.ListIndex = 0
    End If
End Sub

Private Function RequestCode() As Boolean
    RequestCode = gint住院诊断输入 = 2 Or (gint住院诊断输入 = 3 And Val("" & mrsPatiInfo!险类) <> 0)
End Function

'问题28139 by lesfeng 2010-03-02
Private Sub initvfgHeadTitle(ByVal vsGrid As VSFlexGrid, ByVal intFlag As Integer)
    Dim strHead As String
    If intFlag = 1 Then
        strHead = "序号,500,4,1;诊断描述,2200,1,1;ICD编码,1000,1,1;出院情况,1000,1,1;疑诊,800,4,0;诊断ID,0,1,-1;疾病ID,0,1,-1"
    Else
        strHead = "序号,500,4,1;诊断描述,2800,1,1;中医编码,1200,1,1;出院情况,1000,1,1;诊断ID,0,1,-1;疾病ID,0,1,-1"
    End If
        Call SetVsFlexGridChangeHead(strHead, vsGrid, 1)
End Sub

Private Sub SetVfgNo(ByVal vsGrid As VSFlexGrid)
    Dim i As Long
    With vsGrid
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("诊断描述"))) <> "" Then
                .TextMatrix(i, .ColIndex("序号")) = i
            End If
        Next
    End With
End Sub

Private Sub SetInitVfgFormat(ByVal vsGrid As VSFlexGrid, ByVal intFlag As Integer)
    Dim i As Long
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "Select 编码,名称,编码||'-'||Nvl(名称,'') as 项目,Nvl(缺省标志,0) as 缺省 From 治疗结果 Order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

    With vsGrid
        .ColComboList(.ColIndex("出院情况")) = .BuildComboList(rsTemp, "项目", "编码")
        If rsTemp.RecordCount = 1 Then
            .ColData(.ColIndex("出院情况")) = Nvl(rsTemp!编码) & ";" & Nvl(rsTemp!项目)
        Else
            rsTemp.Filter = "缺省=1"
            If rsTemp.EOF = False Then
                .ColData(.ColIndex("出院情况")) = Nvl(rsTemp!编码) & ";" & Nvl(rsTemp!项目)
            Else
                .ColData(.ColIndex("出院情况")) = ";"
            End If
        End If
        .ExplorerBar = flexExSortShowAndMove
        .SelectionMode = flexSelectionByRow
        If (txt中医诊断.Enabled And intFlag = 2) Or intFlag = 1 Then
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If
    End With
    rsTemp.Close
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadVfgData(ByVal vsGrid As VSFlexGrid, ByVal intFlag As Integer)
    Dim strTemp As String
    Dim i As Long
    Dim rsDiagnosisOther As ADODB.Recordset

    If intFlag = 1 Then
        Set rsDiagnosisOther = GetDiagnosticOtherInfo(mlng病人ID, mlng主页ID, "1,2,3", "2,3")
    Else
        Set rsDiagnosisOther = GetDiagnosticOtherInfo(mlng病人ID, mlng主页ID, "11,12,13", "2,3")
    End If
            
    With vsGrid
        .Clear
        Call initvfgHeadTitle(vsGrid, intFlag)
        If Not rsDiagnosisOther Is Nothing Then
            If intFlag = 1 Then
                'a.西医诊断
                rsDiagnosisOther.Filter = "诊断类型=3 and 记录来源=3"            '先取首页整理的出院诊断
                If Not rsDiagnosisOther.EOF Then
                    .Rows = IIf(rsDiagnosisOther.EOF, 0, rsDiagnosisOther.RecordCount) + 1
                Else
                    '问题28483 by lesfeng 2010-03-01
                    rsDiagnosisOther.Filter = "诊断类型=3 and 记录来源=2"        '再取入院登记的出院诊断
                    If Not rsDiagnosisOther.EOF Then
                        .Rows = IIf(rsDiagnosisOther.EOF, 0, rsDiagnosisOther.RecordCount) + 1
                    Else
                        '问题28138 by lesfeng 2010-03-01 增加默认诊断的判断 不获取门诊诊断及入院诊断
                        If mint默认诊断 = 1 Then
                            rsDiagnosisOther.Filter = "诊断类型=2 and 记录来源=2"        '再取入院登记的入院诊断
                            If Not rsDiagnosisOther.EOF Then
                                .Rows = IIf(rsDiagnosisOther.EOF, 0, rsDiagnosisOther.RecordCount) + 1
                            Else
                                rsDiagnosisOther.Filter = "诊断类型=1 and 记录来源=2"    '最后取入院登记的门诊诊断
                                If Not rsDiagnosisOther.EOF Then
                                    .Rows = IIf(rsDiagnosisOther.EOF, 0, rsDiagnosisOther.RecordCount) + 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                'b.中医诊断
                rsDiagnosisOther.Filter = "诊断类型=13 and 记录来源=3"            '先取首页整理的出院诊断
                If Not rsDiagnosisOther.EOF Then
                    .Rows = IIf(rsDiagnosisOther.EOF, 0, rsDiagnosisOther.RecordCount) + 1
                Else
                    '问题28483 by lesfeng 2010-03-01
                    rsDiagnosisOther.Filter = "诊断类型=13 and 记录来源=2"        '再取入院登记的出院诊断
                    If Not rsDiagnosisOther.EOF Then
                        .Rows = IIf(rsDiagnosisOther.EOF, 0, rsDiagnosisOther.RecordCount) + 1
                    Else
                        '问题28138 by lesfeng 2010-03-01 增加默认诊断的判断 不获取门诊诊断及入院诊断
                        If mint默认诊断 = 1 Then
                            rsDiagnosisOther.Filter = "诊断类型=12 and 记录来源=2"        '再取入院登记的入院诊断
                            If Not rsDiagnosisOther.EOF Then
                                .Rows = IIf(rsDiagnosisOther.EOF, 0, rsDiagnosisOther.RecordCount) + 1
                            Else
                                rsDiagnosisOther.Filter = "诊断类型=11 and 记录来源=2"    '最后取入院登记的门诊诊断
                                If Not rsDiagnosisOther.EOF Then
                                    .Rows = IIf(rsDiagnosisOther.EOF, 0, rsDiagnosisOther.RecordCount) + 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            
            '诊断类型,记录来源,诊断描述,疾病ID,诊断ID,出院情况,记录日期,是否疑诊
            If Not rsDiagnosisOther.EOF Then
                For i = 1 To .Rows - 1
                    .TextMatrix(i, .ColIndex("诊断描述")) = IIf(IsNull(rsDiagnosisOther!诊断描述), "", rsDiagnosisOther!诊断描述)
                    .TextMatrix(i, .ColIndex("出院情况")) = IIf(IsNull(rsDiagnosisOther!出院情况), "", rsDiagnosisOther!出院情况)
                    If intFlag = 1 Then
                       .TextMatrix(i, .ColIndex("ICD编码")) = IIf(IsNull(rsDiagnosisOther!编码), "", rsDiagnosisOther!编码)
                        .TextMatrix(i, .ColIndex("疑诊")) = IIf(IsNull(rsDiagnosisOther!是否疑诊), "", IIf(rsDiagnosisOther("是否疑诊") = 1, "？", ""))
                    Else
                        .TextMatrix(i, .ColIndex("中医编码")) = IIf(IsNull(rsDiagnosisOther!编码), "", rsDiagnosisOther!编码)
                    End If
                    .TextMatrix(i, .ColIndex("疾病ID")) = IIf(IsNull(rsDiagnosisOther!疾病ID), 0, rsDiagnosisOther!疾病ID)
                    .TextMatrix(i, .ColIndex("诊断ID")) = IIf(IsNull(rsDiagnosisOther!诊断ID), 0, rsDiagnosisOther!诊断ID)
                    rsDiagnosisOther.MoveNext
                Next
                .Rows = .Rows + 1
    
'            Else
'                .Rows = .Rows + 1
            End If
            
'            If .Rows > 1 Then
'                .Select 1, .ColIndex("vsGrid")
'            End If
        End If
    End With
    Call SetVfgNo(vsGrid)
    Call SetInitVfgFormat(vsGrid, intFlag)
    Call RestoreHead(vsGrid, intFlag)
End Sub

Private Sub vfg西医_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String
    With vfg西医
        Select Case Col
           Case .ColIndex("诊断描述")
                strValue = Trim(.TextMatrix(Row, .ColIndex("诊断描述")))
                If Not IsNull(strValue) Then
                    If Not GetDiagnosis(vfg西医, strValue, 1) Then
                        .Select Row, Col
                        .TextMatrix(Row, .ColIndex("诊断描述")) = mstrOldName
                        mstrOldName = ""
                        Exit Sub
                    End If
                    Call SetVfgNo(vfg西医)
                End If
            Case .ColIndex("出院情况")
                If .ComboIndex < 0 Then Exit Sub
                .Cell(flexcpData, Row, Col) = .ComboData(.ComboIndex)
                .TextMatrix(Row, Col) = .ComboItem(.ComboIndex)
        End Select
    End With
End Sub

Private Sub vfg西医_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vfg西医.Editable = flexEDNone Then
        Cancel = True
        Exit Sub
    End If
    Select Case Col
        Case vfg西医.ColIndex("诊断描述")
            mstrOldName = Trim(vfg西医.TextMatrix(Row, vfg西医.ColIndex("诊断描述")))
            Cancel = False
            Exit Sub
        Case vfg西医.ColIndex("出院情况")  ', vfg西医.ColIndex("疑诊")
            Cancel = False
            Exit Sub
        Case Else
            Cancel = True
            Exit Sub
    End Select
End Sub

Private Sub vfg西医_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    Select Case Col
        Case vfg西医.ColIndex("序号")
            Position = -1
            Exit Sub
    End Select
    If Position = 0 Then
        Position = Col
    End If
End Sub

Private Sub vfg西医_BeforeSort(ByVal Col As Long, Order As Integer)
    Call zl_VsGridBeforeSort(vfg西医, Col, Order)
    Call SetVfgNo(vfg西医)
End Sub

Private Sub vfg西医_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim RecReturn As Recordset
    Dim i As Long
    Dim j As Long
    Dim strTemp As String
    Dim lngRow As Long
    Dim lngCurrRow As Long
    Dim blnRow As Boolean
    Dim strValue As String
    
    strValue = ""
'    If InStr(vfg西医.Cell(flexcpText, 0, Col), "ICD编码") > 0 Then ' And mintDblick = 0
'         Err = 0: On Error GoTo ErrHand:
'        If Not GetDiagnosis(vfg西医, strValue, 2) Then
'            vfg西医.Select Row, Col
'            Exit Sub
'        End If
'        Exit Sub
'    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vfg西医_DblClick()
    Dim strTemp As String
    Dim intCount As Integer

    If vfg西医.Editable = flexEDKbdMouse Then
        With vfg西医
            If .Row > 0 Then
                If .Col = .ColIndex("疑诊") Then
                    If Trim(.TextMatrix(.Row, .ColIndex("疑诊"))) = "" Then
                        .TextMatrix(.Row, .ColIndex("疑诊")) = "？"
                    Else
                        .TextMatrix(.Row, .ColIndex("疑诊")) = ""
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub vfg西医_EnterCell()
    Dim strTemp As String
    Dim intCount As Integer
    Dim strKey As String
    Dim strValue As String
    
    If vfg西医.Editable = flexEDKbdMouse Then
        With vfg西医
            If .Row > 0 Then
                If .Col = .ColIndex("出院情况") Then
                    strTemp = .TextMatrix(.Row, .Col)
                    strKey = .ColData(.ColIndex("出院情况"))
                    strValue = Trim(.TextMatrix(.Row, .ColIndex("诊断描述")))
                    If strTemp = "" And strValue <> "" Then
                        .TextMatrix(.Row, .Col) = Mid(strKey, InStr(1, strKey, ";") + 1)
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub vfg西医_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
     If vfg西医.Editable = flexEDKbdMouse Then
        If KeyCode = vbKeyDelete Then
            If vfg西医.Row > 0 Then
                If MsgBox("真要删除当前记录吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                    vfg西医.RemoveItem (vfg西医.Row)
                    If vfg西医.Row = 0 Then
                        vfg西医.Rows = vfg西医.Rows + 1
                        vfg西医.Select vfg西医.Rows - 1, vfg西医.Col
                    End If
                End If
            End If
        End If
        
        If KeyCode = vbKeyInsert Then
            With vfg西医
                If MsgBox("真要增加记录吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                   vfg西医.Rows = vfg西医.Rows + 1
                   .Select vfg西医.Rows - 1, vfg西医.Col
                End If
            End With
        End If
    End If
    If KeyCode = vbKeyReturn Then
        lngRow = vfg西医.Row
        If vfg西医.Editable = flexEDKbdMouse Then ''诊断描述,2200,1,1;出院情况,1000,1,1;疑诊
            Call zlPvVsMoveGridCell(vfg西医, vfg西医.ColIndex("诊断描述"), vfg西医.ColIndex("疑诊"), True, lngRow, SetHeadCodeData(vfg西医))
        Else
            Call zlPvVsMoveGridCell(vfg西医, vfg西医.ColIndex("诊断描述"), vfg西医.ColIndex("疑诊"), False, lngRow, SetHeadCodeData(vfg西医))
        End If
    End If
    Call SetVfgNo(vfg西医)
'    If KeyCode <> vbKeyReturn Then
'        vfg西医.ColComboList(vfg西医.ColIndex("ICD编码")) = ""
'    End If
End Sub

Private Sub vfg西医_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim lngRow As Long
    If KeyCode = vbKeyReturn Then
        lngRow = Row
        If vfg西医.Editable = flexEDKbdMouse Then
            Call zlPvVsMoveGridCell(vfg西医, vfg西医.ColIndex("诊断描述"), vfg西医.ColIndex("疑诊"), True, lngRow, SetHeadCodeData(vfg西医))
        Else
            Call zlPvVsMoveGridCell(vfg西医, vfg西医.ColIndex("诊断描述"), vfg西医.ColIndex("疑诊"), False, lngRow, SetHeadCodeData(vfg西医))
        End If
    End If
End Sub

Private Sub vfg西医_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0 'Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub vfg西医_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
        Case vbKeyBack, vbKeyEscape, 3, 22: Exit Sub
        Case Else
'            Select Case Col
'                Case vfgInDetail.ColIndex("数量")
'                    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
'                        Exit Sub
'                    Else
'                        KeyAscii = 0
'                    End If
''            End Select

    End Select
End Sub

'Private Sub vfg西医_KeyUp(KeyCode As Integer, Shift As Integer)
'    vfg西医.ColComboList(vfg西医.ColIndex("ICD编码")) = "..."
'End Sub
'
'Private Sub vfg西医_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    vfg西医.ColComboList(vfg西医.ColIndex("ICD编码")) = "..."
'End Sub

Private Sub SaveHead(ByVal vsGrid As VSFlexGrid, ByVal intFlag As Integer)
    If intFlag = 1 Then
        zl_VsGrid_SaveToPara vsGrid, Me.Caption, glngModul, "西医诊断列头信息", True, True
    Else
        zl_VsGrid_SaveToPara vsGrid, Me.Caption, glngModul, "中医诊断列头信息", True, True
    End If
End Sub

Private Sub RestoreHead(ByVal vsGrid As VSFlexGrid, ByVal intFlag As Integer)
    If intFlag = 1 Then
        zl_VsGrid_FromParaRestore vsGrid, Me.Caption, glngModul, "西医诊断列头信息", True, True
    Else
        zl_VsGrid_FromParaRestore vsGrid, Me.Caption, glngModul, "中医诊断列头信息", True, True
    End If
End Sub

Private Function SetHeadCodeData(ByRef vsGrid As VSFlexGrid) As String
    Dim i As Long
    Dim strTemp As String
    
    SetHeadCodeData = ""
    With vsGrid
        For i = 0 To .Cols - 1
            If vsGrid.Editable = flexEDKbdMouse Then
'                If i = .ColIndex("ICD编码") Then
                    If IsNull(strTemp) Or strTemp = "" Then
                        strTemp = i & "||0"
                    Else
                        strTemp = strTemp & ";" & i & "||0"
                    End If
'                End If
            End If
        Next
    End With
    SetHeadCodeData = strTemp
End Function

Private Function GetDiagnosis(ByRef vsGrid As VSFlexGrid, ByVal strSearch As String, ByVal intFlag As Integer) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '功能:根据条件,请数据
    '参数:strSearch-搜索条件值,
    '返回:当只满足一个值时返回True,否则返回False
    '--------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim lngHeigth As Long
    Dim lngTop As Long
    Dim vRect As RECT
    Dim blnCancel As Boolean
    Dim StrCodeName As String
    Dim lng诊断Id As Long
    Dim strInput As String
    Dim strSex As String
    Dim str类别 As String
    
    Dim lngRow As Long
    Dim i As Long
    Dim j As Long
    
    GetDiagnosis = False
    If strSearch = "" Then Exit Function
    strInput = UCase(strSearch)
    
    If intFlag = 1 Then
        str类别 = "D"
    Else
        str类别 = "B"
    End If
    
    On Error GoTo errHandle
    
    If Not RequestCode Then
        If gint住院诊断输入 = 1 Then
            strSex = zlCommFun.GetNeedName(txt性别.Text)
            If zlCommFun.IsCharChinese(strInput) Then
                strSQL = "名称 Like [2] or '('||编码||')'||名称 Like [2]" '输入汉字时只匹配名称
            Else
                strSQL = "编码 Like [1] Or 名称 Like [2] Or " & IIf(gbytCode = 0, "简码", "五笔码") & " Like [2]"
            End If
            
            strSQL = _
                    " Select ID,ID as 项目ID,编码,附码,名称," & IIf(gbytCode = 0, "简码", "五笔码 as 简码") & ",说明" & _
                    " From 疾病编码目录 Where Instr([3],类别)>0 And (" & strSQL & ")" & _
                    IIf(strSex <> "", " And (性别限制=[4] Or 性别限制 is NULL)", "") & _
                    " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Order by 编码"
            '自由录入时有多个匹配(汉字)不进行选择,数字及字母则进行选择
            If zlCommFun.IsCharChinese(strInput) Then
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", gstrLike & strInput & "%", str类别, strSex, gbytCode + 1)
                If rsTemp.EOF Then
                    Set rsTemp = Nothing
                ElseIf rsTemp.RecordCount > 1 Then
                    Set rsTemp = Nothing '自由录入时有多个匹配不进行选择
                End If
            Else
                vRect = zlControl.GetControlRect(vsGrid.hWnd)
                lngTop = vRect.Top + vsGrid.CellTop + vsGrid.CellHeight
                Set rsTemp = GetDiseaseCodeNew(Me, blnCancel, strInput, strSex, str类别, vRect.Left - 15, lngTop, lngHeigth)
'                A.ID,A.编码,A.附码,A.名称,A.简码,A.五笔码,A.说明,A.性别限制
                If Not rsTemp Is Nothing Then
                    If rsTemp.EOF Then
                        Set rsTemp = Nothing
                    End If
                End If
            End If
            If Not rsTemp Is Nothing Then
                '数据库中只有一个匹配项目，则以该匹配的项目为准
                i = 1
                With rsTemp
                    If UCase(TypeName(vsGrid)) = "VSFLEXGRID" Then
                        With vsGrid
                            lng诊断Id = IIf(IsNull(rsTemp!ID), 0, rsTemp!ID)
                            '核对重复
                            If Not ExamineInputRepeat(vsGrid, lng诊断Id, intFlag, .Row) Then
                               
                                If Not IsNull(.TextMatrix(j, .ColIndex("诊断描述"))) And .TextMatrix(j, .ColIndex("诊断描述")) <> "" Then
                                    If i <> 1 Then
                                        .Row = .Rows - 1
                                        If .Row + 1 = .Rows Then .Rows = .Rows + 1
                                    End If
                                Else
                                    If .Row + 1 = .Rows Then .Rows = .Rows + 1
                                End If
                    
                                If intFlag = 1 Then
                                    .TextMatrix(.Row, .ColIndex("ICD编码")) = IIf(IsNull(rsTemp!编码), "", rsTemp!编码)
                                Else
                                    .TextMatrix(.Row, .ColIndex("中医编码")) = IIf(IsNull(rsTemp!编码), "", rsTemp!编码)
                                    .TextMatrix(.Row, .ColIndex("诊断ID")) = IIf(IsNull(rsTemp!ID), 0, rsTemp!ID)
                                End If
                                .TextMatrix(.Row, .ColIndex("疾病ID")) = IIf(IsNull(rsTemp!ID), 0, rsTemp!ID)
                                .TextMatrix(.Row, .ColIndex("诊断描述")) = IIf(IsNull(rsTemp!名称), "", rsTemp!名称)
                                If .Row + 1 = .Rows Then .Rows = .Rows + 1
                    '                    .Row = .Row + 1
                    
                                If intFlag = 1 Then
                                    .Select .Row, .ColIndex("ICD编码")
                                Else
                                    .Select .Row, .ColIndex("中医编码")
                                End If
                            Else
                                .TextMatrix(.Row, .ColIndex("诊断描述")) = mstrOldName
                            End If
                        End With
                    Else
                        .Close
                        If vsGrid.Enabled Then vsGrid.SetFocus
                        zlCommFun.PressKey vbKeyTab
                    End If
                    .Close
                End With
            Else
                '多项或者无匹配项目时才以输入的为准
                GetDiagnosis = True
                Exit Function
            End If
        End If
    ElseIf strSearch = mstrOldName Then
'        Call zlCommFun.PressKey(vbKeyTab)
    Else
        strSex = zlCommFun.GetNeedName(txt性别.Text)
        
        vRect = zlControl.GetControlRect(vsGrid.hWnd)
        lngTop = vRect.Top + vsGrid.CellTop + vsGrid.CellHeight
        Set rsTemp = GetDiseaseCode(Me, blnCancel, strInput, strSex, str类别, vRect.Left - 15, lngTop, lngHeigth)
        If Not rsTemp Is Nothing Then
            i = 1
            With rsTemp
                If UCase(TypeName(vsGrid)) = "VSFLEXGRID" Then
                    With vsGrid
                        lng诊断Id = IIf(IsNull(rsTemp!ID), 0, rsTemp!ID)
                        If Not ExamineInputRepeat(vsGrid, lng诊断Id, intFlag, .Row) Then
                           
                            If Not IsNull(.TextMatrix(j, .ColIndex("诊断描述"))) And .TextMatrix(j, .ColIndex("诊断描述")) <> "" Then
                                If i <> 1 Then
                                    .Row = .Rows - 1
                                    If .Row + 1 = .Rows Then .Rows = .Rows + 1
                                End If
                            Else
                                If .Row + 1 = .Rows Then .Rows = .Rows + 1
                            End If
                
                            If intFlag = 1 Then
                                .TextMatrix(.Row, .ColIndex("ICD编码")) = IIf(IsNull(rsTemp!编码), "", rsTemp!编码)
                            Else
                                .TextMatrix(.Row, .ColIndex("中医编码")) = IIf(IsNull(rsTemp!编码), "", rsTemp!编码)
                                '50337:刘鹏飞,2012-09-18,给诊断ID列赋值，不然不能判断诊断重复
                                .TextMatrix(.Row, .ColIndex("诊断ID")) = IIf(IsNull(rsTemp!ID), 0, rsTemp!ID)
                            End If
                            .TextMatrix(.Row, .ColIndex("疾病ID")) = IIf(IsNull(rsTemp!ID), 0, rsTemp!ID)
                            .TextMatrix(.Row, .ColIndex("诊断描述")) = IIf(IsNull(rsTemp!名称), "", rsTemp!名称)
                            If .Row + 1 = .Rows Then .Rows = .Rows + 1
                '                    .Row = .Row + 1
                
                            If intFlag = 1 Then
                                .Select .Row, .ColIndex("ICD编码")
                            Else
                                .Select .Row, .ColIndex("中医编码")
                            End If
                        Else
                            .TextMatrix(.Row, .ColIndex("诊断描述")) = mstrOldName
                        End If
                    End With
                Else
                    .Close
                    If vsGrid.Enabled Then vsGrid.SetFocus
                    zlCommFun.PressKey vbKeyTab
                End If
                .Close
            End With
        Else
            If Not blnCancel Then
                MsgBox "没有找到匹配的疾病编码。", vbInformation, gstrSysName
            End If
            If vsGrid.Enabled Then
                GetDiagnosis = False
                Exit Function
            End If
        End If
    End If
    GetDiagnosis = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ExamineInputRepeat(ByRef vsGrid As VSFlexGrid, ByVal lng诊断Id As Long, ByVal intFlag As Integer, ByVal CurrRow As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:验证器数据的是否有重复
    '返回:有重复返回true,否则false
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
           
    ExamineInputRepeat = True
    
    With vsGrid
        For i = 1 To .Rows - 1
            If i <> CurrRow Then
                If intFlag = 1 Then
                    If Val(.TextMatrix(i, .ColIndex("疾病ID"))) = lng诊断Id Then
                        MsgBox "录入诊断在列表中第" & i & "的诊断相同，请录入其它诊断的数据！", vbInformation, gstrSysName
                        Exit Function
                    End If
                Else
                    If Val(.TextMatrix(i, .ColIndex("诊断ID"))) = lng诊断Id Then
                        MsgBox "录入诊断在列表中第" & i & "的诊断相同，请录入其它诊断的数据！", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        Next
    End With
    ExamineInputRepeat = False
End Function

Private Sub vfg中医_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String
    With vfg中医
        Select Case Col
           Case .ColIndex("诊断描述")
                strValue = Trim(.TextMatrix(Row, .ColIndex("诊断描述")))
                If Not IsNull(strValue) Then
                    If Not GetDiagnosis(vfg中医, strValue, 2) Then
                        .Select Row, Col
                        .TextMatrix(Row, .ColIndex("诊断描述")) = mstrOldName
                        mstrOldName = ""
                        Exit Sub
                    End If
                    Call SetVfgNo(vfg中医)
                End If
            Case .ColIndex("出院情况")
                If .ComboIndex < 0 Then Exit Sub
                .Cell(flexcpData, Row, Col) = .ComboData(.ComboIndex)
                .TextMatrix(Row, Col) = .ComboItem(.ComboIndex)
        End Select
    End With
End Sub

Private Sub vfg中医_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vfg中医.Editable = flexEDNone Then
        Cancel = True
        Exit Sub
    End If
    Select Case Col
        Case vfg中医.ColIndex("诊断描述")
            mstrOldName = Trim(vfg中医.TextMatrix(Row, vfg中医.ColIndex("诊断描述")))
            Cancel = False
            Exit Sub
        Case vfg中医.ColIndex("出院情况") ', vfg中医.ColIndex("疑诊")
            Cancel = False
            Exit Sub
        Case Else
            Cancel = True
            Exit Sub
    End Select
End Sub

Private Sub vfg中医_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    Select Case Col
        Case vfg中医.ColIndex("序号")
            Position = -1
            Exit Sub
    End Select
    If Position = 0 Then
        Position = Col
    End If
End Sub

Private Sub vfg中医_BeforeSort(ByVal Col As Long, Order As Integer)
    Call zl_VsGridBeforeSort(vfg中医, Col, Order)
    Call SetVfgNo(vfg中医)
End Sub

Private Sub vfg中医_EnterCell()
    Dim strTemp As String
    Dim intCount As Integer
    Dim strKey As String
    Dim strValue As String
    
    If vfg中医.Editable = flexEDKbdMouse Then
        With vfg中医
            If .Row > 0 Then
                If .Col = .ColIndex("出院情况") Then
                    strTemp = .TextMatrix(.Row, .Col)
                    strKey = .ColData(.ColIndex("出院情况"))
                    strValue = Trim(.TextMatrix(.Row, .ColIndex("诊断描述")))
                    If strTemp = "" And strValue <> "" Then
                        .TextMatrix(.Row, .Col) = Mid(strKey, InStr(1, strKey, ";") + 1)
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub vfg中医_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
     If vfg中医.Editable = flexEDKbdMouse Then
        If KeyCode = vbKeyDelete Then
            If vfg中医.Row > 0 Then
                If MsgBox("真要删除当前记录吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                    vfg中医.RemoveItem (vfg中医.Row)
                    If vfg中医.Row = 0 Then
                        vfg中医.Rows = vfg中医.Rows + 1
                        vfg中医.Select vfg中医.Rows - 1, vfg中医.Col
                    End If
                End If
            End If
        End If
        
        If KeyCode = vbKeyInsert Then
            With vfg中医
                If MsgBox("真要增加记录吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                   vfg中医.Rows = vfg中医.Rows + 1
                   .Select vfg中医.Rows - 1, vfg中医.Col
                End If
            End With
        End If
    End If
    If KeyCode = vbKeyReturn Then
        lngRow = vfg中医.Row
        If vfg中医.Editable = flexEDKbdMouse Then ''诊断描述,2200,1,1;出院情况,1000,1,1;疑诊
            Call zlPvVsMoveGridCell(vfg中医, vfg中医.ColIndex("诊断描述"), vfg中医.ColIndex("出院情况"), True, lngRow, SetHeadCodeData(vfg中医))
        Else
            Call zlPvVsMoveGridCell(vfg中医, vfg中医.ColIndex("诊断描述"), vfg中医.ColIndex("出院情况"), False, lngRow, SetHeadCodeData(vfg中医))
        End If
    End If
    Call SetVfgNo(vfg中医)
End Sub

Private Sub vfg中医_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim lngRow As Long
    If KeyCode = vbKeyReturn Then
        lngRow = Row
        If vfg中医.Editable = flexEDKbdMouse Then
            Call zlPvVsMoveGridCell(vfg中医, vfg中医.ColIndex("诊断描述"), vfg中医.ColIndex("出院情况"), True, lngRow, SetHeadCodeData(vfg中医))
        Else
            Call zlPvVsMoveGridCell(vfg中医, vfg中医.ColIndex("诊断描述"), vfg中医.ColIndex("出院情况"), False, lngRow, SetHeadCodeData(vfg中医))
        End If
    End If
End Sub

Private Sub vfg中医_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0 'Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub vfg中医_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
        Case vbKeyBack, vbKeyEscape, 3, 22: Exit Sub
        Case Else
'            Select Case Col
'                Case vfgInDetail.ColIndex("数量")
'                    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
'                        Exit Sub
'                    Else
'                        KeyAscii = 0
'                    End If
''            End Select

    End Select
End Sub
'问题28612 by lesfeng 2010-07-05
Private Function GetdeathTime(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Date
'功能：获取指定病人是否存在死亡医嘱，存在出院时间为死亡时间加1秒
'说明：用于获取病人死亡时间为出院时间
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    GetdeathTime = CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"))
    
    On Error GoTo errH
    '47955:刘鹏飞,2012-09-18,改变条件中的A.序号=0为A.婴儿=0
    '59094:刘鹏飞,2013-04-24,修改为只加1s,原来为1m
    strSQL = "Select Max(Nvl(A.执行终止时间, Nvl(A.上次执行时间, A.开始执行时间)) + 1 / 24 / 60 / 60 ) As 时间 " & _
             "  From 病人医嘱记录 A, 诊疗项目目录 B " & _
             " Where A.诊疗类别 = B.类别 And A.诊疗项目id = B.ID And B.操作类型 = 11 And B.类别 = 'Z' And A.医嘱状态 In (3, 8, 9) And nvl(A.婴儿,0)=0 And " & _
             "       A.病人ID = [1] And A.主页ID = [2] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
    
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!时间) Then
            GetdeathTime = rsTmp!时间
            mintDeath = 1
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ShowMe(frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strPrivs As String) As Boolean
    On Error Resume Next
    Set mfrmParent = frmParent
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mstrPrivs = strPrivs
    
    Me.Show IIf(gfrmMain Is Nothing, 0, 1), frmParent
    
    ShowMe = gblnOK
End Function
