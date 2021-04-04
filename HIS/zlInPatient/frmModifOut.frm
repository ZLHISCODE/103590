VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmModifOut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "新的出院时间"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8520
   Icon            =   "frmModifOut.frx":0000
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
      TabIndex        =   16
      Top             =   795
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7185
      TabIndex        =   15
      Top             =   360
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   7185
      TabIndex        =   17
      Top             =   4710
      Width           =   1100
   End
   Begin VB.Frame fraInfo 
      Height          =   5535
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   7035
      Begin VB.TextBox txt随诊单位 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   6480
         MaxLength       =   3
         TabIndex        =   31
         Top             =   4440
         Width           =   405
      End
      Begin VB.ComboBox cbo出院情况 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         Left            =   5550
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   660
         Width           =   1350
      End
      Begin VB.CheckBox chk随诊 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "随诊"
         Enabled         =   0   'False
         Height          =   195
         Left            =   4875
         TabIndex        =   12
         Top             =   4500
         Width           =   660
      End
      Begin VB.TextBox txt随诊 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   6000
         MaxLength       =   3
         TabIndex        =   13
         Top             =   4440
         Width           =   405
      End
      Begin VB.CheckBox chk尸检 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "尸检"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2910
         TabIndex        =   10
         Top             =   4980
         Width           =   660
      End
      Begin VB.TextBox txt住院号 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   1290
      End
      Begin VB.TextBox txt年龄 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4170
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   690
      End
      Begin VB.TextBox txt性别 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox txt姓名 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txt出院诊断 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         Left            =   960
         MaxLength       =   200
         TabIndex        =   4
         Top             =   660
         Width           =   3660
      End
      Begin VB.CheckBox chk疑诊 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "确诊"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2235
         TabIndex        =   9
         Top             =   4500
         Width           =   660
      End
      Begin VB.TextBox txt中医诊断 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         Left            =   960
         MaxLength       =   200
         TabIndex        =   6
         Top             =   2580
         Width           =   3660
      End
      Begin VB.ComboBox cbo中医出院情况 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         Left            =   5550
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2580
         Width           =   1350
      End
      Begin VB.ComboBox cbo出院方式 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   4440
         Width           =   1230
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   300
         Left            =   960
         TabIndex        =   11
         Top             =   4950
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   14737632
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   19
         Format          =   "yyyy-MM-dd hh:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtNewDate 
         Height          =   300
         Left            =   4920
         TabIndex        =   14
         Top             =   4950
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   529
         _Version        =   393216
         ForeColor       =   16711680
         AutoTab         =   -1  'True
         MaxLength       =   19
         Format          =   "yyyy-MM-dd hh:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VSFlex8Ctl.VSFlexGrid vfg西医 
         Height          =   1455
         Left            =   960
         TabIndex        =   32
         Top             =   1080
         Width           =   5895
         _cx             =   10398
         _cy             =   2566
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
         TabIndex        =   34
         Top             =   3000
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
         Left            =   3000
         TabIndex        =   36
         Top             =   4440
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   14737632
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
         TabIndex        =   35
         Top             =   3045
         Width           =   720
      End
      Begin VB.Label lbl西医其它 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "其它诊断"
         Height          =   180
         Left            =   180
         TabIndex        =   33
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "调整出院时间"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   3720
         TabIndex        =   30
         Top             =   5010
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院情况"
         Height          =   180
         Left            =   4755
         TabIndex        =   29
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院时间"
         Height          =   180
         Left            =   180
         TabIndex        =   28
         Top             =   5010
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "期限"
         Height          =   180
         Left            =   5580
         TabIndex        =   27
         Top             =   4500
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   180
         Left            =   3765
         TabIndex        =   26
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   180
         Left            =   2385
         TabIndex        =   25
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   525
         TabIndex        =   24
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "住院号"
         Height          =   180
         Left            =   4920
         TabIndex        =   23
         Top             =   300
         Width           =   540
      End
      Begin VB.Label lbl出院诊断 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院诊断"
         Height          =   180
         Left            =   180
         TabIndex        =   22
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lbl中医诊断 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "中医诊断"
         Height          =   180
         Left            =   180
         TabIndex        =   21
         Top             =   2640
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院情况"
         Height          =   180
         Left            =   4755
         TabIndex        =   20
         Top             =   2640
         Width           =   720
      End
      Begin VB.Label lbl出院方式 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院方式"
         Height          =   180
         Left            =   180
         TabIndex        =   19
         Top             =   4500
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmModifOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Public mstrPrivs As String
Public mlng病人ID As Long, mlng主页ID As Long
Private mint默认诊断 As Integer
Private mrsPatiInfo As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, strSQL As String, str床号 As String
    Dim dMax As Date, int原因 As Integer, int险类 As Integer
    Dim rsDiagnosis As ADODB.Recordset
    Dim str出院情况 As String, str中医出院情况 As String, strTmp As String
    Dim str出院方式 As String
    Dim int随诊标志 As Integer
    
    
    
    gblnOK = False
    Set mrsPatiInfo = GetPatiInfoModiOut(mlng病人ID, mlng主页ID)
    mint默认诊断 = Val(zlDatabase.GetPara("默认诊断", glngSys, glngModul))
    If mrsPatiInfo.EOF Then
        MsgBox "病人出院信息不存在，请核查 ！", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    With mrsPatiInfo
        txt姓名.Text = !姓名
        txt性别.Text = "" & !性别
        txt年龄.Text = "" & !年龄
        txt住院号.Text = "" & !住院号
        txtDate.Text = "" & !出院日期
        txt中医诊断.Enabled = (InStr(1, "," & GetDepCharacter(Val("" & !出院科室id)) & ",", ",中医科,") > 0)
        txt中医诊断.ToolTipText = "只有当病人所在科室的性质为中医科时才允许输入中医诊断!"
        cbo中医出院情况.Enabled = txt中医诊断.Enabled
        '问题28982 by lesfeng 2010-06-09
        chk疑诊.Value = IIf(IsNull(!是否确诊), 0, IIf(!是否确诊 = 1, 1, 0))
        chk尸检.Value = IIf(IsNull(!尸检标志), 0, IIf(!尸检标志 = 1, 1, 0))
        chk随诊.Value = IIf(IsNull(!随诊标志), 0, IIf(!随诊标志 >= 1, 1, 0))
        txt随诊.Text = IIf(IsNull(!随诊期限), "", !随诊期限)
        int随诊标志 = IIf(IsNull(!随诊标志), 0, !随诊标志)
        '问题28982 by lesfeng 2010-06-09
        If chk疑诊.Value = 1 And Not IsNull(!确诊日期) Then txtOkDate.Text = IIf(IsNull(!确诊日期), "3000-01-01 00:00:00", Format(!确诊日期, "yyyy-MM-dd HH:mm:ss"))
        Select Case int随诊标志
            Case 0
                txt随诊单位.Text = ""
            Case 1
                txt随诊单位.Text = "月"
            Case 2
                txt随诊单位.Text = "年"
            Case 3
                txt随诊单位.Text = "周"
            Case 4
                txt随诊单位.Text = "天"
            Case 9
                txt随诊单位.Text = "终身"
        End Select
    End With
    
    txtNewDate.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    dMax = GetMaxOutDate(mlng病人ID, mlng主页ID, int原因)
    If int原因 = 10 Then
        '59094:刘鹏飞,2013-04-24,修改为只加1s,原来为1m
        txtNewDate.Text = Format(dMax + 1 / 24 / 60 / 60, "yyyy-MM-dd HH:mm:ss")
    Else
        If dMax > CDate(txtNewDate.Text) Then
            txtNewDate.Text = Format(dMax + 1 / 24 / 60, "yyyy-MM-dd HH:mm:ss")
        End If
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
    str出院方式 = IIf(IsNull(mrsPatiInfo!出院方式), "", mrsPatiInfo!出院方式)
    strSQL = "Select 编码,名称,Nvl(缺省标志,0) as 缺省 From 出院方式 Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo出院方式.AddItem rsTmp!编码 & "-" & rsTmp!名称
            If str出院方式 = "" Then
                If rsTmp!缺省 = 1 Then cbo出院方式.ListIndex = cbo出院方式.NewIndex
            Else
                '问题31294 by lesfeng 2010-07-07 rsTmp!编码 改为 rsTmp!名称
                If rsTmp!名称 = str出院方式 Then cbo出院方式.ListIndex = cbo出院方式.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If
    '问题28139 by lesfeng 2010-03-02
    Call LoadVfgData(vfg西医, 1)
    Call LoadVfgData(vfg中医, 2)
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
    
    On Error GoTo errH
    
    If Not IsDate(txtNewDate.Text) Then
        MsgBox "请输入正确的病人新的出院时间！", vbInformation, gstrSysName
        txtNewDate.SetFocus: Exit Sub
    End If
    
    '时间不能超过当前时间太长(一周)
    Curdate = zlDatabase.Currentdate
    If CDate(txtNewDate.Text) > Curdate Then
        If CDate(txtNewDate.Text) - Curdate > 7 Then
            MsgBox "出院时间比当前时间大得过多,请检查！", vbInformation, gstrSysName
            txtNewDate.SetFocus: Exit Sub
        End If
        If MsgBox("出院时间大于了当前系统时间,确实要修改出院时间吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txtNewDate.SetFocus: Exit Sub
        End If
    End If
       
    dMax = GetMaxOutDate(mlng病人ID, mlng主页ID)
    If Format(txtNewDate.Text, "yyyyMMddHHmmss") <= Format(dMax, "yyyyMMddHHmmss") Then
        MsgBox "病人出院时间必须大于该病人上次变动时间 " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & " ！", vbInformation, gstrSysName
        txtNewDate.SetFocus: Exit Sub
    End If
    
    dMax = GetLastAdviceTime(mlng病人ID, mlng主页ID)
    If Format(txtNewDate.Text, "yyyyMMddHHmmss") < Format(dMax, "yyyyMMddHHmmss") Then
        If MsgBox("出院时间小于该病人最后有效医嘱的时间 " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & ",确实需要修改出院时间吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txtNewDate.SetFocus: Exit Sub
        End If
    End If
             
    strSQL = "Zl_病人变动记录_ModifOut(" & mlng病人ID & "," & mlng主页ID & ",To_Date('" & txtNewDate.Text & "','YYYY-MM-DD HH24:MI:SS')," & _
        "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
''
    gcnOracle.BeginTrans
        blnTrans = True
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    gcnOracle.CommitTrans
    blnTrans = False
''
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

Private Sub txtNewDate_GotFocus()
    zlControl.TxtSelAll txtNewDate
End Sub

Private Sub txtNewDate_LostFocus()
    If Not IsDate(txtNewDate.Text) Then txtNewDate.SetFocus
End Sub

Private Function GetMaxOutDate(lng病人ID As Long, lng主页ID As Long, Optional int原因 As Integer) As Date
'功能：获取转科病人最大的上次变动时间
'参数：int原因=返回上次变动的原因
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    GetMaxOutDate = #1/1/1900#
    int原因 = 0
    
    strSQL = "Select 开始时间,开始原因 From 病人变动记录" & _
        " Where 开始时间 is Not NULL And 终止时间 is not  NULL AND 终止原因 = 1 " & _
        " And 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then
        GetMaxOutDate = IIf(IsNull(rsTmp!开始时间), GetMaxOutDate, rsTmp!开始时间)
        int原因 = Nvl(rsTmp!开始原因, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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
        .Editable = flexEDNone
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
                    rsDiagnosisOther.Filter = "诊断类型=3 and 记录来源=2"        '再取入院登记的出院诊断
                    If Not rsDiagnosisOther.EOF Then
                        .Rows = IIf(rsDiagnosisOther.EOF, 0, rsDiagnosisOther.RecordCount) + 1
                    Else
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
            Else
                'b.中医诊断
                rsDiagnosisOther.Filter = "诊断类型=13 and 记录来源=3"            '先取首页整理的出院诊断
                If Not rsDiagnosisOther.EOF Then
                    .Rows = IIf(rsDiagnosisOther.EOF, 0, rsDiagnosisOther.RecordCount) + 1
                Else
                    rsDiagnosisOther.Filter = "诊断类型=13 and 记录来源=2"        '再取入院登记的出院诊断
                    If Not rsDiagnosisOther.EOF Then
                        .Rows = IIf(rsDiagnosisOther.EOF, 0, rsDiagnosisOther.RecordCount) + 1
                    Else
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
    
            Else
                .Rows = .Rows + 1
            End If
            
            If .Rows > 1 Then
                .Select 1, .ColIndex("出院情况")
            End If
        End If
    End With
    Call SetVfgNo(vsGrid)
    Call SetInitVfgFormat(vsGrid, intFlag)
    Call RestoreHead(vsGrid, intFlag)
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
End Sub

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

Private Sub vfg中医_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
     If vfg中医.Editable = flexEDKbdMouse Then
        If KeyCode = vbKeyDelete Then
            If vfg中医.Row > 0 Then
                If MsgBox("真要删除当前记录吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                    vfg西医.RemoveItem (vfg中医.Row)
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

Public Function ShowMe(frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strPrivs As String) As Boolean
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mstrPrivs = strPrivs
    
    Me.Show IIf(gfrmMain Is Nothing, 0, 1), frmParent
    
    ShowMe = gblnOK
End Function
