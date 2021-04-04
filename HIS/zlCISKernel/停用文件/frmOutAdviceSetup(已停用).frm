VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOutAdviceSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "门诊医嘱选项"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   Icon            =   "frmOutAdviceSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab tabPar 
      Height          =   6090
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   10742
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   617
      WordWrap        =   0   'False
      TabCaption(0)   =   "医嘱下达(&1)"
      TabPicture(0)   =   "frmOutAdviceSetup.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl可用药房"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl卫材"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "vsfDrugStore"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cbo卫材"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraLine"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraPurMed"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "医嘱发送(&2)"
      TabPicture(1)   =   "frmOutAdviceSetup.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraPBPSet"
      Tab(1).Control(1)=   "frmPoint"
      Tab(1).Control(2)=   "fraSendNO"
      Tab(1).Control(3)=   "chk关闭医嘱"
      Tab(1).Control(4)=   "chk执行"
      Tab(1).Control(5)=   "fraBillPrint"
      Tab(1).Control(6)=   "fra诊断"
      Tab(1).Control(7)=   "Frame4"
      Tab(1).ControlCount=   8
      Begin VB.Frame fraPBPSet 
         Height          =   1140
         Left            =   -70110
         TabIndex        =   42
         Top             =   1560
         Width           =   4560
         Begin VB.CommandButton cmdPBPSet 
            Caption         =   "支付票据打印设置"
            Height          =   300
            Left            =   390
            TabIndex        =   44
            Top             =   450
            Width           =   1620
         End
         Begin VB.CheckBox chkSendPay 
            Caption         =   "发送时银行卡无密支付(诊间支付)"
            Height          =   360
            Left            =   135
            TabIndex        =   43
            Top             =   -60
            Width           =   3015
         End
      End
      Begin VB.Frame frmPoint 
         Caption         =   "发送后,指引单"
         Height          =   1350
         Left            =   -67680
         TabIndex        =   38
         Top             =   4590
         Width           =   2145
         Begin VB.OptionButton optPoint 
            Caption         =   "自动打印"
            Height          =   180
            Index           =   2
            Left            =   135
            TabIndex        =   41
            Top             =   900
            Width           =   1560
         End
         Begin VB.OptionButton optPoint 
            Caption         =   "选择是否打印"
            Height          =   180
            Index           =   1
            Left            =   135
            TabIndex        =   40
            Top             =   600
            Width           =   1560
         End
         Begin VB.OptionButton optPoint 
            Caption         =   "不打印"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   39
            Top             =   300
            Value           =   -1  'True
            Width           =   1560
         End
      End
      Begin VB.Frame fraPurMed 
         Caption         =   "抗菌药物缺省用药目的"
         Height          =   765
         Left            =   5250
         TabIndex        =   33
         Top             =   2190
         Width           =   4215
         Begin VB.OptionButton optPurMed 
            Caption         =   "下达时确定"
            Height          =   180
            Index           =   0
            Left            =   270
            TabIndex        =   45
            Top             =   360
            Width           =   1560
         End
         Begin VB.OptionButton optPurMed 
            Caption         =   "预防"
            Height          =   180
            Index           =   1
            Left            =   1890
            TabIndex        =   35
            Top             =   360
            Width           =   680
         End
         Begin VB.OptionButton optPurMed 
            Caption         =   "治疗"
            Height          =   180
            Index           =   2
            Left            =   3120
            TabIndex        =   34
            Top             =   360
            Value           =   -1  'True
            Width           =   680
         End
      End
      Begin VB.Frame fraLine 
         Height          =   30
         Left            =   5280
         TabIndex        =   31
         Top             =   1200
         Width           =   4215
      End
      Begin VB.ComboBox cbo卫材 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   6540
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1335
         Width           =   2310
      End
      Begin VB.Frame fraSendNO 
         Caption         =   "单据产生规则"
         Height          =   4395
         Left            =   -74880
         TabIndex        =   22
         Top             =   1545
         Width           =   4605
         Begin VB.CheckBox chkTimeDef 
            Caption         =   "开始时间不是同一天的分别产生单据"
            Height          =   180
            Left            =   240
            TabIndex        =   46
            Top             =   600
            Width           =   3480
         End
         Begin VB.CheckBox chkNOType 
            Caption         =   "不同诊断的医嘱分别产生单据"
            Height          =   180
            Left            =   240
            TabIndex        =   37
            Top             =   315
            Width           =   2760
         End
         Begin VB.OptionButton optSendNO 
            Caption         =   "所有类别医嘱在相同执行科室只产生一张单据"
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   36
            Top             =   1200
            Width           =   4140
         End
         Begin VB.CheckBox chk一并给药发送 
            Caption         =   "一并给药的即使处方笺不同也发送为一张单据"
            Height          =   255
            Left            =   465
            TabIndex        =   32
            Top             =   3045
            Value           =   1  'Checked
            Width           =   3975
         End
         Begin VB.OptionButton optSendNO 
            Caption         =   "每次发送医嘱只产生一张单据"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   25
            Top             =   915
            Width           =   3060
         End
         Begin VB.OptionButton optSendNO 
            Caption         =   "以下同一类别医嘱相同执行科室只产生一张单据"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   24
            Top             =   1515
            Width           =   4140
         End
         Begin VB.ListBox lstSendNO 
            Columns         =   4
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   1110
            IMEMode         =   3  'DISABLE
            Left            =   465
            Style           =   1  'Checkbox
            TabIndex        =   23
            Top             =   1815
            Width           =   3660
         End
         Begin VB.Label lblPrompt 
            Caption         =   $"frmOutAdviceSetup.frx":0044
            Height          =   825
            Left            =   465
            TabIndex        =   26
            Top             =   3360
            Width           =   3735
         End
      End
      Begin VB.CheckBox chk关闭医嘱 
         Caption         =   "发送完成之后自动关闭发送窗口"
         Height          =   195
         Left            =   -69945
         TabIndex        =   13
         Top             =   1095
         Width           =   2940
      End
      Begin VB.CheckBox chk执行 
         Caption         =   "发送时将本科执行的填为已执行"
         Height          =   195
         Left            =   -69945
         TabIndex        =   12
         Top             =   675
         Width           =   2820
      End
      Begin VB.Frame fraBillPrint 
         Caption         =   "发送后,诊疗单据"
         Height          =   1350
         Left            =   -70080
         TabIndex        =   18
         Top             =   4590
         Width           =   2235
         Begin VB.OptionButton optPrint 
            Caption         =   "不打印"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   9
            Top             =   300
            Width           =   1560
         End
         Begin VB.OptionButton optPrint 
            Caption         =   "选择是否打印"
            Height          =   180
            Index           =   1
            Left            =   135
            TabIndex        =   10
            Top             =   600
            Value           =   -1  'True
            Width           =   1560
         End
         Begin VB.OptionButton optPrint 
            Caption         =   "自动打印"
            Height          =   180
            Index           =   2
            Left            =   135
            TabIndex        =   11
            Top             =   900
            Width           =   1560
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " 医嘱操作 "
         Height          =   2565
         Left            =   5280
         TabIndex        =   17
         Top             =   3360
         Width           =   4215
         Begin VB.CommandButton cmdBloodTip 
            Caption         =   "输血申请注意事项设置"
            Height          =   350
            Left            =   105
            TabIndex        =   47
            Top             =   2100
            Width           =   2490
         End
         Begin VB.CheckBox chkMustAddAgent 
            Caption         =   "下达毒麻和第一类精神药品时必须登记代办人"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   1770
            Width           =   3960
         End
         Begin VB.CheckBox chk单量 
            Caption         =   "下达药品医嘱时必须录入药品单量"
            Height          =   195
            Left            =   120
            TabIndex        =   0
            Top             =   480
            Width           =   3360
         End
         Begin VB.CheckBox chk天数 
            Caption         =   "下达药品医嘱时可以指定用药天数"
            Height          =   195
            Left            =   120
            TabIndex        =   1
            Top             =   915
            Width           =   3360
         End
         Begin VB.CheckBox chk皮试 
            Caption         =   "自动增加皮试并根据结果限制医嘱发送"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   1335
            Width           =   3360
         End
      End
      Begin VB.Frame fra诊断 
         Height          =   1530
         Left            =   -70080
         TabIndex        =   20
         Top             =   2850
         Width           =   4530
         Begin VB.ListBox lst诊断 
            Columns         =   3
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   900
            IMEMode         =   3  'DISABLE
            Left            =   360
            Style           =   1  'Checkbox
            TabIndex        =   8
            Top             =   390
            Width           =   2580
         End
         Begin VB.CheckBox chk诊断 
            Caption         =   "发送以下类别时检查诊断填写"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   0
            Width           =   2640
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " 发送单据 "
         Height          =   960
         Left            =   -74880
         TabIndex        =   19
         Top             =   480
         Width           =   4605
         Begin VB.CheckBox chk单位记帐 
            Caption         =   "只有合约单位病人的医嘱才可以发送为记帐单"
            Height          =   195
            Left            =   255
            TabIndex        =   6
            Top             =   630
            Width           =   3960
         End
         Begin VB.OptionButton optSend 
            Caption         =   "发送时再确定"
            Height          =   180
            Index           =   2
            Left            =   2565
            TabIndex        =   5
            Top             =   330
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton optSend 
            Caption         =   "记帐单据"
            Height          =   180
            Index           =   1
            Left            =   1395
            TabIndex        =   4
            Top             =   330
            Width           =   1020
         End
         Begin VB.OptionButton optSend 
            Caption         =   "收费单据"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   3
            Top             =   330
            Width           =   1020
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDrugStore 
         Height          =   5445
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   5055
         _cx             =   8916
         _cy             =   9604
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   14737632
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmOutAdviceSetup.frx":00E8
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
      Begin VB.Label lbl卫材 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缺省发料部门"
         Height          =   180
         Left            =   5280
         TabIndex        =   30
         Top             =   1380
         Width           =   1080
      End
      Begin VB.Label lbl可用药房 
         Caption         =   $"frmOutAdviceSetup.frx":0195
         Height          =   615
         Left            =   5280
         TabIndex        =   28
         Top             =   480
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8625
      TabIndex        =   15
      Top             =   6345
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7530
      TabIndex        =   14
      Top             =   6345
      Width           =   1100
   End
End
Attribute VB_Name = "frmOutAdviceSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mMainPrivs As String
Public mbln医技站 As Boolean
Private Const VsPubBackColor = &HFAEADA

Private Sub chkSendPay_Click()
    cmdPBPSet.Enabled = chkSendPay.value
End Sub

Private Sub chk诊断_Click()
    lst诊断.Enabled = chk诊断.value = 1 And lst诊断.Tag = ""
End Sub

Private Sub cmdBloodTip_Click()
    Dim strPar As String
    strPar = cmdBloodTip.Tag
    Call frmInputBox.InputBox(Me, "输血申请注意事项", "内容：", 4000, 6, True, True, strPar)
    cmdBloodTip.Tag = strPar
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim str诊断 As String, strSendNO As String
    Dim i As Long, bytType As Long
    Dim arr可用药房(3) As String, arr缺省药房(3) As String, arrTmp() As String
    Dim blnSetup As Boolean
    Dim str西药房窗口 As String, str成药房窗口 As String, str中药房窗口 As String
    
    '不检查是否指定了缺省药房，因为可能没有参数设置权限，参数类型是可自定义的。
    
    If mbln医技站 = False Then
        If chk诊断.value = 1 Then
            For i = 0 To lst诊断.ListCount - 1
                If lst诊断.Selected(i) Then
                    str诊断 = str诊断 & Chr(lst诊断.ItemData(i))
                End If
            Next
            If str诊断 = "" Then
                MsgBox "请至少选择一种要检查诊断的医嘱类别。", vbInformation, gstrSysName
                tabPar.Tab = 1: lst诊断.SetFocus: Exit Sub
            End If
        End If
    End If
        
    '允许不选择
    strSendNO = ""
    For i = 0 To lstSendNO.ListCount - 1
        If lstSendNO.Selected(i) Then
            strSendNO = strSendNO & Chr(lstSendNO.ItemData(i))
        End If
    Next
    
    '----------------------------------------------------------------------------------------------------
    '药房
    With vsfDrugStore
        For i = .FixedRows To .Rows - 1
            Select Case .TextMatrix(i, .ColIndex("类别"))
            Case "西药房"
                bytType = 0
                If .TextMatrix(i, .ColIndex("发药窗口")) <> "自动分配" Then
                    str西药房窗口 = str西药房窗口 & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("发药窗口"))
                End If
            Case "成药房"
                bytType = 1
                If .TextMatrix(i, .ColIndex("发药窗口")) <> "自动分配" Then
                    str成药房窗口 = str成药房窗口 & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("发药窗口"))
                End If
            Case "中药房"
                bytType = 2
                If .TextMatrix(i, .ColIndex("发药窗口")) <> "自动分配" Then
                    str中药房窗口 = str中药房窗口 & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("发药窗口"))
                End If
            End Select
            If .TextMatrix(i, .ColIndex("可用")) <> 0 Then arr可用药房(bytType) = arr可用药房(bytType) & "," & .RowData(i)
            If .TextMatrix(i, .ColIndex("缺省")) = "√" Then arr缺省药房(bytType) = .RowData(i)
        Next
    End With
    
    blnSetup = InStr(GetInsidePrivs(p门诊医嘱下达), ";医嘱选项设置;") > 0
    arrTmp = Split("西药房,成药房,中药房", ",")
    For bytType = 0 To UBound(arrTmp)
        Call zlDatabase.SetPara("门诊可用" & arrTmp(bytType), Mid(arr可用药房(bytType), 2), glngSys, p门诊医嘱下达, blnSetup)
        Call zlDatabase.SetPara("门诊缺省" & arrTmp(bytType), arr缺省药房(bytType), glngSys, p门诊医嘱下达, blnSetup)
    Next
    Call zlDatabase.SetPara("西药房窗口", Mid(str西药房窗口, 2), glngSys, p门诊医嘱下达, blnSetup)
    Call zlDatabase.SetPara("成药房窗口", Mid(str成药房窗口, 2), glngSys, p门诊医嘱下达, blnSetup)
    Call zlDatabase.SetPara("中药房窗口", Mid(str中药房窗口, 2), glngSys, p门诊医嘱下达, blnSetup)
          
    Call zlDatabase.SetPara("门诊缺省发料部门", IIF(cbo卫材.ListIndex = 0, "0", cbo卫材.ItemData(cbo卫材.ListIndex)), glngSys, p门诊医嘱下达, blnSetup)
    
    '必须录入药品单量
    Call zlDatabase.SetPara("必须录入药品单量", chk单量.value, glngSys, p门诊医嘱下达, blnSetup)
    
    '医嘱执行天数
    Call zlDatabase.SetPara("医嘱执行天数", chk天数.value, glngSys, p门诊医嘱下达, blnSetup)
    
    '抗菌药物缺省用药目的
    For i = 0 To 2
        If optPurMed(i).value Then
            Call zlDatabase.SetPara("抗菌药物缺省用药目的", i & "", glngSys, p门诊医嘱下达, blnSetup)
            Exit For
        End If
    Next
    
    '----------------------------------------------------------------------------------------------------
    '发送选项
    Call zlDatabase.SetPara("发送单据类型", IIF(optSend(0).value, 0, IIF(optSend(1).value, 1, 2)), glngSys, p门诊医嘱下达, blnSetup)
        
    '仅合约单位病人发送为记帐单
    Call zlDatabase.SetPara("单位记帐", chk单位记帐.value, glngSys, p门诊医嘱下达, blnSetup)

    '下达毒麻和第一类精神药品医嘱时必须登记代办人
    Call zlDatabase.SetPara("要求登记代办人", chkMustAddAgent.value, glngSys, p门诊医嘱下达, blnSetup)
        
    '本科执行自动完成
    Call zlDatabase.SetPara("门诊本科自动执行", chk执行.value, glngSys, p门诊医嘱下达, blnSetup)

    '关闭医嘱窗体
    Call zlDatabase.SetPara("发送完成后关闭医嘱窗体", chk关闭医嘱.value, glngSys, p门诊医嘱下达, blnSetup)
    
    If mbln医技站 = False Then
        '自动处理皮试
        Call zlDatabase.SetPara("自动处理皮试", chk皮试.value, glngSys, p门诊医嘱下达, blnSetup)
        
        '单据打印:0-不打印,1-手工打印,2-自动打印
        Call zlDatabase.SetPara("门诊发送单据打印", IIF(optPrint(0).value, 0, IIF(optPrint(1).value, 1, 2)), glngSys, p门诊医嘱下达, blnSetup)
        
        '要求输入门诊诊断
        Call zlDatabase.SetPara("要求输入门诊诊断", str诊断, glngSys, p门诊医嘱下达, blnSetup)
    End If
     
    '不同诊断的医嘱分别产生单据
    Call zlDatabase.SetPara("不同诊断的医嘱分别产生单据", chkNOType.value, glngSys, p门诊医嘱下达, blnSetup)
    '开始时间不是同一天的分别产生单据
    Call zlDatabase.SetPara("开始时间不是同一天的分别产生单据", chkTimeDef.value, glngSys, p门诊医嘱下达, blnSetup)
    
    '发送单据号
    Call zlDatabase.SetPara("发送单据号规则", IIF(optSendNO(0).value, 1, IIF(optSendNO(2).value, 2, 0)), glngSys, p门诊医嘱下达, blnSetup) '0-多个,1-单个,2-所有
    
    '产生为同一单据的医嘱类别
    Call zlDatabase.SetPara("产生为同一单据的医嘱类别", strSendNO, glngSys, p门诊医嘱下达, blnSetup)
    
    Call zlDatabase.SetPara("一并给药发送为一张", chk一并给药发送.value, glngSys, p门诊医嘱下达, blnSetup)

    '指引单打印
    Call zlDatabase.SetPara("指引单打印方式", IIF(optPoint(0).value, 0, IIF(optPoint(1).value, 1, 2)), glngSys, p门诊医嘱下达, blnSetup)
    
    '诊间支付
    Call zlDatabase.SetPara("启用诊间支付", chkSendPay.value, glngSys, p门诊医嘱下达, blnSetup)
    
    Call zlDatabase.SetPara("输血申请注意事项", cmdBloodTip.Tag, glngSys, p门诊医嘱下达, blnSetup)
    gblnOK = True
    Unload Me
End Sub

Private Sub cmdPBPSet_Click()
    On Error Resume Next
    If gobjSquareCard Is Nothing Then
        Set gobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If gobjSquareCard.zlInitComponents(Me, p门诊医嘱下达, glngSys, gstrDBUser, gcnOracle, False) = False Then
            Set gobjSquareCard = Nothing
            MsgBox "医疗卡部件（zl9CardSquare）初始化失败!", vbInformation, gstrSysName
            err.Clear: Exit Sub
        End If
    End If
    Call gobjSquareCard.zlCliniqueRoomPayPrintSet(Me)
    err.Clear: On Error GoTo 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then Call cmdHelp_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strPar As String
    Dim blnSetup As Boolean, arrTmp() As String
    Dim strDSIDs As String, strDefault As String, lngBackColor As Long, bytLockEdit As Byte
    Dim intType1 As Integer, intType2 As Integer, lngRow As Long
    Dim str窗口 As String, j As Integer
    
    On Error GoTo errH
    
    gblnOK = False
    
    If mbln医技站 Then
        chk皮试.Visible = False
        fraBillPrint.Visible = False
        fra诊断.Visible = False
    End If
    
    blnSetup = InStr(GetInsidePrivs(p门诊医嘱下达), "医嘱选项设置") > 0
    '------------------------------------------------------------------------------------------------------------------------
    '药房与发料部门
    strSQL = _
        "Select Distinct A.ID,A.编码,A.名称,B.工作性质 " & _
        " From 部门表 A,部门性质说明 B " & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " AND B.部门ID=A.ID And B.服务对象 IN(1,3) and B.工作性质 in('中药房','西药房','成药房','发料部门')" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " Order by 工作性质,编码"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    With vsfDrugStore
        .Rows = .FixedRows
        .Editable = flexEDKbdMouse
        .MergeCol(.ColIndex("类别")) = True
        .MergeCells = flexMergeFixedOnly
        
        rsTmp.Filter = "工作性质<>'发料部门'"
        If Not rsTmp.EOF Then
            .Rows = .FixedRows + rsTmp.RecordCount
            lngRow = .FixedRows
            arrTmp = Split("西药房,成药房,中药房", ",")
            For i = 0 To UBound(arrTmp)
                rsTmp.Filter = "工作性质='" & arrTmp(i) & "'"
                strDefault = zlDatabase.GetPara("门诊缺省" & arrTmp(i), glngSys, p门诊医嘱下达, , , , intType1)
                strDSIDs = "," & zlDatabase.GetPara("门诊可用" & arrTmp(i), glngSys, p门诊医嘱下达, , , , intType2) & ","
                '发药窗口
                str窗口 = zlDatabase.GetPara(arrTmp(i) & "窗口", glngSys, p门诊医嘱下达, , , blnSetup)
                Do While Not rsTmp.EOF
                    .TextMatrix(lngRow, .ColIndex("类别")) = arrTmp(i)
                    .TextMatrix(lngRow, .ColIndex("药房")) = rsTmp!名称
                    .RowData(lngRow) = Val(rsTmp!ID)
                    
                    If Val(rsTmp!ID) = Val(strDefault) Then
                        .TextMatrix(lngRow, .ColIndex("缺省")) = "√"
                        .TextMatrix(lngRow, .ColIndex("可用")) = -1   'true
                    Else
                        .TextMatrix(lngRow, .ColIndex("缺省")) = ""
                        .TextMatrix(lngRow, .ColIndex("可用")) = IIF(InStr(strDSIDs, "," & rsTmp!ID & ",") > 0, -1, 0)
                    End If
                    
                    '缺省单元格
                    'intType-'返回参数类型：1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
                    bytLockEdit = 0
                    If InStr(1, ",1,3,15,", "," & intType1 & ",") > 0 Then
                        lngBackColor = IIF(blnSetup, VsPubBackColor, &H8000000F)      '授权限控制
                        bytLockEdit = IIF(blnSetup, 0, 1)
                    ElseIf intType1 = 5 Then
                        lngBackColor = VsPubBackColor       '公共模块,但不授权限控制
                    Else
                        lngBackColor = &H80000005     '正常编辑
                    End If
                    .Cell(flexcpBackColor, lngRow, .ColIndex("缺省")) = lngBackColor
                    .Cell(flexcpData, lngRow, .ColIndex("缺省")) = bytLockEdit
                     
                    '可用单元格
                    bytLockEdit = 0
                    If InStr(1, ",1,3,15,", "," & intType2 & ",") > 0 Then
                        lngBackColor = IIF(blnSetup, VsPubBackColor, &H8000000F)      '授权限控制
                        bytLockEdit = IIF(blnSetup, 0, 1)
                    ElseIf intType2 = 5 Then
                        lngBackColor = VsPubBackColor       '公共模块,但不授权限控制
                    Else
                        lngBackColor = &H80000005     '正常编辑
                    End If
                    .Cell(flexcpBackColor, lngRow, .ColIndex("可用")) = lngBackColor
                    .Cell(flexcpData, lngRow, .ColIndex("可用")) = bytLockEdit
                    
                    '发药窗口
                    For j = 0 To UBound(Split(str窗口, ","))
                        If Val(.RowData(lngRow)) = Val(Split(Split(str窗口, ",")(j), ":")(0)) Then
                            .TextMatrix(lngRow, .ColIndex("发药窗口")) = Split(Split(str窗口, ",")(j), ":")(1)
                            Exit For
                        End If
                    Next
                    If .TextMatrix(lngRow, .ColIndex("发药窗口")) = "" Then .TextMatrix(lngRow, .ColIndex("发药窗口")) = "自动分配"
                    .Cell(flexcpBackColor, lngRow, .ColIndex("发药窗口")) = lngBackColor
                    .Cell(flexcpData, lngRow, .ColIndex("发药窗口")) = bytLockEdit
                    
                    lngRow = lngRow + 1
                    rsTmp.MoveNext
                Loop
                If lngRow < .Rows - 1 Then  '划分隔线
                    .Select lngRow, .FixedCols, lngRow, .Cols - 1
                    .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
                End If
            Next
        End If
    End With
    
    cbo卫材.AddItem "人工选择"
    rsTmp.Filter = "工作性质='发料部门'"
    Do While Not rsTmp.EOF
        cbo卫材.AddItem rsTmp!名称
        cbo卫材.ItemData(cbo卫材.ListCount - 1) = rsTmp!ID
        rsTmp.MoveNext
    Loop
    strPar = zlDatabase.GetPara("门诊缺省发料部门", glngSys, p门诊医嘱下达, , Array(lbl卫材, cbo卫材), blnSetup)
    zlControl.CboLocate cbo卫材, strPar, True
        
    '必须录入药品单量
    chk单量.value = Val(zlDatabase.GetPara("必须录入药品单量", glngSys, p门诊医嘱下达, , Array(chk单量), blnSetup))
    
    '医嘱执行天数
    chk天数.value = Val(zlDatabase.GetPara("医嘱执行天数", glngSys, p门诊医嘱下达, , Array(chk天数), blnSetup))
    
    '抗菌药物缺省用药目的
    strPar = zlDatabase.GetPara("抗菌药物缺省用药目的", glngSys, p门诊医嘱下达, "0")
    If strPar = "3" Then strPar = "0"
    optPurMed(Val(strPar)).value = True
    
    '------------------------------------------------------------------------------------------------------------------------
    '发送选项
    optSend(Val(zlDatabase.GetPara("发送单据类型", glngSys, p门诊医嘱下达, , Array(optSend(0), optSend(1), optSend(2)), blnSetup))).value = True
        
    '仅合约单位病人发送为记帐单
    chk单位记帐.value = Val(zlDatabase.GetPara("单位记帐", glngSys, p门诊医嘱下达, , Array(chk单位记帐), blnSetup))
    
    '要求登记代办人
    chkMustAddAgent.value = Val(zlDatabase.GetPara("要求登记代办人", glngSys, p门诊医嘱下达, "1", Array(chkMustAddAgent), blnSetup))
    
    '本科执行自动完成
    chk执行.value = Val(zlDatabase.GetPara("门诊本科自动执行", glngSys, p门诊医嘱下达, , Array(chk执行), blnSetup))
    
    '关闭医嘱窗体
    chk关闭医嘱.value = Val(zlDatabase.GetPara("发送完成后关闭医嘱窗体", glngSys, p门诊医嘱下达, , Array(chk关闭医嘱), blnSetup))
    
    '指引单打印
    optPoint(Val(zlDatabase.GetPara("指引单打印方式", glngSys, p门诊医嘱下达, , Array(optPoint(0), optPoint(1), optPoint(2)), blnSetup))).value = True
    
    '诊间支付
    chkSendPay.value = Val(zlDatabase.GetPara("启用诊间支付", glngSys, p门诊医嘱下达, , Array(chkSendPay), blnSetup))
    '诊间支付才需要设置发药窗口
    If chkSendPay.value = 0 Then
        vsfDrugStore.ColHidden(vsfDrugStore.ColIndex("发药窗口")) = True
        vsfDrugStore.ColWidth(vsfDrugStore.ColIndex("药房")) = vsfDrugStore.ColWidth(vsfDrugStore.ColIndex("药房")) + vsfDrugStore.ColWidth(vsfDrugStore.ColIndex("发药窗口"))
    End If
    
    cmdPBPSet.Enabled = chkSendPay.value
    
    If mbln医技站 = False Then
            
        '自动处理皮试
        chk皮试.value = Val(zlDatabase.GetPara("自动处理皮试", glngSys, p门诊医嘱下达, , Array(chk皮试), blnSetup))
                        
        '单据打印:0-不打印,1-手工打印,2-自动打印
        optPrint(Val(zlDatabase.GetPara("门诊发送单据打印", glngSys, p门诊医嘱下达, , Array(optPrint(0), optPrint(1), optPrint(2)), blnSetup))).value = True
        
        '要求输入门诊诊断
        strPar = zlDatabase.GetPara("要求输入门诊诊断", glngSys, p门诊医嘱下达, , Array(chk诊断, lst诊断), blnSetup)
        If Not chk诊断.Enabled Then lst诊断.Tag = "1" '固定标识为不可用
        If strPar <> "" Then
            chk诊断.value = 1
            Call chk诊断_Click
        End If
        strSQL = "Select 编码,名称 From 诊疗项目类别 Where 编码 Not IN('4','5','6','7','8','9') Union ALL Select '5','药品' From Dual Order by 编码"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        With lst诊断
            Do While Not rsTmp.EOF
                .AddItem rsTmp!编码 & "-" & rsTmp!名称
                .ItemData(.NewIndex) = Asc(rsTmp!编码)
                
                If strPar <> "" Then
                    If InStr(strPar, Chr(.ItemData(.NewIndex))) > 0 Then
                        .Selected(.NewIndex) = True
                    End If
                End If
                rsTmp.MoveNext
            Loop
            .ListIndex = 0
        End With
        cmdBloodTip.Tag = zlDatabase.GetPara("输血申请注意事项", glngSys, p门诊医嘱下达, , Array(cmdBloodTip), blnSetup)
    Else
        strSQL = "Select 编码,名称 From 诊疗项目类别 Where 编码 Not IN('4','5','6','7','8','9') Order by 编码"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        
    End If
    
    '控件bug，须加30才能显示三行四列（加了后高度690仍然没有变）
    lstSendNO.Height = lstSendNO.Height + 30
    
    '不同诊断的医嘱分别产生单据
    chkNOType.value = Val(zlDatabase.GetPara("不同诊断的医嘱分别产生单据", glngSys, p门诊医嘱下达, 0, Array(chkNOType), blnSetup))
    
    '开始时间不是同一天的分别产生单据
    chkTimeDef.value = Val(zlDatabase.GetPara("开始时间不是同一天的分别产生单据", glngSys, p门诊医嘱下达, 0, Array(chkTimeDef), blnSetup))
    '发送单据号
    i = Val(zlDatabase.GetPara("发送单据号规则", glngSys, p门诊医嘱下达, , Array(optSendNO(0), optSendNO(1), optSendNO(2), lstSendNO), blnSetup)) '0-多个,1-单个，2-所有
    i = IIF(i = 0, 1, IIF(i = 2, 2, 0))
    optSendNO(i).value = True
    Call optSendNO_Click(i)
    
    chk一并给药发送.value = Val(zlDatabase.GetPara("一并给药发送为一张", glngSys, p门诊医嘱下达, 1, Array(chk一并给药发送), blnSetup))
    
    '执行科室相同时产生为同一单据的医嘱类别
    strPar = zlDatabase.GetPara("产生为同一单据的医嘱类别", glngSys, p门诊医嘱下达, , Array(lstSendNO), blnSetup)
    With lstSendNO
        If rsTmp.RecordCount > 0 Then rsTmp.Filter = "编码<>'5'"
        Do While Not rsTmp.EOF
            .AddItem rsTmp!编码 & "-" & rsTmp!名称
            .ItemData(.NewIndex) = Asc(rsTmp!编码)
            
            If strPar <> "" Then
                If InStr(strPar, Chr(.ItemData(.NewIndex))) > 0 Then
                    .Selected(.NewIndex) = True
                End If
            End If
            rsTmp.MoveNext
        Loop
        .ListIndex = 0
    End With
    
    cmdCancel.Left = Me.Left + Me.Width - cmdCancel.Width - 200
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 200
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mMainPrivs = ""
    mbln医技站 = False
End Sub

Private Sub optSend_Click(Index As Integer)
    chk单位记帐.Enabled = Index <> 0
End Sub

Private Sub optSend_GotFocus(Index As Integer)
    tabPar.Tab = 1
End Sub

Private Sub optSendNO_Click(Index As Integer)
    lstSendNO.Enabled = optSendNO(1).value
    chk一并给药发送.Enabled = optSendNO(1).value
End Sub

Private Sub vsfDrugStore_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsfDrugStore.ColIndex("可用") Then
        Call Set可用药房(Row, True)
    ElseIf Col = vsfDrugStore.ColIndex("可用") Then
        Call Set缺省药房
    End If
    If Col <> vsfDrugStore.ColIndex("发药窗口") Then Cancel = True
End Sub

Private Sub vsfDrugStore_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfDrugStore
        Select Case Col
        Case .ColIndex("可用")
            Cancel = Val(.Cell(flexcpData, Row, Col)) <> 0
        Case .ColIndex("缺省")
            Cancel = Val(.Cell(flexcpData, Row, Col)) <> 0
        Case .ColIndex("发药窗口")
            Cancel = Val(.Cell(flexcpData, Row, Col)) <> 0
        Case Else
            Cancel = True
            Exit Sub
        End Select
    End With
End Sub

Private Sub vsfDrugStore_DblClick()
    With vsfDrugStore
        If .MouseCol = .ColIndex("缺省") Then
            Call Set缺省药房
        ElseIf .MouseCol = .ColIndex("药房") Then
            Call Set可用药房(.Row, True)
        ElseIf .MouseCol = .ColIndex("可用") And .MouseRow = .FixedRows - 1 Then
            Dim i As Long
            For i = .FixedRows To .Rows - 1
                Call Set可用药房(i)
            Next
        End If
    End With
End Sub

Private Sub vsfDrugStore_EnterCell()
    Dim rsTmp As ADODB.Recordset, strList As String
    With vsfDrugStore
        If .Row > 0 Then
            If .Col = .ColIndex("发药窗口") Then
                Set rsTmp = Read发药窗口(.RowData(.Row))
                .ColComboList(.Col) = "自动分配|" & .BuildComboList(rsTmp, "名称")
                .FocusRect = flexFocusSolid
            Else
                .FocusRect = flexFocusLight
            End If
        End If
    End With
End Sub

Private Sub vsfDrugStore_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If vsfDrugStore.Col = vsfDrugStore.ColIndex("缺省") Then
            Call Set缺省药房
        End If
    End If
End Sub

Private Sub Set缺省药房()
'功能：设置当前行的缺省药房，同时处理相同类型的其他行的缺省药房
    Dim i As Long
    
    With vsfDrugStore
        If Val("" & .Cell(flexcpData, .Row, .ColIndex("缺省"))) = 0 Then  '该参数允许修改的情况下
            If .TextMatrix(.Row, .ColIndex("缺省")) = "√" Then
                .TextMatrix(.Row, .ColIndex("缺省")) = ""
            Else
                '当没有有权限修改可用时且可用为0（false)时不允许设置缺省
                If Not (Val(.TextMatrix(.Row, .ColIndex("可用"))) = 0 And Val("" & .Cell(flexcpData, .Row, .ColIndex("可用"))) = 1) Then
                    '同类别的其他行取消缺省
                    For i = .FixedRows To .Rows - 1
                        If .TextMatrix(.Row, .ColIndex("类别")) = .TextMatrix(i, .ColIndex("类别")) Then
                            If .TextMatrix(i, .ColIndex("缺省")) = "√" Then .TextMatrix(i, .ColIndex("缺省")) = ""
                        End If
                    Next
                    .TextMatrix(.Row, .ColIndex("可用")) = -1    '自动设置为可用
                    .TextMatrix(.Row, .ColIndex("缺省")) = "√"
                Else
                    MsgBox "设置当前药房为缺省时，会同时将当前药房设置为可用，" & vbNewLine & "你没有修改可用药房的权限。", vbInformation, gstrSysName
                End If
            End If
        Else
            MsgBox "你没有修改缺省药房的权限。", vbInformation, gstrSysName
        End If
    End With
End Sub

Private Sub Set可用药房(ByVal lngRow As Long, Optional ByVal blnAsk As Boolean = False)
'功能：设置当前行的可用药房，同时处理当前行的缺省药房

    With vsfDrugStore
        If Val("" & .Cell(flexcpData, lngRow, .ColIndex("可用"))) = 0 Then   '该参数允许修改的情况下
            If Val(.TextMatrix(lngRow, .ColIndex("可用"))) = -1 Then
                '当前科室勾选可用
                If Not (Val("" & .Cell(flexcpData, lngRow, .ColIndex("缺省"))) = 1 And .TextMatrix(lngRow, .ColIndex("缺省")) = "√") Then
                    .TextMatrix(lngRow, .ColIndex("可用")) = 0
                    .TextMatrix(lngRow, .ColIndex("缺省")) = ""
                Else
                    If blnAsk Then
                        MsgBox "取消当前药房可用时，会同时取消当前药房缺省，" & vbNewLine & "你没有修改缺省药房的权限。", vbInformation, gstrSysName
                    End If
                End If
            Else
                .TextMatrix(lngRow, .ColIndex("可用")) = -1    '自动设置为可用
            End If
        Else
            If blnAsk Then
                MsgBox "你没有修改可用药房的权限。", vbInformation, gstrSysName
            End If
        End If
    End With
End Sub

