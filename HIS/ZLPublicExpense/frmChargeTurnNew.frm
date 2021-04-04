VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChargeTurnNew 
   AutoRedraw      =   -1  'True
   Caption         =   "门(急)诊费用转住院"
   ClientHeight    =   8436
   ClientLeft      =   60
   ClientTop       =   312
   ClientWidth     =   12468
   Icon            =   "frmChargeTurnNew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8436
   ScaleWidth      =   12468
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBill 
      BorderStyle     =   0  'None
      Height          =   4710
      Left            =   60
      ScaleHeight     =   4716
      ScaleWidth      =   11352
      TabIndex        =   7
      Top             =   660
      Width           =   11355
      Begin VSFlex8Ctl.VSFlexGrid vsfBill 
         Height          =   4440
         Left            =   60
         TabIndex        =   8
         Top             =   60
         Width           =   11040
         _cx             =   19473
         _cy             =   7832
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483633
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
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
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   492
      ScaleWidth      =   12468
      TabIndex        =   5
      Top             =   0
      Width           =   12465
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Index           =   4
         Left            =   4620
         TabIndex        =   6
         Top             =   180
         Width           =   840
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2145"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   14
         Left            =   5430
         TabIndex        =   16
         Top             =   180
         Width           =   420
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12岁"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   13
         Left            =   3750
         TabIndex        =   14
         Top             =   180
         Width           =   420
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Index           =   3
         Left            =   3150
         TabIndex        =   13
         Top             =   180
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "男"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   12
         Left            =   2340
         TabIndex        =   12
         Top             =   180
         Width           =   210
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Index           =   2
         Left            =   1740
         TabIndex        =   11
         Top             =   180
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "王二小"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   11
         Left            =   840
         TabIndex        =   10
         Top             =   180
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   180
         Width           =   630
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   555
      Left            =   0
      ScaleHeight     =   552
      ScaleWidth      =   12468
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7524
      Width           =   12465
      Begin VB.TextBox txtSum 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2580
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   138
         Width           =   1245
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   9660
         TabIndex        =   1
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   8490
         TabIndex        =   0
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   210
         TabIndex        =   2
         Top             =   120
         Width           =   1100
      End
      Begin VB.Label lblSum 
         AutoSize        =   -1  'True
         Caption         =   "转出合计:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1500
         TabIndex        =   15
         Top             =   190
         Width           =   1020
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   8070
      Width           =   12465
      _ExtentX        =   21992
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2350
            MinWidth        =   882
            Picture         =   "frmChargeTurnNew.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16955
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmChargeTurnNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mlng挂号ID As Long
Private mbln独立执行 As Boolean
Private mblnOk As Boolean
Private mblnRefreshData As Boolean

Private Enum idx_Lable
    lblName = 1
    txtName = 11
    lblSex = 2
    txtSex = 12
    lblAge = 3
    txtAge = 13
    lblInNumber = 4
    txtInNumber = 14
End Enum

Private mrsPerson As ADODB.Recordset '人员
Private mrsDepartment As ADODB.Recordset '部门
Private mrsChargeitem As ADODB.Recordset '收费项目

Private mobjPati As clsPatientInfo, mlng病人ID As Long '病人信息
Private mrsFeeBill  As ADODB.Recordset '费用信息

Public Function ShowMe(frmMain As Object, ByVal lng挂号ID As Long, _
    Optional ByVal bln独立执行 As Boolean = True, Optional ByRef blnRefreshData As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:新门诊病人门诊费用转住院费用
    '入参:
    '   lng挂号ID:挂号ID
    '   bln独立执行:是否独立执行，如果是独立执行则会提交数据到数据库，否则由 ExecuteTurn 接口单独执行
    '出参:
    '返回:成功返回True,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlng挂号ID = lng挂号ID
    mbln独立执行 = bln独立执行
    
    mblnOk = False
    On Error Resume Next
    Me.Show vbModal, frmMain
    ShowMe = mblnOk
    blnRefreshData = mblnRefreshData
End Function

Public Function ExecuteTurn(ByVal frmMain As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal str住院号 As String, ByVal dat入院时间 As Date, ByVal lng入院科室ID As Long, ByVal lng入院病区ID As Long, _
    ByRef strErrmsg_Out As String, Optional ByRef blnReflashData_Out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行门诊费用转住院，仅非独立执行时调用
    '入参:
    '出参:
    '   strErrMsg_Out=失败时返回错误原因
    '   blnReflashData_Out=是否有数据转入
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strData As String, blnTrans As Boolean
    Dim strSql As String, lng医疗小组ID As Long
    Dim strPreNo As String, strNo As String, int序号 As Integer
    Dim str登记时间 As String
    Dim str账单ID As String, cllPro As Collection
    
    On Error GoTo ErrHandler
    blnReflashData_Out = False
    If mrsFeeBill Is Nothing Then
        ExecuteTurn = Not mbln独立执行: Exit Function
    End If
    If mrsFeeBill.RecordCount = 0 Then
        ExecuteTurn = Not mbln独立执行: Exit Function
    End If
    
    If mlng病人ID <> lng病人ID Then
        strErrmsg_Out = "本次转入费用所属病人与当前病人不同，不允许执行门诊费用转住院。": Exit Function
    End If
    
    zlCommFun.ShowFlash "正在进行门诊费用转住院处理，请稍后...", frmMain
    str账单ID = ""
    str登记时间 = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:MM:ss")
    
    Set cllPro = New Collection
    With mrsFeeBill
        .Sort = "单据号 Asc"
        Do While Not .EOF
            If InStr(str账单ID, "{""outp_bill_id"":" & Nvl(mrsFeeBill!账单ID) & "}") = 0 Then
                str账单ID = str账单ID & ",{""outp_bill_id"":" & Nvl(mrsFeeBill!账单ID) & "}"
            End If
            
            If Nvl(!单据号) <> strPreNo Then
                strNo = zlDatabase.NextNo(14)
                int序号 = 1
                strPreNo = Nvl(!单据号)
                lng医疗小组ID = ZlGetMedicalGroupID(lng病人ID, lng主页ID, Nvl(!开单科室ID), Nvl(!开单人), dat入院时间)
            End If
            
            'Zl_门诊费用转住院_三方转入_S(
            strSql = "Zl_门诊费用转住院_三方转入_S("
            '  No_In         住院费用记录.No%Type,
            strSql = strSql & "'" & strNo & "',"
            '  序号_In       住院费用记录.序号%Type,
            strSql = strSql & "" & int序号 & ","
            '  病人id_In     住院费用记录.病人id%Type,
            strSql = strSql & "" & mobjPati.病人ID & ","
            '  主页id_In     住院费用记录.主页id%Type,
            strSql = strSql & "" & ZVal(lng主页ID) & ","
            '  标识号_In     住院费用记录.标识号%Type,
            strSql = strSql & "" & ZVal(str住院号) & ","
            '  姓名_In       住院费用记录.姓名%Type,
            strSql = strSql & "'" & mobjPati.姓名 & "',"
            '  性别_In       住院费用记录.性别%Type,
            strSql = strSql & "'" & mobjPati.性别 & "',"
            '  年龄_In       住院费用记录.年龄%Type,
            strSql = strSql & "'" & mobjPati.年龄 & "',"
            '  床号_In       住院费用记录.床号%Type,
            strSql = strSql & "'" & mobjPati.床号 & "',"
            '  费别_In       住院费用记录.费别%Type,
            strSql = strSql & "'" & Nvl(!费别) & "',"
            '  病区id_In     住院费用记录.病人病区id%Type,
            strSql = strSql & "" & ZVal(lng入院病区ID) & ","
            '  科室id_In     住院费用记录.病人科室id%Type,
            strSql = strSql & "" & ZVal(lng入院科室ID) & ","
            '  开单部门id_In 住院费用记录.开单部门id%Type,
            strSql = strSql & "" & Nvl(!开单科室ID) & ","
            '  开单人_In     住院费用记录.开单人%Type,
            strSql = strSql & "'" & Nvl(!开单人) & "',"
            '  从属父号_In   住院费用记录.从属父号%Type,
            strSql = strSql & "" & "NULL" & ","
            '  收费细目id_In 住院费用记录.收费细目id%Type,
            strSql = strSql & "" & Nvl(!收费细目ID) & ","
            '  收费类别_In   住院费用记录.收费类别%Type,
            strSql = strSql & "'" & Nvl(!类别) & "',"
            '  计算单位_In   住院费用记录.计算单位%Type,
            strSql = strSql & "'" & Nvl(!单位) & "',"
            '  付数_In       住院费用记录.付数%Type,
            strSql = strSql & "" & Nvl(!付数) & ","
            '  数次_In       住院费用记录.数次%Type,
            strSql = strSql & "" & Nvl(!数量) & ","
            '  执行部门id_In 住院费用记录.执行部门id%Type,
            strSql = strSql & "" & Nvl(!执行科室ID) & ","
            '  价格父号_In   住院费用记录.价格父号%Type,
            strSql = strSql & "" & "NULL" & ","
            '  收入项目id_In 住院费用记录.收入项目id%Type,
            strSql = strSql & "" & Nvl(!收入项目ID) & ","
            '  收据费目_In   住院费用记录.收据费目%Type,
            strSql = strSql & "'" & Nvl(!收据费目) & "',"
            '  标准单价_In   住院费用记录.标准单价%Type,
            strSql = strSql & "" & Nvl(!单价) & ","
            '  应收金额_In   住院费用记录.应收金额%Type,
            strSql = strSql & "" & Nvl(!应收金额) & ","
            '  实收金额_In   住院费用记录.实收金额%Type,
            strSql = strSql & "" & Nvl(!实收金额) & ","
            '  发生时间_In   住院费用记录.发生时间%Type,
            strSql = strSql & "To_Date('" & Format(dat入院时间, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
            '  登记时间_In   住院费用记录.登记时间%Type,
            strSql = strSql & "To_Date('" & str登记时间 & "','yyyy-mm-dd hh24:mi:ss'),"
            '  划价_In       Number,
            strSql = strSql & "" & IIf(Nvl(!单据状态) = 1, 0, 1) & ","
            '  操作员编号_In 住院费用记录.操作员编号%Type,
            strSql = strSql & "'" & Nvl(!操作员编号) & "',"
            '  操作员姓名_In 住院费用记录.操作员姓名%Type,
            strSql = strSql & "'" & Nvl(!操作员姓名) & "',"
            '  执行人_In     住院费用记录.执行人%Type,
            strSql = strSql & "'" & Nvl(!执行人) & "',"
            '  执行时间_In   住院费用记录.执行时间%Type,
            strSql = strSql & "To_Date('" & Nvl(!执行时间) & "','yyyy-mm-dd hh24:mi:ss'),"
            '  医嘱序号_In   住院费用记录.医嘱序号%Type:=Null,
            strSql = strSql & "" & Nvl(!医嘱id, "NULL") & ","
            '  医疗小组id_In 住院费用记录.医疗小组id%Type,
            strSql = strSql & "" & ZVal(lng医疗小组ID) & ","
            '  审核标志_In   Number,
            strSql = strSql & "" & mobjPati.审核标志 & ","
            '  住院状态_In Number
            strSql = strSql & "" & mobjPati.住院状态 & ")"
            cllPro.Add strSql
            
            int序号 = int序号 + 1
            mrsFeeBill.MoveNext
        Loop
    End With
    
    str账单ID = "{""input"":{""head"":{""bizno"":""RJ003"",""sysno"":""ZLDAYROOM"",""time"":"""",""action_no"":"""",""tarno"":""03""}," & _
        """bill_list"":[" & Mid(str账单ID, 2) & "]}}"
    
    gcnOracle.BeginTrans: blnTrans = True
        zlDatabase.ExecuteProcedureBeach cllPro, "执行门诊费用转住院", False, False
        
        '调用新门诊“门诊费用转住院确认”服务
        '输入   编码            名称        说明    数据类型        备注
        '       outp_bill_id    门诊账单ID          Number(18)      非空
        '输出   编码        名称        说明                数据类型        备注
        '       result      执行结果    1-成功；-1-失败     Number(1)       非空
        '       errmsg      错误消息    失败时返回错误消息  Varchar2(200)
        Call Sys.NewSystemSvr("新门诊系统", "门诊费用转住院费用确认", str账单ID, strData)
        If strData = "" Then strData = "{}"
        If Val(zlstr.JSONParse("result", strData)) <> 1 Then
            gcnOracle.RollbackTrans
            zlCommFun.StopFlash
            strErrmsg_Out = zlstr.JSONParse("errmsg", strData)
            Exit Function
        End If
    gcnOracle.CommitTrans: blnTrans = False
    
    zlCommFun.StopFlash
    blnReflashData_Out = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    zlCommFun.StopFlash
    strErrmsg_Out = Err.Description
End Function

Private Sub Form_Load()
    Dim strData As String
    
    zlCommFun.ShowFlash "正在获取可转入的门诊费用列表，请稍后...", Me
    If GetBillData(mlng挂号ID, strData) = False Then GoTo ErrExit:
    If InitData(mlng挂号ID) = False Then GoTo ErrExit:
    If AnalyzeData(strData, mrsFeeBill) = False Then GoTo ErrExit:
    If InitFace() = False Then GoTo ErrExit:
    If ShowBills(mrsFeeBill) = False Then GoTo ErrExit:
    zlCommFun.StopFlash
    Exit Sub
ErrExit:
    zlCommFun.StopFlash
    Unload Me: Exit Sub
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picTop.Move 0, 0, Me.ScaleWidth, picTop.Height
    
    sta.Move 0, Me.ScaleHeight - sta.Height, Me.ScaleWidth, sta.Height
    picBottom.Move 0, sta.Top - picBottom.Height, Me.ScaleWidth, picBottom.Height
    
    picBill.Move 0, picTop.Height, Me.ScaleWidth, picBottom.Top - picTop.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    zl_vsGrid_Para_Save 1131, vsfBill, Me.Caption, "门诊转住院列表_New", True
End Sub

Private Sub cmdOK_Click()
    Dim strErrMsg As String
    
    On Error GoTo ErrHander
    If mrsFeeBill Is Nothing Then
        MsgBox "当前无需要转入的费用。", vbInformation, gstrSysName
        Exit Sub
    End If
    If mrsFeeBill.RecordCount = 0 Then
        MsgBox "当前无需要转入的费用。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mbln独立执行 Then
        cmdOk.Enabled = False
        If ExecuteTurn(Me, mobjPati.病人ID, mobjPati.主页ID, mobjPati.住院号, _
            mobjPati.入院日期, mobjPati.当前科室id, mobjPati.当前病区id, strErrMsg, mblnRefreshData) = False Then
            If strErrMsg <> "" Then MsgBox strErrMsg, vbInformation, gstrSysName
            cmdOk.Enabled = True
            Exit Sub
        End If
        cmdOk.Enabled = True
    End If
    
    mblnOk = True
    Unload Me
    Exit Sub
ErrHander:
    cmdOk.Enabled = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub picBill_Resize()
    On Error Resume Next
    vsfBill.Move 0, 0, picBill.ScaleWidth, picBill.ScaleHeight
End Sub

Private Sub picBottom_Resize()
    On Error Resume Next
    With picBottom
        cmdCancel.Left = .ScaleLeft + .ScaleWidth - cmdCancel.Width - 1000
        cmdOk.Left = cmdCancel.Left - cmdOk.Width - 100
    End With
End Sub

Private Function InitFace() As Boolean
    '初始化界面
    Dim strHead As String
    Dim varHead As Variant, varItem As Variant
    Dim i As Long
    
    On Error GoTo ErrHandler
    With vsfBill
        .Redraw = flexRDNone
        .RowHeightMin = 300
        .Clear
        .Rows = 2
        .FixedRows = 1: .FixedCols = 0
        
        strHead = "单据号,1,0|账单ID,1,0|开单科室,1,0|开单人,1,0|费别,1,0|单据,1,0|" & _
                "类别,1,800|名称,1,2100|规格,1,1400|单位,1,600|数量,7,800|单价,7,1000|" & _
                "应收金额,7,1000|实收金额,7,1000|执行科室,1,1000|说明,1,850|费用时间,1,1800"
        varHead = Split(strHead, "|")
        .Cols = UBound(varHead) + 1
        For i = 0 To UBound(varHead)
            varItem = Split(varHead(i), ",")
            .TextMatrix(0, i) = varItem(0)
            .ColKey(i) = varItem(0)
            .ColAlignment(i) = varItem(1)
            .ColWidth(i) = varItem(2)
            If .ColWidth(i) = 0 Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        zl_vsGrid_Para_Restore 1131, vsfBill, Me.Caption, "门诊转住院列表_New", True
        .Redraw = flexRDBuffered
    End With
    
    lbl(txtName).Caption = mobjPati.姓名
    lbl(txtSex).Caption = mobjPati.性别
    lbl(txtAge).Caption = mobjPati.年龄
    lbl(txtInNumber).Caption = mobjPati.住院号
    Call SetPatiControl
    InitFace = True
    Exit Function
ErrHandler:
    vsfBill.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitData(ByVal lng挂号ID As Long) As Boolean
    '初始化数据
    Dim strSql As String
    
    On Error GoTo ErrHandler
    '人员
    strSql = _
        "Select ID, 编号, 姓名" & vbNewLine & _
        "From 人员表" & vbNewLine & _
        "Where (撤档时间 Is Null Or 撤档时间 = To_Date('3000-01-01','yyyy-mm-dd'))"
    Set mrsPerson = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    '部门
    strSql = _
        "Select ID, 编码, 名称" & vbNewLine & _
        "From 部门表" & vbNewLine & _
        "Where (撤档时间 Is Null Or 撤档时间 = To_Date('3000-01-01','yyyy-mm-dd'))"
    Set mrsDepartment = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    '收费项目
    strSql = _
        "Select Distinct a.Id, a.名称, a.规格, a.计算单位, a.类别, d.名称 As 类别名称, b.收入项目ID, c.收据费目" & vbNewLine & _
        "From 收费项目目录 A, 收费价目 B, 收入项目 C, 收费项目类别 D" & vbNewLine & _
        "Where a.Id = b.收费细目id And b.收入项目id = c.Id And a.类别 = d.编码" & vbNewLine & _
        "      And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01','yyyy-mm-dd'))"
    Set mrsChargeitem = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    '病人信息
    If GetPatiInfoByPage(mobjPati, lng挂号ID) = False Then
        MsgBox "未找到病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
    mlng病人ID = mobjPati.病人ID
    InitData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowBills(ByVal rsBill As ADODB.Recordset) As Boolean
    '显示病人可门诊转住院的费用
    Dim i As Integer
    Dim str单据 As String
    
    On Error GoTo ErrHandler
    With vsfBill
        .Redraw = flexRDNone
        .Clear 1
        .Rows = rsBill.RecordCount + 1
        i = 1
        If rsBill.RecordCount > 0 Then rsBill.MoveFirst
        Do While Not rsBill.EOF
            .TextMatrix(i, .ColIndex("账单ID")) = Nvl(rsBill!账单ID)
            .TextMatrix(i, .ColIndex("单据号")) = Nvl(rsBill!单据号)
            .TextMatrix(i, .ColIndex("开单科室")) = Nvl(rsBill!开单科室名称)
            .TextMatrix(i, .ColIndex("开单人")) = Nvl(rsBill!开单人)
            .TextMatrix(i, .ColIndex("费别")) = Nvl(rsBill!费别)
            
            If Val(Nvl(rsBill!单据类型)) = 2 Then
                str单据 = IIf(Val(Nvl(rsBill!单据状态)) = 0, "记账划价单", "记账单")
            Else
                str单据 = IIf(Val(Nvl(rsBill!单据状态)) = 0, "收费划价单", "收费单")
            End If
            .TextMatrix(i, .ColIndex("单据")) = str单据
            .TextMatrix(i, .ColIndex("类别")) = Nvl(rsBill!类别名称) & IIf(i Mod 2 = 1, "", " ")
            .TextMatrix(i, .ColIndex("名称")) = Nvl(rsBill!项目名称)
            .TextMatrix(i, .ColIndex("规格")) = Nvl(rsBill!规格)
            .TextMatrix(i, .ColIndex("单位")) = Nvl(rsBill!单位)
            .TextMatrix(i, .ColIndex("数量")) = FormatEx(Nvl(rsBill!付数) * Nvl(rsBill!数量), 6, , , 2)
            .TextMatrix(i, .ColIndex("单价")) = FormatEx(Nvl(rsBill!单价), 6, , , 2)
            
            .TextMatrix(i, .ColIndex("应收金额")) = FormatEx(Nvl(rsBill!应收金额), 6, , , 2)
            .TextMatrix(i, .ColIndex("实收金额")) = FormatEx(Nvl(rsBill!实收金额), 6, , , 2)
            .TextMatrix(i, .ColIndex("执行科室")) = Nvl(rsBill!执行科室名称)
            .TextMatrix(i, .ColIndex("说明")) = IIf(Nvl(rsBill!执行人) = "", "未执行", "完全执行")
            .TextMatrix(i, .ColIndex("费用时间")) = Format(Nvl(rsBill!发生时间), "yyyy-mm-dd hh:MM:ss")
            
            i = i + 1
            rsBill.MoveNext
        Loop
        .Redraw = flexRDBuffered
    End With
   
    Call SetSumMoney '转出合计
    Call SplitGroupShow '分组显示
    ShowBills = True
    Exit Function
ErrHandler:
    vsfBill.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SplitGroupShow()
    '费用列表信息进行分组显示
    Dim i As Integer, j As Integer
    Dim strTemp As String

    On Error GoTo ErrHandler
    With vsfBill
        .OutlineBar = flexOutlineBarComplete
        .Subtotal flexSTClear
        .MultiTotals = True
        
        .Subtotal flexSTSum, .ColIndex("单据号"), .ColIndex("应收金额"), gSysPara.Money_Decimal.strFormt_VB, &H8000000F, , False, "%s", , True
        .Subtotal flexSTSum, .ColIndex("单据号"), .ColIndex("实收金额"), gSysPara.Money_Decimal.strFormt_VB, &H8000000F, , False, "%s", , True
        .SubtotalPosition = flexSTAbove

        .Outline .ColIndex("类别")
        .OutlineCol = .ColIndex("类别")

        For i = 1 To .Rows - 1
            .MergeRow(i) = False
            If .IsSubtotal(i) Then
                .IsCollapsed(i) = flexOutlineExpanded
                .MergeRow(i) = True
                .RowHeight(i) = 350

                strTemp = .TextMatrix(i + 1, .ColIndex("单据号")) & "(" & .TextMatrix(i + 1, .ColIndex("单据")) & ")"
                strTemp = strTemp & Space(2) & "费别:" & .TextMatrix(i + 1, .ColIndex("费别"))
                strTemp = strTemp & Space(2) & "开单科室:" & .TextMatrix(i + 1, .ColIndex("开单科室"))
                strTemp = strTemp & Space(2) & "开单人:" & .TextMatrix(i + 1, .ColIndex("开单人"))
                
                For j = 0 To .Cols - 1
                   If j >= .ColIndex("类别") And j < .ColIndex("应收金额") Then
                       .Cell(flexcpText, i, j) = strTemp
                   ElseIf .ColIndex("应收金额") = j Then
                       .TextMatrix(i, j) = FormatEx(Val(.TextMatrix(i, j)), 6, , , 2)
                   ElseIf .ColIndex("实收金额") = j Then
                       .TextMatrix(i, j) = " " & FormatEx(Val(.TextMatrix(i, j)), 6, , , 2)
                   End If
                Next
            End If
        Next
        
        .MergeCells = flexMergeRestrictRows
        For i = 0 To .Cols - 1
            If i < .ColIndex("应收金额") Then
                .MergeCol(i) = True
            Else
                .MergeCol(i) = False
            End If
        Next
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetSumMoney()
    '设置和显示费用合计
    Dim i As Long, dblSum As Double
    
    On Error GoTo ErrHander
    With vsfBill
        For i = .FixedRows To .Rows - 1
            dblSum = dblSum + Val(.TextMatrix(i, .ColIndex("实收金额")))
        Next
    End With
    txtSum.Text = Format(dblSum, "###0.00;-###0.00;0.00;0.00")
    Exit Sub
ErrHander:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetPatiControl()
    '设置病人信息控件位置
    Dim sngSplit As Single
    
    sngSplit = 600
    On Error Resume Next
    lbl(txtName).Left = lbl(lblName).Left + lbl(lblName).Width
    
    lbl(lblSex).Left = lbl(txtName).Left + lbl(txtName).Width + sngSplit
    lbl(txtSex).Left = lbl(lblSex).Left + lbl(lblSex).Width
    
    lbl(lblAge).Left = lbl(txtSex).Left + lbl(txtSex).Width + sngSplit
    lbl(txtAge).Left = lbl(lblAge).Left + lbl(lblAge).Width
    
    lbl(lblInNumber).Left = lbl(txtAge).Left + lbl(txtAge).Width + sngSplit
    lbl(txtInNumber).Left = lbl(lblInNumber).Left + lbl(lblInNumber).Width
End Sub

Private Function GetBillData(ByVal lng挂号ID As Long, ByRef strData As String) As Boolean
    '通过服务获取数据
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim strJsonIn As String, strNo As String
    
    On Error GoTo ErrHandler
    strSql = "Select NO From 病人挂号记录 Where ID =  [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng挂号ID)
    If rsTemp.EOF Then
        MsgBox "未找到病人挂号记录，无法确定挂号单据号！", vbInformation, gstrSysName
        Exit Function
    End If
    strNo = Nvl(rsTemp!NO)
    
    '调用新门诊“门诊费用转住院”服务
    '输入    编码       名称        说明    数据类型        备注
    '        rgst_no    挂号单号            Varchar2(8)      非空
    '输出    编码                   名称            说明                    数据类型    备注
    '        result                 执行结果        1-成功；-1-失败         Number(1)   非空
    '        errmsg                 错误消息        失败时返回错误消息      Varchar2(200)
    '        outp_bill_id           门诊账单ID      ZLHIS返回确认用         Number(18)  非空
    '        outp_bill_no           单据号          账单号，区分单据        VARCHAR2(20)    非空
    '        outp_kacnt_sign        单据类型        1-收费单;2-记账单       Number(1)   非空
    '        pricing_sign           单据状态        0-划价收费单/划价记账单;1-门诊收费单/门诊记账单 Number(1)   非空
    '        plcdept_id             开单科室ID      部门表.ID               Number(18)  非空
    '        placer                 开单医生        人员表.姓名             VARCHAR2(70)    非空
    '        outp_bill_time         单据时间        门诊账单时间            Date    非空
    '        order_id               医嘱ID          病人医嘱记录.ID         Number(18)
    '        outp_bill_creator_id   操作员ID        门诊账单创建者ID         Number(18)  非空
    '        category_id            费别                                    VARCHAR2(20)
    '        fee_id                 收费项目ID      收费项目目录.ID         Number(18)  非空
    '        acntsubj_id            收入项目ID      收入项目.ID             Number(18)  非空
    '        crx_qunt               付数            中草药总剂数            NUMBER(4)   非空
    '        outp_bill_detail_qunt  数量                                    NUMBER(18,5)    非空
    '        fee_now_disct_price    单价                                    NUMBER(18,4)    非空
    '        outp_bill_detail_chrg  应收金额        付数*数量*单价          NUMBER(18,3)    非空
    '        outp_bill_detail_disct_chrg 实收金额   应收金额-折扣金额       NUMBER(18,3)    非空
    '        exedept_id             执行科室ID      部门表.ID               Number(18)  非空
    '        exetr                  执行人          人员表.姓名；为空表示未执行，不为空表示完全执行 VARCHAR2(70)
    '        exetime                执行时间                                Date
    strJsonIn = "{""head"":{""bizno"":""RJ002"",""sysno"":""ZLDAYROOM"",""time"":"""",""action_no"":"""",""tarno"":""03""}"
    strJsonIn = "{""input"":" & strJsonIn & ",""rgst_no"":""" & strNo & """}}"
    Call Sys.NewSystemSvr("新门诊系统", "门诊费用转住院费用", strJsonIn, strData)
    If strData = "" Then strData = "{}"
    If Val(zlstr.JSONParse("result", strData)) <> 1 Then
        MsgBox "获取可门诊费用转住院的费用信息时出错！" & vbCrLf & _
            zlstr.JSONParse("errmsg", strData), vbInformation, gstrSysName
        Exit Function
    End If
    GetBillData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function AnalyzeData(ByVal strData As String, ByRef rsRecord As ADODB.Recordset) As Boolean
    '从JSON字符串中解析数据
    '入参：
    '   strData JSON字符串
    '出参：
    '   rsRecord 费用记录
    '返回：解析成功，返回True,否则返回False
    Dim i As Integer
    Dim objScript As Object '用于解析JSON
    
    On Error GoTo ErrHandler
    Set objScript = CreateObject("MSScriptControl.ScriptControl")
    objScript.Language = "JScript"
    objScript.AddCode "var obj=" & strData & ";"
    
    Set rsRecord = CreateBillRecord()
    
    With rsRecord
        For i = 0 To objScript.Eval("obj.bill_list.length") - 1
            .AddNew
            !账单ID = objScript.Eval("obj.bill_list[" & i & "].outp_bill_id")
            !单据号 = objScript.Eval("obj.bill_list[" & i & "].outp_bill_no")
            !单据类型 = objScript.Eval("obj.bill_list[" & i & "].outp_kacnt_sign")
            !单据状态 = objScript.Eval("obj.bill_list[" & i & "].pricing_sign")
            !开单科室ID = objScript.Eval("obj.bill_list[" & i & "].plcdept_id")
            mrsDepartment.Filter = "ID=" & !开单科室ID
            If mrsDepartment.EOF Then
                MsgBox "未找到单据【" & Nvl(!单据号) & "】的开单科室信息！", vbInformation, gstrSysName
                Exit Function
            Else
                !开单科室名称 = Nvl(mrsDepartment!名称)
            End If
            !开单人 = objScript.Eval("obj.bill_list[" & i & "].placer")
            !费别 = objScript.Eval("obj.bill_list[" & i & "].category_id")
            If Nvl(!费别) = "" Then !费别 = mobjPati.费别
            !发生时间 = objScript.Eval("obj.bill_list[" & i & "].outp_bill_time")
            mrsPerson.Filter = "ID=" & Val(objScript.Eval("obj.bill_list[" & i & "].outp_bill_creator_id"))
            If mrsPerson.EOF Then
                MsgBox "未找到单据【" & Nvl(!单据号) & "】的操作员信息！", vbInformation, gstrSysName
                Exit Function
            Else
                !操作员姓名 = Nvl(mrsPerson!姓名)
                !操作员编号 = Nvl(mrsPerson!编号)
            End If

            !收费细目ID = objScript.Eval("obj.bill_list[" & i & "].fee_id")
            !收入项目ID = objScript.Eval("obj.bill_list[" & i & "].acntsubj_id")
            mrsChargeitem.Filter = "ID=" & !收费细目ID & " And 收入项目ID=" & !收入项目ID
            If mrsChargeitem.EOF Then
                MsgBox "未找到单据【" & Nvl(!单据号) & "】的收费项目信息！", vbInformation, gstrSysName
                Exit Function
            Else
                !类别 = Nvl(mrsChargeitem!类别)
                !类别名称 = Nvl(mrsChargeitem!类别名称)
                !项目名称 = Nvl(mrsChargeitem!名称)
                !规格 = Nvl(mrsChargeitem!规格)
                !单位 = Nvl(mrsChargeitem!计算单位)
                !收据费目 = Nvl(mrsChargeitem!收据费目)
            End If
            !付数 = objScript.Eval("obj.bill_list[" & i & "].crx_qunt")
            If Val(Nvl(!付数)) = 0 Then !付数 = 1
            !数量 = objScript.Eval("obj.bill_list[" & i & "].outp_bill_detail_qunt")
            !单价 = objScript.Eval("obj.bill_list[" & i & "].fee_now_disct_price")
            !应收金额 = objScript.Eval("obj.bill_list[" & i & "].outp_bill_detail_chrg")
            !实收金额 = objScript.Eval("obj.bill_list[" & i & "].outp_bill_detail_disct_chrg")

            !执行科室ID = objScript.Eval("obj.bill_list[" & i & "].exedept_id")
            mrsDepartment.Filter = "ID=" & !执行科室ID
            If mrsDepartment.EOF Then
                MsgBox "未找到单据【" & Nvl(!单据号) & "】的执行科室信息！", vbInformation, gstrSysName
                Exit Function
            Else
                !执行科室名称 = Nvl(mrsDepartment!名称)
            End If
            !执行人 = objScript.Eval("obj.bill_list[" & i & "].exetr")
            !执行时间 = objScript.Eval("obj.bill_list[" & i & "].exetime")
            !医嘱id = objScript.Eval("obj.bill_list[" & i & "].order_id")
        Next
        .UpdateBatch '批量更新
    End With
    Set objScript = Nothing
    AnalyzeData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CreateBillRecord() As ADODB.Recordset
    '创建记录集对象
    Dim rsRecord As ADODB.Recordset
    
    On Error GoTo ErrHandler
    Set rsRecord = New ADODB.Recordset
    rsRecord.Fields.Append "账单ID", adBigInt, , adFldIsNullable
    rsRecord.Fields.Append "单据号", adVarChar, 20, adFldIsNullable
    rsRecord.Fields.Append "单据类型", adInteger, , adFldIsNullable
    rsRecord.Fields.Append "单据状态", adInteger, , adFldIsNullable
    rsRecord.Fields.Append "开单科室ID", adBigInt, , adFldIsNullable
    rsRecord.Fields.Append "开单科室名称", adVarChar, 100, adFldIsNullable
    rsRecord.Fields.Append "开单人", adVarChar, 100, adFldIsNullable
    rsRecord.Fields.Append "费别", adVarChar, 50, adFldIsNullable
    rsRecord.Fields.Append "发生时间", adVarChar, 20, adFldIsNullable
    rsRecord.Fields.Append "操作员编号", adVarChar, 10, adFldIsNullable
    rsRecord.Fields.Append "操作员姓名", adVarChar, 100, adFldIsNullable
    
    rsRecord.Fields.Append "类别", adVarChar, 10, adFldIsNullable
    rsRecord.Fields.Append "类别名称", adVarChar, 20, adFldIsNullable
    rsRecord.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    rsRecord.Fields.Append "项目名称", adVarChar, 200, adFldIsNullable
    rsRecord.Fields.Append "规格", adVarChar, 100, adFldIsNullable
    rsRecord.Fields.Append "单位", adVarChar, 20, adFldIsNullable
    rsRecord.Fields.Append "收入项目ID", adBigInt, , adFldIsNullable
    rsRecord.Fields.Append "收据费目", adVarChar, 50, adFldIsNullable
    rsRecord.Fields.Append "付数", adDouble, , adFldIsNullable
    rsRecord.Fields.Append "数量", adDouble, , adFldIsNullable
    rsRecord.Fields.Append "单价", adDouble, , adFldIsNullable
    rsRecord.Fields.Append "应收金额", adDouble, , adFldIsNullable
    rsRecord.Fields.Append "实收金额", adDouble, , adFldIsNullable
    
    rsRecord.Fields.Append "执行科室ID", adBigInt, , adFldIsNullable
    rsRecord.Fields.Append "执行科室名称", adVarChar, 100, adFldIsNullable
    rsRecord.Fields.Append "执行人", adVarChar, 100, adFldIsNullable
    rsRecord.Fields.Append "执行时间", adVarChar, 20, adFldIsNullable
    rsRecord.Fields.Append "医嘱ID", adBigInt, , adFldIsNullable
    
    rsRecord.CursorLocation = adUseClient
    rsRecord.LockType = adLockOptimistic
    rsRecord.CursorType = adOpenStatic
    rsRecord.Open
    
    Set CreateBillRecord = rsRecord
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPatiInfoByPage(objPati As clsPatientInfo, ByVal lng挂号ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:从病案主页中获取病人信息
    '入参:
    '   lng挂号ID-病案主页.挂号ID
    '出参:
    '   objPati-返回病人信息对象
    '返回:成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo errHandle
    If zlGetServiceObject(objService) = False Then Exit Function
     
    If objService.ZlCissvr_GetPatiPageInfo(1, "", rsTemp, , , , , , lng挂号ID) = False Then Exit Function
    If rsTemp Is Nothing Then Exit Function
    If rsTemp.EOF Then Exit Function
    
    Set objPati = New clsPatientInfo
    With objPati
        .病人ID = Nvl(rsTemp!病人ID)
        .主页ID = Nvl(rsTemp!主页ID)
        .姓名 = Nvl(rsTemp!姓名)
        .性别 = Nvl(rsTemp!性别)
        .年龄 = Nvl(rsTemp!年龄)
        .费别 = Nvl(rsTemp!费别)
        .医疗付款方式 = Nvl(rsTemp!医疗付款方式名称)
        .医疗付款方式编码 = Nvl(rsTemp!医疗付款方式编码)
        .险类 = Val(Nvl(rsTemp!险类))
        .险类名称 = GetInsureName(Val(Nvl(rsTemp!险类)))
        .病人类型 = Nvl(rsTemp!病人类型)
        .当前病区id = Val(Nvl(rsTemp!当前病区id))
        .当前病区名称 = Nvl(rsTemp!当前病区名称)
        .当前科室id = Val(Nvl(rsTemp!当前科室id))
        .当前科室名称 = Nvl(rsTemp!当前科室名称)
        .床号 = Nvl(rsTemp!当前床号)
        .住院号 = Nvl(rsTemp!住院号)
        .病人性质 = Val(Nvl(rsTemp!病人性质))
        .入院日期 = Nvl(rsTemp!入院时间)
        .出院日期 = Nvl(rsTemp!出院时间)
        .住院医师 = Nvl(rsTemp!住院医师)
        .病人备注 = Nvl(rsTemp!病人备注)
        .住院状态 = Val(Nvl(rsTemp!住院状态))
        .审核标志 = Val(Nvl(rsTemp!审核标志))
        .编目日期 = Nvl(rsTemp!编目日期)
        .医保号 = Nvl(rsTemp!医保号)
        .挂号ID = Val(Nvl(rsTemp!挂号ID))
    End With
    GetPatiInfoByPage = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

