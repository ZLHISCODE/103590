VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChargeTurnNew 
   AutoRedraw      =   -1  'True
   Caption         =   "门(急)诊费用转住院"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   12465
   Icon            =   "frmChargeTurnNew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   12465
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBill 
      BorderStyle     =   0  'None
      Height          =   4710
      Left            =   60
      ScaleHeight     =   4710
      ScaleWidth      =   11355
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
      ScaleHeight     =   495
      ScaleWidth      =   12465
      TabIndex        =   5
      Top             =   0
      Width           =   12465
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
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
            Size            =   10.5
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
            Size            =   10.5
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
            Size            =   10.5
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
            Size            =   10.5
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
            Size            =   10.5
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
            Size            =   10.5
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
            Size            =   10.5
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
      ScaleHeight     =   555
      ScaleWidth      =   12465
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7515
      Width           =   12465
      Begin VB.TextBox txtSum 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
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
            Size            =   10.5
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
      _ExtentX        =   21987
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmChargeTurnNew.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16907
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
Private mcllExecSql As Collection
Private mstr账单ID As String 'JSON串
Private mblnOk As Boolean

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

Private mrsPatient As ADODB.Recordset '病人信息
Private mrsFeeBill  As ADODB.Recordset '费用信息

Public Function ShowMe(objParent As Object, ByVal lng挂号ID As Long, _
    Optional ByVal bln独立执行 As Boolean = True, _
    Optional ByRef cllExecSql As Collection, Optional ByRef str账单ID As String) As Boolean
    '功能:新门诊病人门诊费用转住院费用
    '入参:
    '   lng挂号ID - 挂号ID
    '   bln独立执行 - 是否独立执行，如果是独立执行则会提交数据到数据库，否则由cllSql返回执行SQL
    '出参:
    '   cllExecSql bln独立执行为False时，传回执行SQL
    '   str账单ID 成功进行门诊费用转住院的新门诊账单ID,JSON格式
    '返回:成功返回True,否则返回False
    mlng挂号ID = lng挂号ID
    mbln独立执行 = bln独立执行
    Set cllExecSql = Nothing: mstr账单ID = ""
    mblnOk = False
    
    On Error Resume Next
    Me.Show vbModal, objParent
    ShowMe = mblnOk
    If Not bln独立执行 Then
        Set cllExecSql = mcllExecSql
        str账单ID = mstr账单ID
    End If
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
    
    Set mrsPerson = Nothing
    Set mrsDepartment = Nothing
    Set mrsChargeitem = Nothing
    
    Set mrsPatient = Nothing
    Set mrsFeeBill = Nothing
End Sub

Private Sub cmdOk_Click()
    Dim strSql As String
    Dim blnTrans As Boolean, i As Integer
    Dim strData As String
    Dim strPreNo As String, strNO As String, int序号 As Integer
    Dim str登记时间 As String
    
    On Error GoTo ErrHander
    mstr账单ID = ""
    Set mcllExecSql = New Collection
    
    str登记时间 = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:MM:ss")
    With mrsFeeBill
        .Sort = "单据号 Asc"
        Do While Not .EOF
            If InStr(mstr账单ID, "{""outp_bill_id"":" & NVL(mrsFeeBill!账单ID) & "}") = 0 Then
                mstr账单ID = mstr账单ID & ",{""outp_bill_id"":" & NVL(mrsFeeBill!账单ID) & "}"
            End If
            
            If NVL(!单据号) <> strPreNo Then
                strNO = zlDatabase.NextNo(14)
                int序号 = 1
                strPreNo = NVL(!单据号)
            End If
            
            'Zl_门诊费用转住院_三方转入(
            strSql = "Zl_门诊费用转住院_三方转入("
            '  No_In         住院费用记录.No%Type,
            strSql = strSql & "'" & strNO & "',"
            '  序号_In       住院费用记录.序号%Type,
            strSql = strSql & "" & int序号 & ","
            '  病人id_In     住院费用记录.病人id%Type,
            strSql = strSql & "" & NVL(mrsPatient!病人ID) & ","
            '  主页id_In     住院费用记录.主页id%Type,
            strSql = strSql & "" & "{『主页ID』}" & ","
            '  标识号_In     住院费用记录.标识号%Type,
            strSql = strSql & "" & "{『住院号』}" & ","
            '  姓名_In       住院费用记录.姓名%Type,
            strSql = strSql & "'" & NVL(mrsPatient!姓名) & "',"
            '  性别_In       住院费用记录.性别%Type,
            strSql = strSql & "'" & NVL(mrsPatient!性别) & "',"
            '  年龄_In       住院费用记录.年龄%Type,
            strSql = strSql & "'" & NVL(mrsPatient!年龄) & "',"
            '  床号_In       住院费用记录.床号%Type,
            strSql = strSql & "'" & NVL(mrsPatient!床号) & "',"
            '  费别_In       住院费用记录.费别%Type,
            strSql = strSql & "'" & NVL(!费别) & "',"
            '  病区id_In     住院费用记录.病人病区id%Type,
            strSql = strSql & "" & "{『病人病区ID』}" & ","
            '  科室id_In     住院费用记录.病人科室id%Type,
            strSql = strSql & "" & "{『病人科室ID』}" & ","
            '  开单部门id_In 住院费用记录.开单部门id%Type,
            strSql = strSql & "" & NVL(!开单科室ID) & ","
            '  开单人_In     住院费用记录.开单人%Type,
            strSql = strSql & "'" & NVL(!开单人) & "',"
            '  从属父号_In   住院费用记录.从属父号%Type,
            strSql = strSql & "" & "NULL" & ","
            '  收费细目id_In 住院费用记录.收费细目id%Type,
            strSql = strSql & "" & NVL(!收费细目ID) & ","
            '  收费类别_In   住院费用记录.收费类别%Type,
            strSql = strSql & "'" & NVL(!类别) & "',"
            '  计算单位_In   住院费用记录.计算单位%Type,
            strSql = strSql & "'" & NVL(!单位) & "',"
            '  付数_In       住院费用记录.付数%Type,
            strSql = strSql & "" & NVL(!付数) & ","
            '  数次_In       住院费用记录.数次%Type,
            strSql = strSql & "" & NVL(!数量) & ","
            '  执行部门id_In 住院费用记录.执行部门id%Type,
            strSql = strSql & "" & NVL(!执行科室ID) & ","
            '  价格父号_In   住院费用记录.价格父号%Type,
            strSql = strSql & "" & "NULL" & ","
            '  收入项目id_In 住院费用记录.收入项目id%Type,
            strSql = strSql & "" & NVL(!收入项目ID) & ","
            '  收据费目_In   住院费用记录.收据费目%Type,
            strSql = strSql & "'" & NVL(!收据费目) & "',"
            '  标准单价_In   住院费用记录.标准单价%Type,
            strSql = strSql & "" & NVL(!单价) & ","
            '  应收金额_In   住院费用记录.应收金额%Type,
            strSql = strSql & "" & NVL(!应收金额) & ","
            '  实收金额_In   住院费用记录.实收金额%Type,
            strSql = strSql & "" & NVL(!实收金额) & ","
            '  发生时间_In   住院费用记录.发生时间%Type,
            strSql = strSql & "To_Date('{『入院时间』}','yyyy-mm-dd hh24:mi:ss'),"
            '  登记时间_In   住院费用记录.登记时间%Type,
            strSql = strSql & "To_Date('" & str登记时间 & "','yyyy-mm-dd hh24:mi:ss'),"
            '  划价_In       Number,
            strSql = strSql & "" & IIf(NVL(!单据状态) = 1, 0, 1) & ","
            '  操作员编号_In 住院费用记录.操作员编号%Type,
            strSql = strSql & "'" & NVL(!操作员编号) & "',"
            '  操作员姓名_In 住院费用记录.操作员姓名%Type,
            strSql = strSql & "'" & NVL(!操作员姓名) & "',"
            '  执行人_In     住院费用记录.执行人%Type,
            strSql = strSql & "'" & NVL(!执行人) & "',"
            '  执行时间_In   住院费用记录.执行时间%Type,
            strSql = strSql & "To_Date('" & NVL(!执行时间) & "','yyyy-mm-dd hh24:mi:ss'),"
            '  医嘱序号_In   住院费用记录.医嘱序号%Type:=Null
            strSql = strSql & "" & NVL(!医嘱ID, "NULL") & ")"
            mcllExecSql.Add strSql
            
            int序号 = int序号 + 1
            mrsFeeBill.MoveNext
        Loop
    End With
    
    If mstr账单ID <> "" Then
        mstr账单ID = Mid(mstr账单ID, 2)
        mstr账单ID = "[" & mstr账单ID & "]"
    End If
    mstr账单ID = "{""input"":{""head"":{""bizno"":""RJ003"",""sysno"":""ZLDAYROOM"",""time"":"""",""action_no"":"""",""tarno"":""03""}," & _
        """bill_list"":" & mstr账单ID & "}}"
    
    If mbln独立执行 Then
        zlCommFun.ShowFlash "正在进行门诊费用转住院处理，请稍后...", Me
        cmdOK.Enabled = False
        gcnOracle.BeginTrans: blnTrans = True
        For i = 1 To mcllExecSql.Count
            strSql = Replace(mcllExecSql(i), "{『住院号』}", Val(NVL(mrsPatient!住院号)))
            strSql = Replace(strSql, "{『主页ID』}", Val(NVL(mrsPatient!主页ID)))
            strSql = Replace(strSql, "{『入院时间』}", Format(mrsPatient!入院日期, "yyyy-mm-dd hh:MM:ss"))
            strSql = Replace(strSql, "{『病人科室ID』}", Val(NVL(mrsPatient!病人科室ID)))
            strSql = Replace(strSql, "{『病人病区ID』}", Val(NVL(mrsPatient!病人病区ID)))
            zlDatabase.ExecuteProcedure strSql, Me.Caption
        Next
        
        '调用新门诊“门诊费用转住院确认”服务
        '输入   编码            名称        说明    数据类型        备注
        '       outp_bill_id    门诊账单ID          Number(18)      非空
        '输出   编码        名称        说明                数据类型        备注
        '       result      执行结果    1-成功；-1-失败     Number(1)       非空
        '       errmsg      错误消息    失败时返回错误消息  Varchar2(200)
        Call Sys.NewSystemSvr("新门诊系统", "门诊费用转住院费用确认", mstr账单ID, strData)
        If strData = "" Then strData = "{}"
        If Val(zlStr.JSONParse("result", strData)) <> 1 Then
            gcnOracle.RollbackTrans
            zlCommFun.StopFlash: cmdOK.Enabled = True
            MsgBox zlStr.JSONParse("errmsg", strData), vbInformation, gstrSysName
            Exit Sub
        End If
        gcnOracle.CommitTrans: blnTrans = False
        
        zlCommFun.StopFlash: cmdOK.Enabled = True
    End If
    
    mblnOk = True
    Unload Me
    Exit Sub
ErrHander:
    zlCommFun.StopFlash
    cmdOK.Enabled = True
    If blnTrans Then gcnOracle.RollbackTrans
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
        cmdCancel.Left = .ScaleLeft + .ScaleWidth - cmdCancel.width - 1000
        cmdOK.Left = cmdCancel.Left - cmdOK.width - 100
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
    
    lbl(txtName).Caption = NVL(mrsPatient!姓名)
    lbl(txtSex).Caption = NVL(mrsPatient!性别)
    lbl(txtAge).Caption = NVL(mrsPatient!年龄)
    lbl(txtInNumber).Caption = NVL(mrsPatient!住院号)
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
   strSql = _
        "Select a.病人id, a.主页id, a.住院号," & vbNewLine & _
        "       Nvl(a.姓名, b.姓名) As 姓名, Nvl(a.性别, b.性别) As 性别, Nvl(a.年龄, b.年龄) As 年龄," & vbNewLine & _
        "       Nvl(a.费别, b.费别) As 费别,Nvl(a.出院病床, b.当前床号) As 床号, a.入院日期," & vbNewLine & _
        "       a.入院病区ID As 病人病区id, a.入院科室id As 病人科室id" & vbNewLine & _
        "From 病案主页 A, 病人信息 B" & vbNewLine & _
        "Where a.病人id = b.病人id And a.挂号id = [1]"
    Set mrsPatient = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng挂号ID)
    If mrsPatient.RecordCount = 0 Then
        MsgBox "未找到病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
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
            .TextMatrix(i, .ColIndex("账单ID")) = NVL(rsBill!账单ID)
            .TextMatrix(i, .ColIndex("单据号")) = NVL(rsBill!单据号)
            .TextMatrix(i, .ColIndex("开单科室")) = NVL(rsBill!开单科室名称)
            .TextMatrix(i, .ColIndex("开单人")) = NVL(rsBill!开单人)
            .TextMatrix(i, .ColIndex("费别")) = NVL(rsBill!费别)
            
            If Val(NVL(rsBill!单据类型)) = 2 Then
                str单据 = IIf(Val(NVL(rsBill!单据状态)) = 0, "记账划价单", "记账单")
            Else
                str单据 = IIf(Val(NVL(rsBill!单据状态)) = 0, "收费划价单", "收费单")
            End If
            .TextMatrix(i, .ColIndex("单据")) = str单据
            .TextMatrix(i, .ColIndex("类别")) = NVL(rsBill!类别名称) & IIf(i Mod 2 = 1, "", " ")
            .TextMatrix(i, .ColIndex("名称")) = NVL(rsBill!项目名称)
            .TextMatrix(i, .ColIndex("规格")) = NVL(rsBill!规格)
            .TextMatrix(i, .ColIndex("单位")) = NVL(rsBill!单位)
            .TextMatrix(i, .ColIndex("数量")) = FormatEx(NVL(rsBill!付数) * NVL(rsBill!数量), 6, , , 2)
            .TextMatrix(i, .ColIndex("单价")) = FormatEx(NVL(rsBill!单价), 6, , , 2)
            
            .TextMatrix(i, .ColIndex("应收金额")) = FormatEx(NVL(rsBill!应收金额), 6, , , 2)
            .TextMatrix(i, .ColIndex("实收金额")) = FormatEx(NVL(rsBill!实收金额), 6, , , 2)
            .TextMatrix(i, .ColIndex("执行科室")) = NVL(rsBill!执行科室名称)
            .TextMatrix(i, .ColIndex("说明")) = IIf(NVL(rsBill!执行人) = "", "未执行", "完全执行")
            .TextMatrix(i, .ColIndex("费用时间")) = Format(NVL(rsBill!发生时间), "yyyy-mm-dd hh:MM:ss")
            
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
        
        .Subtotal flexSTSum, .ColIndex("单据号"), .ColIndex("应收金额"), gstrDec, &H8000000F, , False, "%s", , True
        .Subtotal flexSTSum, .ColIndex("单据号"), .ColIndex("实收金额"), gstrDec, &H8000000F, , False, "%s", , True
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
    lbl(txtName).Left = lbl(lblName).Left + lbl(lblName).width
    
    lbl(lblSex).Left = lbl(txtName).Left + lbl(txtName).width + sngSplit
    lbl(txtSex).Left = lbl(lblSex).Left + lbl(lblSex).width
    
    lbl(lblAge).Left = lbl(txtSex).Left + lbl(txtSex).width + sngSplit
    lbl(txtAge).Left = lbl(lblAge).Left + lbl(lblAge).width
    
    lbl(lblInNumber).Left = lbl(txtAge).Left + lbl(txtAge).width + sngSplit
    lbl(txtInNumber).Left = lbl(lblInNumber).Left + lbl(lblInNumber).width
End Sub

Private Function GetBillData(ByVal lng挂号ID As Long, ByRef strData As String) As Boolean
    '通过服务获取数据
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim strJsonIn As String, strNO As String
    
    On Error GoTo ErrHandler
    strSql = "Select NO From 病人挂号记录 Where ID =  [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng挂号ID)
    If rsTemp.EOF Then
        MsgBox "未找到病人挂号记录，无法确定挂号单据号！", vbInformation, gstrSysName
        Exit Function
    End If
    strNO = NVL(rsTemp!NO)
    
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
    strJsonIn = "{""input"":" & strJsonIn & ",""rgst_no"":""" & strNO & """}}"
    Call Sys.NewSystemSvr("新门诊系统", "门诊费用转住院费用", strJsonIn, strData)
    If strData = "" Then strData = "{}"
    If Val(zlStr.JSONParse("result", strData)) <> 1 Then
        MsgBox "获取可门诊费用转住院的费用信息时出错！" & vbCrLf & _
            zlStr.JSONParse("errmsg", strData), vbInformation, gstrSysName
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
                MsgBox "未找到单据【" & NVL(!单据号) & "】的开单科室信息！", vbInformation, gstrSysName
                Exit Function
            Else
                !开单科室名称 = NVL(mrsDepartment!名称)
            End If
            !开单人 = objScript.Eval("obj.bill_list[" & i & "].placer")
            !费别 = objScript.Eval("obj.bill_list[" & i & "].category_id")
            If NVL(!费别) = "" Then !费别 = NVL(mrsPatient!费别)
            !发生时间 = objScript.Eval("obj.bill_list[" & i & "].outp_bill_time")
            mrsPerson.Filter = "ID=" & Val(objScript.Eval("obj.bill_list[" & i & "].outp_bill_creator_id"))
            If mrsPerson.EOF Then
                MsgBox "未找到单据【" & NVL(!单据号) & "】的操作员信息！", vbInformation, gstrSysName
                Exit Function
            Else
                !操作员姓名 = NVL(mrsPerson!姓名)
                !操作员编号 = NVL(mrsPerson!编号)
            End If

            !收费细目ID = objScript.Eval("obj.bill_list[" & i & "].fee_id")
            !收入项目ID = objScript.Eval("obj.bill_list[" & i & "].acntsubj_id")
            mrsChargeitem.Filter = "ID=" & !收费细目ID & " And 收入项目ID=" & !收入项目ID
            If mrsChargeitem.EOF Then
                MsgBox "未找到单据【" & NVL(!单据号) & "】的收费项目信息！", vbInformation, gstrSysName
                Exit Function
            Else
                !类别 = NVL(mrsChargeitem!类别)
                !类别名称 = NVL(mrsChargeitem!类别名称)
                !项目名称 = NVL(mrsChargeitem!名称)
                !规格 = NVL(mrsChargeitem!规格)
                !单位 = NVL(mrsChargeitem!计算单位)
                !收据费目 = NVL(mrsChargeitem!收据费目)
            End If
            !付数 = objScript.Eval("obj.bill_list[" & i & "].crx_qunt")
            If Val(NVL(!付数)) = 0 Then !付数 = 1
            !数量 = objScript.Eval("obj.bill_list[" & i & "].outp_bill_detail_qunt")
            !单价 = objScript.Eval("obj.bill_list[" & i & "].fee_now_disct_price")
            !应收金额 = objScript.Eval("obj.bill_list[" & i & "].outp_bill_detail_chrg")
            !实收金额 = objScript.Eval("obj.bill_list[" & i & "].outp_bill_detail_disct_chrg")

            !执行科室ID = objScript.Eval("obj.bill_list[" & i & "].exedept_id")
            mrsDepartment.Filter = "ID=" & !执行科室ID
            If mrsDepartment.EOF Then
                MsgBox "未找到单据【" & NVL(!单据号) & "】的执行科室信息！", vbInformation, gstrSysName
                Exit Function
            Else
                !执行科室名称 = NVL(mrsDepartment!名称)
            End If
            !执行人 = objScript.Eval("obj.bill_list[" & i & "].exetr")
            !执行时间 = objScript.Eval("obj.bill_list[" & i & "].exetime")
            !医嘱ID = objScript.Eval("obj.bill_list[" & i & "].order_id")
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
