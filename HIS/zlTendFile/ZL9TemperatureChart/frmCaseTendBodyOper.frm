VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCaseTendBodyOper 
   Caption         =   "设置手术/分娩"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7785
   Icon            =   "frmCaseTendBodyOper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   7785
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picStb 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      FillColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   -15
      ScaleHeight     =   360
      ScaleWidth      =   2415
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4995
      Width           =   2415
      Begin VB.Label lblStb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   45
      End
   End
   Begin VB.PictureBox picOper 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   4995
      Left            =   135
      ScaleHeight     =   4965
      ScaleWidth      =   8100
      TabIndex        =   1
      Top             =   255
      Width           =   8130
      Begin zl9TemperatureChart.VsfGrid vsfOper 
         Height          =   3810
         Left            =   -15
         TabIndex        =   2
         Top             =   510
         Width           =   7665
         _ExtentX        =   7011
         _ExtentY        =   1005
      End
      Begin VB.PictureBox picDate 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   135
         ScaleHeight     =   360
         ScaleWidth      =   2505
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   2505
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   300
            Left            =   510
            TabIndex        =   6
            Top             =   30
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   122880003
            CurrentDate     =   42285
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            Caption         =   "日期:"
            Height          =   180
            Left            =   45
            TabIndex        =   7
            Top             =   75
            Width           =   450
         End
         Begin VB.Image imgDefault 
            Height          =   255
            Left            =   600
            Top             =   840
            Width           =   255
         End
         Begin VB.Image imgbtn 
            Height          =   240
            Index           =   0
            Left            =   2250
            Picture         =   "frmCaseTendBodyOper.frx":6852
            Top             =   45
            Width           =   240
         End
         Begin VB.Image imgbtn 
            Height          =   240
            Index           =   1
            Left            =   1920
            Picture         =   "frmCaseTendBodyOper.frx":7254
            Stretch         =   -1  'True
            Top             =   45
            Width           =   255
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5445
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCaseTendBodyOper.frx":7C56
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10821
            Key             =   "ZLNOTE"
            Object.ToolTipText     =   "消息提示信息"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2
            MinWidth        =   2
            Text            =   "数据类型"
            TextSave        =   "数据类型"
            Key             =   "ZLDataType"
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
   Begin MSComctlLib.ImageList ilsDate 
      Left            =   9015
      Top             =   510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyOper.frx":84EA
            Key             =   "preGreen"
            Object.Tag             =   "preGreen"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyOper.frx":8EFC
            Key             =   "preGray"
            Object.Tag             =   "preGray"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyOper.frx":990E
            Key             =   "nextGreen"
            Object.Tag             =   "nextGreen"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyOper.frx":A320
            Key             =   "nextGray"
            Object.Tag             =   "nextGray"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyOper.frx":AD32
            Key             =   "preLight"
            Object.Tag             =   "preLight"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyOper.frx":B744
            Key             =   "nextLight"
            Object.Tag             =   "nextLight"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   15
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCaseTendBodyOper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type type_Patient
    lng病人ID As Long
    lng主页ID As Long
    lng文件ID As Long
    lng婴儿 As Long
    lng科室ID As Long
    lng护理等级 As Long
    lng病区ID As Long
    lng格式ID As Long
End Type
Private mT_Patient As type_Patient

Private Enum TYPE_Oper
    Col_OperNull = 0
    Col_OperTime = 1
    Col_OperType = 2
End Enum

Private mcbrToolBar As CommandBar
Private mblnChage  As Boolean
Private Const mFontSize As Integer = 9 '定义字体初始大小为9号字体
Private mstrTime As String
Private mstrDate As String
Private mstrBTime As String     '体温单开始时间
Private mstrETime As String     '体温单结束时间
Private mstrOverDate As String
Private mstrPreOutDate As String
Private mintPreDays As Integer
Private mlngHours As Long
Private mstrSQL As String
Private mintBigSize As Integer  '字体大小
Private mblnMove As Boolean     '是否转出
Private mblnFileBack As Boolean '是否归档
Private mbln出院 As Boolean
Private mblnOK As Boolean '刷新体温单作图


Private mrsOper As New ADODB.Recordset '手术


Public Function ShowEditor(ByVal frmParent As Object, ByVal strParam As String, ByVal strTime As String, ByVal strDayTime As String, _
    ByVal int心率应用 As Integer, Optional blnMove As Boolean = False, Optional ByVal bytSize As Byte = 0) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用体温单编辑窗体
    '参数:frmParent 父窗体,strParam 格式:病人ID;主页Id;文件ID;婴儿;科室ID;护理护理等级  strTime 某段时间的时间范围 例如:2011-01-25 00:00:00;2011-01-25 05:59:59
    
    '     strDayTime 一周开始时间; int心率应用=2 表示脉搏和心率公用 blnMove 历史数据是否转移
    '     bytSize 0-9号字体 1-12号字体
    '----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrParam() As String
    If strParam = "" Then Exit Function
    arrParam = Split(strParam, ";")
    If UBound(arrParam) < 3 Then Exit Function
    mT_Patient.lng科室ID = 0
    mT_Patient.lng护理等级 = 3
    mblnMove = False
    mblnOK = False
    mT_Patient.lng病人ID = Val(arrParam(0))
    mT_Patient.lng主页ID = Val(arrParam(1))
    mT_Patient.lng文件ID = Val(arrParam(2))
    mT_Patient.lng婴儿 = Val(arrParam(3))
    
    If UBound(arrParam) > 3 Then mT_Patient.lng科室ID = arrParam(4)
    If UBound(arrParam) > 4 Then mT_Patient.lng护理等级 = arrParam(5)
    
    If mT_Patient.lng病人ID = 0 And mT_Patient.lng主页ID = 0 And mT_Patient.lng科室ID = 0 Then
        MsgBox "文件ID,病人ID,主页ID不能为空,请检查!", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Not OpenPatientInfo Then Exit Function
    mstrDate = strDayTime
    mstrTime = strTime
    If Not ChekPatientOut(mT_Patient.lng文件ID, mT_Patient.lng病人ID, mT_Patient.lng主页ID, mT_Patient.lng婴儿) Then Exit Function
    mintBigSize = bytSize
    Me.Font.Size = IIf(mintBigSize = 0, 9, 12)
    mblnMove = blnMove
    
    '检查文件是否归档
    mblnFileBack = CheckFileBack(mT_Patient.lng文件ID, mblnMove)
    Call InitCommandBars
    '提取数据
    Call InitTabOper
    Call zlRefreshData
    Me.Show 1
    
    ShowEditor = mblnOK
End Function


Public Function ChekPatientOut(ByVal lng文件ID As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal intBaby As Long) As Boolean
    '-----------------------------------------------------------------------------------------------
    '功能:提取体温单开始时间和结束时间 并检查病人是否出院
    '-----------------------------------------------------------------------------------------------
    Dim strSQL As String, strNewSql As String
    Dim strBeginDate As String, strEndDate As String
    Dim rsTemp As New ADODB.Recordset
    Dim strMaxDate As String, strCurrDate As String
    Dim intDay As Integer
    mbln出院 = False
    On Error GoTo Errhand
    
    mintPreDays = Val(zlDatabase.GetPara("超期录入护理数据天数", glngSys, 1255, "1"))
    mlngHours = Val(Mid(Val(zlDatabase.GetPara("数据补录时限", glngSys)), 1, 6))
    If mintPreDays < 0 Then mintPreDays = 0
    
    '提取病人预出院时间
    strSQL = "Select 开始时间 From 病人变动记录 where 病人ID=[1] and 主页ID=[2] And 开始原因=10"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng病人ID, lng主页ID)
    If Not rsTemp.EOF Then mstrPreOutDate = Format(rsTemp!开始时间, "YYYY-MM-DD HH:mm:ss")
    
    '提取婴儿医嘱信息(转科，出院),存在医嘱以医嘱信息为准，否则以母亲出院日期为准
    strNewSql = "(SELECT " & vbNewLine & _
                "        病人ID, 主页ID, 婴儿时间, DECODE(NVL(婴儿, 0), 0, DECODE(NVL(出院日期, ''), '', 0, 1), DECODE(NVL(婴儿时间, ''), '', 0, 1)) 记录" & vbNewLine & _
                "       FROM (SELECT A.病人ID, A.主页ID, B.开始执行时间 婴儿时间, A.出院日期, B.婴儿" & vbNewLine & _
                "              FROM 病案主页 A," & vbNewLine & _
                "                   (SELECT B.病人ID, B.主页ID, B.婴儿, 开始执行时间" & vbNewLine & _
                "                     FROM 病人医嘱记录 B, 诊疗项目目录 C" & vbNewLine & _
                "                     WHERE B.诊疗项目ID + 0 = C.ID AND B.医嘱状态 = 8 AND NVL(B.婴儿, 0) <> 0 AND B.诊疗类别 = 'Z' " & vbNewLine & _
                "                      AND Instr(',3,5,6,11,', ',' || c.操作类型 || ',') > 0 AND B.病人ID = [2] AND B.主页ID = [3] AND B.婴儿(+) = [4]) B" & vbNewLine & _
                "              WHERE A.病人ID = [2] AND A.主页ID = [3] AND A.病人ID = B.病人ID(+) AND A.主页ID = B.主页ID(+)" & vbNewLine & _
                "              ORDER BY B.开始执行时间 DESC)" & vbNewLine & _
                "       WHERE ROWNUM < 2) E"

    '说明:目前有了专科体温单，病人可能同时存在多份体温单。体温单开始时间和终止时间的规则如下:
    '如果文件的开始时间不为空并且大于等于病人入院时间或婴儿出生时间,体温单的开始时间以文件开始时间为准,否则以病人入院时间或婴儿出生时间为准
    '如果文件的终止时间不为空并且小于等于病人或婴儿出院时间（未出院不能大于当前时间）,体温单结束时间以文件开始时间为准，否则体温单结束时间以病人或婴儿出院时间为准(未出院为当前时间)
    '如果文件的终止时间为空,保持原有方式,病人如果已经出院，就已出院时间为准,未出院就已当前时间或数据结束时间为准.
    strSQL = " SELECT  DECODE(D.开始时间,NULL,DECODE(B.出生时间, NULL, A.开始, B.出生时间)," & vbNewLine & _
            "               DECODE(SIGN(D.开始时间 - DECODE(B.出生时间, NULL, A.开始, B.出生时间))," & vbNewLine & _
            "                      1," & vbNewLine & _
            "                      D.开始时间," & vbNewLine & _
            "                      DECODE(B.出生时间, NULL, A.开始, B.出生时间))) AS 开始," & vbNewLine & _
            "       DECODE(D.结束时间," & vbNewLine & _
            "               NULL," & vbNewLine & _
            "               DECODE(E.记录," & vbNewLine & _
            "                      0," & vbNewLine & _
            "                      DECODE(SIGN(NVL(E.婴儿时间, A.终止) - D.发生时间), 1, NVL(E.婴儿时间, A.终止), D.发生时间)," & vbNewLine & _
            "                      NVL(E.婴儿时间, A.终止))," & vbNewLine & _
            "               DECODE(SIGN(NVL(E.婴儿时间, A.终止) - D.结束时间), 1, D.结束时间, NVL(E.婴儿时间, A.终止))) 终止," & vbNewLine & _
            "       DECODE(D.结束时间, NULL, E.记录, 1) 记录" & vbNewLine & _
            " FROM (SELECT 病人ID, 主页ID, MIN(开始时间) AS 开始, MAX(NVL(终止时间, SYSDATE)) AS 终止" & vbNewLine & _
            "       FROM 病人变动记录" & vbNewLine & _
            "       WHERE 开始时间 IS NOT NULL AND 病人ID = [2] AND 主页ID = [3]" & vbNewLine & _
            "       GROUP BY 病人ID, 主页ID) A," & vbNewLine & _
            "     (SELECT 病人ID, 主页ID, 出生时间 FROM 病人新生儿记录 WHERE 病人ID = [2] AND 主页ID = [3] AND 序号 = [4]) B," & vbNewLine & _
            "     (SELECT NVL(发生时间, SYSDATE) 发生时间, 开始时间, 结束时间" & vbNewLine & _
            "       FROM (SELECT MAX(B.发生时间) 发生时间, MAX(A.开始时间) 开始时间, MAX(A.结束时间) 结束时间" & vbNewLine & _
            "              FROM 病人护理文件 A, 病人护理数据 B" & vbNewLine & _
            "              WHERE A.ID = B.文件ID(+) AND A.ID = [1] AND A.病人ID = [2] AND A.主页ID = [3] AND A.婴儿 = [4])) D," & vbNewLine & _
            "  " & strNewSql & vbNewLine & _
            " WHERE A.病人ID = E.病人ID AND A.主页ID = E.主页ID AND A.病人ID = B.病人ID(+) AND A.主页ID = B.主页ID(+)"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng文件ID, lng病人ID, lng主页ID, intBaby)
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        strBeginDate = Format(rsTemp!开始, "YYYY-MM-DD HH:MM:SS")
        strEndDate = Format(rsTemp!终止, "YYYY-MM-DD HH:MM:SS")
        mbln出院 = Not (Val(rsTemp!记录) = 0)
    Else
        MsgBox "无此病人本次住院信息,请检查!", vbInformation, gstrSysName '无数病人变动信息退出
        Exit Function
    End If
    
    strCurrDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")

    mstrBTime = strBeginDate
    mstrOverDate = strEndDate
    mstrETime = strEndDate
    If CDate(mstrETime) < CDate(Format(strCurrDate, "YYYY-MM-DD HH:mm:ss")) And Not mbln出院 Then mstrETime = CDate(Format(strCurrDate, "YYYY-MM-DD HH:mm:ss"))
    If mstrBTime > mstrETime Then mstrBTime = mstrETime
    If mstrDate < mstrBTime Then mstrDate = mstrBTime
    
    '病人出院以出院时间为终止时间
    If mbln出院 = True Then
        '出院时间和入院时间如果在同一列，则将出院时间后移一列（内蒙需求:出院也要录入体温）
        mstrETime = Format(RetrunEndTimeNew(CDate(mstrBTime), CDate(mstrETime), gintHourBegin), "YYYY-MM-DD HH:mm:ss")
        strMaxDate = Format(mstrETime, "YYYY-MM-DD")
    Else
        intDay = mintPreDays - DateDiff("D", CDate(strCurrDate), CDate(mstrETime))
        If intDay < 0 Then intDay = 0
        strMaxDate = Format(DateAdd("d", intDay, CDate(mstrETime)), "yyyy-MM-dd")
        If CDate(mstrETime) < CDate(Format(DateAdd("d", mintPreDays, CDate(strCurrDate)), "yyyy-MM-dd HH:mm:ss")) Then
            mstrETime = Format(DateAdd("d", mintPreDays, CDate(strCurrDate)), "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    
    mstrETime = Format(strMaxDate & " " & Format(mstrETime, "HH:mm:ss"), "yyyy-MM-DD HH:mm:ss")
    
    dtpDate.Value = Format(mstrTime, "YYYY-MM-DD")
    dtpDate.MaxDate = Format(strMaxDate, "YYYY-MM-DD")
    dtpDate.MinDate = Format(mstrBTime, "YYYY-MM-DD")
    
    ChekPatientOut = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function OpenPatientInfo() As Boolean
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo Errhand
    '提取科室信息
    mstrSQL = "Select 出院科室ID from 病案主页 Where 病人id=[1] And 主页id=[2] "
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mT_Patient.lng病人ID, mT_Patient.lng主页ID)
    If rsTmp.BOF = False Then
        mT_Patient.lng科室ID = Val(zlCommFun.Nvl(rsTmp("出院科室ID").Value))
    End If
    
    '提取护理等级
    mstrSQL = "Select zl_PatitTendGrade([1],[2]) As 护理等级 From dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mT_Patient.lng病人ID, mT_Patient.lng主页ID)
    If rsTmp.BOF = False Then mT_Patient.lng护理等级 = zlCommFun.Nvl(rsTmp("护理等级"), 3)

    OpenPatientInfo = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function CheckFileBack(ByVal lngID As Long, ByVal blnMove As Boolean) As Boolean
'---------------------------------------------------------------
'功能:检查文件是否归档
'---------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo Errhand
    
    CheckFileBack = False
    strSQL = "Select 1 From 病人护理文件 Where Id=[1] And 归档人 Is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查文件是否归档", lngID)
    If blnMove = True Then
        strSQL = Replace(strSQL, "病人护理文件", "H病人护理文件")
    End If
    If rsTemp.RecordCount > 0 Then
        CheckFileBack = True
    End If
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub InitCommandBars()
'--------------------------------------------------------------------------------
'功能:初始化工具栏
'--------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrLable As CommandBarControl
    Dim cbrPop As CommandBarControl
    Dim cboChild As CommandBarPopup
    Dim CtlFont As stdFont
    
    On Error GoTo Errhand
    
     '初始设置
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsMain.ActiveMenuBar.Title = "菜单栏"
    cbsMain.ActiveMenuBar.Visible = False
    
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 24, 24
        .SetIconSize True, 16, 16
        .UseSharedImageList = False 'ImageList方式时,因同一App中共享,在AddImageList之前设置为False
        Set CtlFont = .Font
        If CtlFont Is Nothing Then
            Set CtlFont = Me.Font
        End If
        CtlFont.Size = IIf(mintBigSize = 0, 9, 12)
        Set .Font = CtlFont
    End With

  '------------------------------------------------------------------------------------------------------------------
    '工具栏定义
    Set mcbrToolBar = cbsMain.Add("标准", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    mcbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    
    With mcbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "取消")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    imgbtn(1).Picture = ilsDate.ListImages("preGreen").Picture
    imgbtn(0).Picture = ilsDate.ListImages("nextGreen").Picture
    '定位工具栏
    '------------------------------------------------------------------------------------------------------------------
    For Each cbrControl In mcbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With dtpDate
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Width = .Width + .Width * mintBigSize / 3
        .Height = 300 + 300 * mintBigSize / 3
    End With
    
    
    '快键绑定
    With cbsMain.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save '保存
        .Add FCONTROL, Asc("R"), conMenu_Edit_Reuse '取消
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function InitTabOper() As String
    '-------------------------------------------------------
    '功能:初始化手术分娩录入表格
    '-------------------------------------------------------
    Dim intRow As Integer, intCOl As Integer
    On Error Resume Next
    
    With vsfOper
        .Rows = 2
        .Cols = 0
        
        .NewColumn "", 255, 4
        .NewColumn "时间", 1000 + 1000 * mintBigSize / 3, 4, , 4
        .NewColumn "数据", 2000 + 2000 * mintBigSize / 3, 4, "手术|分娩|手术分娩|回室", 1
        .NewColumn "", 255, 4
        .ExtendLastCol = True
        .Body.RowHeightMin = 300 + 300 * mintBigSize / 3
        .FixedCols = 1
        .FixedRows = 1
        
        .Body.Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Body.WordWrap = False
        .Body.AllowUserResizing = flexResizeNone

        .Cell(flexcpAlignment, 0, .FixedCols, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpText, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = ""
    End With
End Function


Private Function zlRefreshData() As Boolean
    Dim strTime As String
    Dim rsTmp  As New ADODB.Recordset
    
    On Error GoTo Errhand
    '功能刷新数据
    gstrFields = "序号," & adDouble & ",18|项目序号," & adDouble & ",18|时间," & adLongVarChar & ",20|原始时间," & adLongVarChar & ",20|记录类型," & adDouble & ",1|内容," & _
            adLongVarChar & ",100|项目名称," & adLongVarChar & ",20|未记说明," & adLongVarChar & ",20|记录组号," & adDouble & ",1|数据来源," & adDouble & ",1|显示," & adDouble & ",1|" & _
             "来源ID," & adDouble & ",18|共用," & adDouble & ",1|状态," & adDouble & ",1"
    Call Record_Init(mrsOper, gstrFields)
    gstrFields = "序号|项目序号|时间|原始时间|记录类型|内容|项目名称|未记说明|记录组号|数据来源|显示|来源ID|共用|状态"

    '提取手术信息
    mstrSQL = "" & _
         " Select C.ID 序号, B.发生时间 AS 时间,C.记录类型,C.项目序号,C.未记说明,C.记录内容,C.记录组号,C.项目名称,C.数据来源,C.显示,C.来源ID,C.共用" & _
         " FROM 病人护理文件 A, 病人护理数据 B, 病人护理明细 C" & _
         " Where A.ID=B.文件ID and  B.ID = C.记录ID AND A.ID=[1]  AND Nvl(A.婴儿, 0)=[4] AND a.病人id=[2] AND a.主页id=[3] And c.终止版本 Is Null" & _
         " AND c.记录类型=4  AND B.发生时间 BETWEEN [5]  And [6]"

    If mblnMove Then
        mstrSQL = Replace(mstrSQL, "病人护理文件", "H病人护理文件")
        mstrSQL = Replace(mstrSQL, "病人护理数据", "H病人护理数据")
        mstrSQL = Replace(mstrSQL, "病人护理明细", "H病人护理明细")
    End If

    strTime = CDate(Format(mstrTime, "YYYY-MM-DD") & " 23:59:59")
    If CDate(strTime) > CDate(mstrETime) Then strTime = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")

    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, "读取手术、上下标等信息", mT_Patient.lng文件ID, mT_Patient.lng病人ID, mT_Patient.lng主页ID, _
        mT_Patient.lng婴儿, Int(CDate(Format(mstrTime, "YYYY-MM-DD"))), CDate(strTime))
    With rsTmp
        Do While Not .EOF
            gstrValues = zlCommFun.Nvl(!序号) & "|" & zlCommFun.Nvl(!项目序号, 0) & "|" & Format(zlCommFun.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "|" & Format(zlCommFun.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "|" & zlCommFun.Nvl(!记录类型) & "|" & _
                zlCommFun.Nvl(!记录内容) & "|" & zlCommFun.Nvl(!项目名称) & "|" & Nvl(!未记说明) & "|" & zlCommFun.Nvl(!记录组号, 0) & "|" & Val(zlCommFun.Nvl(!数据来源, 0)) & "|" & _
                Val(zlCommFun.Nvl(!显示, 0)) & "|" & Val(zlCommFun.Nvl(!来源ID, 0)) & "|" & Val(zlCommFun.Nvl(!共用, 0)) & "|0"
            Call Record_Add(mrsOper, gstrFields, gstrValues)
        .MoveNext
        Loop
    End With
    
    '添加手术信息
    mrsOper.Filter = 0
    mrsOper.Sort = "时间"
    With mrsOper
        vsfOper.Rows = vsfOper.FixedRows
        Do While Not .EOF
            vsfOper.Rows = vsfOper.Rows + 1
            vsfOper.TextMatrix(vsfOper.Rows - 1, Col_OperTime) = Format(!时间, "HH:mm")
            vsfOper.TextMatrix(vsfOper.Rows - 1, Col_OperType) = Nvl(!项目名称, "手术")
            If InStr(1, ",0,3,9,", "," & Val(zlCommFun.Nvl(!数据来源)) & ",") = 0 Then
                vsfOper.Cell(flexcpForeColor, vsfOper.Rows - 1, Col_OperTime, vsfOper.Rows - 1, Col_OperType) = 255
            Else
                vsfOper.Cell(flexcpForeColor, vsfOper.Rows - 1, Col_OperTime, vsfOper.Rows - 1, Col_OperType) = &H80000012
            End If
            vsfOper.RowData(vsfOper.Rows - 1) = Val(!序号)
        .MoveNext
        Loop
        vsfOper.Rows = vsfOper.Rows + 1
    End With
        vsfOper.Row = 1
        vsfOper.Col = 1
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function UpData(ByVal intRow As Integer, ByVal intCOl As Integer, _
    Optional blnComList As Boolean = False) As Boolean
    Dim strName As String
    Dim strTime As String
    Dim strValue As String
    Dim lngNo As String
    Dim lngID As Long

    
    On Error GoTo Errhand
    lngNo = 4
    If blnComList = True Then
        strName = vsfOper.EditText
        strTime = Format(vsfOper.TextMatrix(intRow, Col_OperTime), "HH:mm")
    Else
        strName = vsfOper.TextMatrix(intRow, Col_OperType)
        strTime = Format(vsfOper.EditText, "HH:mm")
        If Not IsDate(strTime) Then strTime = ""
    End If
    mrsOper.Filter = "记录类型=" & lngNo & " And 序号=" & Val(vsfOper.RowData(intRow))
    If mrsOper.RecordCount <> 0 Then
        If Val(mrsOper!状态) <> 1 And Val(mrsOper!状态) <> 3 Then 'his提取的数据
            mrsOper!状态 = 2
            If Trim(strTime) = "" Or strName = "" Then
                mrsOper!项目名称 = ""
                mrsOper!内容 = ""
            ElseIf Trim(strTime) <> "" And strName <> "" Then
                mrsOper!项目名称 = strName
                mrsOper!内容 = strName
            End If
            If Trim(strTime) <> "" Then mrsOper!时间 = SetDate(Format(Format(dtpDate.Value, "YYYY-MM-DD") & " " & Trim(strTime) & ":00", "YYYY-MM-DD HH:mm:ss"))
        Else
            If Trim(strTime) = "" Or strName = "" Then
                mrsOper!状态 = 3
                mrsOper!项目名称 = ""
                mrsOper!内容 = ""
            Else
                mrsOper!状态 = 1
                mrsOper!项目名称 = strName
                mrsOper!内容 = strName
            End If
            If Trim(strTime) <> "" Then mrsOper!时间 = SetDate(Format(Format(dtpDate.Value, "YYYY-MM-DD") & " " & Trim(strTime) & ":00", "YYYY-MM-DD HH:mm:ss"))
        End If
        mrsOper.Update
    Else
        If Trim(strTime) = "" Or strName = "" Then
            strValue = ""
        Else
            strValue = 1
            strTime = SetDate(Format(Format(dtpDate.Value, "YYYY-MM-DD") & " " & strTime & ":00", "YYYY-MM-DD HH:mm:ss"))
        End If
        
        If strValue <> "" Then
            strValue = strName
            lngID = GetMaxID(mrsOper)
            gstrFields = "序号|项目序号|时间|原始时间|记录类型|内容|项目名称|未记说明|记录组号|数据来源|显示|来源ID|共用|状态"
            gstrValues = lngID & "|" & 0 & "|" & strTime & "|" & strTime & "|" & lngNo & "|" & strValue & "|" & strName & "||0|0|0|0|0|1"
            vsfOper.RowData(intRow) = lngID
            Call Record_Add(mrsOper, gstrFields, gstrValues)
        End If
    End If
    
    If strName <> vsfOper.TextMatrix(intRow, Col_OperType) Or Format(strTime, "HH:mm") <> Format(vsfOper.TextMatrix(intRow, Col_OperTime), "HH:mm") Then
        mblnChage = True
    End If
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function SetDate(ByVal strTime As String) As String
'---------------------------------------------------------
' 检查日期
'---------------------------------------------------------
    Dim strVTime As String
    If Not IsDate(strTime) Then Exit Function
    strVTime = Format(strTime, "YYYY-MM-DD HH:mm:ss")
    If CDate(strTime) < CDate(mstrBTime) Then
        strVTime = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
    End If
    
    If CDate(strTime) > CDate(mstrETime) Then
        strVTime = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
    End If
    SetDate = strVTime
End Function


Private Function GetMaxID(ByVal rsTmp As ADODB.Recordset) As Long
'----------------------------------------------------
'功能:获取记录集中的最大序号
'----------------------------------------------------
    rsTmp.Filter = 0
    rsTmp.Sort = "序号 Desc"
    If rsTmp.RecordCount = 0 Then
        GetMaxID = 1
    Else
        GetMaxID = Val(rsTmp!序号) + 1
    End If
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case conMenu_Edit_Save '保存
            If Not SaveData Then Exit Sub
            Call zlRefreshData
        Case conMenu_Edit_Reuse '取消
            Call zlRefreshData
            mblnChage = False
        Case conMenu_Help_Help '帮助
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_File_Exit '退出
            Unload Me
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    On Error Resume Next
    picOper.Height = 5000 + 5000 * mintBigSize / 3
    Bottom = stbThis.Height
    
    With picDate
        .Left = 0
        .Top = 0
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Width = (lblDate.Width + dtpDate.Width + 520) + (lblDate.Width + dtpDate.Width + 520) * mintBigSize / 3
        .Height = 300 + 300 * mintBigSize / 3
    End With
    
    With lblDate
        .Left = 30
        .Top = 60
        .Height = picDate.Height
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    With dtpDate
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Width = .Width + .Width * mintBigSize / 3
        .Height = 300 + 300 * mintBigSize / 3
        .Top = 0
        .Left = lblDate.Left + lblDate.Width
    End With

    With imgbtn(1)
        .Width = 240 + 240 * mintBigSize / 3
        .Height = 240 + 240 * mintBigSize / 3
        .Top = 30
        .Left = lblDate.Width + dtpDate.Width + 20
    End With
    
    With imgbtn(0)
        .Width = 240 + 240 * mintBigSize / 3
        .Height = 240 + 240 * mintBigSize / 3
        .Top = 30
        .Left = lblDate.Width + dtpDate.Width + imgbtn(1).Width + 30
    End With
    
    With picStb
        .Top = stbThis.Top + 50
        .Left = stbThis.Panels(2).Left + 50
        .Height = stbThis.Height - 50
        .Width = stbThis.Panels(2).Width - 50
    End With
    
    With lblStb
        .Font.Size = 9 + 9 * mintBigSize / 3
        .Height = TextHeight("刘")
        .Top = (picStb.Height - .Height) \ 2
        .Left = 10
    End With


End Sub

Private Sub cbsMain_Resize()
    On Error Resume Next
    
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long  '客户区域的大小
    
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    With picOper
        .Top = lngTop
        .Left = lngLeft
        .Width = lngRight - lngLeft
        .Height = lngBottom - lngTop
    End With
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.Id
        Case conMenu_Edit_Save, conMenu_Edit_Reuse
             Control.Enabled = IIf(mblnChage = True, True, False)
    End Select
    
    If dtpDate.Value = dtpDate.MinDate Then
        imgbtn(1).Picture = ilsDate.ListImages("preGray").Picture
        imgbtn(1).Enabled = False
    End If
    If dtpDate.Value = dtpDate.MaxDate Then
        imgbtn(0).Picture = ilsDate.ListImages("nextGray").Picture
        imgbtn(0).Enabled = False
    End If
    
End Sub

Private Sub dtpDate_Change()
    Dim strDate As String
    If Not dtpDateChageDate(Format(dtpDate.Value, "YYYY-MM-DD")) Then Exit Sub
    imgbtn(1).Enabled = True
    imgbtn(0).Enabled = True
    If dtpDate.Value = dtpDate.MinDate Then
        imgbtn(1).Picture = ilsDate.ListImages("preGray").Picture
        imgbtn(1).Enabled = False
    Else
        imgbtn(1).Picture = ilsDate.ListImages("preGreen").Picture
    End If
    If dtpDate.Value = dtpDate.MaxDate Then
        imgbtn(0).Picture = ilsDate.ListImages("nextGray").Picture
        imgbtn(0).Enabled = False
    Else
        imgbtn(0).Picture = ilsDate.ListImages("nextGreen").Picture
    End If
End Sub


Private Function dtpDateChageDate(ByVal strValue As String) As Boolean
'------------------------------------------------------------------------------
'补录时间合法时，发生变化就刷新数据
'------------------------------------------------------------------------------
    Dim strErrMsg As String
    Dim strDate As String, strTime As String
    Dim i As Integer
    Dim strCurrDate As String
    Dim intBound As Integer
    Dim strBegin As String, strEnd As String
    Dim intCOl As Integer
    Dim strCurDate As String
    Dim intDay As Integer
    Dim strBTime As String
    On Error GoTo Errhand
    
    lblStb.Tag = lblStb.Caption
    
    If Format(strValue, "YYYY-MM-DD") > Format(mstrETime, "YYYY-MM-DD") Then
        If mbln出院 = False Then
            strErrMsg = "录入的日期已超出参数[超期录入天数：" & mintPreDays & "天]所指定的范围！"
        Else
            strErrMsg = "录入的日期不能大于[病人出院时间或文件结束时间：" & Format(mstrETime, "YYYY-MM-DD HH:mm:ss") & "]!"
        End If
        GoTo ErrInfo
    End If
    
    If Format(strValue, "YYYY-MM-DD") < Format(mstrBTime, "YYYY-MM-DD") Then
        strErrMsg = "录入的日期不能小于[体温单开始时间：" & Format(mstrBTime, "YYYY-MM-DD HH:mm:ss") & "]！"
        GoTo ErrInfo
    End If
    
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    strCurDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    If Format(strValue, "YYYY-MM-DD") = mstrETime Then
        strDate = Format(Format(mstrETime, "YYYY-MM-DD") & " 00:00:00", "YYYY-MM-DD HH:mm:ss")
        strTime = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
    ElseIf Format(strValue, "YYYY-MM-DD") = mstrBTime Then
        strDate = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
        strTime = strDate
    Else
        strDate = Format(Format(strValue, "YYYY-MM-DD") & " 00:00:00", "YYYY-MM-DD HH:mm:ss")
        strTime = Format(Format(strValue, "YYYY-MM-DD") & " 23:59:00", "YYYY-MM-DD HH:mm:ss")
    End If
    
    If Not IsAllowInput(mT_Patient.lng病人ID, mT_Patient.lng主页ID, mT_Patient.lng婴儿, strTime, strCurrDate) Then
        strErrMsg = "录入的时间[" & strValue & "]有误！[超过数据补录的有效时限:" & mlngHours & "小时]"
        GoTo ErrInfo
    End If
    
    mstrTime = Format(dtpDate.Value, "YYYY-MM-DD hh:mm:ss")
    If mblnChage Then
        mblnChage = False
        If MsgBox("数据已经发生改变,请问是否进行保存?", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
            If Not SaveData Then Exit Function
        End If
    End If
    Call zlRefreshData
    dtpDateChageDate = True
    Exit Function
ErrInfo:
    If strErrMsg <> "" Then
        lblStb.Caption = strErrMsg
        lblStb.ForeColor = 255
    End If
Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub dtpDate_CloseUp()
    vsfOper.SetFocus
End Sub

Private Sub dtpDate_Validate(Cancel As Boolean)
    If Not dtpDateChageDate(Format(dtpDate.Value, "YYYY-MM-DD")) Then
        Cancel = True
    End If
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChage = True Then
        If MsgBox("病人体温数据已经发生改变,请问是否需要保存？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Cancel = True
            Exit Sub
        End If
    End If

    mblnChage = False
    mblnMove = False
    mbln出院 = False
    
    If Not (mrsOper Is Nothing) Then Set mrsOper = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub imgbtn_Click(Index As Integer)
    Select Case Index
        Case 1
            dtpDate.Value = dtpDate.Value - 1
            Call dtpDate_Change
        Case 0
            dtpDate.Value = dtpDate.Value + 1
            Call dtpDate_Change
    End Select
    vsfOper.SetFocus
End Sub

Private Sub picOper_Paint()
     picOper.BackColor = &H8000000F
End Sub

Private Sub picOper_Resize()
    On Error Resume Next
    With vsfOper
        .Left = 5
        .Top = picDate.Top + picDate.Height + 20
        .Width = picOper.Width
        .Height = picOper.Height - .Top
        .Body.Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
End Sub

Private Function SaveData() As Boolean
    '--------------------------------------------------------
    '功能:进行数据修改保存
    '--------------------------------------------------------
    Dim lngItemCode As Long
    Dim strTime As String
    Dim strEnd As String
    Dim strMarkTime As String
    Dim strSQL As String
    Dim strValue As String
    Dim int检查科室 As Integer
    Dim int项目首次 As Integer
    Dim i As Integer
    Dim blnTran As Boolean
    Dim arrSQL() As String
    
    On Error GoTo Errhand
    Screen.MousePointer = 11
    
    ReDim Preserve arrSQL(1 To 1)
    With mrsOper
        .Filter = 0
        .Sort = "时间"
        '先删除掉修改的手术信息,一天可以设置多次手术，如果手术时间和体温数据时间相同，更新手术时间的话，会导致体温数据时间发生变化
        Do While Not .EOF
            If Val(!状态) <> 3 And Val(!状态) <> 0 Then
                lngItemCode = 4
                If Val(!状态) = 2 Then
                    strTime = Format(!原始时间, "YYYY-MM-DD HH:mm:ss")
                    strEnd = strTime
                    strMarkTime = strTime
                    int检查科室 = IIf(ISCheckDept(strMarkTime) = True, 1, 0)
                    strMarkTime = "To_Date('" & strMarkTime & "','yyyy-mm-dd hh24:mi:ss')"
                    
                    '更新数据信息
                    strSQL = "Zl_体温单数据_Update("
                    '文件id_In   In 病人护理文件.Id%Type,  --病人护理文件ID
                    strSQL = strSQL & Val(mT_Patient.lng文件ID) & ","
                    '发生时间_In In 病人护理数据.发生时间%Type, --护理数据的发生时间
                    strSQL = strSQL & strMarkTime & ","
                    '记录类型_In In 病人护理明细.记录类型%Type, --护理项目=1，上标说明=2，入出转标记=3，手术日标记=4,下标说明=6
                    strSQL = strSQL & lngItemCode & ","
                    '项目序号_In In 病人护理明细.项目序号%Type, --护理项目的序号，非护理项目固定为0
                    strSQL = strSQL & 0 & ","
                    '记录内容_In In 病人护理明细.记录内容%Type := Null, --记录内容，如果内容为空，即清除以前的内容  36或36/37
                    strSQL = strSQL & "NULL" & ","
                    '体温部位_In In 病人护理明细.体温部位%Type := Null, --删除数据时不用填写部位 除活动项目外
                    strSQL = strSQL & "NULL,"
                    '复试合格_In In Number := 0,
                    strSQL = strSQL & "NULL,"
                    '未记说明_In In 病人护理明细.未记说明%Type := Null, --未记说明
                    strSQL = strSQL & "NULL" & ","
                    '他人记录_In In Number := 1,
                    strSQL = strSQL & "1,"
                    '数据来源_In In 病人护理明细.数据来源%Type := 0,
                    strSQL = strSQL & Val(!数据来源) & ","
                    '来源id_In   In 病人护理明细.来源id%Type := Null,
                    strSQL = strSQL & IIf(Val(!来源ID) = 0, "NULL", !来源ID) & ","
                    '共用_In     In 病人护理明细.共用%Type := 0,
                    strSQL = strSQL & Val(!共用) & ","
                    '项目首次_In In Number := 0,--汇总项目使用，保存数据前是否先删除一段时间内的数据信息。 1 删除
                    strSQL = strSQL & 0 & ","
                    '开始时间_In In 病人护理数据.发生时间%Type := Null,
                    strSQL = strSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                    '结束时间_In In 病人护理数据.发生时间%Type := Null --本记录有效跨度的终止时间，单独记录为每分钟，体温表为4小时,时间跨度内的相同项目记录要删除
                    strSQL = strSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss')"
                    '  操作员_IN  IN 病人护理数据.保存人%TYPE := NULL,
                    '  检查科室_IN IN Number :=1
                    strSQL = strSQL & ",NULL," & int检查科室 & ")"
                    arrSQL(ReDimArray(arrSQL)) = strSQL
                End If
                
                strTime = Format(!时间, "YYYY-MM-DD HH:mm:ss")
                strEnd = strTime
                strMarkTime = strTime
                int检查科室 = IIf(ISCheckDept(strMarkTime) = True, 1, 0)
                strMarkTime = "To_Date('" & strMarkTime & "','yyyy-mm-dd hh24:mi:ss')"
                strValue = Trim(zlCommFun.Nvl(!内容))
                If strValue <> "" Then
                    '更新数据信息
                    strSQL = "Zl_体温单数据_Update("
                    '文件id_In   In 病人护理文件.Id%Type,  --病人护理文件ID
                    strSQL = strSQL & Val(mT_Patient.lng文件ID) & ","
                    '发生时间_In In 病人护理数据.发生时间%Type, --护理数据的发生时间
                    strSQL = strSQL & strMarkTime & ","
                    '记录类型_In In 病人护理明细.记录类型%Type, --护理项目=1，上标说明=2，入出转标记=3，手术日标记=4,下标说明=6
                    strSQL = strSQL & lngItemCode & ","
                    '项目序号_In In 病人护理明细.项目序号%Type, --护理项目的序号，非护理项目固定为0
                    strSQL = strSQL & 0 & ","
                    '记录内容_In In 病人护理明细.记录内容%Type := Null, --记录内容，如果内容为空，即清除以前的内容  36或36/37
                    strSQL = strSQL & "'" & strValue & "',"
                    '体温部位_In In 病人护理明细.体温部位%Type := Null, --删除数据时不用填写部位 除活动项目外
                    strSQL = strSQL & "NULL,"
                    '复试合格_In In Number := 0,
                    strSQL = strSQL & IIf(strValue = "回室", "1", "NULL") & ","
                    '未记说明_In In 病人护理明细.未记说明%Type := Null, --未记说明
                    strSQL = strSQL & IIf(lngItemCode <> 4, "'" & Nvl(!未记说明) & "'", "NULL") & ","
                    '他人记录_In In Number := 1,
                    strSQL = strSQL & "1,"
                    '数据来源_In In 病人护理明细.数据来源%Type := 0,
                    strSQL = strSQL & Val(!数据来源) & ","
                    '来源id_In   In 病人护理明细.来源id%Type := Null,
                    strSQL = strSQL & IIf(Val(!来源ID) = 0, "NULL", !来源ID) & ","
                    '共用_In     In 病人护理明细.共用%Type := 0,
                    strSQL = strSQL & Val(!共用) & ","
                    '项目首次_In In Number := 0,--汇总项目使用，保存数据前是否先删除一段时间内的数据信息。 1 删除
                    strSQL = strSQL & int项目首次 & ","
                    '开始时间_In In 病人护理数据.发生时间%Type := Null,
                    strSQL = strSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                    '结束时间_In In 病人护理数据.发生时间%Type := Null --本记录有效跨度的终止时间，单独记录为每分钟，体温表为4小时,时间跨度内的相同项目记录要删除
                    strSQL = strSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss')"
                    '  操作员_IN  IN 病人护理数据.保存人%TYPE := NULL,
                    '  检查科室_IN IN Number :=1
                    strSQL = strSQL & ",NULL," & int检查科室 & ")"
                    arrSQL(ReDimArray(arrSQL)) = strSQL
                End If
            End If
        .MoveNext
        Loop
    End With
    
    gcnOracle.BeginTrans
    blnTran = True
    '在执行数据变化
    For i = 1 To UBound(arrSQL)
        If arrSQL(i) <> "" Then Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "保存体温数据"):
'        Debug.Print CStr(arrSQL(i))
    Next
    gcnOracle.CommitTrans
    
    mblnChage = False
    mblnOK = True
    
    SaveData = True
    Screen.MousePointer = 0
    
    Exit Function
Errhand:
    If blnTran = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Screen.MousePointer = 0
    Call SaveErrLog
End Function


Private Function ISCheckDept(ByVal str发生时间 As String) As Boolean
'功能：是否在Zl_体温单数据_Update中进行科室检查
    'mstrOverDate<=mstrETime 并且病人已经出院，肯定是病人出院时间和入院时间在一列（程序处理后的结果）
    If mbln出院 = True And Format(mstrOverDate, "YYYY-MM-DD HH:mm:ss") < Format(mstrETime, "YYYY-MM-DD HH:mm:ss") Then
        If Format(str发生时间, "YYYY-MM-DD HH:mm:ss") > Format(mstrOverDate, "YYYY-MM-DD HH:mm:ss") And Format(str发生时间, "YYYY-MM-DD HH:mm:ss") <= Format(mstrETime, "YYYY-MM-DD HH:mm:ss") Then
            ISCheckDept = False
        Else
            ISCheckDept = True
        End If
    Else
        ISCheckDept = True
    End If
End Function

Private Sub vsfOper_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    lblStb.Caption = ""
    vsfOper.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
    vsfOper.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
End Sub

Private Sub vsfOper_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '检查是否是同步过来的数据
    Dim lngID As Long, intState As Integer
    lngID = Val(vsfOper.RowData(Row))
    If lngID > 0 Then
        mrsOper.Filter = "记录类型=4 And 序号=" & lngID
        intState = mrsOper!状态
        If InStr(1, ",0,3,9,", "," & Val(Nvl(mrsOper!数据来源, 0)) & ",") = 0 Then
            Cancel = True
            lblStb.Caption = "同步过来的数据,不允许进行数据删除."
            lblStb.ForeColor = 255
            vsfOper.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
        End If
        
        '完成数据的删除操作
        If intState = 0 Or intState = 2 Then '表示是原有数据
            mrsOper!内容 = ""
            mrsOper!项目名称 = ""
            mrsOper!状态 = 2
        Else '表示新增数据
            mrsOper.Delete
        End If
        mrsOper.Update
        mblnChage = True
    End If
End Sub

Private Sub vsfOper_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Dim intRow As Integer
    '如果上一列没有录入时间和手术信息 不能进行下一行
    If Row >= vsfOper.FixedRows And Col >= vsfOper.FixedCols Then
        If vsfOper.TextMatrix(Row, Col_OperTime) = "" Or (vsfOper.TextMatrix(Row, Col_OperType) = "" And vsfOper.EditText = "") Then Cancel = True
    End If
End Sub

Private Sub vsfOper_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    With vsfOper
        If .EditMode(NewCol) = 1 Then
            .Body.FocusRect = flexFocusSolid
        Else
            .Body.FocusRect = flexFocusLight
        End If
    End With
End Sub

Private Sub vsfOper_ComboCloseUp(Row As Long, Col As Long, FinishEdit As Boolean)
    If Trim(vsfOper.TextMatrix(Row, Col_OperTime)) <> "" Then
        Call UpData(Row, Col, True)
    End If
End Sub

Private Sub vsfOper_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mblnFileBack = True Then
        Cancel = True
        vsfOper.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
        lblStb.Caption = "病人体温数据已经归档,不允许进行数据修改."
        lblStb.ForeColor = 255
        vsfOper.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
    End If
    
    '检查是否是同步过来的数据
    If Val(vsfOper.RowData(Row)) > 0 Then
        mrsOper.Filter = "记录类型=4 And 序号=" & Val(vsfOper.RowData(Row))
        If InStr(1, ",0,3,9,", "," & Val(Nvl(mrsOper!数据来源, 0)) & ",") = 0 Then
            Cancel = True
            lblStb.Caption = "同步过来的数据,不允许进行数据修改."
            lblStb.ForeColor = 255
            vsfOper.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
        End If
    End If
End Sub

Private Sub vsfOper_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '进行数据合法性检查
    Dim strText As String
    Dim strInfo As String, strDate As String
    Dim rsTemp As New ADODB.Recordset
    
    If Row < vsfOper.FixedRows Then Exit Sub
    If vsfOper.EditText = vsfOper.TextMatrix(Row, Col) Then Exit Sub
    With vsfOper
        strText = .EditText
        If Col = Col_OperTime Then
            If Trim(strText) = "" Then
                .TextMatrix(Row, Col_OperType) = ""
                GoTo ErrEnd
            End If
            Select Case Len(strText)
            Case 3, 4
                strText = String(4 - Len(strText), "0") & strText
                strText = Mid(strText, 1, 2) & ":" & Mid(strText, 3)
            Case Is < 3
                strText = String(2 - Len(strText), "0") & strText
                strText = Format(Now, "HH") & ":" & strText
            End Select
            
            '合法性检查
            If Mid(strText, 3, 1) <> ":" Then
                strInfo = "录入的时点格式非法！[小时:分钟]"
                GoTo ErrInfo
            End If
            If Mid(strText, 1, 2) < 0 Or Mid(strText, 1, 2) > 23 Then
                strInfo = "录入的时点格式非法！[小时应在0至23之间]"
                GoTo ErrInfo
            End If
            If Mid(strText, 4, 2) < 0 Or Mid(strText, 4, 2) > 59 Then
                strInfo = "录入的时点格式非法！[分钟应在0至59之间]"
                GoTo ErrInfo
            End If
            .EditText = Format(strText, "HH:mm")
            
            '检查录入的时间是否已经存在了手术信息
            strDate = Format(dtpDate.Value & " " & strText, "YYYY-MM-DD HH:mm:ss")
            gstrSQL = "select 1 from 病人护理文件 A,病人护理数据 B,病人护理明细 C" & _
                " Where A.ID=B.文件ID And B.ID=C.记录ID And A.ID=[1] And B.发生时间=[2] And C.记录类型=4"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "是否存在手术", mT_Patient.lng文件ID, CDate(strDate))
            If rsTemp.RecordCount > 0 Then
                strInfo = "该时间已经存在手术信息，请检查！ 时间[" & strDate & "]"
                GoTo ErrInfo
            End If
            If Not CheckDateTime(Row, "时间", Format(dtpDate.Value & " " & strText, "YYYY-MM-DD HH:mm:ss")) Then
                Cancel = True
            End If
ErrEnd:
            If Cancel = False Then Call UpData(Row, Col, IIf(Col = Col_OperType, True, False))
        End If
    End With
    
    Exit Sub
ErrInfo:
    lblStb.Caption = strInfo
    lblStb.ForeColor = 255
    Cancel = True
End Sub


Private Function CheckDateTime(ByVal lngRow As Long, ByVal strName As String, ByVal strTime As String) As Boolean
'------------------------------------------------------------------
'功能:补录数据时检查数据设置范围
'------------------------------------------------------------------
    Dim strErrMsg As String
    Dim strDate As String
    Dim strCurrDate As String
    Dim strInfo As String
    
    On Error GoTo Errhand
    If lngRow <> 0 Then
        strInfo = "第" & lngRow & "行"
    ElseIf strName <> "" Then
        strInfo = strInfo & "[" & strName & "]"
    Else
        strInfo = ""
    End If
    
    If Format(strTime, "YYYY-MM-DD HH:mm") > Format(mstrETime, "YYYY-MM-DD HH:mm") Then
        If mbln出院 = False Then
            strErrMsg = strInfo & "记录数据时间已超出参数[超期录入天数：" & mintPreDays & "天]所指定的范围! "
        Else
            strErrMsg = strInfo & "记录数据时间不能大于[病人出院时间或文件结束时间：" & Format(mstrETime, "YYYY-MM-DD HH:mm:ss") & "]!"
        End If
        GoTo ErrInfo
    End If
    
    If Format(strTime, "YYYY-MM-DD HH:mm") < Format(mstrBTime, "YYYY-MM-DD HH:mm") Then
        strErrMsg = strInfo & "记录数据时间不能小于[体温单开始时间：" & Format(mstrBTime, "YYYY-MM-DD HH:mm:ss") & "]!"
        GoTo ErrInfo
    End If
    
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    If Not IsAllowInput(mT_Patient.lng病人ID, mT_Patient.lng主页ID, mT_Patient.lng婴儿, strTime, strCurrDate) Then
        strErrMsg = strInfo & "记录数据时间[" & strTime & "]有误![超过数据补录的有效时限:" & mlngHours & "小时]"
        GoTo ErrInfo
    End If
    
    CheckDateTime = True
    Exit Function
ErrInfo:
    If strErrMsg <> "" Then
        lblStb.Caption = strErrMsg
        lblStb.ForeColor = 255
    End If
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function IsAllowInput(ByVal lng病人ID As Long, ByVal lng主页ID As Long, lng婴儿 As Long, ByVal strTime As String, ByVal strCurTime As String) As Boolean
    '取出指定病人在指定时间之后关键点的时间
    Dim rsTemp As New ADODB.Recordset
    Dim strBabyOutTime As String
    On Error GoTo Errhand
    
    IsAllowInput = True
    If lng婴儿 <> 0 And mbln出院 = True Then
        strBabyOutTime = GetAdviceOutTime(lng病人ID, lng主页ID, lng婴儿)
        If strBabyOutTime <> "" Then
            strTime = Format(DateAdd("H", mlngHours, strBabyOutTime), "yyyy-MM-dd HH:mm")
            GoTo GONext
        End If
    End If
    gstrSQL = "" & _
              " SELECT DECODE(终止原因,1,'出院',3,'转科',10,'预出院',15,'转病区',DECODE(开始原因,10,'出院','未定义')) AS 类型,终止时间 AS 时间" & _
              " From 病人变动记录" & _
              " WHERE (终止原因 IN (1,3,10,15) OR 开始原因=10) And 病人ID=[1] And 主页ID=[2] And [3] <= 终止时间" & _
              " ORDER BY 终止时间"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取出指定病人在指定时间之后关键点的时间", lng病人ID, lng主页ID, CDate(strTime))
    If rsTemp.RecordCount = 0 Then Exit Function
    '只取第一条符合的记录
    strTime = Format(DateAdd("H", mlngHours, rsTemp!时间), "yyyy-MM-dd HH:mm")
GONext:
    strCurTime = Format(strCurTime, "yyyy-MM-dd HH:mm")
    
    If strTime < strCurTime Then IsAllowInput = False
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
