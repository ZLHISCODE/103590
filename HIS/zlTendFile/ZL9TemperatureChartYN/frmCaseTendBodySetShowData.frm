VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmCaseTendBodySetShowData 
   Caption         =   "体温数据显示"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10350
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCaseTendBodySetShowData.frx":0000
   LockControls    =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   10350
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
      Left            =   1440
      ScaleHeight     =   360
      ScaleWidth      =   2415
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4800
      Width           =   2415
      Begin VB.Label lblStb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   45
      End
   End
   Begin VB.PictureBox picThis 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
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
      Height          =   3015
      Left            =   240
      ScaleHeight     =   3015
      ScaleWidth      =   4935
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   4935
      Begin VB.PictureBox picShow 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   0
         ScaleHeight     =   1215
         ScaleWidth      =   4335
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1560
         Width           =   4335
         Begin VSFlex8Ctl.VSFlexGrid vfgShow 
            Height          =   615
            Left            =   0
            TabIndex        =   4
            Top             =   360
            Width           =   3975
            _cx             =   7011
            _cy             =   1085
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
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
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   0
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   2
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   1
            RowHeightMin    =   0
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
         Begin VB.Label lblTmp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   90
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgThis 
         Height          =   1095
         Left            =   0
         TabIndex        =   2
         Top             =   360
         Width           =   3735
         _cx             =   6588
         _cy             =   1931
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
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
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   2
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   7
         FixedRows       =   2
         FixedCols       =   2
         RowHeightMin    =   255
         RowHeightMax    =   5000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
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
         OwnerDraw       =   1
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
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "时间:2011-02-25"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1350
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6180
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCaseTendBodySetShowData.frx":08CA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15346
            Object.ToolTipText     =   "打印机信息"
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCaseTendBodySetShowData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnTemType As Boolean 'TRUE 专科体温单,FALSE 标准体温单
Private mstrCurveItem As String  '专科体温单的体温曲线项目
'病人护理信息
Private Type type_Patient
    lng病人ID As Long
    lng主页ID As Long
    lng文件ID As Long
    lng婴儿 As Long
    lng科室ID As Long
    lng护理等级 As Long
    lng格式ID As Long
End Type
Private mT_Patient As type_Patient

'工具栏:
Private mcbrToolBar As CommandBar
Private mrsPoint As New ADODB.Recordset
Private mrs部位 As New ADODB.Recordset
Private mrsCopy As New ADODB.Recordset '用于还原数据信息

Private Const mFontSize As Integer = 9 '定义字体初始大小为9号字体
Private mintBigSize As Integer
Private mstrActiveItem As String
Private mint心率应用 As Integer
Private marrTime() As String
Private mDTime As Date
Private mDEndTime As Date
Private mblnChage As Boolean
Private mblnOK As Boolean
Private mblnMove As Boolean
Private mstrSQL As String
Private mblnInit As Boolean
Private mintColSel As Integer
Private mblnFileBack As Boolean
Private mbln出院 As Boolean '病人出院或文件结束为TRUE
Private mbln脉搏共用显示 As Boolean

Public Function ShowEdit(ByVal frmParent As Object, ByVal strParam As String, ByVal DTime As Date, ByVal DEndTime As Date, _
    ByVal int心率应用 As Integer, Optional blnMove As Boolean = False, Optional ByVal bytSize As Byte = 0) As Boolean
'----------------------------------------------------------------------------------------------------------
'功能:调用体温单编辑窗体
'参数:frmParent 父窗体,strParam 格式:病人ID;主页Id;文件ID;婴儿;科室ID;护理护理等级
'     Dtime 要编辑体温单的时间 格式为 YYYY-MM-DD HH:mm:ss:DEndTime 体温单结束时间 ; int心率应用=2 表示脉搏和心率公用 blnMove 历史数据是否转移
'bytSize 0-9号字体 1-12号字体
'----------------------------------------------------------------------------------------------------------
    Dim arrParam() As String
    If strParam = "" Then Exit Function
    arrParam = Split(strParam, ";")
    If UBound(arrParam) < 3 Then Exit Function
    
    mblnChage = False
    mblnMove = False
    mblnInit = False
    mblnOK = False
    mblnFileBack = False
    mT_Patient.lng科室ID = 0
    mT_Patient.lng护理等级 = 3
    
    mT_Patient.lng病人ID = arrParam(0)
    mT_Patient.lng主页ID = arrParam(1)
    mT_Patient.lng文件ID = arrParam(2)
    mT_Patient.lng婴儿 = arrParam(3)
    If UBound(arrParam) > 3 Then mT_Patient.lng科室ID = arrParam(4)
    If UBound(arrParam) > 4 Then mT_Patient.lng护理等级 = arrParam(5)
    
    If mT_Patient.lng病人ID = 0 And mT_Patient.lng主页ID = 0 And mT_Patient.lng科室ID = 0 Then
        MsgBox "文件ID,病人ID,主页ID不能为空,请检查!", vbInformation, gstrSysName
        Exit Function
    End If
    
    mDTime = DTime
    mDEndTime = DEndTime
    mbln脉搏共用显示 = (Val(zlDatabase.GetPara("脉搏短绌以(心率/脉搏)方式录入", glngSys, 1255, 0)) = 1)
    mint心率应用 = int心率应用
    mblnMove = blnMove
    
    If Not OpenPatientInfo Then Exit Function
    If Not ChekPatientOut(mT_Patient.lng文件ID, mT_Patient.lng病人ID, mT_Patient.lng主页ID, mT_Patient.lng婴儿) Then Exit Function
    
    mintBigSize = bytSize   'zldatabase.GetPara("护理文件显示模式", glngSys, 1255, 0)
    Me.Font.Size = IIf(mintBigSize = 0, 9, 12)
    '检查文件是否归档
    mblnFileBack = CheckFileBack(mT_Patient.lng文件ID, mblnMove)
    If mblnFileBack = True Then lblStb.Caption = "病人体温数据已经归档,不允许进行数据修改.": lblStb.ForeColor = 255

    Call InitCommandBars
    Call GetTableRowName
    
    mblnInit = True
    Me.Show 1
    ShowEdit = mblnOK
End Function

Public Function ChekPatientOut(ByVal lng文件ID As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal intBaby As Long) As Boolean
'-----------------------------------------------------------------------------------------------
'功能:提取体温单开始时间和结束时间 并检查病人是否出院
'-----------------------------------------------------------------------------------------------
    Dim strSQL As String, strNewSql As String
    Dim strBeginDate As String, strEndDate As String
    Dim rsTemp As New ADODB.Recordset
    Dim strCurrDate As String
    
    mbln出院 = False
    On Error GoTo Errhand
        
    '提取婴儿医嘱信息(转科，出院),存在医嘱以医嘱信息为准，否则以母亲出院日期为准
    strNewSql = "(SELECT " & vbNewLine & _
                "        病人ID, 主页ID, 婴儿时间, DECODE(NVL(婴儿, 0), 0, DECODE(NVL(出院日期, ''), '', 0, 1), DECODE(NVL(婴儿时间, ''), '', 0, 1)) 记录" & vbNewLine & _
                "       FROM (SELECT A.病人ID, A.主页ID, B.开始执行时间 婴儿时间, A.出院日期, B.婴儿" & vbNewLine & _
                "              FROM 病案主页 A," & vbNewLine & _
                "                   (SELECT B.病人ID, B.主页ID, B.婴儿, 开始执行时间" & vbNewLine & _
                "                     FROM 病人医嘱记录 B, 诊疗项目目录 C" & vbNewLine & _
                "                     WHERE B.诊疗项目ID + 0 = C.ID AND B.医嘱状态 = 8 AND NVL(B.婴儿, 0) <> 0 AND C.类别 = 'Z' AND EXISTS" & vbNewLine & _
                "                      (SELECT 1" & vbNewLine & _
                "                            FROM TABLE(CAST(F_STR2LIST('3,5,11') AS ZLTOOLS.T_STRLIST))" & vbNewLine & _
                "                            WHERE C.操作类型 = COLUMN_VALUE) AND B.病人ID = [2] AND B.主页ID = [3] AND B.婴儿(+) = [4]) B" & vbNewLine & _
                "              WHERE A.病人ID = [2] AND A.主页ID = [3] AND A.病人ID = B.病人ID(+) AND A.主页ID = B.主页ID(+)" & vbNewLine & _
                "              ORDER BY B.开始执行时间 DESC)" & vbNewLine & _
                "       WHERE ROWNUM < 2) E"

    '说明:目前有了专科体温单，病人可能同时存在多份体温单。体温单开始时间和终止时间的规则如下:
    '如果文件的开始时间不为空并且大于等于病人入院时间或婴儿出生时间,体温单的开始时间以文件开始时间为准,否则以病人入院时间或婴儿出生时间为准
    '如果文件的终止时间不为空并且小于等于病人或婴儿出院时间（未出院不能不能大于当前时间）,体温单结束时间以文件开始时间为准，否则体温单结束时间以病人或婴儿出院时间为准(未出院为当前时间)
    '如果文件的终止时间为空,保持原有方式,病人如果已经出院，就已出院时间为准,未出院就已当前时间或数据结束时间为准.
    strSQL = " SELECT /*+ RULE */ DECODE(D.开始时间,NULL,DECODE(B.出生时间, NULL, A.开始, B.出生时间)," & vbNewLine & _
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
        MsgBox "无此病人本次住院信息,请检查!", vbInformation, gstrSysName
        Exit Function '无数病人变动信息退出
    End If
    
    strCurrDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
    mDEndTime = strEndDate
    If CDate(mDEndTime) < CDate(Format(strCurrDate, "YYYY-MM-DD HH:mm:ss")) And Not mbln出院 Then mDEndTime = CDate(Format(strCurrDate, "YYYY-MM-DD HH:mm:ss"))
    
    '病人出院已出院时间为终止时间
    If mbln出院 = True Then
        '出院时间和入院时间如果在同一列，则将出院时间后移一列（内蒙需求:出院也要录入体温）
        If mblnTemType = False Then '标准体温单
            mDEndTime = Format(RetrunEndTime(CDate(strBeginDate), CDate(mDEndTime), gintHourBegin), "YYYY-MM-DD HH:mm:ss")
        Else '专科体温单
            mDEndTime = Format(RetrunEndTimeNew(CDate(strBeginDate), CDate(mDEndTime), gintHourBegin), "YYYY-MM-DD HH:mm:ss")
        End If
    End If
        
    If Not (CDate(mDTime) >= CDate(strBeginDate) And CDate(mDTime) <= CDate(mDEndTime)) Then
        If Int(CDate(strBeginDate)) = Int(CDate(mDEndTime)) Then
            mDTime = Format(strBeginDate, "YYYY-MM-DD HH:mm:ss")
        Else
            mDTime = Format(Int(CDate(mDEndTime)), "YYYY-MM-DD HH:mm:ss")
        End If
    End If
    
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
    
    '提取体温单信息
    mblnTemType = False
    mstrSQL = "Select B.子类,B.ID From 病人护理文件 A,病历文件列表 B Where A.格式ID=B.ID And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mT_Patient.lng文件ID)
    If rsTmp.BOF = False Then
        mblnTemType = (Nvl(rsTmp!子类) = "1")
        mT_Patient.lng格式ID = rsTmp!Id
    End If
    
    If mblnTemType = True Then
        gintHourBegin = T_BodyStyle.lng开始时点
    Else
        gintHourBegin = zlDatabase.GetPara("体温开始时间", glngSys, 1255, 4)
        T_BodyStyle.lng开始时点 = gintHourBegin
        T_BodyStyle.lng时间间隔 = 4
        T_BodyStyle.lng监测次数 = 6
        T_BodyStyle.lng天数 = 7
    End If
    
    OpenPatientInfo = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitCommandBars()
'--------------------------------------------------------------------------------
'功能:初始化工具栏
'--------------------------------------------------------------------------------
    Dim cbrControl As CommandBarButton
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrLable As CommandBarControl
    Dim CtlFont As StdFont
    
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Curve, "曲线编辑"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CurveTable, "表格编辑")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    
    '设置工具栏文本和图表显示方式
    For Each cbrControl In mcbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next

    '快键绑定
    With cbsMain.KeyBindings
        .Add FCONTROL, Asc("Q"), conMenu_Edit_Curve
        .Add FCONTROL, Asc("T"), conMenu_Edit_CurveTable
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("R"), conMenu_Edit_Reuse
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function GetTableRowName() As String
    Dim rsTemp As New ADODB.Recordset
    Dim strTmpName0 As String
    Dim strTmpCurve As String
    Dim arrItem() As Variant, i As Integer
    
    On Error GoTo Errhand
    
    '提取所有体温项目
    mstrCurveItem = ""
    If mblnTemType = False Then
        mstrSQL = _
                " Select A.记录法,A.记录名 as 项目名称,A.项目序号 as 项目号,A.单位" & _
                " From 体温记录项目 A,诊治所见项目 B,护理记录项目 C " & _
                " Where c.项目ID=B.ID(+) And A.项目序号=C.项目序号 And 项目性质=1 And (nvl(A.记录法,1)<>2 Or (nvl(A.记录法,1)=2 and A.项目序号=3)) And Nvl(C.应用方式,0)=1 AND C.护理等级>=[1] And Nvl(C.适用病人,0) In (0,[3]) " & _
                " And (c.适用科室=1 Or (c.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=c.项目序号 And D.科室id=[2]))) " & _
                " Order by Decode(A.项目序号,1,0,1),A.排列序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mT_Patient.lng护理等级, mT_Patient.lng科室ID, IIf(mT_Patient.lng婴儿 = 0, 1, 2))
    Else '专科体温单
        mstrCurveItem = T_BodyItem.str曲线项目
        If InStr(1, "," & mstrCurveItem & ",", "," & gint呼吸 & ",") = 0 Then
            arrItem = Array(T_BodyItem.str表格内容)
            For i = 0 To UBound(arrItem)
                If Val(arrItem(i)) = gint呼吸 Then
                    mstrCurveItem = mstrCurveItem & "," & gint呼吸
                    Exit For
                End If
            Next
        End If
        If Left(mstrCurveItem, 1) = "," Then mstrCurveItem = Mid(mstrCurveItem, 2)
        mstrSQL = _
                " Select /*+ RULE */ A.记录法,A.记录名 as 项目名称,A.项目序号 as 项目号,A.单位" & _
                " From 体温记录项目 A,诊治所见项目 B,护理记录项目 C,Table(Cast(f_num2list([1]) As zlTools.t_Numlist)) D " & _
                " Where C.项目ID=B.ID(+) And A.项目序号=C.项目序号 And (A.记录法<>2 OR (A.记录法=2 And A.项目序号=3)) And NOT (C.应用方式=2 And C.项目序号=-1)" & _
                " And C.项目序号=D.COLUMN_VALUE Order by Decode(A.项目序号,1,0,1),A.排列序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mstrCurveItem)
    End If
    strTmpName0 = ""
    With rsTemp
        Do While Not .EOF
            strTmpName0 = strTmpName0 & ";" & zlCommFun.Nvl(!项目号) & "'" & zlCommFun.Nvl(!项目名称) & IIf(zlCommFun.Nvl(!单位) = "", "", "(" & zlCommFun.Nvl(!单位) & ")") & "'" & zlCommFun.Nvl(!项目名称)
        .MoveNext
        Loop
    End With
    
    If Left(strTmpName0, 1) = ";" Then strTmpName0 = Mid(strTmpName0, 2)
    
    Call InitTable(strTmpName0)
    '刷新数据
    Call zlRefreshData
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitTable(ByVal strTmpName As String)
    Dim intCOl As Integer, intRow As Integer
    Dim strColName As String
    Dim arrColName() As String, arrColTime() As String
    
    strColName = InitTime
    arrColName = Split(strColName, "[LPF]")
    
    On Error GoTo Errhand
    
    With vfgThis
        .Clear
        .FixedCols = 3
        .FixedRows = 2
        .Rows = 3
        .Cols = .FixedCols + T_BodyStyle.lng监测次数
        .ColHidden(0) = True
        .ColWidth(0) = 0
        .ColHidden(1) = True
        .ColWidth(1) = 0
        .MergeCellsFixed = flexMergeRestrictColumns
        .MergeCol(2) = True
        .MergeRow(0) = True
        .MergeRow(1) = True
        .Col = .FixedCols: .Row = .FixedRows
        .ColSel = .Col
        .RowSel = .Row
    
        vfgThis.Font.Size = mFontSize + mFontSize * mintBigSize / 3
       
        For intRow = 0 To .FixedRows - 1
            arrColTime = Split(arrColName(intRow), ";")
            For intCOl = .FixedCols - 1 To .Cols - 1
                .TextMatrix(intRow, intCOl) = arrColTime(intCOl + 1 - .FixedCols)
            Next intCOl
            If intRow = 0 Then
                .RowHeight(intRow) = 300 + 300 * mintBigSize / 3
            Else
                .RowHeight(intRow) = 400 + 400 * mintBigSize / 3
            End If
        Next intRow
        
        '设置列宽
        For intCOl = .FixedCols - 1 To .Cols - 1
            If intCOl = .FixedCols - 1 Then
                .ColWidth(intCOl) = 1300 + 1300 * mintBigSize / 3
            Else
                .ColWidth(intCOl) = 1200 + 1200 * mintBigSize / 3
            End If
        Next intCOl
        '固定表头格式居中
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpForeColor, 0, 0, .FixedRows - 1, .Cols - 1) = RGB(0, 0, 255)
        .Cell(flexcpBackColor, 0, 0, .FixedRows - 1, .Cols - 1) = &H8000000F
        
        '加载列的头部信息
        arrColName = Split(strTmpName, ";")
        .Rows = UBound(arrColName) + .FixedRows + 1
        For intRow = .FixedRows To .Rows - 1
            arrColName(intRow - .FixedRows) = arrColName(intRow - .FixedRows) & String(3 - UBound(Split(arrColName(intRow - .FixedRows), "'")), "'")
            .RowHeight(intRow) = 300 + 300 * mintBigSize / 3
            .TextMatrix(intRow, 0) = Split(arrColName(intRow - .FixedRows), "'")(0)
            .TextMatrix(intRow, 1) = Split(arrColName(intRow - .FixedRows), "'")(2)
            .TextMatrix(intRow, 2) = Split(arrColName(intRow - .FixedRows), "'")(1)
        Next intRow
        .Cell(flexcpBackColor, .FixedRows, .FixedCols - 1, .Rows - 1, .FixedCols - 1) = &H8000000F
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .ScrollBars = flexScrollBarBoth
    End With
    
    vfgThis.Cell(flexcpText, vfgThis.FixedRows, vfgThis.FixedCols, vfgThis.Rows - 1, vfgThis.Cols - 1) = ""
    
    With vfgShow
        .RowHeight(-1) = 300 + 300 * mintBigSize / 3
        .ColWidth(-1) = 1200 + 1200 * mintBigSize / 3
        .FixedRows = 0
        .FixedCols = 1
        .Rows = 2
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = &H0&
        .ScrollBars = flexScrollBarBoth
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function zlRefreshData() As Boolean
'---------------------------------------------------------------
'功能:提取病人某天内的所有数据信息
'---------------------------------------------------------------
    '序号 为病人护理明细的ID    ID为物理降温或脉搏短轴时心率的数据 ,标注记录信息数据库中是否为显示
    gstrFields = "序号," & adDouble & ",18|数值," & adLongVarChar & ",400|部位," & adLongVarChar & ",200|" & _
         "标记," & adDouble & ",1|时间," & adLongVarChar & ",20|项目序号," & adDouble & ",18|项目名称," & adLongVarChar & ",40|" & _
         "复查," & adDouble & ",1|数据来源," & adDouble & ",1|显示," & adDouble & ",1|标注," & adDouble & ",1|状态," & adDouble & ",1|时间段," & adLongVarChar & ",20|列号," & _
         adDouble & ",1|ID," & adDouble & ",18"
    Call Record_Init(mrsPoint, gstrFields)
    gstrFields = "序号|数值|部位|标记|时间|项目序号|项目名称|复查|数据来源|显示|标注|状态|时间段|列号|ID"
    
    Dim rsTmp As New ADODB.Recordset
    Dim strFidlds As String, strParam As String, strPart As String
    Dim arrValue() As String
    Dim lng项目序号 As Long, lngCol As Long
    Dim str项目名称 As String
    Dim int显示 As Integer, int标注 As Integer
    Dim blnAdd As Boolean, blnAllow As Boolean
    Dim intRow As Integer, intCOl As Integer
    Dim strTime As String
    Dim int标记 As Integer
    Dim strEndTime As String
    
    On Error GoTo Errhand
    
    lblTime.Caption = "时间:" & Format(mDTime, "YYYY-MM-DD")
    
    '提取部位
    mstrSQL = "Select 项目序号,部位,缺省项 From 体温部位"
    Call zlDatabase.OpenRecordset(mrs部位, mstrSQL, Me.Caption)
    
    If CDate(Format(mDTime, "YYYY-MM-DD")) = CDate(Format(mDEndTime, "YYYY-MM-DD")) Then
        strEndTime = Format(CDate(mDEndTime), "YYYY-MM-DD HH:mm:ss")
    Else
        strEndTime = Format((Format(mDTime, "YYYY-MM-DD") & " 23:59:59"), "YYYY-MM-DD HH:mm:ss")
    End If
    
    '提取某时间段的所有体温曲线数据
    If mblnTemType = False Then
        mstrSQL = _
        " SELECT C.ID 序号,A.发生时间 As 时间,C.显示,c.记录内容 As 数值,c.体温部位,c.复试合格,D.记录名,E.保留项目,D.项目序号,DECODE(D.项目序号,-1,1,C.记录标记) 记录标记,C.未记说明,C.数据来源" & vbNewLine & _
        "                    FROM 病人护理文件 B,病人护理数据 A,病人护理明细 C,体温记录项目 D,护理记录项目 E" & vbNewLine & _
        "                    Where B.ID=A.文件ID" & vbNewLine & _
        "                        AND A.ID = C.记录ID" & vbNewLine & _
        "                        AND B.ID=[1]" & vbNewLine & _
        "                        AND Nvl(B.婴儿,0)=[4]" & vbNewLine & _
        "                        AND B.病人id=[2]" & vbNewLine & _
        "                        AND B.主页id=[3]" & vbNewLine & _
        "                        AND D.项目序号=C.项目序号" & vbNewLine & _
        "                        AND C.记录类型=1" & vbNewLine & _
        "                        AND E.项目序号=D.项目序号" & vbNewLine & _
        "                        AND E.护理等级>=[7]" & vbNewLine & _
        "                        AND (nvl(D.记录法,1)<>2 Or (nvl(D.记录法,1)=2 and D.项目序号=3))" & _
        "                        And A.发生时间 BETWEEN [5] And [6] And C.终止版本 Is Null" & vbNewLine & _
        "                        AND (nvl(E.应用方式,0)=1 OR ( -1=[10] and nvl(E.应用方式,0)=2))" & vbNewLine & _
        "                        AND nvl(E.适用病人,0) in (0,[8]) AND (E.适用科室=1 or ( E.适用科室=2 AND Exists (select 1 from 护理适用科室 D where D.项目序号=E.项目序号 and D.科室ID=[9])))" & vbNewLine & _
        "                    Order By A.发生时间,DECODE(D.项目序号,-1,1,0),DECODE(D.项目序号,-1,1,C.记录标记),D.记录法"
    
        If mblnMove Then
            mstrSQL = Replace(mstrSQL, "病人护理文件", "H病人护理文件")
            mstrSQL = Replace(mstrSQL, "病人护理数据", "H病人护理数据")
            mstrSQL = Replace(mstrSQL, "病人护理明细", "H病人护理明细")
        End If
            
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mT_Patient.lng文件ID, mT_Patient.lng病人ID, mT_Patient.lng主页ID, mT_Patient.lng婴儿, _
             CDate(mDTime), CDate(strEndTime), mT_Patient.lng护理等级, IIf(mT_Patient.lng婴儿 = 0, 1, 2), mT_Patient.lng科室ID, IIf(mint心率应用 = 2, -1, 0))
    Else '专科体温单
        mstrSQL = _
        " SELECT /*+ RULE */ C.ID 序号,A.发生时间 As 时间,C.显示,c.记录内容 As 数值,c.体温部位,c.复试合格,D.记录名,E.保留项目,D.项目序号,DECODE(D.项目序号,-1,1,C.记录标记) 记录标记,C.未记说明,C.数据来源" & vbNewLine & _
        "                    FROM 病人护理文件 B,病人护理数据 A,病人护理明细 C,体温记录项目 D,护理记录项目 E,Table(Cast(f_num2list([7]) As zlTools.t_Numlist)) F" & vbNewLine & _
        "                    Where B.ID=A.文件ID" & vbNewLine & _
        "                        AND A.ID = C.记录ID" & vbNewLine & _
        "                        AND B.ID=[1]" & vbNewLine & _
        "                        AND Nvl(B.婴儿,0)=[4]" & vbNewLine & _
        "                        AND B.病人id=[2]" & vbNewLine & _
        "                        AND B.主页id=[3]" & vbNewLine & _
        "                        AND D.项目序号=C.项目序号" & vbNewLine & _
        "                        AND C.记录类型=1" & vbNewLine & _
        "                        AND E.项目序号=D.项目序号" & vbNewLine & _
        "                        AND E.项目序号=F.COLUMN_VALUE" & vbNewLine & _
        "                        AND (NVL(D.记录法,1)<>2 Or (NVL(D.记录法,1)=2 and D.项目序号=3))" & _
        "                        AND A.发生时间 BETWEEN [5] And [6] And C.终止版本 Is Null" & vbNewLine & _
        "                    Order By A.发生时间,DECODE(D.项目序号,-1,1,0),DECODE(D.项目序号,-1,1,C.记录标记),D.记录法"
    
        If mblnMove Then
            mstrSQL = Replace(mstrSQL, "病人护理文件", "H病人护理文件")
            mstrSQL = Replace(mstrSQL, "病人护理数据", "H病人护理数据")
            mstrSQL = Replace(mstrSQL, "病人护理明细", "H病人护理明细")
        End If
            
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mT_Patient.lng文件ID, mT_Patient.lng病人ID, mT_Patient.lng主页ID, mT_Patient.lng婴儿, _
             CDate(mDTime), CDate(strEndTime), mstrCurveItem)
    End If
    '1--处理体温数据
    '--------------------------------------------------------------------------------------
    With rsTmp
        Do While Not .EOF
            lng项目序号 = zlCommFun.Nvl(!项目序号)
            Select Case lng项目序号
                Case gint心率
                    int标记 = 1
                Case Else
                    int标记 = Val(Nvl(!记录标记))
            End Select
            lngCol = GetTimeCOL(Format(zlCommFun.Nvl(!时间), "HH:mm:ss"))
            blnAllow = False: blnAdd = False: int显示 = 0
            '心率和脉搏公用时，检查脉搏对应的时间是否存在心率
            If mint心率应用 = 2 And lng项目序号 = -1 Then
                mrsPoint.Filter = "项目序号=2 and 时间='" & Format(zlCommFun.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "'"
                If mrsPoint.RecordCount > 0 Then
                    strParam = "序号|" & mrsPoint("序号")
                    strFidlds = "数值|ID"
                    
                    '脉搏短轴时心率未未记说明，脉搏为未记说明时就显示未记说明
                    If UBound(Split(mrsPoint("数值"), "/")) <> -1 Then
                        If IsNumeric(zlCommFun.Nvl(!数值)) Then
                            If mbln脉搏共用显示 Then
                                gstrValues = zlCommFun.Nvl(!数值) & "/" & Split(mrsPoint("数值"), "/")(0) & "|" & Val(zlCommFun.Nvl(!序号))
                            Else
                                gstrValues = Split(mrsPoint("数值"), "/")(0) & "/" & zlCommFun.Nvl(!数值) & "|" & Val(zlCommFun.Nvl(!序号))
                            End If
                            
                        Else
                            gstrValues = zlCommFun.Nvl(!数值) & "|" & Val(zlCommFun.Nvl(!序号))
                        End If
                    Else
                        gstrValues = mrsPoint("数值") & "|" & Val(zlCommFun.Nvl(!序号))
                    End If
                        
                    Call Record_Update(mrsPoint, strFidlds, gstrValues, strParam)
                    blnAllow = True
                Else
                    lng项目序号 = 2
                End If
            End If
            
            '处理物理降温
            If lng项目序号 = 1 And zlCommFun.Nvl(!记录标记) = 1 Then
                mrsPoint.Filter = "项目序号=1 and 时间='" & Format(zlCommFun.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "' and 标记<>1"
                If mrsPoint.RecordCount > 0 Then
                    strParam = "序号|" & mrsPoint("序号")
                    strFidlds = "数值|ID"
                    gstrValues = Split(mrsPoint("数值"), "/")(0) & "/" & zlCommFun.Nvl(!数值) & "|" & Val(zlCommFun.Nvl(!序号))
                    Call Record_Update(mrsPoint, strFidlds, gstrValues, strParam)
                End If
                blnAllow = True
            End If
            
            If blnAllow = False Then
                '进行曲线显示处理
                mrsPoint.Filter = "项目序号=" & lng项目序号 & " and 列号=" & lngCol & " and 显示=1"
                If mrsPoint.RecordCount > 0 Then
                    If Val(zlCommFun.Nvl(!显示)) = 1 And Val(mrsPoint!标注) <> 1 Then
                        blnAllow = True
                    ElseIf (Val(zlCommFun.Nvl(!显示)) = 1 And Val(mrsPoint!标注) = 1) Or (Val(zlCommFun.Nvl(!显示)) <> 1 And Val(mrsPoint!标注) <> 1) Then
                        blnAllow = CheckShow(mrsPoint("时间"), Format(zlCommFun.Nvl(!时间), "YYYY-MM-DD HH:mm:ss"), lngCol)
                    Else
                        blnAllow = False
                    End If
                    
                    int显示 = IIf(blnAllow = True, 1, 0)
                    int标注 = Val(zlCommFun.Nvl(!显示, 0))
                    
                    If blnAllow = True Then
                        Call Record_Update(mrsPoint, "显示", "0", "序号|" & mrsPoint!序号)
                    End If
                Else
                    int显示 = 1
                    int标注 = Val(zlCommFun.Nvl(!显示, 0))
                End If
                
                strPart = GetPart(lng项目序号)
                
                gstrValues = zlCommFun.Nvl(!序号) & "|" & zlCommFun.Nvl(!数值, zlCommFun.Nvl(!未记说明, "拒测")) & "|" & _
                    zlCommFun.Nvl(!体温部位, strPart) & "|" & int标记 & "|" & _
                    Format(zlCommFun.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "|" & lng项目序号 & "|" & zlCommFun.Nvl(!记录名) & "|" & Val(zlCommFun.Nvl(!复试合格)) & "|" & _
                    Val(zlCommFun.Nvl(!数据来源, 0)) & "|" & int显示 & "|" & int标注 & "|0|" & vfgThis.TextMatrix(0, vfgThis.FixedCols + lngCol - 1) & "|" & lngCol & "|0"
         
                Call Record_Add(mrsPoint, gstrFields, gstrValues)
            End If
        .MoveNext
        Loop
    End With
    
    '复制数据信息
    Set mrsCopy = CopyNewRs(mrsPoint)
        
    '展示数据信息
    Call ShowData
    
    zlRefreshData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CopyNewRs(ByVal rsData As ADODB.Recordset) As ADODB.Recordset
'-------------------------------------------------
'功能:复制新的记录集信息
'-------------------------------------------------
    Dim i As Integer
    Dim rsNew As New ADODB.Recordset
    On Error GoTo Errhand
    
    rsData.Filter = 0

    With rsNew
        '复制字段
        For i = 0 To rsData.Fields.Count - 1
            .Fields.Append rsData.Fields(i).Name, rsData.Fields(i).Type, rsData.Fields(i).DefinedSize, adFldIsNullable
        Next i
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        '复制数据信息
        rsData.Filter = 0
        Do While Not rsData.EOF
            .AddNew
            For i = 0 To rsData.Fields.Count - 1
                .Fields(rsData.Fields(i).Name).Value = rsData.Fields(i).Value
            Next i
            .Update
        rsData.MoveNext
        Loop
    End With
    
    rsNew.Filter = 0
    
    Set CopyNewRs = rsNew
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub ShowData()
'---------------------------------------------------
'功能:展示数据信息
'---------------------------------------------------
    Dim intRow As Integer, intCOl As Integer
    Dim strPart As String

    '检查是否存在显示为2的记录
    For intRow = vfgThis.FixedRows To vfgThis.Rows - 1
        For intCOl = vfgThis.FixedCols To vfgThis.Cols - 1
            mrsPoint.Filter = 0
            mrsPoint.Filter = "项目序号=" & Val(vfgThis.TextMatrix(intRow, 0)) & " and 标注=2 and 列号=" & (intCOl - vfgThis.FixedCols + 1)
            If mrsPoint.RecordCount > 0 Then
                '更新显示为2的记录
                Do While Not mrsPoint.EOF
                    mrsPoint!显示 = 2
                    mrsPoint.Update
                mrsPoint.MoveNext
                Loop
                '更新显示不为2的记录
                mrsPoint.Filter = "项目序号=" & Val(vfgThis.TextMatrix(intRow, 0)) & " and 标注<>2 and 列号=" & (intCOl - vfgThis.FixedCols + 1)
                Do While Not mrsPoint.EOF
                    mrsPoint!显示 = 0
                    mrsPoint.Update
                mrsPoint.MoveNext
                Loop
            End If
        Next intCOl
    Next intRow
    
    mrsPoint.Filter = 0
    '显示体温数据
    mrsPoint.Filter = "显示=1"
    mrsPoint.Sort = "序号,时间"
    With mrsPoint
        Do While Not .EOF
            For intRow = vfgThis.FixedRows To vfgThis.Rows - 1
                If Val(vfgThis.TextMatrix(intRow, 0)) = !项目序号 Then
                    strPart = GetPart(!项目序号)
                    If Nvl(!部位) = "" Then
                        vfgThis.TextMatrix(intRow, vfgThis.FixedCols + Val(!列号) - 1) = !数值
                    Else
                        vfgThis.TextMatrix(intRow, vfgThis.FixedCols + Val(!列号) - 1) = IIf(Trim(strPart) <> Trim(!部位), Trim(!部位) & ":" & !数值, !数值)
                    End If
                End If
            Next intRow
        .MoveNext
        Loop
    End With
    mblnInit = True
    Call vfgThis.Select(vfgThis.Row, vfgThis.Col)
    Call vfgThis_AfterRowColChange(vfgThis.Row, vfgThis.Col, vfgThis.Row, vfgThis.Col)
    mblnInit = False
End Sub

Private Function SaveData() As Boolean
'------------------------------------------------
'功能:保存数据信息
'------------------------------------------------
    Dim blnTran As Boolean
    Dim lngID As Long
    Dim strSQL As String
    Dim arrSQL() As String
    Dim i As Integer, lngItemCode As Long
    
    On Error GoTo Errhand
    Screen.MousePointer = 11
    
    ReDim Preserve arrSQL(1 To 1)
    
    With mrsPoint
        .Filter = 0
        Do While Not .EOF
            If Val(!状态) = 2 Then
                lngID = Val(!序号)
                lngItemCode = Val(!项目序号)
                
                If InStr(1, !数值, "/") = 0 Then
                    strSQL = "ZL_体温单数据_设置显示("
                    strSQL = strSQL & lngID & ","
                    strSQL = strSQL & Val(!显示) & ")"
                    
                    arrSQL(ReDimArray(arrSQL)) = strSQL
                Else
                    lngID = Val(!序号)
                    
                    strSQL = "ZL_体温单数据_设置显示("
                    strSQL = strSQL & lngID & ","
                    strSQL = strSQL & Val(!显示) & ")"
                    
                    arrSQL(ReDimArray(arrSQL)) = strSQL
                    
                    If InStr(1, ",1,2,", "," & lngItemCode & ",") <> 0 Then
                        lngID = Val(!Id)
                        
                        strSQL = "ZL_体温单数据_设置显示("
                        strSQL = strSQL & lngID & ","
                        strSQL = strSQL & Val(!显示) & ")"
                        
                        arrSQL(ReDimArray(arrSQL)) = strSQL
                    End If
                End If
            End If
        .MoveNext
        Loop
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '循环执行SQL保存数据
    'Debug.Print "----保存开始:" & Now
    gcnOracle.BeginTrans
    blnTran = True
    For i = 1 To UBound(arrSQL)
        If arrSQL(i) <> "" Then Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "保存体温数据"): ' Debug.Print CStr(arrSQL(i))
    Next
    gcnOracle.CommitTrans
    blnTran = False
    'Debug.Print "----保存结束:" & Now
    
    '修改状态=0
    mrsPoint.Filter = 0
    Do While Not mrsPoint.EOF
        mrsPoint!状态 = 0
        mrsPoint.Update
        mrsPoint.MoveNext
    Loop
    
    mblnChage = False
    mblnOK = True
    Screen.MousePointer = 0
    SaveData = True
    Exit Function
Errhand:
    If blnTran = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Screen.MousePointer = 0
End Function

Private Function GetPart(ByVal lng项目序号) As String
'功能:提取默认的体温部位
    Dim strPart As String
    mrs部位.Filter = "项目序号=" & lng项目序号 & " and 缺省项=1"
    If mrs部位.RecordCount > 0 Then strPart = zlCommFun.Nvl(mrs部位("部位"))
    GetPart = strPart
End Function

Private Function CheckShow(ByVal strBegin As String, ByVal strEnd As String, ByVal lngCol As Long) As Boolean
'-------------------------------------------------
'功能：对比两个时间点那个更靠近终点时间
'strbegin 对比的时间  strend当前时间   lngcol-1=时间范围数组的索引
'--------------------------------------------------
    Dim strTime As String
    Dim blnAllow As Boolean
    
    If (lngCol - 1) <= UBound(marrTime) Then
        If gintHourBegin + (lngCol - 1) * T_BodyStyle.lng时间间隔 = 24 Then
            strTime = Format(Format(mDTime, "YYYY-MM-DD") & " " & "23:59:59", "YYYY-MM-DD HH:mm:ss")
        Else
            strTime = Format(Format(mDTime, "YYYY-MM-DD") & " " & gintHourBegin + (lngCol - 1) * T_BodyStyle.lng时间间隔 & ":00:00", "YYYY-MM-DD HH:mm:ss")
        End If
    End If
    
    If Abs(DateDiff("s", CDate(Format(strBegin, "YYYY-MM-DD HH:mm:ss")), CDate(strTime))) > Abs(DateDiff("s", CDate(Format(strEnd, "YYYY-MM-DD HH:mm:ss")), CDate(strTime))) Then
        blnAllow = True
    Else
        blnAllow = False
    End If
    
    CheckShow = blnAllow
End Function

Private Function GetTimeCOL(ByVal strTime As String) As Integer
'--------------------------------------------------
'根据传入的时间计算改时间输入那段时间
'-------------------------------------------------
    Dim i As Integer
    Dim strValue As String
    
    strValue = Format(strTime, "HH:mm")
    For i = 0 To UBound(marrTime) - 1
        If strValue >= Format(Split(marrTime(i), ",")(0), "HH:mm") And strValue <= Format(Split(marrTime(i), ",")(1), "HH:mm") Then
            Exit For
        End If
    Next i
    
    GetTimeCOL = i + 1
End Function

Private Function InitTime() As String
'--------------------------------------------------------
'功能:提取一天的时间段信息
'--------------------------------------------------------
    Dim i As Integer
    Dim strName As String, strTime As String
    
    Call InitDateTimeRange(marrTime, gintHourBegin, T_BodyStyle.lng监测次数, T_BodyStyle.lng时间间隔)
    For i = 0 To UBound(marrTime) - 1
        strName = strName & ";" & Format(Split(marrTime(i), ",")(0), "HH:mm") & "-" & Format(Split(marrTime(i), ",")(1), "HH:mm")
    Next i
    If Left(strName, 1) = ";" Then strName = Mid(strName, 2)
    strName = "项目\时间范围" & ";" & strName
    
    For i = 0 To T_BodyStyle.lng监测次数 - 1
        strTime = strTime & ";" & gintHourBegin + i * T_BodyStyle.lng时间间隔
    Next i
    If Left(strTime, 1) = ";" Then strTime = Mid(strTime, 2)
    strTime = "项目\时间范围" & ";" & strTime
    
    InitTime = strTime & "[LPF]" & strName
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strParam As String
    Dim intCOl As Integer
    Select Case Control.Id
    
        Case conMenu_Edit_Save '保存
            If Not SaveData Then Exit Sub
            Set mrsCopy = CopyNewRs(mrsPoint)
            '展示数据信息
            Call ShowData
        Case conMenu_Edit_Reuse '取消
            '复制数据信息
            Set mrsPoint = CopyNewRs(mrsCopy)
            '展示数据信息
            Call ShowData
            mblnOK = False
            mblnChage = False
        Case conMenu_Edit_Curve, conMenu_Edit_CurveTable '设置记录
             If mblnChage Then
                If MsgBox("数据已经发生改变,请问是否需要保存?", vbInformation + vbDefaultButton1 + vbYesNo, gstrSysName) = vbYes Then
                    If Not SaveData Then Exit Sub
                End If
            End If
            intCOl = GetTimeCOL(Format(mDTime, "YYYY-MM-DD HH:mm:ss")) - 1
            If intCOl < 0 Then intCOl = 0
            strParam = Format(Format(mDTime, "YYYY-MM-DD") & " " & Split(marrTime(intCOl), ",")(0), "YYYY-MM-DD HH:mm:ss") & ";" & _
                Format(Format(mDTime, "YYYY-MM-DD") & " " & Split(marrTime(intCOl), ",")(1), "YYYY-MM-DD HH:mm:ss")
            '调用显示编辑窗体
            Call gobjTendEditor.BodyEditCur(IIf(Control.Id = conMenu_Edit_Curve, 0, -1), strParam)
            Call GetTableRowName
        Case conMenu_Help_Help '帮助
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_File_Exit '退出
            Unload Me
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    On Error Resume Next
    
    Bottom = stbThis.Height
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
    
    With picThis
        .Top = lngTop
        .Left = lngLeft
        .Width = lngRight - lngLeft
        .Height = lngBottom - lngTop
    End With
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim frmMain As Form
    Dim blnEnable As Boolean
    
     Select Case Control.Id
        Case conMenu_Edit_Save, conMenu_Edit_Reuse
             Control.Enabled = IIf(mblnChage = True, True, False)
        Case conMenu_Edit_Curve, conMenu_Edit_CurveTable
            blnEnable = True
            For Each frmMain In Forms
                If frmMain.Name = "frmCaseTendBodySetData" Then
                    blnEnable = False
                End If
            Next
            Control.Enabled = blnEnable
    End Select
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '卸载所有对象
    mbln出院 = False
    If Not (mrsPoint Is Nothing) Then Set mrsPoint = Nothing
    If Not (mrs部位 Is Nothing) Then Set mrs部位 = Nothing
    If Not (mcbrToolBar Is Nothing) Then Set mcbrToolBar = Nothing
    If Not (mrsCopy Is Nothing) Then Set mrsCopy = Nothing
    mblnChage = False
     '保存窗体
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub picShow_Paint()
    picShow.BackColor = &H8000000F
End Sub

Private Sub picShow_Resize()
    lblTmp.Top = 0
    lblTmp.Left = 0
    With vfgShow
        .Top = lblTmp.Height
        .Left = 0
        .Width = picShow.Width
        .Height = picShow.Height - lblTmp.Height - lblTmp.Top
    End With
End Sub

Private Sub picThis_Paint()
    picThis.BackColor = &H8000000F
End Sub

Private Sub picThis_Resize()
    With lblTime
        .Left = 10
        .Top = 10
        .Caption = "时间:" & Format(mDTime, "YYYY-MM-DD")
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    With vfgThis
        .Left = 5
        .Top = lblTime.Top + lblTime.Height + 20
        .Width = picThis.Width
        .Height = (picThis.Height - .Top) * 0.65
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    With picShow
        .Left = vfgThis.Left
        .Top = vfgThis.Height + vfgThis.Top + 50
        .Width = vfgThis.Width
        .Height = picThis.Height - picShow.Top
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    With lblTmp
        .Top = 10
        .Left = 10
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    With vfgShow
        .Left = 5
        .Top = lblTmp.Top + lblTmp.Height + 20
        .Width = picShow.Width
        .Height = picShow.Height - .Top
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    picShow.Visible = True
    lblTmp.ForeColor = RGB(0, 0, 255)
End Sub

Private Sub vfgShow_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    With vfgShow
        If .Col >= .FixedCols Then
            If NewRow = .Rows - 1 Then
                .FocusRect = flexFocusHeavy
            Else
                .FocusRect = flexFocusLight
            End If
        End If
    End With
End Sub

Private Sub vfgShow_Click()
    vfgShow.ColSel = vfgShow.Col
End Sub

Private Sub vfgShow_DblClick()
    Dim intSate As Integer, int显示 As Integer
    Dim intCOl As Integer, intRow As Integer
    Dim intColSel As Integer
    Dim arrValue() As String
    Dim strPart As String
    Dim lngItemNO As Long
    
    If mblnInit = False Then Exit Sub
    If mblnFileBack = True Then Exit Sub
    
    With vfgShow
        If .Rows - 1 = .Row And .Col >= .FixedCols Then
            '体温曲线项目
            If .TextMatrix(.Row, .Col) = "√" Then
                
                mrsPoint.Filter = 0
                mrsPoint.Filter = "序号=" & Val(.ColData(.Col))
                intSate = Val(mrsPoint!状态)
                intCOl = Val(mrsPoint!列号)
                lngItemNO = Val(mrsPoint!项目序号)
                int显示 = 2
                intSate = 2
                mrsPoint!显示 = int显示
                mrsPoint!状态 = intSate
                mrsPoint!标注 = int显示
                mrsPoint.Update
                .TextMatrix(.Row, .Col) = ""
                mrsPoint.Filter = "项目序号=" & lngItemNO & " And 列号=" & intCOl & " And 序号<>" & Val(.ColData(.Col))
                Do While Not mrsPoint.EOF
                    mrsPoint!显示 = 0
                    mrsPoint!标注 = 0
                    mrsPoint.Update
                mrsPoint.MoveNext
                Loop
            Else
                '处理记录集信息
                For intCOl = .FixedCols To .Cols - 1
                    If .TextMatrix(.Row, intCOl) = "√" Then
                        mrsPoint.Filter = 0
                        mrsPoint.Filter = "序号=" & Val(.ColData(intCOl))
                        intSate = Val(mrsPoint!状态)
                        int显示 = 0
                        Select Case intSate
                            Case 0
                                intSate = 2
                            Case 2
                                intSate = 0
                        End Select
                        mrsPoint!显示 = int显示
                        mrsPoint!状态 = intSate
                        mrsPoint!标注 = int显示
                        mrsPoint.Update
                        .TextMatrix(.Row, intCOl) = ""
                    End If
                Next intCOl
                .TextMatrix(.Row, .Col) = "√"
                mrsPoint.Filter = 0
                mrsPoint.Filter = "序号=" & Val(.ColData(.Col))
                intCOl = Val(mrsPoint!列号)
                lngItemNO = Val(mrsPoint!项目序号)
                intSate = Val(mrsPoint!状态)
                int显示 = 1
                Select Case intSate
                    Case 0
                        intSate = 2
                    Case 2
                        intSate = 0
                End Select
                mrsPoint!显示 = int显示
                mrsPoint!状态 = intSate
                mrsPoint!标注 = int显示
                mrsPoint.Update
                
                mrsPoint.Filter = "项目序号=" & lngItemNO & " And 列号=" & intCOl & " And 显示=2"
                Do While Not mrsPoint.EOF
                    intSate = Val(mrsPoint!状态)
                    int显示 = 0
                    intSate = 2
                    mrsPoint!显示 = int显示
                    mrsPoint!状态 = intSate
                    mrsPoint!标注 = int显示
                    mrsPoint.Update
                mrsPoint.MoveNext
                Loop
            End If
            vfgThis.Cell(flexcpText, vfgThis.FixedRows, vfgThis.FixedCols, vfgThis.Rows - 1, vfgThis.Cols - 1) = ""
            '显示数据
            mrsPoint.Filter = "显示=1"
            mrsPoint.Sort = "序号,时间"
            Do While Not mrsPoint.EOF
                For intRow = vfgThis.FixedRows To vfgThis.Rows - 1
                    If Val(vfgThis.TextMatrix(intRow, 0)) = Val(mrsPoint!项目序号) Then
                        strPart = GetPart(mrsPoint!项目序号)
                        If Trim(mrsPoint!部位) = "" Then
                            vfgThis.TextMatrix(intRow, vfgThis.FixedCols + Val(mrsPoint!列号) - 1) = mrsPoint!数值
                        Else
                            vfgThis.TextMatrix(intRow, vfgThis.FixedCols + Val(mrsPoint!列号) - 1) = IIf(Trim(strPart) <> Trim(mrsPoint!部位), Trim(mrsPoint!部位) & ":" & mrsPoint!数值, mrsPoint!数值)
                        End If
                    End If
                Next intRow
            mrsPoint.MoveNext
            Loop
            mblnChage = True
        End If
    End With
End Sub

Private Sub vfgThis_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim intCOl As Integer, intRow As Integer, i As Integer
    Dim strFind As String, strValue As String, strInfo As String
    intCOl = NewCol
    intRow = NewRow
    If mblnInit = False Then Exit Sub
    lblTmp.Caption = ""
    With vfgShow
        If NewRow >= vfgThis.FixedRows And NewCol >= vfgThis.FixedCols Then
            mintColSel = intCOl
            If vfgThis.TextMatrix(intRow, 0) = 1 Then '体温项目
                .Rows = 4
                .TextMatrix(0, 0) = "时点"
                .TextMatrix(1, 0) = "数值"
                .TextMatrix(2, 0) = "复查"
                .TextMatrix(3, 0) = "显示"
                strFind = " and 列号=" & intCOl - vfgThis.FixedCols + 1
            Else
                .Rows = 3
                .TextMatrix(0, 0) = "时点"
                .TextMatrix(1, 0) = "数值"
                .TextMatrix(2, 0) = "显示"
                strFind = " and 列号=" & intCOl - vfgThis.FixedCols + 1
             End If
             lblTmp.Caption = vfgThis.TextMatrix(0, intCOl) & "之间存在的" & vfgThis.TextMatrix(intRow, 2) & "数据有:"
        
             picShow.Visible = True
             mrsPoint.Filter = "项目序号=" & Val(vfgThis.TextMatrix(intRow, 0)) & strFind
             mrsPoint.Sort = "时间,序号"
             
             .Cols = mrsPoint.RecordCount + .FixedCols
             i = .FixedCols
             Do While Not mrsPoint.EOF
                .ColWidth(-1) = 1200 + 1200 * mintBigSize / 3
                 vfgShow.TextMatrix(0, i) = Format(mrsPoint!时间, "HH:mm")
                 vfgShow.TextMatrix(1, i) = mrsPoint!数值
                 If Val(vfgThis.TextMatrix(intRow, 0)) = 1 Then
                     vfgShow.TextMatrix(2, i) = IIf(mrsPoint!复查 = 1, "√", "")
                     vfgShow.TextMatrix(3, i) = IIf(mrsPoint!显示 = 1, "√", "")
                 Else
                     vfgShow.TextMatrix(2, i) = IIf(mrsPoint!显示 = 1, "√", "")
                 End If
                 vfgShow.ColData(i) = Val(mrsPoint!序号)
                 i = i + 1
             mrsPoint.MoveNext
             Loop
            .RowHeight(-1) = 300 + 300 * mintBigSize / 3
             .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
             .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = &H0&
             vfgThis.Cell(flexcpBackColor, vfgThis.FixedRows, vfgThis.FixedCols, vfgThis.Rows - 1, vfgThis.Cols - 1) = &H80000005
             vfgThis.Cell(flexcpBackColor, NewRow, NewCol, NewRow, NewCol) = &H80000018
        End If
    End With
End Sub

