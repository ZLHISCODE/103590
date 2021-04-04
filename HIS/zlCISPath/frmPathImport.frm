VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPathImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "临床路径选择"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6690
   Icon            =   "frmPathImport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6690
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   6690
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3600
      Width           =   6690
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   5520
         TabIndex        =   5
         Top             =   195
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   4200
         TabIndex        =   4
         Top             =   195
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   10000
         Y1              =   45
         Y2              =   45
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   10000
         Y1              =   30
         Y2              =   30
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPath 
      Height          =   1185
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   6495
      _cx             =   11456
      _cy             =   2090
      Appearance      =   0
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
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   320
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathImport.frx":169B2
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      WordWrap        =   -1  'True
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsDiag 
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   6495
      _cx             =   11456
      _cy             =   2355
      Appearance      =   0
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
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   320
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathImport.frx":16A4A
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      WordWrap        =   -1  'True
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lblDiag 
      Caption         =   "该病人有多个合并症或并发症，请选择一个："
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   400
      Width           =   4575
   End
   Begin VB.Label lblPait 
      Caption         =   "当前病人：周小川,男,46岁"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblPath 
      Caption         =   "请从下表中选择一个适用于该病人的临床路径"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   6495
   End
End
Attribute VB_Name = "frmPathImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngFun As Long '0-首要路径导入,1-跳转路径(阶段评估时)，=2  合并路径导入,=3合并路径取消导入，=4合并路径完成
Private mlngPathID As Long  '路径跳转时，传回选择的路径表ID
Private mlngPathVersion As Long
Private mlngCurPathID As Long '路径跳转时，当前路径ID
Private mlngDiagnosisType As Long '诊断类型:1-西医门诊诊断;2-西医入院诊断;11-中医门诊诊断;12-中医入院诊断
Private mlngDiagnosisSorce As Long '诊断来源1-病历；2-入院登记；3-首页整理;4-病案

Private mPati As TYPE_Pati
Private mrsPath As ADODB.Recordset
Private mrsPati As ADODB.Recordset
Private mrsMerge As ADODB.Recordset    '此对象是传入的内存地址型对象，在Unload中不用卸载。
Private mfrmParent As Object
Private mblnOK As Boolean
Private mlng疾病ID As Long
Private mlng诊断ID As Long
Private mrsDiag As ADODB.Recordset
Private mlng首要路径记录ID As Long
Private mblnTmp As Boolean
Private mblnChoose As Boolean
Private mblnPathSend As Boolean
Private mbln外挂 As Boolean
Private mt_pp As TYPE_PATH_Pati
Private mlngHwnd As Long

Public Function ShowMe(frmParent As Object, t_pati As TYPE_Pati, ByVal lngFun As Long, ByRef t_pp As TYPE_PATH_Pati, _
    Optional ByVal lngCurPathID As Long, Optional ByRef lngPathID As Long, Optional ByRef lngPathVersion As Long, _
     Optional ByVal blnAuto As Boolean, Optional lng首要路径记录ID As Long, Optional rsMerge As Recordset, _
    Optional ByVal blnChoose As Boolean, Optional ByVal lngHwnd As Long, _
    Optional ByRef str名称 As String, Optional ByRef lngDiagnosisType As Long, Optional ByRef lngDiagnosisSorce As Long, _
    Optional ByRef lng疾病ID As Long, Optional ByRef lng诊断ID As Long) As Boolean
'参数：
'       lngCurPathID As Long '路径跳转时，当前路径ID
'       lngPathID,lngPathVersion=路径跳转时，传入之前选择的路径ID和版本，传回选择的路径ID和版本
'       lngFun=0  首要诊断导入，=1  合并路径导入,=2合并路径取消导入，=3合并路径完成, =4合并路径完成
'       lngFun=1时，blnAuto=true 导入首要路径后自动调用检查是否有可导入的合并路径，没有直接退出，不提示
'       rsMerge=合并路径取消或结束时，如果有多个合并路径，则弹出选择框，rsmerge则为所有合并路径的记录,blnChoose=true ,多选
'       lngHwnd=新版病历传入父窗体句柄，默认为0,新版病历不提示导入不成功的原因
    Dim str诊断描述 As String
    Dim rsTmp As ADODB.Recordset, rsNext As ADODB.Recordset
    Dim str疾病IDs As String
    Dim str诊断IDs As String
    Dim blnTmp As Boolean
    Dim bln中医 As Boolean
    Dim i As Long
    
    mPati = t_pati
    mlng首要路径记录ID = lng首要路径记录ID
    Set mfrmParent = frmParent
    mlngFun = lngFun
    
    mblnOK = False
    mlngCurPathID = lngCurPathID
    mlngPathID = lngPathID
    mlngPathVersion = lngPathVersion
    Set mrsMerge = rsMerge
    mblnChoose = blnChoose
    mlngHwnd = lngHwnd
    
    mlngDiagnosisType = 0
    mlngDiagnosisSorce = 0
    mblnPathSend = CheckPathSend(mPati.病人ID, mPati.主页ID)
    If lngHwnd <> 0 Then blnAuto = True
    
    '导入路径
    If rsMerge Is Nothing And lngFun <> 3 And lngFun <> 4 Then
        Set rsTmp = Get病种ID(mPati.病人ID, mPati.主页ID, IIf(lngFun = 2, 1, IIf(mblnPathSend And lngCurPathID = 0, 2, 0)), mPati.科室ID, bln中医)
        If rsTmp.RecordCount > 0 Then
            mlng疾病ID = Val("" & rsTmp!疾病id)
            mlng诊断ID = Val("" & rsTmp!诊断id)
            str诊断描述 = "" & rsTmp!诊断描述
            mlngDiagnosisType = Val("" & rsTmp!诊断类型)
            mlngDiagnosisSorce = Val("" & rsTmp!记录来源)
        End If
        If mlng疾病ID = 0 And mlng诊断ID = 0 Then
            If Not blnAuto Then
                If mlngFun = 0 Then
                    MsgBox "该病人没有填写任何诊断，请先填写后再执行导入。", vbInformation, gstrSysName
                ElseIf mlngFun = 2 Then
                    MsgBox "该病人除导入诊断外，还未填写其他诊断或并发症，请先填写后在导入合并路径。", vbInformation, gstrSysName
                ElseIf mlngFun = 1 Then
                    MsgBox "该病人的诊断记录被删除，无法执行路径跳转。", vbInformation, gstrSysName
                End If
            End If
            Exit Function
        End If
        
        If mlngFun = 0 Or mlngFun = 1 Then
            '第一次导入有效路径时或跳转路径时，按第一病种查询
            If mblnPathSend = False Or mlngFun = 1 Then
                '如果是首要路径则根据第一个疾病ID判断
                If mblnPathSend = False And bln中医 Then
                    '当中医科室入院诊断没有匹配到中医路径时默认判断西医入院诊断
                    rsTmp.Filter = "诊断类型 =12 OR 诊断类型 =2 "
                    For i = 1 To rsTmp.RecordCount
                        mlng疾病ID = Val("" & rsTmp!疾病id)
                        mlng诊断ID = Val("" & rsTmp!诊断id)
                        str诊断描述 = "" & rsTmp!诊断描述
                        mlngDiagnosisType = Val("" & rsTmp!诊断类型)
                        mlngDiagnosisSorce = Val("" & rsTmp!记录来源)
                        
                        Set mrsPath = GetPathTable(mlng疾病ID, mlng诊断ID, mPati.科室ID, lngCurPathID)
                        If mrsPath.RecordCount > 0 Then Exit For
                        rsTmp.MoveNext
                    Next
                    If mrsPath Is Nothing Then
                        If lngHwnd = 0 Then MsgBox "当前科室没有适合于该病人主要诊断[" & str诊断描述 & "]的临床路径。", vbInformation, gstrSysName
                        Exit Function
                    End If
                Else
                     Set mrsPath = GetPathTable(mlng疾病ID, mlng诊断ID, mPati.科室ID, lngCurPathID)
                End If
                
                If mblnPathSend = False And mrsPath.RecordCount = 0 Then
                    Set rsNext = Get病种ID(mPati.病人ID, mPati.主页ID, 3, mPati.科室ID)
                    If rsNext.RecordCount > 0 Then
                        If MsgBox("当前科室没有适合于该病人首要诊断[" & str诊断描述 & "]的临床路径。" & vbCrLf & _
                                "但存在适合于该病人次要诊断[" & rsNext!诊断描述 & "]的临床路径。" & vbCrLf & _
                                "是否导入次要诊断的临床路径？", vbInformation + vbYesNo, gstrSysName) = vbYes Then
                            mlng疾病ID = Val("" & rsNext!疾病id)
                            mlng诊断ID = Val("" & rsNext!诊断id)
                            str诊断描述 = "" & rsNext!诊断描述
                            mlngDiagnosisType = Val("" & rsNext!诊断类型)
                            mlngDiagnosisSorce = Val("" & rsNext!记录来源)
                            Set mrsPath = GetPathTable(mlng疾病ID, mlng诊断ID, mPati.科室ID, lngCurPathID)
                        Else
                            Exit Function
                        End If
                    End If
                End If
            Else
                rsTmp.Filter = ""
                Do While Not rsTmp.EOF
                    If Val(rsTmp!疾病id & "") <> 0 Then
                        str疾病IDs = str疾病IDs & "," & rsTmp!疾病id
                    End If
                    If Val(rsTmp!诊断id & "") <> 0 Then
                        str诊断IDs = str诊断IDs & "," & rsTmp!诊断id
                    End If
                    rsTmp.MoveNext
                Loop
                If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
                Set mrsDiag = rsTmp
                str疾病IDs = Mid(str疾病IDs, 2)
                str诊断IDs = Mid(str诊断IDs, 2)
                Set mrsPath = GetPathTable(0, 0, mPati.科室ID, 0, str疾病IDs, 0, str诊断IDs, mPati.病人ID, mPati.主页ID)
            End If
        Else
            '如果是合并路径，则传入所有非导入病种的其他诊断或并发症
            Do While Not rsTmp.EOF
                If Val(rsTmp!疾病id & "") <> 0 Then
                    str疾病IDs = str疾病IDs & "," & rsTmp!疾病id
                End If
                If Val(rsTmp!诊断id & "") <> 0 Then
                    str诊断IDs = str诊断IDs & "," & rsTmp!诊断id
                End If
                rsTmp.MoveNext
            Loop
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
            Set mrsDiag = rsTmp
            str疾病IDs = Mid(str疾病IDs, 2)
            str诊断IDs = Mid(str诊断IDs, 2)
            Set mrsPath = GetPathTable(0, 0, mPati.科室ID, 0, str疾病IDs, mlng首要路径记录ID, str诊断IDs)
        End If
        
        If mrsPath.RecordCount = 0 Then
            If Not blnAuto Then
                If mlngFun = 0 Then
                    MsgBox "当前科室没有" & IIf(mlngFun = 0, "", "其他") & "适合于该病人主要诊断[" & str诊断描述 & "]的临床路径。", vbInformation, gstrSysName
                Else
                    MsgBox "当前科室没有适合于该病人的临床合并路径。", vbInformation, gstrSysName
                End If
            End If
            Exit Function
        Else
            Set mrsPati = GetPatiInfo(mPati.病人ID, mPati.主页ID)
            If mrsPati.RecordCount = 0 Then
                MsgBox "读取病人当前住院信息失败。", vbInformation, gstrSysName
                Exit Function
            End If
            If mlngFun = 2 And blnAuto And mrsPath.RecordCount = 1 Then
                If MsgBox("当前病人存在可导入的合并路径:""" & mrsPath!名称 & """，是否要导入？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
        End If
    Else
        Set mrsPati = GetPatiInfo(mPati.病人ID, mPati.主页ID)
        If mrsPati.RecordCount = 0 Then
            MsgBox "读取病人当前住院信息失败。", vbInformation, gstrSysName
            Exit Function
        End If
        '只加未执行过的
        If mblnChoose Then
            If mlngFun = 3 Then
                rsMerge.Filter = "是否执行<>1"
                If rsMerge.RecordCount = 0 Then
                    MsgBox "该病人有合并路径都已经生成了项目，请取消生成合并路径的项目后再取消导入。", vbInformation, gstrSysName
                    Exit Function
                End If
            ElseIf mlngFun = 4 Then
                blnTmp = False
                Do While Not rsMerge.EOF
                    If Val(mrsMerge!显示 & "") = 1 Then blnTmp = True
                    
                    rsMerge.MoveNext
                Loop
                If Not blnTmp Then
                    MsgBox "该病人没有达到标准住院日的合并路径，如需提前完成，请选择下一阶段提前。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    End If
    
    On Error Resume Next
    If lngHwnd <> 0 Then
        Me.Show 1
    Else
        Me.Show 1, frmParent
    End If
    
    lngPathID = mlngPathID
    lngPathVersion = mlngPathVersion
    t_pp.病人路径ID = mt_pp.病人路径ID
    t_pp.病人路径状态 = mt_pp.病人路径状态
    t_pp.当前阶段ID = mt_pp.当前阶段ID
    t_pp.当前阶段分支ID = mt_pp.当前阶段分支ID
    t_pp.当前日期 = mt_pp.当前日期
    t_pp.当前天数 = mt_pp.当前天数
    t_pp.合并路径个数 = mt_pp.合并路径个数
    t_pp.阶段父ID = mt_pp.阶段父ID
    t_pp.结束路径控制 = mt_pp.结束路径控制
    t_pp.路径ID = mt_pp.路径ID
    t_pp.未导入原因 = mt_pp.未导入原因
    t_pp.原路径ID = mt_pp.原路径ID
    t_pp.版本号 = mt_pp.版本号
    lngDiagnosisType = mlngDiagnosisType
    lngDiagnosisSorce = mlngDiagnosisSorce
    lng疾病ID = mlng疾病ID
    lng诊断ID = mlng诊断ID
    
    ShowMe = mblnOK
    If mblnOK And Not rsMerge Is Nothing And (mlngFun = 3 Or mlngFun = 4) Then
         Set rsMerge = mrsMerge
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Set导入诊断(ByVal lng路径ID As Long)
'功能：如果是病人当次住院第二次导入路径，需要默认了一个导入诊断，如果有多条对应，则默认顺序是入院、出院。
    Dim strSql As String, rsTmp As Recordset
    Dim str疾病IDs As String, str诊断IDs As String
    
    On Error GoTo errH
    strSql = "Select a.疾病ID,a.诊断id From 临床路径病种 A Where a.路径ID=[1] And a.性质=0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "Set导入诊断", lng路径ID)
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            If Val(rsTmp!疾病id & "") <> 0 Then
                str疾病IDs = str疾病IDs & "," & rsTmp!疾病id
            End If
            If Val(rsTmp!诊断id & "") <> 0 Then
                str诊断IDs = str诊断IDs & "," & rsTmp!诊断id
            End If
            rsTmp.MoveNext
        Loop
        If mrsDiag.RecordCount > 0 Then
            str疾病IDs = Mid(str疾病IDs, 2)
            str诊断IDs = Mid(str诊断IDs, 2)
            mrsDiag.MoveFirst
            Do While Not mrsDiag.EOF
                If InStr("," & str疾病IDs & ",", "," & mrsDiag!疾病id & ",") > 0 And Val(mrsDiag!疾病id & "") <> 0 _
                    Or InStr("," & str诊断IDs & ",", "," & mrsDiag!诊断id & ",") > 0 And Val(mrsDiag!诊断id & "") <> 0 Then
                    mlng疾病ID = Val("" & mrsDiag!疾病id)
                    mlng诊断ID = Val("" & mrsDiag!诊断id)
                    mlngDiagnosisType = Val("" & mrsDiag!诊断类型)
                    mlngDiagnosisSorce = Val("" & mrsDiag!记录来源)
                    mrsDiag.MoveFirst
                    Exit Sub
                End If
                mrsDiag.MoveNext
            Loop
            mrsDiag.MoveFirst
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim t_pp As TYPE_PATH_Pati, arrtmp As Variant
    Dim lng住院天数 As Long, lng标准住院日 As Long
    Dim rsTmp As ADODB.Recordset, str未导入编码 As String, str未导入名称 As String
    Dim lngB As Long, lngE As Long, strUnit As String, strTmp As String, DatCur As Date, lngValue As Long
    Dim i As Long, strFilter As String
    Dim bln外挂判断 As Boolean
    Dim dt入院时间 As Date
    Dim dtDate As Date
    
    If mrsMerge Is Nothing And mlngFun <> 3 And mlngFun <> 4 Then
        If vsPath.Row <= 0 Then
            MsgBox "请选择一个适用于该病人的临床路径.", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
        arrtmp = Split(vsPath.RowData(vsPath.Row), ":")
        t_pp.路径ID = arrtmp(0)
        t_pp.版本号 = arrtmp(1)
        mrsPath.Filter = "ID=" & t_pp.路径ID
        If mlngFun = 0 And mblnPathSend Then
            Call Set导入诊断(t_pp.路径ID)
        End If

        Set rsTmp = GetUnImportReson


        '填了病例分型，并且路径表也定义了要求
        If Not IsNull(mrsPati!病例分型) Then
            If mrsPath!病例分型 <> "无" And mrsPati!病例分型 <> mrsPath!病例分型 Then
                MsgBox "该路径要求的病例分型[" & mrsPath!病例分型 & "]不适合于该病人的病例分型[" & mrsPati!病例分型 & "]", vbInformation, gstrSysName

                str未导入名称 = "病例分型不适用"
                GoTo UnImport
            End If
        End If

        If Not IsNull(mrsPati!当前病况) Then
            If mrsPath!适用病情 <> "通用" And mrsPath!适用病情 <> mrsPati!当前病况 Then
                MsgBox "该路径[" & mrsPath!适用病情 & "]不适合于该病人病情[" & mrsPati!当前病况 & "]", vbInformation, gstrSysName

                str未导入名称 = "病情不适用"
                GoTo UnImport
            End If
        End If
        If Val(mrsPath!适用性别) <> 0 Then
            If Val(mrsPath!适用性别) <> IIf(mrsPati!性别 = "男", 1, IIf(mrsPati!性别 = "女", 2, 0)) Then
                MsgBox "该路径不适合于该病人性别[" & mrsPati!性别 & "]", vbInformation, gstrSysName

                str未导入名称 = "性别不适合"
                GoTo UnImport
            End If
        End If

        If Not IsNull(mrsPath!适用年龄) And Not IsNull(mrsPati!年龄) Then
            lngValue = 0
            lngB = Split(mrsPath!适用年龄, "-")(0)
            strTmp = Split(mrsPath!适用年龄, "-")(1)
            lngE = Mid(strTmp, 1, Len(strTmp) - 1)
            strUnit = Mid(strTmp, Len(strTmp))

            strTmp = mrsPati!年龄           '特殊：2岁3月等
            If strUnit = Mid(strTmp, Len(strTmp)) And IsNumeric(Mid(strTmp, 1, Len(strTmp) - 1)) Or IsNumeric(strTmp) Then
                '相同年龄单位才做比较
                lngValue = Val(strTmp)
            ElseIf Not IsNull(mrsPati!出生日期) Then
                DatCur = zlDatabase.Currentdate
                lngValue = DateDiff(IIf(strUnit = "岁", "yyyy", IIf(strUnit = "月", "m", "d")), CDate(mrsPati!出生日期), DatCur)
                If lngValue = 0 Then lngValue = 1
            End If
            If lngValue <> 0 Then
                If lngValue < lngB Or lngValue > lngE Then
                    MsgBox "该路径不适合于该病人年龄[" & mrsPati!年龄 & "]", vbInformation, gstrSysName

                    str未导入名称 = "年龄不适合"
                    GoTo UnImport
                End If
            End If
        End If


        '住院日不能大于路径的标准住院日和确诊天数(如果没有设置了确诊天数，则不限制)
        dt入院时间 = GetPatiInDate(mPati, lng住院天数)
        dtDate = zlDatabase.Currentdate

        If InStr(mrsPath!标准住院日, "-") > 0 Then
            lng标准住院日 = Split(mrsPath!标准住院日, "-")(1)
        Else
            lng标准住院日 = Val(mrsPath!标准住院日)
        End If
        '如果是合并路径或者是当次住院有已经生成过路径项目的，不检查住院天数超出标准住院日
        '104002:设置了确诊天数,住院天数超过确诊天数禁止导入路径;确诊天数未设置或为0时,则住院天数大于标准住院日时禁止导入路径
        If mlngFun = 0 Or mlngFun = 1 Then
            If Not CheckPathSend(mPati.病人ID, mPati.主页ID) Then
                If mrsPath!确诊天数 <> 0 Then
                    If dtDate > Format(DateAdd("d", Val(mrsPath!确诊天数), dt入院时间), "yyyy-MM-DD HH:mm:ss") Then
                        MsgBox "该病人已入院" & lng住院天数 & "天，超过了规定的确诊天数(" & mrsPath!确诊天数 & "天)。", vbInformation, gstrSysName
                        str未导入名称 = "超过确诊天数"
                        GoTo UnImport
                    End If
                Else
                    If lng住院天数 > lng标准住院日 Then
                        MsgBox "该病人已入院" & lng住院天数 & "天，超过了该路径的标准住院日(" & lng标准住院日 & "天)。", vbInformation, gstrSysName
                        str未导入名称 = "超过标准住院日"
                        GoTo UnImport
                    End If
                End If
            End If
        End If

        If mlngFun = 0 Or mlngFun = 2 Then
            Me.Hide
            bln外挂判断 = True
            '临床路径导入前调用外挂口
            If CreatePlugInOK(p临床路径应用) Then
                On Error Resume Next
                bln外挂判断 = gobjPlugIn.PathImportBefore(glngSys, p临床路径应用, mPati.病人ID, mPati.主页ID, t_pp.路径ID, t_pp.版本号, , mlngDiagnosisSorce, mlng疾病ID, mlng诊断ID)
                '如果接口不存在，不影响原有逻辑
                If Not bln外挂判断 And Err.Number <> 0 Then bln外挂判断 = True
                Call zlPlugInErrH(Err, "PathImportBefore")
                Err.Clear: On Error GoTo 0
                If Not bln外挂判断 Then
                    mbln外挂 = True
                    mblnOK = True
                    Unload Me
                    Exit Sub
                End If
            End If
            
            If mlngHwnd = 0 Then
                mblnOK = frmEvaluate.ShowMe(mfrmParent, 0, 1, mPati, t_pp, mrsPath!名称, mlngDiagnosisType, mlngDiagnosisSorce, mlng疾病ID, mlng诊断ID, IIf(mlngFun = 0, 0, 1), mlng首要路径记录ID)
            Else
                mblnOK = True
                mt_pp.病人路径ID = t_pp.病人路径ID
                mt_pp.病人路径状态 = t_pp.病人路径状态
                mt_pp.当前阶段ID = t_pp.当前阶段ID
                mt_pp.当前阶段分支ID = t_pp.当前阶段分支ID
                mt_pp.当前日期 = t_pp.当前日期
                mt_pp.当前天数 = t_pp.当前天数
                mt_pp.合并路径个数 = t_pp.合并路径个数
                mt_pp.阶段父ID = t_pp.阶段父ID
                mt_pp.结束路径控制 = t_pp.结束路径控制
                mt_pp.路径ID = t_pp.路径ID
                mt_pp.未导入原因 = t_pp.未导入原因
                mt_pp.原路径ID = t_pp.原路径ID
                mt_pp.版本号 = t_pp.版本号
            End If
            '临床路径导后前调用外挂口
            If CreatePlugInOK(p临床路径应用) Then
                On Error Resume Next
                Call gobjPlugIn.PathImportAfter(glngSys, p临床路径应用, mPati.病人ID, mPati.主页ID, t_pp.路径ID, t_pp.版本号)
                Call zlPlugInErrH(Err, "PathImportAfter")
                Err.Clear: On Error GoTo 0
            End If
            
            If cmdOK.Tag <> "Unload" Then Unload Me
        Else
            mlngPathID = t_pp.路径ID
            mlngPathVersion = t_pp.版本号
            mblnOK = True
            Unload Me
        End If
    Else
        With vsPath
            On Error Resume Next
            If mblnChoose Then
                If mlngFun = 3 Then
                    For i = .FixedRows To .Rows - 1
                        If .Cell(flexcpChecked, i, 0) = 1 Then
                            strFilter = strFilter & " Or ID=" & .RowData(i)
                        End If
                    Next
                    If strFilter = "" Then
                        MsgBox "请至少选择一个需要取消的合并路径。", vbInformation, gstrSysName
                        Exit Sub
                    Else
                        mrsMerge.Filter = Mid(strFilter, 5)
                        mblnOK = True
                        Me.Hide
                    End If
                Else
                    For i = .FixedRows To .Rows - 1
                        mrsMerge.Filter = "ID=" & .RowData(i)
                        If .Cell(flexcpChecked, i, 0) = 1 Then
                            mrsMerge!选择 = 1
                        Else
                            mrsMerge!选择 = 0
                        End If
                    Next
                    mrsMerge.Update
                    mrsMerge.Filter = 0
                    mrsMerge.MoveFirst
                    mblnOK = True
                    Unload Me
                End If
            Else
                If .Row > 0 Then
                    mrsMerge.Filter = "ID=" & .RowData(.Row)
                    mblnOK = True
                    Me.Hide
                Else
                    MsgBox "请选择一个合并路径。", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End With
    End If
    Exit Sub

UnImport:
    '首要诊断才保存未导入原因
    If mlngFun = 0 Then
        rsTmp.Filter = "名称='" & str未导入名称 & "'"
        If rsTmp.RecordCount = 0 Then
            str未导入编码 = ""
        Else
            str未导入编码 = rsTmp!编码
        End If

        Call SaveUnImport(mPati, t_pp, str未导入编码, str未导入名称)
        mblnOK = True
        If cmdOK.Tag <> "Unload" Then
            Unload Me
        End If
    End If
End Sub

Private Sub SaveUnImport(mPati As TYPE_Pati, mPP As TYPE_PATH_Pati, strVariationCode As String, strVariationTitle As String)
'功能：保存未导入原因
'参数：strVariationCode=未导入原因编码,strVariationTitle=未导入原因名称
    Dim strSql As String, strID As String, DateInPath As Date
    Dim str符合导入 As String

    strID = zlDatabase.GetNextId("病人临床路径")
    If CheckPathSend(mPati.病人ID, mPati.主页ID) Then
        DateInPath = zlDatabase.Currentdate
    Else
        DateInPath = GetPatiInDate(mPati)
    End If
    str符合导入 = "0"
    
    
    strSql = "Zl_病人路径导入_Insert(" & mPati.病人ID & "," & mPati.主页ID & "," & mPati.科室ID & "," & _
            mPP.路径ID & "," & mPP.版本号 & "," & strID & ",'" & UserInfo.姓名 & "','" & strVariationTitle & "'," & _
            str符合导入 & ",To_Date('" & Format(DateInPath, "yyyy-MM-DD HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),'" & _
            strVariationCode & "'," & mlngDiagnosisType & "," & mlngDiagnosisSorce & "," & IIf(mlng疾病ID = 0, "NULL", mlng疾病ID) & "," & IIf(mlng诊断ID = 0, "NULL", mlng诊断ID) & ",Null,1)"
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSql, "未导入原因")
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetUnImportReson() As ADODB.Recordset
'功能：读取固定的几项未导入原因
    Dim strSql As String
 
    strSql = "Select 编码, 名称 From 变异常见原因 Where 性质 = 0 And 末级 = 1"
    On Error GoTo errH
    Set GetUnImportReson = zlDatabase.OpenSQLRecord(strSql, "未导入原因")

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Activate()
    If cmdOK.Tag = "Unload" Then
        cmdOK.Tag = ""
        Unload Me
    Else
        If vsPath.Rows = vsPath.FixedRows + 1 Then vsPath.Row = vsPath.Rows - 1: vsPath.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim strFilter As String
    Dim blnVisble As Boolean
    
    lblPait.Caption = "当前病人：" & mrsPati!姓名 & "," & mrsPati!性别 & "," & mrsPati!年龄
    
    If mrsMerge Is Nothing And mlngFun <> 3 And mlngFun <> 4 Then
        If mlngFun = 2 Then
            With vsDiag
                '先查找有多少个合并症或并发症有路径对应
                Do While Not mrsPath.EOF
                    If Val(mrsPath!疾病id & "") <> 0 Then
                        If InStr(strFilter & "", " Or 疾病ID=" & mrsPath!疾病id) = 0 Then
                            strFilter = strFilter & " Or 疾病ID=" & mrsPath!疾病id
                        End If
                    ElseIf Val(mrsPath!诊断id & "") <> 0 Then
                        If InStr(strFilter & "", " Or 诊断ID=" & mrsPath!诊断id) = 0 Then
                            strFilter = strFilter & " Or 诊断ID=" & mrsPath!诊断id
                        End If
                    End If
                    mrsPath.MoveNext
                Loop
                strFilter = Mid(strFilter, 5)
                If strFilter <> "" Then
                    mrsDiag.Filter = strFilter
                End If
                
                If mrsDiag.RecordCount > 0 Then
                    If mrsDiag.RecordCount = 1 Then
                        blnVisble = False
                    Else
                        blnVisble = True
                    End If
                    .Rows = .FixedRows
                    Do While Not mrsDiag.EOF
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 0) = IIf(Val(mrsDiag!诊断类型 & "") = 10, "并发症", "合并症")
                        .TextMatrix(.Rows - 1, 1) = Decode(Val(mrsDiag!记录来源 & ""), 1, "病历", 2, "入院登记", 3, "首页整理", "")
                        .TextMatrix(.Rows - 1, 2) = Decode(Val(mrsDiag!诊断类型 & ""), 10, "并发症", 1, "西医门诊诊断", 2, "西医入院诊断", 3, "西医出院诊断", 11, "中医门诊诊断", 12, "中医入院诊断", 13, "中医出院诊断")
                        .TextMatrix(.Rows - 1, 3) = mrsDiag!诊断描述 & ""
                        .RowData(.Rows - 1) = Val(mrsDiag!疾病id & "")
                        .TextMatrix(.Rows - 1, 4) = Val(mrsDiag!诊断id & "")
                        mrsDiag.MoveNext
                    Loop
                    .Row = .FixedRows
                    '只有一个诊断和一个路径表
                    If .Rows = .FixedRows + 1 And vsPath.Rows = vsPath.FixedRows + 1 Then
                        vsPath.Row = vsPath.Rows - 1
                        cmdOK.Tag = "Unload"
                        cmdOK_Click
                    End If
                End If
            End With
        Else
            blnVisble = False
            If mrsPath.RecordCount > 0 Then mrsPath.MoveFirst
        
            With vsPath
                .Rows = .FixedRows
                For i = 1 To mrsPath.RecordCount
                    '缺省不选择任何一行
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = mrsPath!编码
                    .TextMatrix(.Rows - 1, 2) = mrsPath!名称
                    .TextMatrix(.Rows - 1, 3) = "" & mrsPath!说明
                    .RowData(.Rows - 1) = mrsPath!ID & ":" & mrsPath!最新版本
                    If mlngFun = 1 Then
                        If mrsPath!ID = mlngPathID Then .Row = i
                    End If
                    
                    mrsPath.MoveNext
                Next
                
                If mlngFun = 0 Then
                    cmdCancel.Visible = False
                    cmdOK.Left = cmdCancel.Left
                    
                    '只有一个路径表
                    If .Rows = .FixedRows + 1 Then
                        .Row = .Rows - 1
                        cmdOK.Tag = "Unload"
                        cmdOK_Click
                    End If
                End If
            End With
        End If
    Else
        With vsPath
            .Rows = .FixedRows
            If mlngFun = 3 Then
                For i = 1 To mrsMerge.RecordCount
                    '缺省不选择任何一行
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = mrsMerge!编码
                    .TextMatrix(.Rows - 1, 2) = mrsMerge!名称
                    .TextMatrix(.Rows - 1, 3) = "" & mrsMerge!说明
                    .RowData(.Rows - 1) = mrsMerge!ID & ""
                    mrsMerge.MoveNext
                Next
            ElseIf mlngFun = 4 Then
                If mrsMerge.RecordCount > 0 Then mrsMerge.MoveFirst
                For i = 1 To mrsMerge.RecordCount
                    If Val(mrsMerge!显示 & "") = 1 Then
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 1) = mrsMerge!编码
                        .TextMatrix(.Rows - 1, 2) = mrsMerge!名称
                        .TextMatrix(.Rows - 1, 3) = "" & mrsMerge!说明
                        .TextMatrix(.Rows - 1, 4) = Val(mrsMerge!允许修改 & "")
                        .RowData(.Rows - 1) = mrsMerge!ID & ""
                        If Val(mrsMerge!选择 & "") = 1 Then
                            .Cell(flexcpChecked, .Rows - 1, 0) = 1
                        End If
                    End If
                    mrsMerge.MoveNext
                Next
            End If
            lblPath.Caption = "请从下列合并路径中选择:"
            If mblnChoose Then
                vsPath.Editable = flexEDKbdMouse
                .ColHidden(0) = False
                .ColWidth(2) = .ColWidth(2) - .ColWidth(0)
            End If
            '只有一个路径表,完成不自动勾选
            If .Rows = .FixedRows + 1 And mlngFun = 3 Then
                .Row = .Rows - 1
                If mblnChoose Then
                    .Cell(flexcpChecked, .Row, 0) = 1
                End If
                cmdOK.Tag = "Unload"
                cmdOK_Click
            End If
        End With
    End If
    If mbln外挂 Then mbln外挂 = False: Exit Sub
    If Not blnVisble Then
        '只有一个合并症或者是导入首要诊断或跳转则隐藏诊断列表
        lblDiag.Visible = False
        vsDiag.Visible = False
        lblPath.Top = lblDiag.Top
        vsPath.Top = vsDiag.Top
        vsPath.Height = vsPath.Height + 800
        Me.Height = vsPath.Top + vsPath.Height + picBottom.Height + 450
    End If
            
End Sub
Private Sub Form_Unload(Cancel As Integer)
    '导入首要诊断时隐藏了取消按钮，只能点确定
    If mlngFun = 0 And mblnOK = False And mrsMerge Is Nothing Then
        Cancel = 1
    Else
        Set mrsPath = Nothing
        Set mrsPati = Nothing
        Set mfrmParent = Nothing
        Set mrsDiag = Nothing
        mlng疾病ID = 0
        mlng诊断ID = 0
    End If
End Sub

Private Sub vsDiag_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Long
    
    With vsPath
        If mlngFun = 2 And NewRow >= .FixedRows Then
            .Rows = .FixedRows
            mrsPath.Filter = "疾病ID=" & Val(vsDiag.RowData(NewRow) & "") & " Or 诊断ID=" & Val(vsDiag.TextMatrix(NewRow, 4))
            If mrsPath.RecordCount > 0 Then mrsPath.MoveFirst: mlng疾病ID = Val(vsDiag.RowData(NewRow) & ""): mlng诊断ID = Val(vsDiag.TextMatrix(NewRow, 4))
            For i = 1 To mrsPath.RecordCount
                '缺省不选择任何一行
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = mrsPath!编码
                .TextMatrix(.Rows - 1, 2) = mrsPath!名称
                .TextMatrix(.Rows - 1, 3) = "" & mrsPath!说明
                .RowData(.Rows - 1) = mrsPath!ID & ":" & mrsPath!最新版本
                
                mrsPath.MoveNext
            Next
        End If
    End With
End Sub

Private Sub vsPath_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If mblnTmp Then
        If Row > 0 And Not mrsMerge Is Nothing And Col = 0 And mblnChoose And (mlngFun = 3 Or mlngFun = 4) Then
            If vsPath.Cell(flexcpChecked, Row, 0) = 1 Then
                vsPath.Cell(flexcpChecked, Row, 0) = 0
            Else
                vsPath.Cell(flexcpChecked, Row, 0) = 1
            End If
        End If
        mblnTmp = False
    End If
End Sub

Private Sub vsPath_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Sub vsPath_DblClick()
    Call cmdOK_Click
End Sub

Private Sub vsPath_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
        If vsPath.Row > 0 And Not mrsMerge Is Nothing And mblnChoose And (mlngFun = 3 Or mlngFun = 4) Then
            If vsPath.Cell(flexcpChecked, vsPath.Row, 0) = 1 Then
                If vsPath.TextMatrix(vsPath.Row, 4) <> "1" Then
                    vsPath.Cell(flexcpChecked, vsPath.Row, 0) = 0
                End If
            Else
                vsPath.Cell(flexcpChecked, vsPath.Row, 0) = 1
            End If
            mblnTmp = True
            '先鼠标勾选一次，在按空格，则会继续调用AfterEdit，如果选择其他行按空格则没有问题。
        End If
    End If
End Sub



Private Sub vsPath_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsPath.TextMatrix(Row, 4) = "1" Then
        MsgBox "该合并路径已经达到标准住院日,如需不结束，请选择下一阶段延后。", vbInformation, "合并路径结束"
        Cancel = True
        Exit Sub
    End If
End Sub


