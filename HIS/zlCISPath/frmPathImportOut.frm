VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPathImportOut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "门诊临床路径选择"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6690
   Icon            =   "frmPathImportOut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
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
      Top             =   2145
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
      Top             =   720
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
      FormatString    =   $"frmPathImportOut.frx":169B2
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
      Top             =   480
      Width           =   6495
   End
End
Attribute VB_Name = "frmPathImportOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mPati               As TYPE_Pati
Private mfrmParent          As Object

Private mlngDiagnosisType   As Long             '诊断类型:1-西医门诊诊断;11-中医门诊诊断
Private mlngDiagnosisSorce  As Long             '诊断来源 1-病历；3-首页整理
Private mlng疾病ID          As Long
Private mlng诊断ID          As Long

Private mblnOK              As Boolean

Private mrsDiag             As ADODB.Recordset
Private mrsPath             As ADODB.Recordset
Private mrsPati             As ADODB.Recordset

Public Function ShowMe(frmParent As Object, t_pati As TYPE_Pati, Optional blnImport As Boolean) As Boolean
'参数: blnImport:true-点击按钮导入;false-自动导入
'功能：设置一个参数，在填写完诊断之后自动导入的功能，只计算主要诊断，不计算次要诊断，点击导入路径按钮的时候再计算次要诊断
    Dim str首要诊断描述 As String
    Dim str次要诊断描述 As String
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim bln中医 As Boolean
    Dim i As Long
    Dim blnPath As Boolean
    Dim blnSecondPath As Boolean
    Dim blnPathSend As Boolean
    
    mPati = t_pati
    Set mfrmParent = frmParent
    
    mblnOK = False
    
    mlngDiagnosisType = 0
    mlngDiagnosisSorce = 0
    '检查该病人是否生成过项目
    blnPathSend = CheckOutPathSend(mPati.挂号ID)

    If blnPathSend Then
        MsgBox "该病人已经导入了临床路径，不允许再次导入。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '检查该病人是否在本科室有未完成的临床路径，是否可以继续路径
    If Not blnPathSend Then
        strSql = " Select 1 From 病人门诊路径 Where 病人ID=[1] and 科室ID+0 = [2] and 状态 = 1 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "CheckOutPathSend", mPati.病人ID, mPati.科室ID)
        If rsTmp.RecordCount > 0 Then
            MsgBox "该病人在本科室存在未完成的临床路径，不能够继续导入路径。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '导入路径
    '取出所有诊断，没有的话提示；
    '判断第一诊断的是否有合适的临床路径，有的话导入，没有的话不提示；
    '判断其余的诊断是否有符合的临床路径，有的话询问导入，没有的话提示。
    Set rsTmp = Get门诊病种ID(mPati.病人ID, mPati.挂号ID, 0, mPati.科室ID, bln中医)
    Set mrsDiag = rsTmp
    If rsTmp.RecordCount = 0 Then
        MsgBox "该病人没有填写任何诊断，请先填写后再执行导入。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Not blnPathSend Then
'        rsTmp.Filter = "诊断次序 = 1"
        For i = 1 To rsTmp.RecordCount                      '之所以用循环，是因为有可能中医一个，西医一个
            mlng疾病ID = Val("" & rsTmp!疾病id)
            mlng诊断ID = Val("" & rsTmp!诊断id)
            str首要诊断描述 = "" & rsTmp!诊断描述
            mlngDiagnosisType = Val("" & rsTmp!诊断类型)
            mlngDiagnosisSorce = Val("" & rsTmp!记录来源)
            Set mrsPath = GetOutPathTable(mlng疾病ID, mlng诊断ID, mPati.科室ID)
            If mrsPath.RecordCount > 0 Then
                blnPath = True
                Exit For
            End If
        Next
        
'        If Not blnPath And blnImport Then
'            rsTmp.Filter = "诊断次序 <> 1"
'            For i = 1 To rsTmp.RecordCount
'                mlng疾病ID = Val("" & rsTmp!疾病id)
'                mlng诊断ID = Val("" & rsTmp!诊断id)
'                str次要诊断描述 = "" & rsTmp!诊断描述
'                mlngDiagnosisType = Val("" & rsTmp!诊断类型)
'                mlngDiagnosisSorce = Val("" & rsTmp!记录来源)
'                Set mrsPath = GetOutPathTable(mlng疾病ID, mlng诊断ID, mPati.科室ID)
'                If mrsPath.RecordCount > 0 Then
'                    blnSecondPath = True
'                    Exit For
'                End If
'            Next
'        End If
        
        If Not blnPath And Not blnSecondPath Then
            MsgBox "当前科室没有适合于该病人主要诊断[" & str首要诊断描述 & "]的门诊临床路径。", vbInformation, gstrSysName
            Exit Function
        ElseIf (Not blnPath And blnSecondPath And blnImport) Then
            If MsgBox("当前科室没有适合于该病人首要诊断[" & str首要诊断描述 & "]的门诊临床路径。" & vbCrLf & _
                            "但存在适合于该病人次要诊断[" & str次要诊断描述 & "]的门诊临床路径。" & vbCrLf & _
                            "是否导入次要诊断的临床路径？", vbInformation + vbYesNo, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
        
        Set mrsPati = GetPatiInfoOut(mPati.病人ID, mPati.挂号ID)
        If mrsPati.RecordCount = 0 Then
            MsgBox "读取病人当前就诊信息失败。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    On Error Resume Next
    
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim t_pp As TYPE_PATH_Pati, arrtmp As Variant
    Dim lng就诊天数 As Long, lng标准治疗时间 As Long
    Dim rsTmp As ADODB.Recordset, str未导入编码 As String, str未导入名称 As String
    Dim lngB As Long, lngE As Long, strUnit As String, strTmp As String, DatCur As Date, lngValue As Long
    Dim i As Long, strFilter As String
    
    If vsPath.Row <= 0 Then
        MsgBox "请选择一个适用于该病人的门诊临床路径。", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    arrtmp = Split(vsPath.RowData(vsPath.Row), ":")
    t_pp.路径ID = arrtmp(0)
    t_pp.版本号 = arrtmp(1)
    mrsPath.Filter = "ID=" & t_pp.路径ID

    Set rsTmp = GetUnImportReson

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

    Me.Hide
    mblnOK = frmEvaluateOut.ShowMe(mfrmParent, 0, 1, mPati, t_pp, mrsPath!名称, mlngDiagnosisType, mlngDiagnosisSorce, mlng疾病ID, mlng诊断ID)
    
    If cmdOK.Tag <> "Unload" Then Unload Me
    Exit Sub
UnImport:
    '首要诊断才保存未导入原因
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
End Sub

Private Sub SaveUnImport(mPati As TYPE_Pati, mPP As TYPE_PATH_Pati, strVariationCode As String, strVariationTitle As String)
'功能：保存未导入原因
'参数：strVariationCode=未导入原因编码,strVariationTitle=未导入原因名称
    Dim strSql As String, strID As String, DateInPath As Date
    Dim str符合导入 As String
    
    On Error GoTo errH
    
    strID = zlDatabase.GetNextId("病人门诊路径")
    If CheckOutPathSend(mPati.挂号ID) Then
        DateInPath = zlDatabase.Currentdate
    Else
        DateInPath = GetPatiInDateOut(mPati)
    End If
    str符合导入 = "0"
    
    strSql = "Zl_病人门诊路径导入_Insert(" & mPati.病人ID & "," & mPati.挂号ID & "," & mPati.科室ID & "," & _
            mPP.路径ID & "," & mPP.版本号 & "," & strID & ",'" & UserInfo.姓名 & "','" & strVariationTitle & "'," & _
            str符合导入 & ",To_Date('" & Format(DateInPath, "yyyy-MM-DD HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),'" & _
            strVariationCode & "'," & mlngDiagnosisType & "," & mlngDiagnosisSorce & "," & IIf(mlng疾病ID = 0, "NULL", mlng疾病ID) & "," & IIf(mlng诊断ID = 0, "NULL", mlng诊断ID) & ",Null,1)"
    
    Call zlDatabase.ExecuteProcedure(strSql, "未导入原因")
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetUnImportReson() As ADODB.Recordset
'功能：读取固定的几项未导入原因
    Dim strSql As String
 
    strSql = "Select 编码, 名称 From 门诊变异常见原因 Where 性质 = 0 And 末级 = 1"
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
        If vsPath.Rows = vsPath.FixedRows + 1 Then
            vsPath.Row = vsPath.Rows - 1
            vsPath.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long

    lblPait.Caption = "当前病人：" & mrsPati!姓名 & "," & mrsPati!性别 & "," & mrsPati!年龄

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
            mrsPath.MoveNext
        Next
        
        cmdCancel.Visible = False
        cmdOK.Left = cmdCancel.Left
        
        '只有一个路径表
        If .Rows = .FixedRows + 1 Then
            .Row = .Rows - 1
            cmdOK.Tag = "Unload"
            cmdOK_Click
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '导入首要诊断时隐藏了取消按钮，只能点确定
    If mblnOK = False Then
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

Private Sub vsPath_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Sub vsPath_DblClick()
    Call cmdOK_Click
End Sub
