VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmModuleCheck 
   Caption         =   "模块对象权限检查"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10815
   Icon            =   "frmModuleCheck.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   10815
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.Frame fraCheck 
      Height          =   6650
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   10755
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   1
         Left            =   0
         TabIndex        =   7
         Top             =   6000
         Width           =   10680
      End
      Begin VB.CommandButton cmdRepair 
         Caption         =   "修复(&R)"
         Height          =   350
         Left            =   8160
         TabIndex        =   6
         Top             =   6200
         Width           =   1100
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "退出(&X)"
         Height          =   350
         Left            =   9360
         TabIndex        =   5
         Top             =   6200
         Width           =   1100
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   0
         Left            =   15
         TabIndex        =   3
         Top             =   690
         Width           =   10800
      End
      Begin VSFlex8Ctl.VSFlexGrid vsCheckResult 
         Height          =   5100
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   10395
         _cx             =   18336
         _cy             =   8996
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
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483628
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   100
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmModuleCheck.frx":74F2
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
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
         Editable        =   2
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
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "模块授权后任然存在无效标识符错误，请使用该功能进行检查和修复。"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   180
         Width           =   5580
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "以下模块的对象权限存在问题："
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   435
         Width           =   2520
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6630
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmModuleCheck.frx":75BE
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15028
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "11:13"
            Key             =   "STANUM"
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
Attribute VB_Name = "frmModuleCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsCheck As ADODB.Recordset
Private mblnOK As Boolean

Private Enum ErrInfoCol
    Col_自动修复 = 0
    Col_系统 = 1
    Col_模块 = 2
    Col_所有者 = 3
    Col_对象 = 4
    Col_权限 = 5
    Col_错误信息 = 6
End Enum

Public Function ShowMe(Optional ByVal lngSys As Long, Optional ByRef blnHaveErr As Boolean, Optional ByVal blnCheck As Boolean) As Boolean
'功能：ShowMe入口
    Dim strSQL As String
    Dim rsCheck As ADODB.Recordset
    If Not blnCheck Then
        strSQL = "Select 序号, 标题, 系统, 所有者, 对象, 权限,b.OWNER 实际所有者,b.OBJECT_NAME 实际对象,b.OBJECT_TYPE 对象类型" & vbNewLine & _
                    "From (Select L.序号, L.标题, L.系统, P.功能, P.对象, P.所有者, P.权限" & vbNewLine & _
                    "From Zlprograms l, Zlprogprivs p," & vbNewLine & _
                    "     (Select Table_Schema As 所有者, Table_Name As 对象, Privilege As 权限 From All_Tab_Privs Where Grantable = 'YES'  Union" & vbNewLine & _
                    "       Select User, Object_Name, 'ALTER' From User_Objects Where Object_Type In ('TABLE', 'SEQUENCE')  Union" & vbNewLine & _
                    "       Select User, Object_Name, 'DELETE' From User_Objects Where Object_Type In ('TABLE', 'VIEW') Union" & vbNewLine & _
                    "       Select User, Object_Name, 'EXECUTE' From User_Objects Where Object_Type In ('PACKAGE', 'PROCEDURE', 'FUNCTION', 'TYPE') Union" & vbNewLine & _
                    "       Select User, Object_Name, 'INDEX' From User_Objects Where Object_Type = 'TABLE'  Union" & vbNewLine & _
                    "       Select User, Object_Name, 'INSERT' From User_Objects Where Object_Type In ('TABLE', 'VIEW')  Union" & vbNewLine & _
                    "       Select User, Object_Name, 'REFERENCES' From User_Objects Where Object_Type = 'TABLE' Union" & vbNewLine & _
                    "       Select User, Object_Name, 'SELECT' From User_Objects Where Object_Type In ('TABLE', 'VIEW', 'SEQUENCE') Union" & vbNewLine & _
                    "       Select User, Object_Name, 'UPDATE' From User_Objects Where Object_Type In ('TABLE', 'VIEW') ) r" & vbNewLine & _
                    "Where Nvl(L.系统, 0) = Nvl(P.系统, 0) And L.序号 = P.序号 And Upper(P.对象) = R.对象(+) And Upper(P.所有者) = R.所有者(+) And" & vbNewLine & _
                    "      Upper(P.权限) = R.权限(+) And R.对象 Is Null And P.对象 Is Not Null " & IIf(lngSys <> 0, "And nvl(L.系统,0)=" & lngSys, "") & ") a,(select * from  All_Objects b where b.OBJECT_TYPE<>'SYNONYM') b" & vbNewLine & _
                    "Where  Upper(A.对象) = B.Object_Name(+)" & vbNewLine & _
                    "order by 所有者,对象"
    Else
        strSQL = "Select 1" & vbNewLine & _
                    "From Zlprograms l, Zlprogprivs p," & vbNewLine & _
                    "(Select Table_Schema As 所有者, Table_Name As 对象, Privilege As 权限 From All_Tab_Privs Where Grantable = 'YES'  Union" & vbNewLine & _
                    "Select User, Object_Name, 'ALTER' From User_Objects Where Object_Type In ('TABLE', 'SEQUENCE')  Union" & vbNewLine & _
                    "Select User, Object_Name, 'DELETE' From User_Objects Where Object_Type In ('TABLE', 'VIEW') Union" & vbNewLine & _
                    "Select User, Object_Name, 'EXECUTE' From User_Objects Where Object_Type In ('PACKAGE', 'PROCEDURE', 'FUNCTION', 'TYPE') Union" & vbNewLine & _
                    "Select User, Object_Name, 'INDEX' From User_Objects Where Object_Type = 'TABLE'  Union" & vbNewLine & _
                    "Select User, Object_Name, 'INSERT' From User_Objects Where Object_Type In ('TABLE', 'VIEW')  Union" & vbNewLine & _
                    "Select User, Object_Name, 'REFERENCES' From User_Objects Where Object_Type = 'TABLE' Union" & vbNewLine & _
                    "Select User, Object_Name, 'SELECT' From User_Objects Where Object_Type In ('TABLE', 'VIEW', 'SEQUENCE') Union" & vbNewLine & _
                    "Select User, Object_Name, 'UPDATE' From User_Objects Where Object_Type In ('TABLE', 'VIEW') ) r" & vbNewLine & _
                    "Where Nvl(L.系统, 0) = Nvl(P.系统, 0) And L.序号 = P.序号 And Upper(P.对象) = R.对象(+) And Upper(P.所有者) = R.所有者(+) And" & vbNewLine & _
                    "Upper(P.权限) = R.权限(+) And R.对象 Is Null And P.对象 Is Not Null"
    End If
    Set rsCheck = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "模块对象权限检查")
    mblnOK = False
    If rsCheck.RecordCount = 0 Then
        blnHaveErr = False
        If blnCheck Then
            ShowMe = False
        End If
        Exit Function
    Else
        blnHaveErr = True
        If blnCheck Then
            ShowMe = True
            Exit Function
        End If
    End If
    Set mrsCheck = CopyNewRec(rsCheck, False, , Array("自动修复", adInteger, 1, 0, "错误提示", adVarChar, 2000, Empty, "修复SQL", adVarChar, 2000, Empty))
    Me.Show 1
    ShowMe = mblnOK
End Function

Private Sub cmdExit_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdRepair_Click()
    Dim arrTmp As Variant
    
    mrsCheck.Filter = "自动修复=1"
    mrsCheck.Sort = "系统,序号,所有者,对象"
    Do While Not mrsCheck.EOF
        If mrsCheck!修复SQL & "" <> "" Then
            arrTmp = Split(mrsCheck!修复SQL & "", ";")
            gcnOracle.Execute arrTmp(0)
            gcnOracle.Execute arrTmp(1)
        End If
        mrsCheck.MoveNext
    Loop
    mrsCheck.Filter = "自动修复<>1"
    mblnOK = mrsCheck.RecordCount = 0
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strInfo As String
    Dim strSQL As String
    Dim i As Long
    
    Call ApplyOEM(stbThis)
    With mrsCheck
        .Filter = "实际对象=Null"
        Do While Not .EOF
            strInfo = IIf(JudgeName(!对象 & ""), "对象名称中存在意外字符（除字母、数字、下划线、汉字之外的字符）!", "")
            If IsNull(!系统) Then '可能是报表
                .Update "错误提示", "该报表涉及的对象”" & !所有者 & "." & !对象 & "“" & IIf(strInfo = "", "可能已经不存在！", strInfo)
            Else
                .Update "错误提示", "该模块涉及的对象”" & !所有者 & "." & !对象 & "“" & IIf(strInfo = "", "可能已经不存在！", strInfo)
            End If
            .MoveNext
        Loop
        .Filter = "实际对象<>Null"
        Do While Not .EOF
            strInfo = ""
            strSQL = GetUpdateSQL(UCase(!对象 & ""), UCase(!所有者 & ""), UCase(!权限 & ""), UCase(!实际所有者 & ""), UCase(!对象类型 & ""), strInfo)
            .Update Array("自动修复", "错误提示", "修复SQL"), Array(IIf(strSQL = "", 0, 1), strInfo, strSQL)
            .MoveNext
        Loop
         .Filter = ""
        .Sort = "系统,序号,所有者,对象"
        vsCheckResult.Rows = 1: vsCheckResult.Rows = .RecordCount + 1
        For i = 1 To .RecordCount
            vsCheckResult.TextMatrix(i, Col_系统) = !系统 & ""
            vsCheckResult.TextMatrix(i, Col_模块) = "【" & Format(!序号 & "", "000000") & "】" & !标题
            vsCheckResult.TextMatrix(i, Col_所有者) = !所有者 & ""
            vsCheckResult.TextMatrix(i, Col_对象) = !对象 & ""
            vsCheckResult.TextMatrix(i, Col_权限) = !权限 & ""
            vsCheckResult.TextMatrix(i, Col_错误信息) = !错误提示 & ""
            vsCheckResult.TextMatrix(i, Col_自动修复) = IIf(Val(!自动修复 & "") = 0, "×", "√")
            .MoveNext
        Next
        .Filter = "自动修复=1"
        cmdRepair.Visible = .RecordCount <> 0
    End With
End Sub

Private Function JudgeName(ByVal strName As String) As Boolean
'功能：判断对象名称是否存在问题
    Dim i As Long, j As Long, strChar As String
    Dim blnExit As Boolean
    
    strName = Trim(strName)
    For i = 1 To Len(strName)
        strChar = Mid(strName, i, 1)
        If strChar = "_" Then
            '允许有下划线
        ElseIf IsNumeric(strChar) Then
            '允许有数字
        Else
            j = Asc(strChar)
            If (j > 64 And j < 91) Or (j > 96 And j < 123) Then
                '允许有字母
            ElseIf j < 0 Then
                '允许有汉字
            Else
                blnExit = True
            End If
        End If
        If blnExit Then Exit For
    Next
    JudgeName = blnExit
End Function

Private Function GetUpdateSQL(ByVal strName As String, ByVal strOwner As String, ByVal strPriv As String, ByVal strActualOwer As String, ByVal strType As String, ByRef strErr As String) As String
    Dim strTmpOwner As String
    Dim strTmpPrivs As String
    Dim strErrTmp As String
    Dim strSQL As String
    
    If strOwner <> strActualOwer Then
        strErr = "该对象实际所有者为：" & strActualOwer & "，ZLProgPrivs中所有者为：" & strOwner
        strTmpOwner = strActualOwer
    Else
        strTmpOwner = strOwner
    End If

    Select Case strType
        Case "TABLE"
            If strPriv = "DELETE" Or strPriv = "SELECT" Or strPriv = "UPDATE" Or strPriv = "INSERT" Or _
                strPriv = "ALTER" Or strPriv = "REFERENCES" Or strPriv = "INDEX" Then
                strTmpPrivs = strPriv
            ElseIf strPriv = "EXECUTE" Then
                strErrTmp = "该对象类型为：" & strType & "，不具有”" & strPriv & "“权限（可具有DELETE,SELECT,UPDATE,INSERT,ALTER,REFERENCES,INDEX权限）"
            End If
        Case "VIEW"
            If strPriv = "DELETE" Or strPriv = "SELECT" Or strPriv = "UPDATE" Or strPriv = "INSERT" Then
                strTmpPrivs = strPriv
            ElseIf strPriv = "EXECUTE" Then
                strTmpPrivs = "SELECT"
                strErrTmp = "该对象类型为：" & strType & "，不具有”" & strPriv & "“权限（可具有DELETE,SELECT,UPDATE,INSERT权限）"
            End If
        Case "SEQUENCE"
            If strPriv = "SELECT" Or strPriv = "ALTER" Then
                strTmpPrivs = strPriv
            Else
                strErrTmp = "该对象类型为：" & strType & "，不具有”" & strPriv & "“权限（可具有SELECT,ALTER权限）"
            End If
        Case "PACKAGE", "PROCEDURE", "FUNCTION", "TYPE"
            If strPriv = "EXECUTE" Then
                strTmpPrivs = strPriv
            Else
                strTmpPrivs = "EXECUTE"
                strErrTmp = "该对象类型为：" & strType & "，不具有”" & strPriv & "“权限（可具有EXECUTE权限）"
            End If
    End Select
    If strErrTmp <> "" Or strErr <> "" Then
        If strErr <> "" Then
            strErr = IIf(strErrTmp <> "", "1、" & strErr & "；2、" & strErrTmp, strErr)
        Else
            strErr = strErrTmp
        End If
        
        If strTmpPrivs <> "" Then
            '删除这种情况，同一模块功能下有两条同一对象权限数据，但是条错误，一条正确
            strSQL = "Delete From Zlprogprivs a" & vbNewLine & _
                        "Where Upper(所有者) ='" & strTmpOwner & "' And Upper(权限) = '" & strTmpPrivs & "' And Upper(对象) = '" & strName & "' And Exists" & vbNewLine & _
                        " (Select 1" & vbNewLine & _
                        "       From Zlprogprivs b" & vbNewLine & _
                        "       Where Nvl(B.系统, 0) = Nvl(A.系统, 0) And B.序号 = A.序号 And A.功能 = B.功能 And Upper(所有者) ='" & strActualOwer & "' And Upper(权限) = '" & strTmpPrivs & "' And Upper(对象) = '" & strName & "') "

            '更新权限可能会违反唯一约束，因此只更新不会违反唯一约束的数据
            strSQL = strSQL & ";" & "Update Zlprogprivs a" & vbNewLine & _
                        "Set 所有者 = '" & strTmpOwner & "', 权限 = '" & strTmpPrivs & "'" & vbNewLine & _
                        "Where Upper(所有者) = '" & strOwner & "' And Upper(对象) = '" & strName & "' And Upper(权限) = '" & strTmpPrivs & "'  And Not Exists" & vbNewLine & _
                        " (Select 1" & vbNewLine & _
                        "       From Zlprogprivs b" & vbNewLine & _
                        "       Where Nvl(B.系统, 0) = Nvl(A.系统, 0) And B.序号 = A.序号 And A.功能 = B.功能 And Upper(B.对象) = '" & strName & "' And" & vbNewLine & _
                        "             Upper(B.所有者) = '" & strActualOwer & "' And Upper(B.权限) ='" & strTmpPrivs & "')"
    
            GetUpdateSQL = strSQL
        End If
    Else
        strErr = "未知错误,对象实际所有者：" & strActualOwer & "，类型：" & strType
    End If
End Function

Private Sub Form_Resize()
    If Me.Height < 6000 Then Me.Height = 6000
    If Me.Width < 5000 Then Me.Width = 5000
    fraCheck.Width = Me.ScaleWidth - fraCheck.Left * 2.5
    fraCheck.Height = Me.ScaleHeight - fraCheck.Top - stbThis.Height - 60
    cmdRepair.Top = fraCheck.Height - 120 - cmdRepair.Height
    cmdExit.Top = cmdRepair.Top
    cmdExit.Left = fraCheck.Width - cmdExit.Width - 120
    cmdRepair.Left = cmdExit.Left - 60 - cmdRepair.Width
    fraSplit(1).Top = cmdExit.Top - 135
    vsCheckResult.Height = fraSplit(1).Top - vsCheckResult.Top - 30
    vsCheckResult.Width = fraCheck.Width - vsCheckResult.Left * 2
    fraSplit(1).Width = fraCheck.Width + 100
    fraSplit(0).Width = fraCheck.Width + 100
End Sub

Private Sub vsCheckResult_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

