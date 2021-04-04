VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm保险项目选择北京 
   AutoRedraw      =   -1  'True
   Caption         =   "医保项目选择"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   Icon            =   "frm保险项目选择北京.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7845
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   7845
      TabIndex        =   4
      Top             =   4350
      Width           =   7845
      Begin VB.CommandButton cmdRequery 
         Caption         =   "更新明细"
         Height          =   350
         Left            =   3900
         TabIndex        =   11
         Top             =   150
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "打印列表"
         Height          =   350
         Left            =   2790
         TabIndex        =   9
         Top             =   150
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   6
         Top             =   175
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   6660
         TabIndex        =   8
         Top             =   150
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Default         =   -1  'True
         Height          =   350
         Left            =   5400
         TabIndex        =   7
         Top             =   150
         Width           =   1100
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "明细查找(&F)"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Top             =   240
         Width           =   990
      End
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   2340
      MousePointer    =   9  'Size W E
      ScaleHeight     =   930
      ScaleWidth      =   45
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1590
      Width           =   45
   End
   Begin MSComctlLib.ListView lvwDetail 
      Height          =   4050
      Left            =   3060
      TabIndex        =   2
      Top             =   270
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   7144
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img16"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   45
      Top             =   3915
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目选择北京.frx":0E42
            Key             =   "Detail"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目选择北京.frx":1C94
            Key             =   "Class"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   4050
      Left            =   0
      TabIndex        =   10
      Top             =   270
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   7144
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin VB.Label lblClass 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "项目大类(&K)"
      Height          =   240
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   2970
   End
   Begin VB.Label lblDetail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "项目明细(&D)"
      Height          =   240
      Left            =   3060
      TabIndex        =   1
      Top             =   30
      Width           =   4710
   End
End
Attribute VB_Name = "frm保险项目选择北京"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint险类 As Integer
Private mstrCode As String
Private mstrName As String
Private mcnYB As New ADODB.Connection
Private mobjStream As TextStream
Private mobjFileSystem As New FileSystemObject
Private mstrErr As String
Private Const strFile = "C:\DOWNLOAD_ERR.LOG"
Private mErrFile As TextStream
Private mblnOK As Boolean

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsTemp As New ADODB.Recordset
    If lvwDetail.SelectedItem Is Nothing Then
        MsgBox "没有选择项目！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '返回选择项目编码
    mstrCode = lvwDetail.SelectedItem.Text
    
    mblnOK = True
    Unload Me
End Sub

Private Sub GetValueByCol(ByVal strColumnName As String, strValue As String)
    Dim lngCount As Long, lngIndex As Long

    For lngCount = 1 To lvwDetail.ColumnHeaders.Count
        If lvwDetail.ColumnHeaders(lngCount).Text = strColumnName Then
            lngIndex = lngCount
            Exit For
        End If
    Next
    
    If lngIndex > 0 Then
        strValue = lvwDetail.SelectedItem.SubItems(lngIndex - 1)
    End If
End Sub

Public Function GetCode(strCode As String, STRNAME As String, ByVal int险类 As Integer) As Boolean
'功能：获得一个收费项目的医保编码
'参数：strCode 既作为输入参数，又输出
'返回：成功返回True
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    Dim nod As Node
    
    mblnOK = False
    mstrCode = strCode
    mstrName = STRNAME
    mint险类 = int险类
    
    On Error GoTo ErrH
    
    '首先读出参数，打开连接
    gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, int险类)
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        Select Case rsTemp("参数名")
            Case "医保服务器"
                strServer = strTemp
            Case "医保用户名"
                strUser = strTemp
            Case "医保用户密码"
                strPass = strTemp
        End Select
        rsTemp.MoveNext
    Loop
    If OraDataOpen(mcnYB, strServer, strUser, strPass) = False Then
        Exit Function
    End If
    
    '按收费类别进行分类显示医保项目
    gstrSQL = "" & _
        " SELECT *" & _
        " FROM (" & _
        "     SELECT B.编码,B.名称,'0' AS 上级ID" & _
        "     FROM 指标主表 A,指标体系对照表 B" & _
        "     WHERE A.名称='收费类别' AND A.类别=B.类别" & _
        "     AND SUBSTR(B.编码,3,2)='00'" & _
        "     Union" & _
        "     SELECT B.编码,B.名称,SUBSTR(B.编码,1,2)||'00' AS 上级ID" & _
        "     FROM 指标主表 A,指标体系对照表 B" & _
        "     WHERE A.名称='收费类别' AND A.类别=B.类别" & _
        "     AND SUBSTR(B.编码,3,2)<>'00')" & _
        " ORDER BY 上级ID,编码"
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, mcnYB, adOpenStatic, adLockReadOnly
    
    tvwClass.Nodes.Clear
    Do Until rsTemp.EOF
        If rsTemp("上级id") = 0 Then
            Set nod = tvwClass.Nodes.Add(, , "K" & rsTemp("编码"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), "Class", "Class")
        Else
            Set nod = tvwClass.Nodes.Add("K" & rsTemp("上级ID"), tvwChild, "K" & rsTemp("编码"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), "Class", "Class")
        End If
        nod.Sorted = True
        rsTemp.MoveNext
    Loop
    
    tvwClass.Nodes(2).Selected = True
    Call FillList
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Call RestoreWinState(Me, App.ProductName)
    
    frm保险项目选择北京.Show vbModal, frm保险项目
    '返回值
    If mblnOK = True Then strCode = mstrCode
    GetCode = mblnOK
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub FillList()
'功能：显示当前类别下的医保明细
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem, fld As ADODB.Field
    Dim str类别代码 As String, blnColSet As Boolean
    Dim lngCol  As Long
    Dim varValue As Variant
    
    Me.MousePointer = vbHourglass
    
    On Error GoTo errHandle
    With tvwClass.SelectedItem
        str类别代码 = Mid(.Text, 2, InStr(.Text, "】") - 2)
    End With
    
    rsTemp.CursorLocation = adUseClient
    '暂时让列表不能刷新
    LockWindowUpdate lvwDetail.hwnd
    lvwDetail.ListItems.Clear
    
    If str类别代码 < "0400" Then
        '当前选择是的一个药品类别
        Me.lvwDetail.Tag = "Y"
        gstrSQL = "" & _
            " Select A.编码,A.类目,A.名称,A.助记码,A.剂量单位 AS 单位,B.名称 As 特殊病,H.名称 AS 项目等级,A.标准单价,A.自付比例,0 限价," & _
            " C.名称 AS 处方药,F.名称 AS 剂型,A.用法,A.日常规用量,D.名称 AS 药品分类,G.名称 AS 产地,E.名称 AS 使用限制等级,A.备注,A.生效日期" & _
            " From 药品目录 A," & _
            "      (Select B.编码,B.名称" & _
            "      FROM 指标主表 A,指标体系对照表 B" & _
            "      Where A.名称='特殊用药标识' and A.类别=B.类别) B," & _
            "      (Select B.编码,B.名称" & _
            "      FROM 指标主表 A,指标体系对照表 B" & _
            "      Where A.名称='处方药标志' and A.类别=B.类别) C," & _
            "      (Select B.编码,B.名称" & _
            "      FROM 指标主表 A,指标体系对照表 B" & _
            "      Where A.名称='药品分类' and A.类别=B.类别) D," & _
            "      (Select B.编码,B.名称" & _
            "      FROM 指标主表 A,指标体系对照表 B" & _
            "      Where A.名称='使用限制等级' and A.类别=B.类别) E,"
        gstrSQL = gstrSQL & _
            "      (Select B.编码,B.名称" & _
            "      FROM 指标主表 A,指标体系对照表 B" & _
            "      Where A.名称='剂型' and A.类别=B.类别) F," & _
            "      (Select B.编码,B.名称" & _
            "      FROM 指标主表 A,指标体系对照表 B" & _
            "      Where A.名称='产地' and A.类别=B.类别) G," & _
            "      (Select B.编码,B.名称" & _
            "      FROM 指标主表 A,指标体系对照表 B" & _
            "      Where A.名称='收费项目等级' and A.类别=B.类别) H" & _
            " Where A.特殊病 =B.编码(+) And A.处方药=C.编码(+) And A.药品分类 =D.编码(+)" & _
            " And A.使用限制等级=E.编码(+) And A.剂型=F.编码(+) And A.产地=G.编码(+) AND A.药品等级=H.编码(+)" & _
            " And A.收费类别='" & str类别代码 & "'"
    Else
        '当前选择是的一个诊疗类别
        Me.lvwDetail.Tag = "Z"
        gstrSQL = "" & _
            " Select A.编码,A.名称,A.助记码,A.单位,B.名称 AS 特殊病,C.名称 AS 项目等级,A.标准单价,A.自付比例,A.限价,A.备注,A.生效日期" & _
            "      From 诊疗目录 A," & _
            "      (Select B.编码,B.名称" & _
            "      FROM 指标主表 A,指标体系对照表 B" & _
            "      Where A.名称='特殊用药标识' and A.类别=B.类别) B," & _
            "      (Select B.编码,B.名称" & _
            "      FROM 指标主表 A,指标体系对照表 B" & _
            "      Where A.名称='收费项目等级' and A.类别=B.类别) C" & _
            " Where A.特殊病 =B.编码(+) And A.项目等级=C.编码(+)" & _
            " And A.收费类别='" & str类别代码 & "'"
    End If
    rsTemp.Open gstrSQL, mcnYB, adOpenStatic, adLockReadOnly
    
    If lvwDetail.ColumnHeaders.Count <> rsTemp.Fields.Count Then
        '重新处理表头
        blnColSet = True
        lvwDetail.ColumnHeaders.Clear
        For Each fld In rsTemp.Fields
            lvwDetail.ColumnHeaders.Add , , fld.Name, 1000
        Next
    End If
        
    Do Until rsTemp.EOF
        Set lst = lvwDetail.ListItems.Add(, "K" & rsTemp("编码"), rsTemp("编码"), "Detail", "Detail")
        
        '根据ListView的列名从数据库取数
        For lngCol = 2 To lvwDetail.ColumnHeaders.Count
            varValue = rsTemp(lvwDetail.ColumnHeaders(lngCol).Text).Value
            lst.SubItems(lngCol - 1) = IIf(IsNull(varValue), "", varValue)
        Next
        rsTemp.MoveNext
    Loop
    If blnColSet = True Then
        '重新对列进行了处理
        If lvwDetail.ListItems.Count > 0 Then Call zlControl.LvwSetColWidth(lvwDetail)
    End If
    LockWindowUpdate 0
    Me.MousePointer = vbDefault
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    LockWindowUpdate 0
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdPrint_Click()
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As New zlPrintLvw
    
    
    objPrint.Title.Text = "保险项目"
    Set objPrint.Body.objData = lvwDetail
    objPrint.UnderAppItems.Add "医保大类：" & tvwClass.SelectedItem.Text
    objPrint.BelowAppItems.Add "打印人：" & gstrUserName
    objPrint.BelowAppItems.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    Select Case zlPrintAsk(objPrint)
        Case 1
             zlPrintOrViewLvw objPrint, 1
        Case 2
            zlPrintOrViewLvw objPrint, 2
        Case 3
            zlPrintOrViewLvw objPrint, 3
    End Select

End Sub

Private Sub cmdRequery_Click()
    Dim blnExist As Boolean                 '目录是否存在
    Dim blnTrans As Boolean, bln全量 As Boolean, blnSuccess As Boolean
    Dim str医保项目目录 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '选择目录，根据是全量更新还是增量更新，自动选择对应的文件更新
    
    '判断更新目录是否已设定
    gstrSQL = "Select 参数值 From 保险参数 Where 参数名='医保项目目录'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取医保项目目录")
    If Not rsTemp.EOF Then
        If Nvl(rsTemp!参数值) <> "" Then
            blnExist = True
            str医保项目目录 = rsTemp!参数值
        End If
    End If
    If blnExist = False Then
        MsgBox "请先在保险类别的参数设置中，设置医保项目目录！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '初始化医保中间库连接
    If Not 医保初始化_北京(True) Then Exit Sub
    
    '判断更新模式
    bln全量 = True
    gstrSQL = "Select 1 From 药品目录 Where Rownum<2"
    If rsTemp.State = 1 Then rsTemp.Close
    rsTemp.Open gstrSQL, mcnYB
    If Not rsTemp.EOF Then
        bln全量 = (MsgBox("发现已存在医保项目，本次将要进行增量更新，点击是表示增量更新，点击否表示全量更新", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo)
    End If
    
    '删除日志文件
    Set mobjFileSystem = New FileSystemObject
    If mobjFileSystem.FileExists(strFile) Then mobjFileSystem.DeleteFile (strFile)
    Set mErrFile = mobjFileSystem.CreateTextFile(strFile)
    
    mcnYB.BeginTrans
    blnTrans = True
    '准备进行药品目录的下载
    blnSuccess = 药品目录_北京(bln全量, str医保项目目录)
    If blnSuccess = False Then
        mErrFile.Close
        Call zlCommFun.StopFlash
        mcnYB.RollbackTrans
        Me.Caption = "医保项目选择"
        Exit Sub
    End If
    '准备进行诊疗目录的下载
    blnSuccess = 诊疗目录_北京(bln全量, str医保项目目录)
    If blnSuccess = False Then
        mErrFile.Close
        Call zlCommFun.StopFlash
        mcnYB.RollbackTrans
        Me.Caption = "医保项目选择"
        Exit Sub
    End If
    
    mcnYB.CommitTrans
    mErrFile.Close
    Call zlCommFun.StopFlash
    Me.Caption = "医保项目选择"
    blnTrans = False
    
    '重新装入明细
    Call FillList
    MousePointer = vbDefault
    Me.Caption = "医保项目选择"
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    If blnTrans Then mcnYB.RollbackTrans
    mErrFile.Close
    Call zlCommFun.StopFlash
End Sub

Private Function 药品目录_北京(ByVal bln全量 As Boolean, ByVal str目录 As String) As Boolean
    Const str药品目录_全量 As String = "ypml_all.txt"
    Const str药品别名_全量 As String = "ypymml_all.txt"
    Const str药品目录_增量 As String = "ypml.txt"
    Const str药品别名_增量 As String = "ypymml.txt"
    Const int编码 As Integer = 1
    Const int名称 As Integer = 2
    Const strSplit As String = "|"
    Dim strProcedure As String
    Dim arrLine
    Dim lngCol As Long, lngCols As Long
    Dim blnExist As Boolean
    Dim strText As String
    Dim strFileName As String
    Dim strExecute As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    On Error GoTo errHand
    '先处理药品目录
    '解析文件
    strFileName = str目录 & "\" & IIf(bln全量, str药品目录_全量, str药品目录_增量)
    If Not objFileSystem.FileExists(strFileName) Then
        blnExist = False
    Else
        '打开文件进行更新操作
        blnExist = True
        Set objStream = objFileSystem.OpenTextFile(strFileName)
    End If
    
    Call zlCommFun.ShowFlash("正在提取药品项目明细", Me)
    '更新药品目录ForBJ
    If bln全量 Then
        gstrSQL = "zl_药品目录_Clear"
        mcnYB.Execute gstrSQL, , adCmdStoredProc
    End If
    
    strProcedure = "zl_药品目录_"
    If blnExist Then
        Do While Not objStream.AtEndOfStream
            arrLine = Split(objStream.ReadLine, strSplit)
            lngCols = UBound(arrLine)
            gstrSQL = ""
            strExecute = strProcedure & IIf(arrLine(lngCols) = "0", "INSERT", IIf(arrLine(lngCols) = "1", "DELETE", "UPDATE"))
            If arrLine(lngCols) = "1" Then
                gstrSQL = strExecute & "('" & arrLine(int编码) & "')"
            Else
                For lngCol = 0 To lngCols - 1 '最后一个是操作代码（0-新增;1-删除;2-修改）
                    If lngCol <> lngCols - 1 Then
                        gstrSQL = gstrSQL & ",'" & UCase(Replace(arrLine(lngCol), "'", "")) & "'"
                    Else
                        strText = arrLine(lngCol)
                        strText = Mid(strText, 1, 4) & "-" & Mid(strText, 5, 2) & "-" & Mid(strText, 7, 2)
                        gstrSQL = gstrSQL & ",to_Date('" & strText & "','yyyy-MM-dd')"
                    End If
                Next
                gstrSQL = Mid(gstrSQL, 2)
                gstrSQL = strExecute & "(" & gstrSQL & ")"
            End If
            mcnYB.Execute gstrSQL, , adCmdStoredProc
            Me.Caption = "医保项目选择" & Space(10) & "正在更新第" & objStream.Line - 1 & "个药品项目"
        Loop
        objStream.Close
    End If
    
    '再处理药品别名
    '解析文件
    strProcedure = "zl_药品别名_"
    strFileName = str目录 & "\" & IIf(bln全量, str药品别名_全量, str药品别名_增量)
    If Not objFileSystem.FileExists(strFileName) Then
        blnExist = False
    Else
        '打开文件进行更新操作
        blnExist = True
        Set objStream = objFileSystem.OpenTextFile(strFileName)
    End If
    
    Call zlCommFun.ShowFlash("正在提取药品别名", Me)
    '更新药品目录ForBJ
    If bln全量 Then
        gstrSQL = "zl_药品别名_Clear"
        mcnYB.Execute gstrSQL, , adCmdStoredProc
    End If
    
    If blnExist Then
        Do While Not objStream.AtEndOfStream
            arrLine = Split(objStream.ReadLine, strSplit)
            lngCols = UBound(arrLine)
            gstrSQL = ""
            strExecute = strProcedure & IIf(arrLine(lngCols) = "0", "INSERT", IIf(arrLine(lngCols) = "1", "DELETE", "UPDATE"))
            If arrLine(lngCols) = "1" Then
                gstrSQL = strExecute & "('" & arrLine(int编码) & "','" & arrLine(int名称) & "')"
            Else
                For lngCol = 0 To lngCols - 1 '最后一个是操作代码（0-新增;1-删除;2-修改）
                    gstrSQL = gstrSQL & ",'" & arrLine(lngCol) & "'"
                Next
                gstrSQL = Mid(gstrSQL, 2)
                gstrSQL = strExecute & "(" & gstrSQL & ")"
            End If
            mcnYB.Execute gstrSQL, , adCmdStoredProc
            Me.Caption = "医保项目选择" & Space(10) & "正在更新第" & objStream.Line - 1 & "个药品别名项目"
        Loop
        objStream.Close
    End If
    
    药品目录_北京 = True
    Exit Function
errHand:
    mstrErr = "当前行:" & objStream.Line - 1 & "错误号:" & Err.Number & "错误信息:" & Err.Description
    mErrFile.WriteLine mstrErr
    Resume Next
End Function

Private Function 诊疗目录_北京(ByVal bln全量 As Boolean, ByVal str目录 As String) As Boolean
    Const str诊疗目录_全量 As String = "zlml_all.txt"
    Const str服务设施目录_全量 As String = "fwssml_all.txt"
    Const str诊疗目录_增量 As String = "zlml.txt"
    Const str服务设施目录_增量 As String = "fwssml.txt"
    Const int编码 As Integer = 1
    Const int名称 As Integer = 2
    Const strSplit As String = "|"
    Const strProcedure As String = "zl_诊疗目录_"
    Dim arrLine
    Dim lngCol As Long, lngCols As Long
    Dim blnExist As Boolean
    Dim strText As String
    Dim strFileName As String
    Dim strExecute As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    On Error GoTo errHand
    '解析文件
    strFileName = str目录 & "\" & IIf(bln全量, str诊疗目录_全量, str诊疗目录_增量)
    If Not objFileSystem.FileExists(strFileName) Then
        blnExist = False
    Else
        '打开文件进行更新操作
        blnExist = True
        Set objStream = objFileSystem.OpenTextFile(strFileName)
    End If
    
    Call zlCommFun.ShowFlash("正在提取诊疗项目明细", Me)
    '更新诊疗目录ForBJ
    If bln全量 Then
        gstrSQL = "zl_诊疗目录_Clear"
        mcnYB.Execute gstrSQL, , adCmdStoredProc
    End If
    
    If blnExist Then
        Do While Not objStream.AtEndOfStream
            arrLine = Split(objStream.ReadLine, strSplit)
            lngCols = UBound(arrLine)
            gstrSQL = ""
            strExecute = strProcedure & IIf(arrLine(lngCols) = "0", "INSERT", IIf(arrLine(lngCols) = "1", "DELETE", "UPDATE"))
            If arrLine(lngCols) = "1" Then
                gstrSQL = strExecute & "('" & arrLine(int编码) & "')"
            Else
                For lngCol = 0 To lngCols - 1 '最后一个是操作代码（0-新增;1-删除;2-修改）
                    If lngCol <> lngCols - 1 Then
                        gstrSQL = gstrSQL & ",'" & Replace(arrLine(lngCol), "'", "") & "'"
                    Else
                        strText = arrLine(lngCol)
                        strText = Mid(strText, 1, 4) & "-" & Mid(strText, 5, 2) & "-" & Mid(strText, 7, 2)
                        gstrSQL = gstrSQL & ",to_Date('" & strText & "','yyyy-MM-dd')"
                    End If
                Next
                gstrSQL = Mid(gstrSQL, 2)
                gstrSQL = strExecute & "(" & gstrSQL & ")"
            End If
            mcnYB.Execute gstrSQL, , adCmdStoredProc
            Me.Caption = "医保项目选择" & Space(10) & "正在更新第" & objStream.Line - 1 & "个诊疗项目"
        Loop
        objStream.Close
    End If
    
    '解析文件(因服务设施与诊疗保存在一张表中，因此服务设施导入时，可以不清除)
    strFileName = str目录 & "\" & IIf(bln全量, str服务设施目录_全量, str服务设施目录_增量)
    If Not objFileSystem.FileExists(strFileName) Then
        blnExist = False
    Else
        '打开文件进行更新操作
        blnExist = True
        Set objStream = objFileSystem.OpenTextFile(strFileName)
    End If
    
    Call zlCommFun.ShowFlash("正在提取服务设施项目明细", Me)
    '更新服务设施目录ForBJ
    If blnExist Then
        Do While Not objStream.AtEndOfStream
            arrLine = Split(objStream.ReadLine, strSplit)
            lngCols = UBound(arrLine)
            gstrSQL = ""
            strExecute = strProcedure & IIf(arrLine(lngCols) = "0", "INSERT", IIf(arrLine(lngCols) = "1", "DELETE", "UPDATE"))
            If arrLine(lngCols) = "1" Then
                gstrSQL = strExecute & "('" & arrLine(int编码) & "')"
            Else
                For lngCol = 0 To lngCols - 1 '最后一个是操作代码（0-新增;1-删除;2-修改）
                    If lngCol <> lngCols - 1 Then
                        gstrSQL = gstrSQL & ",'" & Replace(arrLine(lngCol), "'", "") & "'"
                    Else
                        strText = arrLine(lngCol)
                        strText = Mid(strText, 1, 4) & "-" & Mid(strText, 5, 2) & "-" & Mid(strText, 7, 2)
                        gstrSQL = gstrSQL & ",to_Date('" & strText & "','yyyy-MM-dd')"
                    End If
                Next
                gstrSQL = Mid(gstrSQL, 2)
                gstrSQL = strExecute & "(" & gstrSQL & ")"
            End If
            mcnYB.Execute gstrSQL, , adCmdStoredProc
            Me.Caption = "医保项目选择" & Space(10) & "正在更新第" & objStream.Line - 1 & "个服务设施项目"
        Loop
        objStream.Close
    End If
    诊疗目录_北京 = True
    Exit Function
errHand:
    mstrErr = "当前行:" & objStream.Line - 1 & "错误号:" & Err.Number & "错误信息:" & Err.Description
    mErrFile.WriteLine mstrErr
    Resume Next
End Function

Private Sub Form_Load()
    cmdRequery.Visible = True
End Sub

Private Sub Form_Resize()
    lblClass.Top = 0: lblClass.Left = 0: lblClass.Width = tvwClass.Width
    
    On Error Resume Next
    
    tvwClass.Left = 0: tvwClass.Top = lblClass.Top + lblClass.Height
    tvwClass.Height = Me.ScaleHeight - lblClass.Height - picCmd.Height
    
    picSplit.Top = tvwClass.Top
    picSplit.Left = tvwClass.Left + tvwClass.Width
    picSplit.Height = tvwClass.Height
    
    lblDetail.Top = lblClass.Top
    If tvwClass.Visible = True Then
        lblDetail.Left = picSplit.Left + picSplit.Width
    Else
        lblDetail.Left = 0
    End If
    lblDetail.Width = Me.ScaleWidth - lblDetail.Left
    
    lvwDetail.Top = tvwClass.Top
    lvwDetail.Left = lblDetail.Left
    lvwDetail.Width = lblDetail.Width
    lvwDetail.Height = tvwClass.Height
End Sub

Private Sub picCmd_Resize()
    cmdCancel.Left = picCmd.ScaleWidth - cmdCancel.Width * 1.4
    cmdOK.Left = cmdCancel.Left - cmdOK.Width * 1.25
    cmdPrint.Top = cmdOK.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwDetail_DblClick()
    cmdOK_Click
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If tvwClass.Width + x < 1000 Or lvwDetail.Width - x < 1000 Then Exit Sub
        picSplit.Left = picSplit.Left + x
        lblClass.Width = lblClass.Width + x
        tvwClass.Width = tvwClass.Width + x
        
        lblDetail.Left = lblDetail.Left + x
        lblDetail.Width = lblDetail.Width - x
        
        lvwDetail.Left = lvwDetail.Left + x
        lvwDetail.Width = lvwDetail.Width - x
    End If
End Sub

Private Sub lvwdetail_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwDetail.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwDetail.SortOrder = lvwDescending
    Else
        lvwDetail.SortOrder = lvwAscending
    End If
    lvwDetail.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvwDetail.SelectedItem Is Nothing Then lvwDetail.SelectedItem.EnsureVisible
End Sub

Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    Call FillList
End Sub

Private Sub txtFind_Change()
'功能：根据用户输入的内容查找匹配的内容
    Dim lst As ListItem, lngIndex As Long, lngSubItems As Long
    Dim strFind As String
    
    strFind = UCase(Trim(txtFind.Text))
    If strFind = "" Then Exit Sub
    If lvwDetail.ListItems.Count = 0 Then Exit Sub

    Set lst = lvwDetail.FindItem(strFind, lvwText, , lvwPartial)
    If Not lst Is Nothing Then
        lst.Selected = True
        lst.EnsureVisible
    Else
        '非文本不能做到部分匹配
        lngSubItems = lvwDetail.ColumnHeaders.Count - 1
        For Each lst In lvwDetail.ListItems
            For lngIndex = 1 To lngSubItems
                If lst.SubItems(lngIndex) Like strFind & "*" Then
                    lst.Selected = True
                    lst.EnsureVisible
                    Exit Sub
                End If
            Next
            
        Next
    End If
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Function ReplaceStr(ByVal StrInput As String) As String
    ReplaceStr = Trim(Replace(StrInput, "'", ""))
    ReplaceStr = Replace(ReplaceStr, """", "")
End Function
