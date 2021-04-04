VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm保险项目选择广元旺苍 
   AutoRedraw      =   -1  'True
   Caption         =   "医保项目选择"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   Icon            =   "frm保险项目选择广元旺苍.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7845
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   3690
      Top             =   2220
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
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
      TabIndex        =   5
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
         TabIndex        =   10
         Top             =   150
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   7
         Top             =   175
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   6660
         TabIndex        =   9
         Top             =   150
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Default         =   -1  'True
         Height          =   350
         Left            =   5400
         TabIndex        =   8
         Top             =   150
         Width           =   1100
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "明细查找(&F)"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   6
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1590
      Width           =   45
   End
   Begin MSComctlLib.ListView lvwDetail 
      Height          =   4050
      Left            =   3060
      TabIndex        =   3
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
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "编码"
         Object.Width           =   2752
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "名称"
         Object.Width           =   2434
      EndProperty
   End
   Begin MSComctlLib.ListView lvwClass 
      Height          =   3990
      Left            =   15
      TabIndex        =   1
      Top             =   285
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   7038
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "编码"
         Object.Width           =   1535
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "名称"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   15
      Top             =   525
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
            Picture         =   "frm保险项目选择广元旺苍.frx":0E42
            Key             =   "Detail"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目选择广元旺苍.frx":1C94
            Key             =   "Class"
         EndProperty
      EndProperty
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
      TabIndex        =   2
      Top             =   30
      Width           =   4710
   End
End
Attribute VB_Name = "frm保险项目选择广元旺苍"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrCode As String '入出参数,医保项目DetailCode
Private mrsDetail As ADODB.Recordset
Private mblnOK As Boolean
Private mint中心 As Integer
Private mint适用地区 As Integer '沈阳专用；0表示其他地区，1表示长春（允许删除已审核的项目）
Private mint险类 As Integer
Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If lvwDetail.SelectedItem Is Nothing Then
        MsgBox "没有选择项目！", vbInformation, gstrSysName
        Exit Sub
    End If
    '返回选择项目编码
    mstrCode = Mid(lvwDetail.SelectedItem.Key, 2)
    mblnOK = True
    Unload Me
End Sub

Public Function GetCode(strCode As String, ByVal int中心 As Integer, ByVal int险类 As Integer) As Boolean
'功能：获得一个收费项目的医保编码
'参数：strCode 既作为输入参数，又输出
'返回：成功返回True
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, objItem As ListItem
    
    mblnOK = False
    mint中心 = int中心
    
    On Error GoTo ErrH
    
    Set rsTmp = New ADODB.Recordset
    Set mrsDetail = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    mrsDetail.CursorLocation = adUseClient
    mint险类 = int险类
    
    gstrSQL = "Select 编码 AS CODE,名称 AS NAME From 保险支付大类 where 险类=[1] order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "保险项目选择", mint险类)
        
    If mint险类 = TYPE_成都德阳 Then
        gstrSQL = "Select 大类编码 as CLASSCODE,编码 AS CODE,名称 AS NAME ,简码," & _
                  "substr(附注,1,instr(附注," & "'|'" & ",1,1)-1) As 费用类别," & _
                  "substr(附注,instr(附注," & "'|'" & ",1,1)+1,instr(附注," & "'|'" & ",1,2)-1-instr(附注," & "'|'" & ",1,1)) As 自付比例," & _
                  "substr(附注,instr(附注," & "'|'" & ",1,2)+1,instr(附注," & "'|'" & ",1,3)-1-instr(附注," & "'|'" & ",1,2)) As 费用项目," & _
                  "substr(附注,instr(附注," & "'|'" & ",1,3)+1,instr(附注," & "'|'" & ",1,4)-1-instr(附注," & "'|'" & ",1,3)) As 限制范围," & _
                  "substr(附注,instr(附注," & "'|'" & ",1,4)+1,instr(附注," & "'|'" & ",1,5)-1-instr(附注," & "'|'" & ",1,4)) As 备注" & _
                  " from 医保收费项目_德阳 where 险类=[1] and 中心=[2] order by 大类编码,编码"
    Else
        gstrSQL = "Select 大类编码 as CLASSCODE,编码 AS CODE,名称 AS NAME ,简码," & _
                  "substr(附注,1,instr(附注," & "'|'" & ",1,1)-1) As 费用类别," & _
                  "substr(附注,instr(附注," & "'|'" & ",1,1)+1,instr(附注," & "'|'" & ",1,2)-1-instr(附注," & "'|'" & ",1,1)) As 自付比例," & _
                  "substr(附注,instr(附注," & "'|'" & ",1,2)+1,instr(附注," & "'|'" & ",1,3)-1-instr(附注," & "'|'" & ",1,2)) As 费用项目," & _
                  "substr(附注,instr(附注," & "'|'" & ",1,3)+1,instr(附注," & "'|'" & ",1,4)-1-instr(附注," & "'|'" & ",1,3)) As 限制范围," & _
                  "substr(附注,instr(附注," & "'|'" & ",1,4)+1,instr(附注," & "'|'" & ",1,5)-1-instr(附注," & "'|'" & ",1,4)) As 备注" & _
                  " from 医保收费项目 where 险类=[1] and 中心=[2] order by 大类编码,编码"
    End If
    Set mrsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "保险项目选择", mint险类, int中心)
    
    '为明细增加多余显示的列
    Dim fld As ADODB.Field
    For Each fld In mrsDetail.Fields
        If fld.Name <> "CLASSCODE" And fld.Name <> "NAME" And fld.Name <> "CODE" Then
            If fld.Name <> "附注" Then
                lvwDetail.ColumnHeaders.Add , , fld.Name, 1000
            End If
        End If
    Next
    
    '初始化大类
    If rsTmp.State = adStateOpen Then
        If Not rsTmp.EOF Then
            lvwClass.ListItems.Clear
            For i = 1 To rsTmp.RecordCount
                Set objItem = lvwClass.ListItems.Add(, "_" & rsTmp("CODE"), rsTmp("CODE"), , "Class")
                objItem.SubItems(1) = IIf(IsNull(rsTmp("NAME")), "", rsTmp("NAME"))
                rsTmp.MoveNext
            Next
        End If
    Else
        '这种情况下是没有大类的
        lblClass.Visible = False
        lvwClass.Visible = False
        picSplit.Visible = False
        Call lvwClass.ListItems.Add(, "_1", "1", , "Class")
    End If
    cmdRequery.Visible = True
    
    If Not mrsDetail.EOF Then
       If mstrCode <> "" Then
            '查找大类编码并定位
            mrsDetail.Filter = "CODE Like '" & UCase(mstrCode) & "%'"
            If Not mrsDetail.EOF Then
                lvwClass.ListItems("_" & mrsDetail("CLASSCODE")).Selected = True
            ElseIf lvwClass.ListItems.Count > 0 Then
                lvwClass.ListItems(1).Selected = True
            End If
            Call lvwClass_ItemClick(lvwClass.SelectedItem)
            lvwClass.SelectedItem.EnsureVisible
        Else
            If lvwClass.ListItems.Count > 0 Then
                lvwClass.ListItems(1).Selected = True
            End If
            Call lvwClass_ItemClick(lvwClass.SelectedItem)
        End If
        
    End If
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Call RestoreWinState(Me, App.ProductName)
    
    
    frm保险项目选择广元旺苍.Show 1
    '返回值
    If mblnOK = True Then
        strCode = mstrCode
    End If
    GetCode = mblnOK
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdPrint_Click()
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As New zlPrintLvw
    
    
    objPrint.Title.Text = "保险项目"
    Set objPrint.Body.objData = lvwDetail
    objPrint.UnderAppItems.Add "医保大类：" & lvwClass.SelectedItem.Text
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
    Dim str费用类型 As String
    Dim str附注 As String
    Dim rsTemp As New ADODB.Recordset
    Dim bln全量 As Boolean
    Dim blnReturn As Boolean
    
    If MsgBox("本操作可能会花比较长的时间，是否继续？" & vbCrLf & vbCrLf & "另外注意，本操作只更新医保项目明细，而不包括对应关系。", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    MousePointer = vbHourglass
      With rsTemp
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .Fields.Append "CLASSCODE", adVarChar, 6   '大类编码
        .Fields.Append "CODE", adVarChar, 20       '编号
        .Fields.Append "NAME", adVarChar, 40       '名称
        .Fields.Append "PY", adVarChar, 10         '简码
        .Fields.Append "FYLB", adVarChar, 10       '费用类别
        .Fields.Append "ZFBL", adVarChar, 6       '自付比例
        .Fields.Append "FYXM", adVarChar, 14       '费用项目
        .Fields.Append "XZFW", adVarChar, 100      '限制范围
        .Fields.Append "BZ", adVarChar, 100        '备注
        .Fields.Append "MEMO", adVarChar, 500      '附注
        .Open
      End With
      
    bln全量 = True
    Me.Caption = "医保项目选择（正在读取从文件或网络读取保险项目明细，请稍候......）"
    If MsgBox("是否清空原来的医保项目？", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) <> vbYes Then
        bln全量 = False
    End If
    blnReturn = 医保项目_成都内江(rsTemp)
    
    If blnReturn = False Then
        MousePointer = vbDefault
        Exit Sub
    End If
    
    Me.Caption = "医保项目选择（正在更新医保项目......）"
    If mint险类 = TYPE_成都德阳 Then
        gcnOracle_成都德阳.BeginTrans
    ElseIf mint险类 = TYPE_南充阆中 Then
        gcnOracle_南充阆中.BeginTrans
    Else
        gcnOracle_广元旺苍.BeginTrans
    End If
    On Error GoTo errHandle
    If bln全量 Then
        gstrSQL = "zl_医保收费项目_Clear(" & mint险类 & "," & mint中心 & ")"
        If mint险类 = TYPE_成都德阳 Then
            Call ExecuteProcedure_成都德阳("医保项目选择")
        ElseIf mint险类 = TYPE_南充阆中 Then
            Call ExecuteProcedure_南充阆中("医保项目选择")
        Else
            Call ExecuteProcedure_广元旺苍("医保项目选择")
        End If
    End If
    
    '更新保险项目
    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    Do Until rsTemp.EOF
        rsTemp("FYLB") = IIf(Trim(rsTemp("FYLB")) = "", "无", rsTemp("FYLB"))
        rsTemp("ZFBL") = IIf(Trim(rsTemp("ZFBL")) = "", "无", rsTemp("ZFBL"))
        rsTemp("FYXM") = IIf(Trim(rsTemp("FYXM")) = "", "无", rsTemp("FYXM"))
        rsTemp("XZFW") = IIf(Trim(rsTemp("XZFW")) = "", "无", rsTemp("XZFW"))
        rsTemp("BZ") = IIf(Trim(rsTemp("BZ")) = "", "无", rsTemp("BZ"))
        rsTemp("MEMO") = IIf(Trim(rsTemp("MEMO")) = "", "无", rsTemp("MEMO"))
        str附注 = rsTemp("FYLB") & "|" & rsTemp("ZFBL") & "|" & rsTemp("FYXM") & "|" & rsTemp("XZFW") & "|" & rsTemp("BZ") & "|" & rsTemp("MEMO")

        '插入保险项目
        gstrSQL = "zl_医保收费项目_Insert(" & mint险类 & "," & mint中心 & ",'" & rsTemp("CODE") & "','" & ToVarchar(rsTemp("NAME"), 40) & _
            "','" & ToVarchar(rsTemp("PY"), 10) & "','" & ToVarchar(rsTemp("CLASSCODE"), 6) & "','" & ToVarchar(str附注, 500) & "')"
        If mint险类 = TYPE_成都德阳 Then
            ExecuteProcedure_成都德阳 ("更新医保项目")
        ElseIf mint险类 = TYPE_南充阆中 Then
            Call ExecuteProcedure_南充阆中("更新医保项目")
        Else
            Call ExecuteProcedure_广元旺苍("更新医保项目")
        End If
        Me.Caption = "医保项目选择（正在更新医保项目，已插入" & rsTemp.AbsolutePosition & "条记录）"
        rsTemp.MoveNext
    Loop
    
    If mint险类 = TYPE_成都德阳 Then
        gcnOracle_成都德阳.CommitTrans
    ElseIf mint险类 = TYPE_南充阆中 Then
        gcnOracle_南充阆中.CommitTrans
    Else
        gcnOracle_广元旺苍.CommitTrans
    End If
    '重新装入明细
    mrsDetail.Requery
    Call lvwClass_ItemClick(lvwClass.SelectedItem)
    MousePointer = vbDefault
    Me.Caption = "医保项目选择"
    MsgBox "更新完成。", vbInformation, gstrSysName
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    If mint险类 = TYPE_成都德阳 Then
        gcnOracle_成都德阳.RollbackTrans
    ElseIf mint险类 = TYPE_南充阆中 Then
        gcnOracle_南充阆中.RollbackTrans
    Else
        gcnOracle_广元旺苍.RollbackTrans
    End If
    MousePointer = vbDefault
End Sub
Private Function 医保项目_成都内江(ByVal rsTemp As ADODB.Recordset) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:导入数据
    '--入参数:
    '--出参数:
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Const COL_项目种类   As Long = 1
    Const COL_编码 As Long = 2
    Const COL_名称 As Long = 3
    Const COL_简码  As Long = 4
    Const COL_费用类别  As Long = 5
    Const COL_自付比例  As Long = 6
    Const COL_费用项目  As Long = 7
    Const COL_限制范围  As Long = 8
    Const COL_备注  As Long = 9
    Const COL_大类  As Long = 10
    Err = 0
    On Error GoTo errHand:
    Dim ObjExcel As Object, ObjCell As Object, strFile As String, strValue As String
    
    '选择指定文件
    On Error Resume Next
    Err = 0
    With dlg
        .Filter = "EXCEL文件(*.xls)|*.xls"
        .flags = cdlOFNFileMustExist Or cdlOFNLongNames
        .ShowOpen
        If Err <> 0 Then Exit Function
        strFile = .FileName
    End With
    
    '创建EXCEL对象
    On Error Resume Next
    Err = 0
    Set ObjExcel = CreateObject("Excel.Application")
    If Err <> 0 Then
        MsgBox "EXCEL未正确安装，请正确安装EXCEL中文版后再运行！", vbInformation, gstrSysName
        Exit Function
    End If
    
    On Error GoTo errHand:
    Me.Caption = "医保项目选择（正在从EXCEL文件中提取数据......）"
    Dim strCode As String
    
    '取EXCEL文件的数据
    With ObjExcel
        .Workbooks.Open strFile
        
        '取各列的值
        Dim lngRow As Long
        lngRow = 2
        Do While True
            If .ActiveSheet.Cells(lngRow, COL_编码) <> "" Then
                rsTemp.AddNew
                rsTemp("Code") = Mid(Trim(.ActiveSheet.Cells(lngRow, COL_编码)), 1, 20)
                rsTemp("Name") = Replace(ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL_名称)), 40), "'", "")
                rsTemp("PY") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL_简码)), 10)
                strCode = Replace(ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL_项目种类)), 6), "'", "")
                strCode = Decode(strCode, "药品", 0, "治疗", 1, "诊疗", 1, "检查", 2, "其他", 3, strCode)
                rsTemp("CLASSCODE") = strCode
                If mint险类 = TYPE_南充阆中 Or TYPE_广元旺苍 Then
                    rsTemp("FYLB") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL_费用类别)), 10)
                    rsTemp("ZFBL") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL_自付比例)), 6)
                    rsTemp("FYXM") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL_费用项目)), 14)
                    rsTemp("XZFW") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL_限制范围)), 100)
                    rsTemp("BZ") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL_备注)), 100)
                End If
                
                rsTemp.Update
                Me.Caption = "医保项目选择（正在从EXCEL文件中提取数据，已获取" & rsTemp.RecordCount & "条记录）"
            Else
                Exit Do
            End If
            lngRow = lngRow + 1
        Loop
    End With
    
    '关闭EXCEL对象
    ObjExcel.quit
    Set ObjExcel = Nothing
    医保项目_成都内江 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Form_Resize()
    lblClass.Top = 0: lblClass.Left = 0: lblClass.Width = lvwClass.Width
    
    On Error Resume Next
    
    lvwClass.Left = 0: lvwClass.Top = lblClass.Top + lblClass.Height
    lvwClass.Height = Me.ScaleHeight - lblClass.Height - picCmd.Height
    
    picSplit.Top = lvwClass.Top
    picSplit.Left = lvwClass.Left + lvwClass.Width
    picSplit.Height = lvwClass.Height
    
    lblDetail.Top = lblClass.Top
    If lvwClass.Visible = True Then
        lblDetail.Left = picSplit.Left + picSplit.Width
    Else
        lblDetail.Left = 0
    End If
    lblDetail.Width = Me.ScaleWidth - lblDetail.Left
    
    lvwDetail.Top = lvwClass.Top
    lvwDetail.Left = lblDetail.Left
    lvwDetail.Width = lblDetail.Width
    lvwDetail.Height = lvwClass.Height
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
        If lvwClass.Width + x < 1000 Or lvwDetail.Width - x < 1000 Then Exit Sub
        picSplit.Left = picSplit.Left + x
        lblClass.Width = lblClass.Width + x
        lvwClass.Width = lvwClass.Width + x
        
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

Private Sub lvwclass_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwClass.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwClass.SortOrder = lvwDescending
    Else
        lvwClass.SortOrder = lvwAscending
    End If
    lvwClass.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvwClass.SelectedItem Is Nothing Then lvwClass.SelectedItem.EnsureVisible
End Sub

Private Sub lvwClass_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer, objItem As ListItem
    Dim lngCount As Long, str列 As String, bln特殊处理 As Boolean
    Dim BLNSEL As Boolean
    Dim varPart As Variant
    
    
    Me.MousePointer = vbHourglass
    lvwDetail.ListItems.Clear
    If Item Is Nothing Then Exit Sub
    
    mrsDetail.Filter = "CLASSCODE='" & Mid(Item.Key, 2) & "'"
    If Not mrsDetail.EOF Then
        For i = 1 To mrsDetail.RecordCount
            Set objItem = lvwDetail.ListItems.Add(, "_" & mrsDetail("CODE"), mrsDetail("CODE"), , "Detail")
            objItem.SubItems(1) = IIf(IsNull(mrsDetail("NAME")), "", mrsDetail("NAME"))
            objItem.Tag = mrsDetail("CLASSCODE")
            
            '显示另外的列
            With lvwDetail.ColumnHeaders
                For lngCount = 3 To lvwDetail.ColumnHeaders.Count
                    str列 = .Item(lngCount).Text
                    '没有进行特殊处理
                    objItem.SubItems(lngCount - 1) = IIf(IsNull(mrsDetail(.Item(lngCount).Text)), "", mrsDetail(.Item(lngCount).Text))
                Next
            End With
                        
            If InStr(mrsDetail("CODE"), mstrCode) > 0 And Not BLNSEL Then
                objItem.Selected = True
                BLNSEL = True
            End If
            mrsDetail.MoveNext
        Next
        If Not BLNSEL And lvwDetail.ListItems.Count > 0 Then lvwDetail.ListItems(1).Selected = True
        lvwDetail.SelectedItem.EnsureVisible
    End If
    Call zlControl.LvwSetColWidth(lvwDetail)
    Me.MousePointer = vbDefault
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
Private Sub AddRecord(rsObj As ADODB.Recordset, ByVal str编码 As String, ByVal str名称 As String, _
str简码 As String, ByVal str备注 As String, ByVal str大类 As String)
    With rsObj
        .AddNew
        !CODE = str编码
        !Name = Replace(str名称, "'", "")
        !py = Replace(str简码, "'", "")
        !Memo = Replace(str备注, "'", "")
        !ClassCode = str大类
        .Update
    End With
End Sub


