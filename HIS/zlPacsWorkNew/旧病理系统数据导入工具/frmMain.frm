VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病理历史数据导入工具"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   3645
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "退 出(&E)"
      Height          =   360
      Left            =   2040
      TabIndex        =   3
      Top             =   3840
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确 定(&O)"
      Height          =   360
      Left            =   720
      TabIndex        =   2
      Top             =   3840
      Width           =   990
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgData 
      Height          =   2775
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   3015
      _cx             =   5318
      _cy             =   4895
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
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
      Editable        =   2
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
   Begin VB.Label lblPrompt 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   3435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "新系统的检查类型"
      Height          =   195
      Left            =   1155
      TabIndex        =   5
      Top             =   600
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "对应至"
      Height          =   195
      Left            =   1560
      TabIndex        =   4
      Top             =   360
      Width           =   540
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请将旧系统的检查类型 "
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1845
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub InitVfgData()
'初始化VfgData格式
    Dim i As Integer
    
    With vfgData
        .Width = 3150
        .Height = 2700
    
        .ColWidth(0) = 300
        .ColWidth(1) = 1400
        .ColWidth(2) = 1400
        .Cols = 3
        
        .TextMatrix(0, 1) = "原系统检查类型"
        .TextMatrix(0, 2) = "新系统检查类型"
        
        .FixedCols = 2
    
    End With
    
    For i = 1 To vfgData.Rows - 1
        vfgData.TextMatrix(i, 0) = i
    Next

End Sub

Private Sub LoadOldPathologyData()
'加载老病理检查类型
    Dim strSql As String
    Dim rsOldPathData As ADODB.Recordset
    Dim i As Integer
    
    strSql = "select 名称 from 影像病理类别"
    
    Set rsOldPathData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    If Not rsOldPathData Is Nothing Then
        For i = 1 To rsOldPathData.RecordCount
            '循环遍历数据集
             vfgData.TextMatrix(i, 1) = rsOldPathData("名称").Value
             
             If Not rsOldPathData.EOF Then
                rsOldPathData.MoveNext
             End If
        Next
        
        '动态设置控件行数
        vfgData.Rows = rsOldPathData.RecordCount + 1
        '设置新系统固定检查方法
        vfgData.ColComboList(2) = "常规|冰冻|细胞|会诊|尸检|快速石蜡|"
        
        If rsOldPathData.RecordCount > 0 Then
          '设置提示信息
            With lblPrompt
                .Font.Bold = True
                .Caption = "数据库连接成功,检查类型已加载!"
            End With
            
            '固定列居中显示
            vfgData.ColAlignment(1) = flexAlignCenterCenter
            
            Exit Sub
        End If
    
        '设置提示信息
        With lblPrompt
            .Font.Bold = True
            .Caption = "数据库连接成功,但检查类型记录为空!"
        End With
    
        
    End If
End Sub

Private Sub cmdCancel_Click()
'关闭数据库后 卸载窗体
On Error GoTo errHandle

    If OraDataClose Then
        Unload Me
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdOK_Click()
On Error GoTo errHandle
    Dim strDecode As String
    Dim intCheckType As Integer
    Dim strSql As String
    Dim rsPathologyData As ADODB.Recordset
    Dim intPathDataCount As Integer
    Dim i As Integer
    
    For i = 1 To vfgData.Rows - 1
        '判断是否进行了选择
        If vfgData.TextMatrix(i, 2) = "" Then
            MsgBox "请选择对应的检查类型！"
            Exit Sub
        End If
        
        '判断新系统检查类型对应的编号
        Select Case vfgData.TextMatrix(i, 2)
            Case "常规"
                intCheckType = 0
            Case "冰冻"
                intCheckType = 1
            Case "细胞"
                intCheckType = 2
            Case "会诊"
                intCheckType = 3
            Case "尸检"
                intCheckType = 4
            Case "快速石蜡"
                intCheckType = 5
        End Select

        strDecode = strDecode + ",'" & vfgData.TextMatrix(i, 1) & "'" & ",'" & intCheckType & "'"
    Next
    
    strDecode = "decode(病理检查类别" & strDecode & ")"
    
    
    If MsgBox("您确定开始进行数据导入操作吗？(此操作不可逆,请慎重!)", vbOKCancel + vbDefaultButton2) = vbOK Then

        '在执行过程中禁用确认和退出按钮
        cmdOK.Enabled = False
        cmdCancel.Enabled = False

        '执行前得到检查记录表记录数
        strSql = "select count(*) as 记录数 from 病理检查信息"
        Set rsPathologyData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        intPathDataCount = Val(rsPathologyData("记录数").Value)
        
        '开始事务
        gcnOracle.BeginTrans
    On Error GoTo errTrans
        
        '提示不同导入状态信息
        With lblPrompt
            .Font.Bold = True
            .Caption = "开始导入病理标本信息数据,请稍后....."
        End With
        
        '执行导入病理数据SQL语句
        Call gcnOracle.Execute("insert into 病理标本信息(标本ID, 医嘱ID, 标本名称,标本类型,数量,接收日期) " & _
                        " select 病理标本信息_标本ID.Nextval,a.医嘱ID,a.标本部位,0,a.块数,b.核收时间 " & _
                        " from 影像病理标本 a, 影像标本核收取材 b where a.医嘱id=b.医嘱id " & _
                        " and not exists(Select 1 From 病理标本信息 where 医嘱ID=a.医嘱ID and 标本名称=a.标本部位 and 数量=a.块数 and 接收日期=b.核收时间)")
                        
        With lblPrompt
            .Font.Bold = True
            .Caption = "开始导入病理检查信息数据,请稍后....."
        End With
            
        Call gcnOracle.Execute("insert into 病理检查信息(病理医嘱ID,病理号,医嘱ID,检查类型,巨检描述,剩余位置) " & _
                               " select 病理检查信息_病理医嘱ID.Nextval,病理号,医嘱ID," & strDecode & ",巨检所见,剩余标本位置 " & _
                               " from 影像标本核收取材 where 医嘱ID not in(select 医嘱ID from 病理检查信息)")
        With lblPrompt
            .Font.Bold = True
            .Caption = "开始导入病理送检信息数据,请稍后....."
        End With
                               
        Call gcnOracle.Execute("insert into 病理送检信息(ID,医嘱ID,送检单位,送检科室,送检人,送检日期,登记人,核收状态,拒收原因,备注) " & _
                               " select 病理送检信息_id.nextval,医嘱ID,'本院','', '未录入',核收时间,decode(核收技师,null,'未录入',核收技师),decode(核收情况,'1','1','0'),拒收原因,备注 " & _
                               " from 影像标本核收取材 a where not exists(Select 1 From 病理送检信息 where 医嘱ID=a.医嘱id and 送检日期=a.核收时间 and 拒收原因=a.拒收原因)")
        
        Call gcnOracle.Execute("update 病理标本信息 a set 送检ID=(select  ID from 病理送检信息 where 医嘱ID=a.医嘱ID and rownum=1)")
        
        '提交事务
        gcnOracle.CommitTrans
        GoTo transOk
errTrans:
   gcnOracle.RollbackTrans
transOk:
        '执行得到导入记录数SQL
        strSql = "select count(*) as 记录数 from 病理检查信息"
        Set rsPathologyData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        
        '计算导入前和导入后的记录差
        intPathDataCount = Val(rsPathologyData("记录数").Value) - intPathDataCount

        With lblPrompt
            .Font.Bold = True
            .Caption = "历史数据已全部导入,共导入" & intPathDataCount & "条记录"
            
        End With

        '在执行完成后启用退出按钮
        cmdOK.Enabled = True
        cmdCancel.Enabled = True

        Exit Sub
        
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
'加载初始化方法
    Call InitVfgData
    Call LoadOldPathologyData
End Sub

