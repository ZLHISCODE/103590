VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm保险项目选择 
   AutoRedraw      =   -1  'True
   Caption         =   "医保项目选择"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   Icon            =   "frm保险项目选择.frx":0000
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
            Picture         =   "frm保险项目选择.frx":0E42
            Key             =   "Detail"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目选择.frx":1C94
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
Attribute VB_Name = "frm保险项目选择"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrCode As String '入出参数,医保项目DetailCode
Private mrsDetail As ADODB.Recordset
Private mblnOK As Boolean
Private mint险类 As Integer
Private mint适用地区 As Integer '沈阳专用；0表示其他地区，1表示长春（允许删除已审核的项目）

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
    If mint险类 = TYPE_兴成核工业 Then
        mstrCode = Mid(lvwDetail.SelectedItem.Key, 2) & "|" & lvwDetail.SelectedItem.SubItems(1) & "|" & Mid(lvwClass.SelectedItem.Key, 2)
    ElseIf mint险类 = TYPE_陕西大兴 Then
        mstrCode = Mid(lvwDetail.SelectedItem.Key, 3) & "|" & lvwDetail.SelectedItem.SubItems(1) & "|" & Mid(lvwClass.SelectedItem.Key, 2)
    Else
        mstrCode = Mid(lvwDetail.SelectedItem.Key, 2)
    End If
    mblnOK = True
    Unload Me
End Sub

Public Function GetCode(strCode As String, ByVal int险类 As Integer) As Boolean
'功能：获得一个收费项目的医保编码
'参数：strCode 既作为输入参数，又输出
'返回：成功返回True
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, objItem As ListItem
    
    mblnOK = False
    mint险类 = int险类
    
    On Error GoTo ErrH
    
    Set rsTmp = New ADODB.Recordset
    Set mrsDetail = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    mrsDetail.CursorLocation = adUseClient
    
    Select Case int险类
        Case TYPE_云南建水
            With gcnSybase
                If .State = adStateOpen Then .Close
                .Provider = "MSDataShape"
                '固定使用该用户、密码和主机字符串
'                .Open "Driver={Microsoft ODBC for Oracle};Server=" & "si2000", "yyzf", "yhcsi2000"
                .Open "Driver={Microsoft ODBC for Oracle};Server=" & "si2000", "his", "his"
                If .State = adStateClosed Then Exit Function
            End With
            
            rsTmp.Open "Select Upper(SFDLBM) as CODE,SFDLMC as NAME From BG01SFXMDL order by CODE", gcnSybase, adOpenKeyset
            mrsDetail.Open "Select Upper(SFDLBM) as CLASSCODE,Upper(SFXMBM) as CODE,XMMC as NAME from v_bg02fwxm order by CLASSCODE,CODE", gcnSybase, adOpenKeyset
        Case TYPE_云南省, TYPE_昆明市
            With gcnSybase
                If .State = adStateOpen Then .Close
                .Provider = "MSDataShape"
                '固定使用该用户、密码和主机字符串
'                .Open "Driver={Microsoft ODBC for Oracle};Server=" & "si2000", "yyzf", "yhcsi2000"
                .Open "Driver={Microsoft ODBC for Oracle};Server=" & "si2000", "his", "his"
                If .State = adStateClosed Then Exit Function
            End With
            
            rsTmp.Open "Select Upper(SFDLBM) as CODE,SFDLMC as NAME From BG01SFXMDL order by CODE", gcnSybase, adOpenKeyset
            mrsDetail.Open "select Upper(SFDLBM) as CLASSCODE,Upper(SFXMBM) as CODE,xmmc NAME,gg 规格,dw 单位,jx 剂型,cd 产地," & _
                           " DECODE(tjdm,1,'门诊特检',2,'甲类公费',3,'乙类挂钩',5,'抢救用药',6,'器官购置',31,'特殊挂钩','全自费') AS 类别 " & _
                           " from v_bg02fwxm Where YAB060 IN ('$$$$'," & IIf(int险类 = TYPE_昆明市, "'0101'", "'0000'") & ") order by CLASSCODE,CODE", gcnSybase, adOpenKeyset
        Case TYPE_成都市
            If 医保初始化_成都 = False Then Exit Function
            
            rsTmp.Open "Select Upper(sfdlbm) as CODE,sfdlmc as NAME From sfxmdl order by CODE", gcnSybase, adOpenKeyset
            mrsDetail.Open "Select Upper(sfdlbm) as CLASSCODE,Upper(sfxmbm) as CODE,xmmc as NAME from ypsfxmb order by CLASSCODE,CODE", gcnSybase, adOpenKeyset
        Case TYPE_贵阳市, type_成都郊县
            gstrSQL = "Select 大类编码 as CLASSCODE,编码 AS CODE,trim(名称) AS NAME ,简码 " & IIf(int险类 = TYPE_贵阳市, ",附注", "") & _
                           " from 保险项目 where 险类=[1] order by 大类编码,编码"
            Set mrsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "保险项目选择", int险类)
        Case TYPE_自贡市, TYPE_成都市农医, TYPE_成都德阳, Is > 900
            '医保大类
            gstrSQL = "Select 编码 AS CODE,名称 AS NAME From 保险支付大类 where 险类=[1] order by 编码"
            Set mrsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "保险项目选择", int险类)
            
            '中心药典
            gstrSQL = "Select 大类编码 as CLASSCODE,编码 AS CODE,名称 AS NAME ,简码,附注 " & _
                           " from 保险项目 where 险类=[1] order by 大类编码,编码"
            Set mrsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "保险项目选择", int险类)
        Case TYPE_泸州市
            gstrSQL = "Select 编码 AS CODE,名称 AS NAME From 保险支付大类 where 险类=" & int险类 & " order by 编码"
            rsTmp.Open gstrSQL, gcn泸州, adOpenStatic, adLockReadOnly
            
            gstrSQL = "SELECT A.编码  AS CODE,A.名称 AS NAME,A.简码,A.单位,A.大类编码 as CLASSCODE,C.名称 AS 剂型 " & _
                      "     ,A.是否是药,A.是否医保,A.最大价格限制,A.首先自付比例,A.价格,A.项目内涵,A.除外内容,A.说明 " & _
                      "  FROM 保险项目 A,剂型 C " & _
                      "  WHERE A.险类=" & TYPE_泸州市 & " AND A.剂型编码=c.编码(+) "
            mrsDetail.Open gstrSQL, gcn泸州, adOpenStatic, adLockReadOnly
        Case TYPE_铜仁
            gstrSQL = "Select 编码 AS CODE,名称 AS NAME From 保险支付大类 where 险类=" & int险类 & " order by 编码"
            rsTmp.Open gstrSQL, gcn铜仁, adOpenStatic, adLockReadOnly
            
            gstrSQL = "SELECT A.编码  AS CODE,A.名称 AS NAME,A.简码,A.单位,A.大类编码 as CLASSCODE,C.名称 AS 剂型 " & _
                      "     ,A.是否是药,A.是否医保,A.最大价格限制,A.首先自付比例,A.价格,A.项目内涵,A.除外内容,A.说明 " & _
                      "  FROM 保险项目 A,剂型 C " & _
                      "  WHERE A.险类=" & TYPE_铜仁 & " AND A.剂型编码=c.编码(+) "
            mrsDetail.Open gstrSQL, gcn铜仁, adOpenStatic, adLockReadOnly
        'Modified by 朱玉宝 20031218 地区：福州
        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市, TYPE_四川眉山, TYPE_沈阳市, TYPE_乐山
            gstrSQL = "Select 编码 AS CODE,名称 AS NAME From 保险支付大类 where 险类=" & int险类 & " order by 编码"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "保险项目选择", int险类)
            If rsTmp.RecordCount = 0 Then
                MsgBox "请先完成保险大类的设置。", vbInformation, gstrSysName
                Exit Function
            End If
            
            gstrSQL = "Select 大类编码 as ClassCode ,编码 AS CODE,名称 AS NAME,简码,附注 From 保险项目 where 险类=[1] order by 编码"
            Set mrsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "保险项目选择", int险类)
            
        Case TYPE_成都内江
            gstrSQL = "Select 编码 AS CODE,名称 AS NAME From 保险支付大类 where 险类=" & TYPE_成都内江 & " order by 编码"
            
            zlDatabase.OpenRecordset rsTmp, gstrSQL, "获取保险项目"
            gstrSQL = "SELECT A.编码  AS CODE,A.名称 AS NAME,A.简码,A.大类编码 as CLASSCODE,附注 " & _
                      "  FROM 保险项目 A " & _
                      "  WHERE A.险类=" & TYPE_成都内江 & _
                      "  order by 编码 "
            zlDatabase.OpenRecordset mrsDetail, gstrSQL, "获取保险项目"
        Case TYPE_开县
            gstrSQL = "Select 编码 AS CODE,名称 AS NAME From 保险支付大类 where 险类=" & TYPE_开县 & " order by 编码"
            
            zlDatabase.OpenRecordset rsTmp, gstrSQL, "获取保险项目"
            gstrSQL = "SELECT A.编码  AS CODE,A.名称 AS NAME,A.简码,A.大类编码 as CLASSCODE,附注 " & _
                      "  FROM 保险项目 A " & _
                      "  WHERE A.险类=" & TYPE_开县 & _
                      "  order by 编码 "
            zlDatabase.OpenRecordset mrsDetail, gstrSQL, "获取保险项目"
        
        Case TYPE_兴成核工业
            
            gstrSQL = "SELECT 0  AS CODE,'药品' AS NAME from dual  union all "
            gstrSQL = gstrSQL & "SELECT 1  AS CODE,'诊疗' AS NAME from dual  union all "
            gstrSQL = gstrSQL & "SELECT 2  AS CODE,'服务' AS NAME from dual   "
            
            zlDatabase.OpenRecordset rsTmp, gstrSQL, "获取保险项目"
            
            gstrSQL = "select 0  附注, '药品' as 类别,xmdm CODE,xmmc  Name,pl 品类,zfbl 自付比例,0 CLASSCODE " & _
                     " from  YB_YD  " & _
                     " union all  " & _
                     " select 1  附注, '诊疗' as 类别,xmdm CODE,xmmc Name,pl 品类,zfbl 自付比例,1 CLASSCODE " & _
                     " from   YB_ZLML " & _
                     " union all  " & _
                     " select 2  附注, '服务' as 类别,xmdm CODE,xmmc Name,pl 品类,zfbl 自付比例,2 CLASSCODE" & _
                     " from  YB_FWSS " & _
                     " "
            mrsDetail.Open gstrSQL, gcnSQLSEVER_兴成, adOpenStatic, adLockReadOnly
            
        Case TYPE_陕西大兴
            
            
            gstrSQL = "SELECT 1  AS CODE,'药品' AS NAME from dual  union all "
            gstrSQL = gstrSQL & "SELECT 2  AS CODE,'辅助项目' AS NAME from dual  "
            zlDatabase.OpenRecordset rsTmp, gstrSQL, "获取保险项目"

            gstrSQL = "" & _
                     " Select LB CLASSCODE, LB 附注,decode(LB,1,'药品','辅助') as 类别,LB||BM CODE,MC Name,PYBM 助记码, " & _
                     "        YPBM1 别名1,PYBM1 别名1助记码,YPBM2 别名2,PYBM2 别名2助记码,YPBM3 别名3,PYBM3 别名3助记码, " & _
                     "        YPJX  剂型,JG 价格,decode(YPLX,1,'中成药',2,'中草药',3,'西药','') 药品类型, " & _
                     "        decode(BXLX,1,'甲类',2,'自费',3,'乙类','') 报销类型,GUIG 药品规格 " & _
                     " From YY_YPFZB " & _
                     " union all  " & _
                     " select 类别 CLASSCODE,类别 附注,decode(类别,1,'药品','辅助') as 类别 ,类别||编码 CODE,名称 Name,简码 助记码, " & _
                     "        '' 别名1,'' 别名1助记码,'' 别名2,'' 别名2助记码,'' 别名3,'' 别名3助记码, " & _
                     "        ''  剂型,0 价格,decode(药品类型,1,'中成药',2,'中草药',3,'西药','') 药品类型, " & _
                     "        decode(报销类型,1,'甲类',2,'自费',3,'乙类','') 报销类型,'' 药品规格 " & _
                     " From 收费项目公用信息"
                        
            
            mrsDetail.Open gstrSQL, gcnOracle_神木大兴, adOpenStatic, adLockReadOnly
            
        Case Else
            Exit Function
    End Select
    
    '为明细增加多余显示的列
    Dim fld As ADODB.Field
    For Each fld In mrsDetail.Fields
        If fld.Name <> "CLASSCODE" And fld.Name <> "NAME" And fld.Name <> "CODE" Then
            If fld.Name <> "附注" Then
                lvwDetail.ColumnHeaders.Add , , fld.Name, 1000
            Else
                '将附注进行分解
                'Modified by 朱玉宝 20031218 地区：福州
                If int险类 = TYPE_福建巨龙 Or int险类 = TYPE_福建省 Or int险类 = TYPE_福州市 Or int险类 = TYPE_南平市 Then
                    lvwDetail.ColumnHeaders.Add , , "规格", 1000
                    lvwDetail.ColumnHeaders.Add , , "单位", 1000
                    lvwDetail.ColumnHeaders.Add , , "发票名称", 600
                    lvwDetail.ColumnHeaders.Add , , "化学名", 1200
                    lvwDetail.ColumnHeaders.Add , , "商品名", 1200
                    lvwDetail.ColumnHeaders.Add , , "厂家", 1500
                    lvwDetail.ColumnHeaders.Add , , "剂型", 600
                    lvwDetail.ColumnHeaders.Add , , "是否医保", 1000, lvwColumnCenter
                ElseIf int险类 = TYPE_云南省 Or int险类 = TYPE_昆明市 Then
                    lvwDetail.ColumnHeaders.Add , , "规格", 1000
                    lvwDetail.ColumnHeaders.Add , , "剂型", 600
                    lvwDetail.ColumnHeaders.Add , , "单位", 1000
                    lvwDetail.ColumnHeaders.Add , , "厂家", 1500
                    lvwDetail.ColumnHeaders.Add , , "类别", 1200
                ElseIf int险类 = TYPE_四川眉山 Then
                    lvwDetail.ColumnHeaders.Add , , "单位", 1000
                    lvwDetail.ColumnHeaders.Add , , "规格", 1000
                    lvwDetail.ColumnHeaders.Add , , "是否医保", 800
                    lvwDetail.ColumnHeaders.Add , , "服务对象", 1000, lvwColumnCenter
                ElseIf int险类 = TYPE_自贡市 Then
                    lvwDetail.ColumnHeaders.Add , , "单位", 1000
                    lvwDetail.ColumnHeaders.Add , , "是否医保", 1000, lvwColumnCenter
                    lvwDetail.ColumnHeaders.Add , , "是否是药", 1000, lvwColumnCenter
                    lvwDetail.ColumnHeaders.Add , , "剂型", 1000
                'Modified By 朱玉宝 地区：长沙
                ElseIf int险类 = TYPE_沈阳市 Then
                    lvwDetail.ColumnHeaders.Add , , "规格", 1000
                    lvwDetail.ColumnHeaders.Add , , "产地", 1000
                    lvwDetail.ColumnHeaders.Add , , "剂型", 1000
                    If mint适用地区 = 1 Then lvwDetail.ColumnHeaders.Add , , "单价", 1000
                ElseIf int险类 = TYPE_乐山 Then
                    lvwDetail.ColumnHeaders.Add , , "费用类型", 1000
                    lvwDetail.ColumnHeaders.Add , , "费用项目类型", 1000
                    lvwDetail.ColumnHeaders.Add , , "特殊项目", 1000
                    lvwDetail.ColumnHeaders.Add , , "医保编码", 1000
                ElseIf int险类 = TYPE_兴成核工业 Then
                    lvwDetail.ColumnHeaders.Add , , "类别", 1000
                    lvwDetail.ColumnHeaders.Add , , "品类", 1000
                    lvwDetail.ColumnHeaders.Add , , "自付比例", 1000
                ElseIf int险类 = TYPE_陕西大兴 Then
                    
                    lvwDetail.ColumnHeaders.Add , , "类别", 1000
                    lvwDetail.ColumnHeaders.Add , , "助记码", 1000
                    lvwDetail.ColumnHeaders.Add , , "别名1", 1000
                    lvwDetail.ColumnHeaders.Add , , "别名1助记码", 800
                    lvwDetail.ColumnHeaders.Add , , "别名2", 1000
                    lvwDetail.ColumnHeaders.Add , , "别名2助记码", 800
                    lvwDetail.ColumnHeaders.Add , , "别名3", 1000
                    lvwDetail.ColumnHeaders.Add , , "别名3助记码", 800
                    lvwDetail.ColumnHeaders.Add , , "剂型", 1000
                    lvwDetail.ColumnHeaders.Add , , "价格", 800
                    lvwDetail.ColumnHeaders.Add , , "药品类型", 800
                    lvwDetail.ColumnHeaders.Add , , "报销类型", 800
                    lvwDetail.ColumnHeaders.Add , , "药品规格", 1000
                ElseIf int险类 = TYPE_贵阳市 Then
                    lvwDetail.ColumnHeaders.Add , , "最高限价", 1000
                    lvwDetail.ColumnHeaders.Add , , "自付比例", 1000
                    lvwDetail.ColumnHeaders.Add , , "生育项目", 1000
                    lvwDetail.ColumnHeaders.Add , , "工伤项目", 1000
                    lvwDetail.ColumnHeaders.Add , , "特殊报销", 1000
                    lvwDetail.ColumnHeaders.Add , , "包干结算", 1000
                End If
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
    If int险类 = TYPE_贵阳市 Or int险类 = type_成都郊县 Or int险类 = TYPE_福建巨龙 Or _
    int险类 = TYPE_四川眉山 Or int险类 = TYPE_乐山 Or int险类 = TYPE_沈阳市 Or _
    int险类 = TYPE_福建省 Or int险类 = TYPE_福州市 Or int险类 = TYPE_南平市 Or int险类 = TYPE_成都内江 Or int险类 = TYPE_开县 Then
        '明细可以更新
        cmdRequery.Visible = True
    End If
    
    
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
    
    
    frm保险项目选择.Show 1
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
        .Fields.Append "CLASSCODE", adVarChar, 6 '大类编码
        'Modified By 朱玉宝 2003-12-09 地区：乐山
        If mint险类 = TYPE_乐山 Then
            .Fields.Append "CODE", adVarChar, 40     '编码
        Else
            .Fields.Append "CODE", adVarChar, 20     '编码
        End If
        .Fields.Append "NAME", adVarChar, 300     '名称
        .Fields.Append "PY", adVarChar, 150       '拼音简码
        .Fields.Append "MEMO", adVarChar, 500     '附注
        .Open
    End With
    
    bln全量 = True
    Me.Caption = "医保项目选择（正在读取从文件或网络读取保险项目明细，请稍候......）"
    If mint险类 = TYPE_贵阳市 Then
        blnReturn = 医保项目_贵阳(rsTemp)
    ElseIf mint险类 = type_成都郊县 Then
        blnReturn = 医保项目_成都郊县(rsTemp)
    ElseIf mint险类 = TYPE_福建巨龙 Or mint险类 = TYPE_福建省 Or mint险类 = TYPE_福州市 Or mint险类 = TYPE_南平市 Then
        blnReturn = 医保项目_福建巨龙(rsTemp)
    ElseIf mint险类 = TYPE_四川眉山 Then
        blnReturn = 医保项目_四川眉山(rsTemp)
    ElseIf mint险类 = TYPE_乐山 Then
        blnReturn = 医保项目_乐山(rsTemp, bln全量)
    ElseIf mint险类 = TYPE_沈阳市 Then
        blnReturn = 医保项目_沈阳市(rsTemp)
    ElseIf mint险类 = TYPE_成都内江 Then
        If MsgBox("是否清空原来的医保项目？", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) <> vbYes Then
            bln全量 = False
        End If
        blnReturn = 医保项目_成都内江(rsTemp)
    ElseIf mint险类 = TYPE_开县 Then
        If MsgBox("是否清空原来的医保项目？", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) <> vbYes Then
            bln全量 = False
        End If
        blnReturn = 医保项目_开县(rsTemp)
    End If
    
    If blnReturn = False Then
        MousePointer = vbDefault
        Exit Sub
    End If
    
    Me.Caption = "医保项目选择（正在更新医保项目......）"
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    If bln全量 Then
        gstrSQL = "zl_保险项目_Clear(" & mint险类 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "医保项目选择")
    End If
    
    '更新保险项目
    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    Do Until rsTemp.EOF
        str附注 = Nvl(rsTemp("MEMO"))
        If mint险类 = TYPE_沈阳市 Then
            str费用类型 = Split(str附注, "^^")(1)
            If Trim(str费用类型) <> "" Then
                '只要不为空，说明是药品项目，更新费用类型
                gstrSQL = "ZL_更新费用类型('" & rsTemp("CODE") & "','" & str费用类型 & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "更新费用类型")
            End If
            str附注 = Split(str附注, "^^")(0)
        End If
        
        '插入保险项目
        gstrSQL = "zl_保险项目_Insert(" & mint险类 & ",'" & rsTemp("CODE") & "','" & ToVarchar(rsTemp("NAME"), 300) & _
            "','" & ToVarchar(rsTemp("PY"), 150) & "','" & ToVarchar(rsTemp("CLASSCODE"), 6) & "','" & ToVarchar(str附注, 500) & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新医保项目")
        Me.Caption = "医保项目选择（正在更新医保项目，已插入" & rsTemp.AbsolutePosition & "条记录）"
        rsTemp.MoveNext
    Loop
    
    '更新保险病种
    If mint险类 = TYPE_沈阳市 Then
        Me.Caption = "医保项目选择（正在读取从文件或网络读取保险疾病明细，请稍候......）"
        If Not 疾病目录_沈阳 Then
            gcnOracle.RollbackTrans
            Exit Sub
        End If
    End If
    gcnOracle.CommitTrans
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
    gcnOracle.RollbackTrans
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
    Const COL_大类  As Long = 5
    
    Err = 0
    On Error GoTo errHand:
    Dim ObjExcel As Object, ObjCell As Object, strFile As String, strValue As String
    
    '选择指定文件
    On Error Resume Next
    Err = 0
    With Dlg
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
                rsTemp("CLASSCODE") = Replace(ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL_项目种类)), 6), "'", "")
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

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    If mint险类 = TYPE_沈阳市 Then
        mint适用地区 = 0
        gstrSQL = "Select 参数值 From 保险参数 Where 参数名='适用地区' And 险类=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取适用地区", TYPE_沈阳市)
        If Not rsTemp.EOF Then
            mint适用地区 = Nvl(rsTemp!参数值, 0)
        End If
    End If
End Sub

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
                    bln特殊处理 = False
                    
                    'Modified by 朱玉宝 20031218 地区：福州
                    If mint险类 = TYPE_福建巨龙 Or mint险类 = TYPE_福建省 Or mint险类 = TYPE_福州市 Or mint险类 = TYPE_南平市 Then
                        '附注的字段内容依次是：规格、发票名称、是否医保
                        If InStr(1, ",规格,单位,发票名称,是否医保,化学名,商品名,厂家,剂型,", str列) <> 0 Then
                            bln特殊处理 = True
                            varPart = Split(IIf(IsNull(mrsDetail("附注")), "", mrsDetail("附注")), "|")
                            Select Case str列
                                Case "规格"
                                    If UBound(varPart) >= 0 Then objItem.SubItems(lngCount - 1) = varPart(0)
                                Case "单位"
                                    If UBound(varPart) >= 1 Then objItem.SubItems(lngCount - 1) = varPart(1)
                                Case "发票名称"
                                    If UBound(varPart) >= 2 Then objItem.SubItems(lngCount - 1) = varPart(2)
                                Case "是否医保"
                                    If UBound(varPart) >= 3 Then objItem.SubItems(lvwDetail.ColumnHeaders.Count - 1) = varPart(3)
                                Case "化学名"
                                    If UBound(varPart) >= 4 Then objItem.SubItems(lngCount - 1) = varPart(4)
                                Case "商品名"
                                    If UBound(varPart) >= 5 Then objItem.SubItems(lngCount - 1) = varPart(5)
                                Case "厂家"
                                    If UBound(varPart) >= 6 Then objItem.SubItems(lngCount - 1) = varPart(6)
                                Case "剂型"
                                    If UBound(varPart) >= 7 Then objItem.SubItems(lngCount - 1) = varPart(7)
                            End Select
                        End If
                    'Modified By 朱玉宝 地区：长沙
                    ElseIf mint险类 = TYPE_沈阳市 Then
                        If str列 = "剂型" Or str列 = "规格" Or str列 = "产地" Or str列 = "单价" Then
                            bln特殊处理 = True
                            varPart = Split(IIf(IsNull(mrsDetail("附注")), "", mrsDetail("附注")), "||")
                            Select Case str列
                                Case "规格"
                                    If UBound(varPart) >= 1 Then objItem.SubItems(lngCount - 1) = varPart(1)
                                Case "产地"
                                    If UBound(varPart) >= 2 Then objItem.SubItems(lngCount - 1) = varPart(2)
                                Case "剂型"
                                    If UBound(varPart) >= 3 Then objItem.SubItems(lngCount - 1) = varPart(3)
                                Case "单价"
                                    If UBound(varPart) >= 4 Then objItem.SubItems(lngCount - 1) = varPart(4)
                            End Select
                        End If
                    ElseIf mint险类 = TYPE_四川眉山 Then
                        '附注的字段内容依次是：规格、发票名称、是否医保
                        If str列 = "规格" Or str列 = "是否医保" Or str列 = "单位" Or str列 = "服务对象" Then
                            bln特殊处理 = True
                            varPart = Split(IIf(IsNull(mrsDetail("附注")), "", mrsDetail("附注")), "|")
                            Select Case str列
                                Case "单位"
                                    If UBound(varPart) >= 0 Then objItem.SubItems(lngCount - 1) = varPart(0)
                                Case "规格"
                                    If UBound(varPart) >= 1 Then objItem.SubItems(lngCount - 1) = varPart(1)
                                Case "是否医保"
                                    If UBound(varPart) >= 2 Then objItem.SubItems(lngCount - 1) = varPart(2)
                                Case "服务对象"
                                    If UBound(varPart) >= 3 Then objItem.SubItems(lngCount - 1) = varPart(3)
                            End Select
                        End If
                    ElseIf mint险类 = TYPE_自贡市 Then
                        '附注的字段内容依次是：剂型编码、是否医保、是否是药、单位
                        If str列 = "单位" Or str列 = "是否是药" Or str列 = "是否医保" Or str列 = "剂型" Then
                            bln特殊处理 = True
                            varPart = Split(IIf(IsNull(mrsDetail("附注")), "", mrsDetail("附注")), "|")
                            If UBound(varPart) >= 4 Then
                                If str列 = "单位" Then
                                    objItem.SubItems(lngCount - 1) = varPart(3)
                                ElseIf str列 = "是否是药" Then
                                    objItem.SubItems(lngCount - 1) = IIf(varPart(2) = "1", "是", "否")
                                ElseIf str列 = "是否医保" Then
                                    objItem.SubItems(lngCount - 1) = IIf(varPart(1) = "1", "是", "否")
                                Else          '"剂型"
                                    objItem.SubItems(lngCount - 1) = varPart(4)
                                End If
                            End If
                        End If
                    ElseIf mint险类 = TYPE_乐山 Then
                        If str列 = "费用类型" Or str列 = "费用项目类型" Or str列 = "特殊项目" Or str列 = "医保编码" Then
                            bln特殊处理 = True
                            varPart = Split(IIf(IsNull(mrsDetail("附注")), "", mrsDetail("附注")), "|")
                            Select Case str列
                                Case "费用类型"
                                    If UBound(varPart) >= 0 Then objItem.SubItems(lngCount - 1) = varPart(0)
                                Case "费用项目类型"
                                    If UBound(varPart) >= 1 Then objItem.SubItems(lngCount - 1) = varPart(1)
                                Case "特殊项目"
                                    If UBound(varPart) >= 2 Then objItem.SubItems(lngCount - 1) = IIf(Val(varPart(2)) = 0, "否", "是")
                                Case "医保编码"
                                    If UBound(varPart) >= 3 Then objItem.SubItems(lngCount - 1) = varPart(3)
                            End Select
                        End If
                    ElseIf mint险类 = TYPE_贵阳市 Then
                        If str列 = "最高限价" Or str列 = "自付比例" Or str列 = "生育项目" Or str列 = "工伤项目" _
                        Or str列 = "特殊报销" Or str列 = "包干结算" Then
                            bln特殊处理 = True
                            varPart = Split(IIf(IsNull(mrsDetail("附注")), "", mrsDetail("附注")), "|")
                            Select Case str列
                                Case "最高限价"
                                    If UBound(varPart) >= 0 Then objItem.SubItems(lngCount - 1) = varPart(0)
                                Case "自付比例"
                                    If UBound(varPart) >= 1 Then objItem.SubItems(lngCount - 1) = varPart(1)
                                Case "生育项目"
                                    If UBound(varPart) >= 2 Then objItem.SubItems(lngCount - 1) = IIf(Val(varPart(2)) = 0, "否", "是")
                                Case "工伤项目"
                                    If UBound(varPart) >= 3 Then objItem.SubItems(lngCount - 1) = IIf(Val(varPart(3)) = 0, "否", "是")
                                Case "特殊报销"
                                    If UBound(varPart) >= 4 Then objItem.SubItems(lngCount - 1) = IIf(Val(varPart(4)) = 0, "普通项目", IIf(Val(varPart(4)) = 1, "允许公务员补助支付的特殊自费项目", "基金直接支付项目"))
                                Case "包干结算" '01-普通项目；02-包干结算加收范围内项目；03-医疗照顾人员特殊项目； 04-基金直接支付项目；
                                                '05-包干结算加收自费项目；06-允许三级医院加上全自费项目
                                    If UBound(varPart) >= 5 Then
                                        objItem.SubItems(lngCount - 1) = IIf(Val(varPart(5)) = 1, "普通项目", _
                                                                         IIf(Val(varPart(5)) = 2, "包干结算加收范围内项目", _
                                                                         IIf(Val(varPart(5)) = 3, "医疗照顾人员特殊项目", _
                                                                         IIf(Val(varPart(5)) = 4, "基金直接支付项目", _
                                                                         IIf(Val(varPart(5)) = 5, "包干结算加收自费项目", "允许三级医院加上全自费项目")))))
                                    End If
                            End Select
                        End If
                    End If
                    
                    If bln特殊处理 = False Then
                        '没有进行特殊处理
                        objItem.SubItems(lngCount - 1) = IIf(IsNull(mrsDetail(.Item(lngCount).Text)), "", mrsDetail(.Item(lngCount).Text))
                    End If
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

Private Function 医保项目_福建巨龙(rsTemp As ADODB.Recordset) As Boolean
'功能：更新福建巨龙的医保项目
    Const cOL编码 As Long = 1
    Const COL收据费目 As Long = 2
    Const cOL名称 As Long = 3
    Const COL规格 As Long = 4
    Const COL单位 As Long = 5
    Const COL是否医保 As Long = 7
    Const COL大类 As Long = 8
    Const COL拼音 As Long = 9
    Const COL化学名 As Long = 10
    Const COL商品名 As Long = 11
    Const COL厂家 As Long = 12
    Const COL剂型 As Long = 13
    
    Dim ObjExcel As Object, ObjCell As Object, strFile As String, strValue As String
    Dim rs大类 As New ADODB.Recordset
    
    
    gstrSQL = "Select 编码,名称 From 保险支付大类 Where 险类=[1]"
    Set rs大类 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint险类)
    
    '选择指定文件
    On Error Resume Next
    Err = 0
    With Dlg
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
    
    On Error GoTo errHandle
    Me.Caption = "医保项目选择（正在从EXCEL文件中提取数据......）"
    
    '取EXCEL文件的数据
    With ObjExcel
        .Workbooks.Open strFile
        
        '取各列的值
        Dim lngRow As Long
        lngRow = 2
        Do While True
            If .ActiveSheet.Cells(lngRow, cOL编码) <> "" Then
                rsTemp.AddNew
                
                rs大类.Filter = "名称='" & Trim(.ActiveSheet.Cells(lngRow, COL大类)) & "'"
                If rs大类.RecordCount > 0 Then
                    rsTemp("ClassCode") = rs大类("编码")
                End If
                rsTemp("Code") = Mid(Trim(.ActiveSheet.Cells(lngRow, cOL编码)), 1, 20)
                rsTemp("Name") = Replace(ToVarchar(Trim(.ActiveSheet.Cells(lngRow, cOL名称)), 40), "'", "")
                rsTemp("PY") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL拼音)), 10)
                rsTemp("MEMO") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL规格)) & _
                                "|" & Trim(.ActiveSheet.Cells(lngRow, COL单位)) & _
                                "|" & Trim(.ActiveSheet.Cells(lngRow, COL收据费目)) & _
                                "|" & Trim(.ActiveSheet.Cells(lngRow, COL是否医保)) & _
                                "|" & Trim(.ActiveSheet.Cells(lngRow, COL化学名)) & _
                                "|" & Trim(.ActiveSheet.Cells(lngRow, COL商品名)) & _
                                "|" & Trim(.ActiveSheet.Cells(lngRow, COL厂家)) & _
                                "|" & Trim(.ActiveSheet.Cells(lngRow, COL剂型)), 1000)
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
    医保项目_福建巨龙 = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function 医保项目_四川眉山(rsTemp As ADODB.Recordset) As Boolean
'功能：更新福建巨龙的医保项目
    Const cOL编码 As Long = 1
    Const cOL名称 As Long = 2
    Const COL单位 As Long = 3
    Const COL规格 As Long = 4
    Const COL是否医保 As Long = 5
    Const COL大类 As Long = 6
    Const cOL简码 As Long = 7
    Const COL服务对象 As Long = 8
    
    Dim ObjExcel As Object, ObjCell As Object, strFile As String, strValue As String
    Dim int服务对象 As Integer
    Dim rs大类 As New ADODB.Recordset
    
    gstrSQL = "Select 编码,名称 From 保险支付大类 Where 险类=[1]"
    Set rs大类 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint险类)
    
    '选择指定文件
    On Error Resume Next
    Err = 0
    With Dlg
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
    
    On Error GoTo errHandle
    Me.Caption = "医保项目选择（正在从EXCEL文件中提取数据......）"
    
    '取EXCEL文件的数据
    With ObjExcel
        .Workbooks.Open strFile
        
        '取各列的值
        Dim lngRow As Long
        lngRow = 2
        Do While True
            If .ActiveSheet.Cells(lngRow, cOL编码) <> "" Then
                rsTemp.AddNew
                
                rs大类.Filter = "名称='" & Trim(.ActiveSheet.Cells(lngRow, COL大类)) & "'"
                If rs大类.RecordCount > 0 Then
                    rsTemp("ClassCode") = rs大类("编码")
                End If
                rsTemp("Code") = Mid(Trim(.ActiveSheet.Cells(lngRow, cOL编码)), 1, 20)
                rsTemp("Name") = Replace(ToVarchar(Trim(.ActiveSheet.Cells(lngRow, cOL名称)), 40), "'", "")
                rsTemp("PY") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, cOL简码)), 10)
                If Trim(.ActiveSheet.Cells(lngRow, COL服务对象)) = "门诊、住院" Then
                    int服务对象 = 3
                Else
                    If Trim(.ActiveSheet.Cells(lngRow, COL服务对象)) = "门诊" Then
                        int服务对象 = 1
                    Else
                        int服务对象 = 2
                    End If
                End If
                rsTemp("MEMO") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL单位)) & _
                                "|" & Trim(.ActiveSheet.Cells(lngRow, COL规格)) & _
                                "|" & Trim(.ActiveSheet.Cells(lngRow, COL是否医保)) & _
                                "|" & int服务对象, 50)
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
    医保项目_四川眉山 = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function 医保项目_开县(rsTemp As ADODB.Recordset) As Boolean
    '功能：更新福建巨龙的医保项目
    Const cOL编码 As Long = 1
    Const cOL名称 As Long = 2
    Const cOL简码 As Long = 3
    Const COL大类 As Long = 4
    Dim ObjExcel As Object, ObjCell As Object, strFile As String, strValue As String
    
    Dim int服务对象 As Integer
    Dim rs大类 As New ADODB.Recordset
    
    gstrSQL = "Select 编码,名称 From 保险支付大类 Where 险类=[1]"
    Set rs大类 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint险类)
    
    '选择指定文件
    On Error Resume Next
    Err = 0
    With Dlg
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
    
    On Error GoTo errHandle
    Me.Caption = "医保项目选择（正在从EXCEL文件中提取数据......）"
    
    '取EXCEL文件的数据
    With ObjExcel
        .Workbooks.Open strFile
        
        '取各列的值
        Dim lngRow As Long
        lngRow = 2
        Do While True
            If .ActiveSheet.Cells(lngRow, cOL编码) <> "" Then
                rsTemp.AddNew
                
                rs大类.Filter = "名称='" & Trim(.ActiveSheet.Cells(lngRow, COL大类)) & "'"
                If rs大类.RecordCount > 0 Then
                    rsTemp("ClassCode") = rs大类("编码")
                End If
                rsTemp("Code") = Mid(Trim(.ActiveSheet.Cells(lngRow, cOL编码)), 1, 20)
                rsTemp("Name") = Replace(ToVarchar(Trim(.ActiveSheet.Cells(lngRow, cOL名称)), 40), "'", "")
                rsTemp("PY") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, cOL简码)), 10)
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
    医保项目_开县 = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function 医保项目_沈阳市(rsTemp As ADODB.Recordset) As Boolean
    Dim str编码 As String, str名称 As String, str简码 As String, str单价 As String
    Dim str规格 As String, str厂家 As String, str剂型 As String, str费用类型 As String
    Dim str大类 As String, int类型 As Integer, strTmp As String
    Dim rs大类 As New ADODB.Recordset
    Dim classInsure As New clsInsure
    '重新获取医保项目
    On Error GoTo errHand
    
    If Not classInsure.InitInsure(gcnOracle, TYPE_沈阳市) Then Exit Function
    
    '如何没有设置大类则退出
    gstrSQL = "Select 编码,名称 From 保险支付大类 Where 险类=[1]"
    Set rs大类 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_沈阳市)
    If rs大类.RecordCount < 4 Then
        MsgBox "请先正确的设置了保险大类后再使用本功能！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '开始导入中心的医保项目
    '----先取诊疗项目----
    If Not 调用接口_准备_沈阳市(Function_沈阳市.项目匹配_取项目信息) Then Exit Function
    '0-诊疗项目;1-药品
    gstrField_沈阳市 = "match_type"
    gstrValue_沈阳市 = "0"
    If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
    If Not 调用接口_执行_沈阳市() Then Exit Function
    If Not 调用接口_指定记录集_沈阳市("diseaseinfo") Then Exit Function
'    (1)当match_type="0"(诊疗项目)时，数据集包含以下内容：
'    序号    字段    字段说明    最大长度    备注
'    1   item_code  项目编码    20
'    2   item_name  项目名称    50
'    3   price      单价        12
'    4   code_wb    五笔码      20
'    5   code_py    拼音码      20
'    (2)当match_type="1"(药品)时，数据集包含以下内容：
'    序号    字段    字段说明    最大长度    备注
'    1   medi_code      药品编码    20
'    2   medi_name      药品名称    50
'    3   model_name     剂型名称    12
'    4   factory        生产厂家    50
'    5   standard       规格        20
'    6   medi_item_type 药品类型    1   "1"：西药   "2"：中成药    "3"：中草药
'    7   Staple_flag    费用类型    1   "1"：甲类   "2"：乙类      "9"：全自费
'    8   medi_item_name 药品类型名称10
'    9   code_wb        五笔码      20
'   10   code_py        拼音码      20
    int类型 = 0
    str大类 = "诊疗项目"
    rs大类.Filter = "名称='" & str大类 & "'"
    If rs大类.RecordCount > 0 Then
        str大类 = rs大类("编码")
    End If
    If 调用接口_记录数_沈阳市 Then
        Do While True
            Call 调用接口_读取数据_沈阳市("item_code", str编码)
            Call 调用接口_读取数据_沈阳市("item_name", str名称)
            Call 调用接口_读取数据_沈阳市("code_py", str简码)
            If mint适用地区 = 1 Then
                Call 调用接口_读取数据_沈阳市("price", str单价)
            End If
            '备注数据格式：类型||规格||厂家||剂型^^匹配序列号
            Call AddRecord(rsTemp, str编码, ToVarchar(str名称, 300), ToVarchar(str简码, 150), "0|| || || ||" & str单价 & "^^", ToVarchar(str大类, 6))
            
            If Not 调用接口_移动记录集_沈阳市(MoveNext) Then Exit Do
        Loop
    End If
    
    '----取药品信息----
    If Not 调用接口_准备_沈阳市(Function_沈阳市.项目匹配_取项目信息) Then Exit Function
    '0-诊疗项目;1-药品
    gstrField_沈阳市 = "match_type"
    gstrValue_沈阳市 = "1"
    If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
    If Not 调用接口_执行_沈阳市() Then Exit Function
    If Not 调用接口_指定记录集_沈阳市("diseaseinfo") Then Exit Function
    If 调用接口_记录数_沈阳市 Then
        Do While True
            Call 调用接口_读取数据_沈阳市("medi_code", str编码)
            Call 调用接口_读取数据_沈阳市("medi_name", str名称)
            Call 调用接口_读取数据_沈阳市("code_py", str简码)
            Call 调用接口_读取数据_沈阳市("standard", str规格)
            Call 调用接口_读取数据_沈阳市("model_name", str剂型)
            Call 调用接口_读取数据_沈阳市("factory", str厂家)
            If mint适用地区 = 1 Then
                Call 调用接口_读取数据_沈阳市("price", str单价)
            End If
            
            '取药品类型及大类信息
            Call 调用接口_读取数据_沈阳市("medi_item_type", strTmp)
            int类型 = Val(strTmp)
            str大类 = IIf(int类型 = 1, "西成药", IIf(int类型 = 2, "中成药", "中草药"))
            rs大类.Filter = "名称='" & str大类 & "'"
            If rs大类.RecordCount > 0 Then
                str大类 = rs大类("编码")
            End If
            
            '取费用类型
            Call 调用接口_读取数据_沈阳市("staple_flag", strTmp)
            If Val(strTmp) = 1 Then
                strTmp = "甲类药品"
            ElseIf Val(strTmp) = 2 Then
                strTmp = "乙类药品"
            Else
                strTmp = "非基本药品"
            End If
            
            Call AddRecord(rsTemp, str编码, ToVarchar(str名称, 300), ToVarchar(str简码, 150), int类型 & "||" & str规格 & "||" & str厂家 & "||" & str剂型 & "||" & str单价 & "^^" & strTmp, ToVarchar(str大类, 6))
            
            If Not 调用接口_移动记录集_沈阳市(MoveNext) Then Exit Do
        Loop
    End If
    
    医保项目_沈阳市 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function 医保项目_乐山(rsTemp As ADODB.Recordset, bln更新方式 As Boolean) As Boolean
'功能：更新福建巨龙的医保项目
    Dim str大类 As String
    Const COL类型 As Long = 1   '0:药品;1-诊疗;2-服务
    Const cOL编码 As Long = 2
    Const COL医保编码 As Long = 3
    Const cOL名称 As Long = 4
    Const cOL简码 As Long = 5
    Const COL费用种类 As Long = 6
    Const COL费用项目种类 As Long = 7
    Const COL特殊项目 As Long = 8
    
    Dim ObjExcel As Object, ObjCell As Object, strFile As String, strValue As String
    Dim rs大类 As New ADODB.Recordset
    
    gstrSQL = "Select 编码,名称 From 保险支付大类 Where 险类=[1]"
    Set rs大类 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint险类)
    
    '选择指定文件
    On Error Resume Next
    Err = 0
    With Dlg
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
    
    On Error GoTo errHandle
    bln更新方式 = Not 下载模式(TYPE_乐山)       '真表示增量，否表示全量，与外层不符，所以取非
    Me.Caption = "医保项目选择（正在从EXCEL文件中提取数据......）"
    
    '取EXCEL文件的数据
    With ObjExcel
        .Workbooks.Open strFile
        
        '取各列的值
        Dim lngRow As Long
        lngRow = 2
        Do While True
            If .ActiveSheet.Cells(lngRow, cOL编码) <> "" Then
                str大类 = Trim(.ActiveSheet.Cells(lngRow, COL类型))
                Select Case str大类
                Case "0"
                    str大类 = "药品"
                Case "1"
                    str大类 = "诊疗"
                Case "2"
                    str大类 = "服务"
                End Select
                
                rsTemp.AddNew
                
                rs大类.Filter = "名称='" & str大类 & "'"
                If rs大类.RecordCount > 0 Then
                    rsTemp("ClassCode") = rs大类("编码")
                End If
                rsTemp("Code") = .ActiveSheet.Cells(lngRow, cOL编码)
                rsTemp("Name") = Replace(ToVarchar(Trim(.ActiveSheet.Cells(lngRow, cOL名称)), 40), "'", "")
                rsTemp("PY") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, cOL简码)), 10)
                rsTemp("MEMO") = Trim(.ActiveSheet.Cells(lngRow, COL费用种类)) & "|" & Trim(.ActiveSheet.Cells(lngRow, COL费用项目种类)) & "|" & Trim(.ActiveSheet.Cells(lngRow, COL特殊项目)) & "|" & Trim(.ActiveSheet.Cells(lngRow, COL医保编码))
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
    医保项目_乐山 = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function 疾病目录_沈阳() As Boolean
    Dim lngRecords As Long
    Dim lngNextID As Long
    Dim str编码 As String, str名称 As String, str简码 As String
    Dim blnInsert As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim classInsure As New clsInsure
    '重新获取医保项目
    On Error GoTo errHand
    
    lngRecords = 1
    If Not classInsure.InitInsure(gcnOracle, TYPE_沈阳市) Then Exit Function
    '----取疾病信息----
    If Not 调用接口_准备_沈阳市(Function_沈阳市.项目匹配_取项目信息) Then Exit Function
    '0-诊疗项目;1-药品;2-疾病
    gstrField_沈阳市 = "match_type"
    gstrValue_沈阳市 = "2"
    If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
    If Not 调用接口_执行_沈阳市() Then Exit Function
    If Not 调用接口_指定记录集_沈阳市("diseaseinfo") Then Exit Function
    If 调用接口_记录数_沈阳市 Then
        '不能删除后重新插入病种信息，因为病种ID和其他表有联系，只有新的病种才能插入，现有病种通过修改实现
'        gstrSQL = "zl_保险病种_DELETEALL(" & TYPE_沈阳市 & ")"
'        Call zlDatabase.ExecuteProcedure(gstrSQL, "删除本险类所有医保疾病")
        gstrSQL = "Select ID,编码 From 保险病种 Where 险类=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取现有病种目录", TYPE_沈阳市)
        
        Do While True
            Call 调用接口_读取数据_沈阳市("icd", str编码)
            Call 调用接口_读取数据_沈阳市("disease", str名称)
            Call 调用接口_读取数据_沈阳市("code_py", str简码)
            str名称 = Replace(str名称, "'", "")
            
            With rsTemp
                .Filter = "编码='" & str编码 & "'"
                blnInsert = (.RecordCount = 0)
            End With
            
            '更新保险疾病
            If blnInsert Then
                lngNextID = zlDatabase.GetNextID("保险病种")
                gstrSQL = "zl_保险病种_INSERT(" & lngNextID & "," & TYPE_沈阳市 & ",'" & str编码 & _
                            "','" & str名称 & "','" & str简码 & "',0,NULL,NULL)"
            Else
                lngNextID = rsTemp!ID
                gstrSQL = "zl_保险病种_UPDATE(" & lngNextID & ",'" & str编码 & _
                            "','" & str名称 & "','" & str简码 & "',0,NULL,NULL)"
            End If
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Me.Caption = "医保项目选择（正在更新医保疾病目录，已插入" & lngRecords & "条记录）"
            lngRecords = lngRecords + 1
            
            If Not 调用接口_移动记录集_沈阳市(MoveNext) Then Exit Do
        Loop
        rsTemp.Filter = 0
    End If
    
    疾病目录_沈阳 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

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

Private Function 下载模式(ByVal lng险类 As Long) As Boolean
    Dim intReturn As Integer
    Dim rsTemp As New ADODB.Recordset
    '检查是否已存在医保项目，返回真表示增量下载，假表示全量下载
    gstrSQL = "Select 1 From 保险项目 Where 险类=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否已存在医保项目", lng险类)
    If rsTemp.RecordCount = 0 Then Exit Function
    
    '说明已存在记录，提示操作员将采取的方式
    intReturn = MsgBox("由于已存在数据，本次是否采取增量下载？" & vbCrLf & "是表示增量下载，否表示全量下载", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName)
    下载模式 = (intReturn = vbYes)
End Function
