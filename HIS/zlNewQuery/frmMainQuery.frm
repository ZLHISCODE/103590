VERSION 5.00
Begin VB.Form frmMainQuery 
   BorderStyle     =   0  'None
   ClientHeight    =   5925
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8850
   Icon            =   "frmMainQuery.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   8850
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrCheckConnect 
      Interval        =   60000
      Left            =   645
      Top             =   1035
   End
   Begin VB.Timer tmrHome 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   210
      Top             =   2535
   End
   Begin zl9NewQuery.ctlDefaultFrame FrameDefault 
      Height          =   4470
      Left            =   1065
      TabIndex        =   0
      Top             =   630
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   7885
   End
End
Attribute VB_Name = "frmMainQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvarlngHome As Long                          '查询页面返回主页的时间间隔
Private mvarBlnFirst As Boolean                      '是否是刚进入本模块

Public mvarHomeInternal As Long
Private mvarHomeLong As Long
Private mvarCheckConnectInternal As Long
Private mvarCheckConnectCounter As Long
Private mobjRegister As Object
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Activate()
    On Error GoTo ErrHandle

    If mvarBlnFirst = False Then Exit Sub
    mvarBlnFirst = False
     
    mvarHomeInternal = 0
    tmrHome.Enabled = IIf(mvarHomeLong = 0, False, True)
    
    mvarCheckConnectInternal = Val(GetPara("检查数据连接间隔时间", "30"))
    tmrCheckConnect.Enabled = IIf(mvarCheckConnectInternal = 0, False, True)
    
    FrameDefault.AllowEdit = (InStr(gstrPrivs, "信息维护") > 0)
    FrameDefault.AllowSelfRegist = (InStr(gstrPrivs, "自助挂号") > 0)
    FrameDefault.AllowSelfPrint = (InStr(gstrPrivs, "自助打印") > 0)
    FrameDefault.AllowFreeRegist = (InStr(gstrPrivs, "自助挂号") > 0)
    
    Set gfrmMain = Me
    
    '2.装载并显示主页面
    Call FrameDefault.InitLoad
    
    Call FrameDefault.ShowHome
    
    DoEvents
    'zyk add 200410
    Call FrameDefault.showwww
    
    Dim wwwurl As String
    wwwurl = GetPara("医院主页", "")
    If Not wwwurl = "" Then
        ShellExecute hwnd, "open", "iexplore.exe", "-k " & wwwurl, "", 1
        'Sleep 5000   'API函数延时5000毫秒
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '如果按了Esc,则退出本显示模块
    Select Case KeyCode
    Case vbKeyEscape
        
        If Shift = vbShiftMask Then
            If Val(GetPara("关闭查询需输入登录口令", "0")) = 1 Then
                If frmExitPsw.ShowPsw(Me) Then
                    Unload Me
                End If
            Else
                Unload Me
            End If
        End If
        
    Case vbKey0, vbKey1, vbKey2, vbKey3, vbKey4, vbKey5, vbKey6, vbKey7, vbKey8, vbKey9, vbKeyNumpad0, vbKeyNumpad1, vbKeyNumpad2, vbKeyNumpad3, vbKeyNumpad4, vbKeyNumpad5, vbKeyNumpad6, vbKeyNumpad7, vbKeyNumpad8, vbKeyNumpad9
        '直接调用病人费用查询
        
        gstrSQL = "select 1 from 咨询页面排列 A,咨询页面目录 B where A.页面=B.页面序号 and B.页面序号=2"
        
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If gRs.BOF = False Then
            '显示费用查询
            Call FrameDefault.ShowSpecPage(2)
            Select Case KeyCode
            Case vbKey0, vbKeyNumpad0
                Call FrameDefault.FirstChar("0")
            Case vbKey1, vbKeyNumpad1
                Call FrameDefault.FirstChar("1")
            Case vbKey2, vbKeyNumpad2
                Call FrameDefault.FirstChar("2")
            Case vbKey3, vbKeyNumpad3
                Call FrameDefault.FirstChar("3")
            Case vbKey4, vbKeyNumpad4
                Call FrameDefault.FirstChar("4")
            Case vbKey5, vbKeyNumpad5
                Call FrameDefault.FirstChar("5")
            Case vbKey6, vbKeyNumpad6
                Call FrameDefault.FirstChar("6")
            Case vbKey7, vbKeyNumpad7
                Call FrameDefault.FirstChar("7")
            Case vbKey8, vbKeyNumpad8
                Call FrameDefault.FirstChar("8")
            Case vbKey9, vbKeyNumpad9
                Call FrameDefault.FirstChar("9")
            End Select
        End If
        
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo errHand
    mvarBlnFirst = True
    
    '1.检查服务器上的图片是否已经更新，如果已经更新，则更新本地图片
    Call CheckPicture
    
'    Me.Width = 12000
'    Me.Height = 9000
    '2.获取返回主页面的时间间隔
    mvarHomeLong = Val(GetPara("返回主页间隔", "0"))
    Exit Sub
    
    
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call ResizeControl(FrameDefault, 0, 0, Me.ScaleWidth, Me.ScaleHeight)
End Sub

Private Sub LoadPageItemList(ByVal PageNo As Long)
'功能:加载页面的每一查询项目
'参数:PageNo            页面序号
'说明:这是查询内容显示的主体部份,显示查询内容
    Dim FileName As String
    Dim W As Single
    Dim H As Single
    Dim vFont As New StdFont
    Dim i As Long
    Dim j As Long
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
    Dim vRs As New ADODB.Recordset
    Dim vNextY As Single
    Dim vNextX As Single
    Dim objDraw As ctlQueryItem
    Dim vWidth As Single
    Dim vHeight As Single
    Dim vTmp As Single
    Dim vTmp1 As Single
    Dim vMaxWidth As Single
    Dim vVisible As Boolean
    Dim strText As String
    
    On Error GoTo errHand
    i = 1
    vNextY = 60 + (i - 1) * 600
    vNextX = 120
    vMaxWidth = 120
            
    ShowFlatFlash "请稍候，正在生成页面...", Me
    DoEvents
    
    Set objDraw = FrameDefault.ClientObj
    objDraw.ClientVisible = False
    Call objDraw.ClearAllPageItem
    
    '读取页面的背景及广告条幅
    gstrSQL = "select B.类型,B.名称 from 咨询页面目录 A,咨询图片元素 B where A.宣传标语=B.序号 and A.页面序号=[1]"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo)
    If gRs.BOF = False Then FrameDefault.AdviceMovie = IIf(IsNull(gRs!名称), "", App.Path & "\图形\" & gRs!名称 & IIf(gRs!类型 <> 2, ".pic", ".swf"))
                    
    '开始生成自定义查询页面
    gstrSQL = "select 页面序号,段落序号,段落类型,标题文本,标题图标,标题隐藏,标题位置,标题字体,返回页首,段落字体,插表序号,插表位置,插图序号,插图位置 from 咨询段落目录 where 页面序号=[1] order by 段落序号"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo)
    If gRs.BOF = False Then
        While Not gRs.EOF
            strTmp = IIf(IsNull(gRs!标题字体), "宋体;12;0;0;0", gRs!标题字体)
            vFont.Name = Split(strTmp, ";")(0)
            vFont.Size = Val(Split(strTmp, ";")(1))
            vFont.Bold = Val(Split(strTmp, ";")(2))
            vFont.Italic = Val(Split(strTmp, ";")(3))
                                    
            FileName = ""
            '1.加载标题内容及标题图标
            vVisible = IIf(IsNull(gRs!标题隐藏), 1, gRs!标题隐藏)
            
            gstrSQL = "select 名称 from 咨询图片元素 where 序号=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(IIf(IsNull(gRs!标题图标), 0, gRs!标题图标)))
            If rs.BOF = False Then
                FileName = GetFileName(IIf(IsNull(gRs!标题图标), 0, gRs!标题图标), W, H)
            End If
            Call objDraw.AddPageItemTitle(i, vNextY, IIf(IsNull(gRs!标题文本), "", gRs!标题文本), Val(Split(strTmp, ";")(4)), vFont, FileName, PageNo, IIf(IsNull(gRs!段落序号), 0, gRs!段落序号), vWidth, vHeight, Not vVisible, IIf(IsNull(gRs!标题位置), 0, gRs!标题位置))
                                                                                    
            If Not vVisible = True Then vNextY = vNextY + vHeight + 150

            Select Case zlCommFun.Nvl(gRs("段落类型").Value, 0)
            '----------------------------------------------------------------------------------------------------------
            Case 0      '纯文本内容
                strTmp = IIf(IsNull(gRs!段落字体), "宋体;12;0;0;0", gRs!段落字体)
                j = objDraw.NextTxtIndex
                
                vWidth = FrameDefault.ClientWidth - 330
                
                'strText = gRs!段落文本
                strText = Sys.ReadLob(glngSys, 29, PageNo & "," & gRs("段落序号").Value, "", 1)
                                
                Call objDraw.AddPageItemTxt(j, vNextX, vNextY, strText & Chr(13) & Chr(10), strTmp, vWidth, vHeight)
                vNextY = vNextY + vHeight + 150
            '----------------------------------------------------------------------------------------------------------
            Case 1      '纯表格内容
                vHeight = 0
                Call InsertGrid(objDraw, IIf(IsNull(gRs!插表序号), 0, gRs!插表序号), vNextX, vNextY, vWidth, vHeight)
                If vHeight > 0 Then vNextY = vNextY + vHeight + 150
            '----------------------------------------------------------------------------------------------------------
            Case 2      '纯图形内容
                FileName = GetFileName(IIf(IsNull(gRs!插图序号), 0, gRs!插图序号), W, H)
                Call objDraw.AddPageItemPic(objDraw.NextPicIndex, 0, vNextY, FileName, vWidth, vHeight, W, H)
                vNextY = vNextY + vHeight + 150
            '----------------------------------------------------------------------------------------------------------
            Case 3      '纯链接内容
                gstrSQL = "select A.链接页面,A.页内段号 from 咨询段落链接 A Where A.页面序号 =[1] And A.段落序号 = [2]"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo, Val(IIf(IsNull(gRs!段落序号), 0, gRs!段落序号)))
                If rs.BOF = False Then
                    While Not rs.EOF
                        If IIf(IsNull(rs!页内段号), 0, rs!页内段号) = 0 Then
                            '只链接到页面，没有指明页面内的具体项目
                            gstrSQL = "select C.页面名称 as 标题文本 from 咨询页面目录 C Where C.页面序号=[1]"
                            Set vRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(IIf(IsNull(rs!链接页面), 0, rs!链接页面)))
                            If vRs.BOF = False Then
                                Call objDraw.AddPageItemConnect(objDraw.NextConnectIndex, vNextX + 150, vNextY, IIf(IsNull(vRs!标题文本), "", vRs!标题文本), IIf(IsNull(rs!链接页面), 0, rs!链接页面), 0, vWidth, vHeight)
                                vNextY = vNextY + 300
                            End If
                        Else
                            '链接到页面内的具体项目
                            gstrSQL = "select C.页面名称||decode(B.标题文本,NULL,'','：'||B.标题文本) as 标题文本 from 咨询段落目录 B,咨询页面目录 C Where C.页面名称<>'专家介绍' and B.页面序号=C.页面序号 and C.页面序号=[1] and B.段落序号=[2]"
                            Set vRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(IIf(IsNull(rs!链接页面), 0, rs!链接页面)), Val(IIf(IsNull(rs!页内段号), 0, rs!页内段号)))
                            If vRs.BOF = False Then
                                Call objDraw.AddPageItemConnect(objDraw.NextConnectIndex, vNextX + 150, vNextY, IIf(IsNull(vRs!标题文本), "", vRs!标题文本), IIf(IsNull(rs!链接页面), 0, rs!链接页面), 0, vWidth, vHeight)
                                vNextY = vNextY + 300
                            Else
                                gstrSQL = "select B.姓名||'('||C.名称||')' as 姓名 from 人员表 B,部门人员 A,部门表 C Where B.id=A.人员id And (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null) and A.部门id=C.id and A.缺省=1 and B.id=[1]"
                                Set vRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(IIf(IsNull(rs!页内段号), 0, rs!页内段号)))
                                If vRs.BOF = False Then
                                    Call objDraw.AddPageItemConnect(objDraw.NextConnectIndex, vNextX + 150, vNextY, IIf(IsNull(vRs!姓名), "", "专家介绍：" & vRs!姓名), IIf(IsNull(rs!链接页面), 0, rs!链接页面), IIf(IsNull(rs!页内段号), 0, rs!页内段号), vWidth, vHeight)
                                    vNextY = vNextY + 300
                                End If
                            End If
                        End If
                        rs.MoveNext
                    Wend
                    vNextY = vNextY + 150
                End If
            '----------------------------------------------------------------------------------------------------------
            Case 4      '文本和表格
                strTmp = IIf(IsNull(gRs!段落字体), "宋体;12;0;0;0", gRs!段落字体)
                j = objDraw.NextTxtIndex
                
                strText = Sys.ReadLob(glngSys, 29, PageNo & "," & gRs("段落序号").Value, "", 1)
                
                Select Case IIf(IsNull(gRs!插表位置), 0, gRs!插表位置)
                Case 0
                    vHeight = 0
                    Call InsertGrid(objDraw, IIf(IsNull(gRs!插表序号), 0, gRs!插表序号), 0, vNextY, vTmp1, vTmp)
                    vWidth = FrameDefault.ClientWidth - vTmp1 - 120 - 120
                    Call objDraw.AddPageItemTxt(j, vNextX + vTmp1 + 60, vNextY, strText & Chr(13) & Chr(10), strTmp, vWidth, vHeight)
                Case 1
                    Call objDraw.AddPageItemTxt(j, vNextX, vNextY, strText & Chr(13) & Chr(10), strTmp, vTmp1, vHeight)
                    Call InsertGrid(objDraw, IIf(IsNull(gRs!插表序号), 0, gRs!插表序号), 1, vNextY, vWidth, vTmp)
                End Select
                vNextY = vNextY + IIf(vTmp > vHeight, vTmp, vHeight) + 150
            '----------------------------------------------------------------------------------------------------------
            Case 5      '文本和图形
            
                FileName = GetFileName(IIf(IsNull(gRs!插图序号), 0, gRs!插图序号), W, H)
                
                strText = Sys.ReadLob(glngSys, 29, PageNo & "," & gRs("段落序号").Value, "", 1)
                
                strTmp = IIf(IsNull(gRs!段落字体), "宋体;12;0;0;0", gRs!段落字体)
                j = objDraw.NextTxtIndex
                Select Case IIf(IsNull(gRs!插图位置), 0, gRs!插图位置)
                Case 0
                    Call objDraw.AddPageItemPic(objDraw.NextPicIndex, 0, vNextY, FileName, vTmp1, vTmp, W, H)
                    vWidth = FrameDefault.ClientWidth - vTmp1 - 120 - 120
                    Call objDraw.AddPageItemTxt(j, vNextX + vTmp1 + 60, vNextY, strText & Chr(13) & Chr(10), strTmp, vWidth, vHeight)
                Case 1
                    Call objDraw.AddPageItemPic(objDraw.NextPicIndex, 1, vNextY, FileName, vWidth, vTmp, W, H)
                    vTmp1 = FrameDefault.ClientWidth - vWidth - 60 - 90
                    Call objDraw.AddPageItemTxt(j, vNextX, vNextY, strText & Chr(13) & Chr(10), strTmp, vTmp1, vHeight)
                End Select
                vNextY = vNextY + IIf(vTmp > vHeight, vTmp, vHeight) + 150
                
            End Select
                      
            '8.设置返回页首标志
            If IIf(IsNull(gRs!返回页首), 0, gRs!返回页首) = 1 Then
                vHeight = 0
                Call objDraw.AddReturnFlag(vNextX, vNextY, vHeight)
                If vHeight > 0 Then vNextY = vNextY + vHeight + 150
            End If
            
            i = i + 1
            gRs.MoveNext
        Wend
    End If
        
    Call objDraw.ResizePage(FrameDefault.ClientWidth, vNextY)
    Call FrameDefault.InitNavigator(FrameDefault.ClientWidth, vNextY)
    
    '获取背景并画出页面背景
    gstrSQL = "select B.类型,B.名称,B.宽度,B.高度 from 咨询页面目录 A,咨询图片元素 B where A.页面背景=B.序号 and A.页面序号=[1]"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo)
    If gRs.BOF = False Then
        Call objDraw.BackPicture(IIf(IsNull(gRs!名称), "", App.Path & "\图形\" & gRs!名称 & IIf(gRs!类型 <> 2, ".pic", ".swf")), IIf(IsNull(gRs!宽度), 0, gRs!宽度) * Screen.TwipsPerPixelX, IIf(IsNull(gRs!高度), 0, gRs!高度) * Screen.TwipsPerPixelY)
    End If
    
    
'    '获取背景音乐文件
'    FrameDefault.MusicFile = ""
'
'    Set gRs = OpenRecord(gRs, "select B.类型,B.名称 from 咨询页面目录 A,咨询图片元素 B where A.背景音乐=B.序号 and A.页面序号=" & PageNo, Me.Caption)
'    If gRs.BOF = False Then
'        If IsNull(gRs!名称) = False Then FrameDefault.MusicFile = App.Path & "\图形\" & gRs!名称 & ".mid"
'    End If
                
    Call objDraw.InitLoad
    objDraw.ClientVisible = True
    
    StopFlatFlash
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub LoadDoctorMsg(ByVal PageNo As Long)
'功能:生成专家介绍页面内容
'参数:PageNo            页面序号
'说明:这是固定内容部份，是由ZLHIS9中提取的人员部份信息
    Dim FileName As String
    Dim W As Single
    Dim H As Single
    
    Dim vFont As New StdFont
    Dim i As Long
    Dim j As Long
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
    Dim vNextY As Single
    Dim vNextX As Single
    Dim objDraw As ctlQueryItem
    Dim vWidth As Single
    Dim vHeight As Single
    Dim vTmp As Single
    Dim vTmp1 As Single
    Dim vMaxWidth As Single
    
    On Error GoTo errHand
    i = 1
    vNextY = 60 + (i - 1) * 600
    vNextX = 120
    vMaxWidth = 120
    
    Set objDraw = FrameDefault.ClientObj
    Call objDraw.ClearAllPageItem
    
    gstrSQL = "select A.人员id,B.姓名||'('||D.名称||')' as 姓名 from 咨询专家清单 A,人员表 B,部门人员 C,部门表 D where B.ID=C.人员id And (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null) and C.部门id=D.ID and C.缺省=1 and A.人员ID=B.ID order by A.序号"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If gRs.BOF = False Then
        While Not gRs.EOF
            strTmp = "黑体;12;1;0;0"
            vFont.Name = Split(strTmp, ";")(0)
            vFont.Size = Val(Split(strTmp, ";")(1))
            vFont.Bold = Val(Split(strTmp, ";")(2))
            vFont.Italic = Val(Split(strTmp, ";")(3))
                                                            
            '1.加载标题内容及标题图标
            Call objDraw.AddPageItemTitle(i, vNextY, IIf(IsNull(gRs!姓名), "", gRs!姓名), Val(Split(strTmp, ";")(4)), vFont, "", PageNo, IIf(IsNull(gRs!人员ID), 0, gRs!人员ID), vWidth, vHeight, True, 0)
            vNextY = vNextY + vHeight + 150
            
            '2.照片和文字混合内容
            gstrSQL = "select A.姓名, A.个人简介 from 人员表 A where   (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) and A.ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(gRs!人员ID))
            strTmp = "宋体;12;0;0;0"
            
            FileName = ""
            vTmp = 0
            If rs.BOF = False Then
                If IsNull(rs!姓名) = False Then FileName = App.Path & "\图形\" & rs!姓名 & ".pic"
                If Dir(FileName) <> "" And FileName <> "" Then
                    
                    '以两寸照片大小显示或进行等比例缩小 高2940*0.6? 宽2280*0.6?   3.33 厘米
                    '照片是如何规定高度和宽度的?
                    
                    Call objDraw.AddPageItemPic(objDraw.NextPicIndex, 0, vNextY, FileName, vTmp1, vTmp, 1368, 1764)
                    
                End If
                                
                vWidth = FrameDefault.ClientWidth - vTmp1 - 120 - 120
                j = objDraw.NextTxtIndex
                Call objDraw.AddPageItemTxt(j, vNextX + vTmp1 + 60, vNextY, IIf(IsNull(rs!个人简介), "", rs!个人简介) & Chr(13) & Chr(10), strTmp, vWidth, vHeight)
                vNextY = vNextY + IIf(vTmp > vHeight, vTmp, vHeight) + 150
            End If
            
'            '8.设置返回页首标志
'
'            vHeight = 0
'            Call objDraw.AddReturnFlag(vNextX, vNextY, vHeight)
'            If vHeight > 0 Then vNextY = vNextY + vHeight + 150
'
            
            i = i + 1
            gRs.MoveNext
        Wend
    End If
    
    Call objDraw.ResizePage(FrameDefault.ClientWidth, vNextY)
    Call FrameDefault.InitNavigator(FrameDefault.ClientWidth, vNextY)
    
    gstrSQL = "select B.类型,B.名称,B.宽度,B.高度 from 咨询页面目录 A,咨询图片元素 B where A.页面背景=B.序号 and A.页面序号=[1]"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo)
    If gRs.BOF = False Then
        Call objDraw.BackPicture(IIf(IsNull(gRs!名称), "", App.Path & "\图形\" & gRs!名称 & IIf(gRs!类型 <> 2, ".pic", ".swf")), IIf(IsNull(gRs!宽度), 0, gRs!宽度) * Screen.TwipsPerPixelX, IIf(IsNull(gRs!高度), 0, gRs!高度) * Screen.TwipsPerPixelY)
    End If
    
'    '获取背景音乐文件
'    FrameDefault.MusicFile = ""
'
''    Call MusicClose
'    Set gRs = OpenRecord(gRs, "select B.类型,B.名称 from 咨询页面目录 A,咨询图片元素 B where A.背景音乐=B.序号 and A.页面序号=" & PageNo, Me.Caption)
'    If gRs.BOF = False Then
'        If IsNull(gRs!名称) = False Then FrameDefault.MusicFile = App.Path & "\图形\" & gRs!名称 & ".mid"
'    End If
                
    
    Call objDraw.InitLoad
        
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call MusicClose
    Set grs挂号诊室 = Nothing   '67045
End Sub

Private Sub FrameDefault_ExitNewQuery(blnCancel As Boolean)
    If GetPara("允许指令退出查询", "0") = "0" Then
        blnCancel = False
    Else
        blnCancel = True
        Unload Me
    End If
End Sub

Private Sub FrameDefault_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub FrameDefault_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'tmrHome.Enabled = False
    'tmrHome.Enabled = True
    'tmrHome.Interval = mvarlngHome
    
    mvarHomeInternal = 0
    mvarCheckConnectCounter = 0
    
End Sub

Private Sub FrameDefault_ShowPage(ByVal PageNo As Long, ByVal CusomFormat As String)
'功能:显示查询页面
'参数:PageNo            页面号
'     CusomFormat       格式,这里暂时只有两种，一是"专家介绍";二是"自定义"

    If CusomFormat = "" Then
        Call LoadPageItemList(PageNo)
    Else
        Call LoadDoctorMsg(PageNo)
    End If
    
End Sub

Private Sub tmrCheckConnect_Timer()
    mvarCheckConnectCounter = mvarCheckConnectCounter + 1
    If mvarCheckConnectCounter >= mvarCheckConnectInternal Then

        '检查数据库连接状态
        If gcnOracle.State = adStateOpen Then gcnOracle.Close

        Dim strErr As String
        
        If gobjRegister Is Nothing Then
            Set gobjRegister = gobjLogin.Register
        End If
        Set gcnOracle = gobjRegister.ReGetConnection(0, strErr)
        InitCommon gcnOracle
    End If
End Sub

Private Sub tmrHome_Timer()
    mvarHomeInternal = mvarHomeInternal + 1
    If mvarHomeInternal < mvarHomeLong Then Exit Sub
    mvarHomeInternal = 0
    
    On Error Resume Next
    Unload frmHelp
    Unload frmCardPass
    Unload frmSelect
    Unload frmIdentify泸州
    On Error GoTo 0
    
    Call FrameDefault.ShowHome
End Sub

Public Sub RefreshParamer(ByVal lngHomeLong As Long, ByVal lngCheckConnect As Long)
    mvarHomeLong = lngHomeLong
    tmrHome.Enabled = IIf(mvarHomeLong = 0, False, True)
    mvarCheckConnectInternal = lngCheckConnect
    
    'zyk add 200410
    Call FrameDefault.showwww
End Sub
