VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   5535
   ClientLeft      =   1845
   ClientTop       =   990
   ClientWidth     =   7875
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "关闭帮助"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6465
      MouseIcon       =   "frmHelp.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   4920
      Width           =   1200
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00FFFFFF&
      Height          =   4800
      Left            =   75
      ScaleHeight     =   4740
      ScaleWidth      =   7665
      TabIndex        =   0
      Top             =   45
      Width           =   7725
      Begin VB.VScrollBar vsb 
         Height          =   3990
         Left            =   7020
         MouseIcon       =   "frmHelp.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   465
         Width           =   345
      End
      Begin VB.HScrollBar hsb 
         Height          =   330
         Left            =   90
         MouseIcon       =   "frmHelp.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   4350
         Width           =   2010
      End
      Begin VB.PictureBox picBack1 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   6360
         ScaleHeight     =   885
         ScaleWidth      =   1050
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   3780
         Width           =   1050
      End
      Begin zl9NewQuery.ctlQueryItem QueryItem 
         Height          =   2820
         Left            =   2100
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   375
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   4974
      End
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "这是帮助查看，您必须按右边的[关闭帮助]来退出。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   150
      TabIndex        =   6
      Top             =   5025
      Width           =   5520
   End
   Begin VB.Shape shp 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   5475
      Left            =   30
      Top             =   15
      Width           =   7800
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnFist As Boolean
Private mvarPageNo As Long
Private mvarSvrDept As String           '保存增加医生的科室
Private mvarSvrDuty As String           '保存增加医生的职务

Private mvarLeftStart As Single

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call InitInternal
End Sub

Private Sub Form_Activate()
    If mblnFist = False Then Exit Sub
    mblnFist = False
            
    DoEvents
    
    Call LoadPageItemList(mvarPageNo)
        
    Call CalcVsb
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call picBack_KeyDown(KeyCode, Shift)
End Sub

Private Sub Form_Load()
    mblnFist = True

    
    QueryItem.Height = Screen.Height
End Sub

Private Sub Form_Resize()
    '根据窗体状态,调整窗体中各控件的显示位置
    
    On Error Resume Next
    
    QueryItem.Width = Screen.Width - 2010 - 45
    Call ResizeControl(shp, 15, 15, Me.ScaleWidth - 30, Me.ScaleHeight - 30)
    
    Call ResizeControl(picBack, 45, 45, Me.ScaleWidth - 90, Me.ScaleHeight - cmdClose.Height - 120)
    Call ResizeControl(QueryItem, 0, 0, QueryItem.Width, QueryItem.Height)
    
    mvarLeftStart = QueryItem.Left
    
    Call ResizeControl(vsb, picBack.ScaleWidth - vsb.Width + 60, 0, vsb.Width, picBack.ScaleHeight - hsb.Height + 60)
    Call ResizeControl(hsb, 0, picBack.ScaleHeight - hsb.Height + 60, picBack.ScaleWidth - vsb.Width + 60, hsb.Height)
    picBack1.Left = vsb.Left
    picBack1.Top = hsb.Top
    
    Call ResizeControl(cmdClose, Me.ScaleWidth - cmdClose.Width - 60, picBack.Top + picBack.Height + 30, cmdClose.Width, cmdClose.Height)
    lbl.Top = cmdClose.Top + 75
    Call CalcVsb
End Sub

Public Function ShowHelp(frmMain As Object, ByVal PageNo As Long, ByVal vWidth As Single, ByVal vHeight As Single)
    mvarPageNo = PageNo
    frmHelp.Width = vWidth
    frmHelp.Height = vHeight
    frmHelp.Show 1, frmMain
End Function

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
    
    Set objDraw = QueryItem
    objDraw.ClientVisible = False
    Call objDraw.ClearAllPageItem
    
    '读取页面的背景及广告条幅
'    Set gRs = OpenRecord(gRs, "select B.类型,B.名称 from 咨询页面目录 A,咨询图片元素 B where A.宣传标语=B.序号 and A.页面序号=" & PageNo)
'    If gRs.BOF = False Then FrameDefault.AdviceMovie = IIf(IsNull(gRs!名称), "", App.Path & "\图形\" & gRs!名称 & IIf(gRs!类型 <> 2, ".pic", ".swf"))
                    
    '开始生成自定义查询页面
    gstrSQL = "select 页面序号,段落序号,标题文本,标题图标,标题隐藏,标题位置,标题字体,返回页首,段落类型,段落字体,插表序号,插表位置,插图序号,插图位置 from 咨询段落目录 where 页面序号=[1] order by 段落序号"
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
            Case 0          '纯文本内容
                strTmp = IIf(IsNull(gRs!段落字体), "宋体;12;0;0;0", gRs!段落字体)
                j = objDraw.NextTxtIndex
                
                vWidth = QueryItem.Width - 330
                
                strText = Sys.ReadLob(glngSys, 29, PageNo & "," & gRs("段落序号").Value, "", 1)
                Call objDraw.AddPageItemTxt(j, vNextX, vNextY, strText & Chr(13) & Chr(10), strTmp, vWidth, vHeight)
                vNextY = vNextY + vHeight + 150
            '----------------------------------------------------------------------------------------------------------
            Case 1          '纯表格内容
                vHeight = 0
                Call InsertGrid(objDraw, IIf(IsNull(gRs!插表序号), 0, gRs!插表序号), vNextX, vNextY, vWidth, vHeight)
                If vHeight > 0 Then vNextY = vNextY + vHeight + 150
            '----------------------------------------------------------------------------------------------------------
            Case 2          '纯图形内容
                FileName = GetFileName(IIf(IsNull(gRs!插图序号), 0, gRs!插图序号), W, H)
                Call objDraw.AddPageItemPic(objDraw.NextPicIndex, 0, vNextY, FileName, vWidth, vHeight, W, H)
                vNextY = vNextY + vHeight + 150
            '----------------------------------------------------------------------------------------------------------
            Case 3          '纯链接内容
                gstrSQL = "select C.页面名称||decode(B.标题文本,NULL,'','：'||B.标题文本) as 标题文本,A.链接页面,A.页内段号 from 咨询段落链接 A,咨询段落目录 B,咨询页面目录 C Where A.链接页面=C.页面序号 and A.链接页面=B.页面序号(+) and A.页内段号=B.段落序号(+) and A.页面序号 = [1] And A.段落序号 = [2]"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo, Val(IIf(IsNull(gRs!段落序号), 0, gRs!段落序号)))
                If rs.BOF = False Then
                    While Not rs.EOF
                        Call objDraw.AddPageItemConnect(objDraw.NextConnectIndex, vNextX + 150, vNextY, IIf(IsNull(rs!标题文本), "", rs!标题文本), IIf(IsNull(rs!链接页面), 0, rs!链接页面), IIf(IsNull(rs!页内段号), 0, rs!页内段号), vWidth, vHeight)
                        vNextY = vNextY + 300
                        rs.MoveNext
                    Wend
                    vNextY = vNextY + 150
                Else
                    '检查是否连接到ZLHIS的人员
                    gstrSQL = "select B.姓名,A.链接页面,A.页内段号 from 咨询段落链接 A,人员表 B Where A.页内段号=B.id And (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null) and A.页面序号 = [1] And A.段落序号 = [2]"
                    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo, Val(IIf(IsNull(gRs!段落序号), 0, gRs!段落序号)))
                    If rs.BOF = False Then
                        While Not rs.EOF
                            Call objDraw.AddPageItemConnect(objDraw.NextConnectIndex, vNextX, vNextY, IIf(IsNull(rs!姓名), "", rs!姓名), IIf(IsNull(rs!链接页面), 0, rs!链接页面), IIf(IsNull(rs!页内段号), 0, rs!页内段号), vWidth, vHeight)
                            vNextY = vNextY + 300
                            rs.MoveNext
                        Wend
                        vNextY = vNextY + 150
                    End If
                End If
            '----------------------------------------------------------------------------------------------------------
            Case 4          '文本和表格
                strTmp = IIf(IsNull(gRs!段落字体), "宋体;12;0;0;0", gRs!段落字体)
                j = objDraw.NextTxtIndex
                
                strText = Sys.ReadLob(glngSys, 29, PageNo & "," & gRs("段落序号").Value, "", 1)
                
                Select Case IIf(IsNull(gRs!插表位置), 0, gRs!插表位置)
                Case 0
                    vHeight = 0
                    Call InsertGrid(objDraw, IIf(IsNull(gRs!插表序号), 0, gRs!插表序号), 0, vNextY, vTmp1, vTmp)
                    vWidth = QueryItem.Width - vTmp1 - 120 - 120
                    Call objDraw.AddPageItemTxt(j, vNextX + vTmp1 + 60, vNextY, strText & Chr(13) & Chr(10), strTmp, vWidth, vHeight)
                Case 1
                    Call objDraw.AddPageItemTxt(j, vNextX, vNextY, strText & Chr(13) & Chr(10), strTmp, vTmp1, vHeight)
                    Call InsertGrid(objDraw, IIf(IsNull(gRs!插表序号), 0, gRs!插表序号), 1, vNextY, vWidth, vTmp)
                End Select
                vNextY = vNextY + IIf(vTmp > vHeight, vTmp, vHeight) + 150
            '----------------------------------------------------------------------------------------------------------
            Case 5          '文本和图形
                FileName = GetFileName(IIf(IsNull(gRs!插图序号), 0, gRs!插图序号), W, H)
                strText = Sys.ReadLob(glngSys, 29, PageNo & "," & gRs("段落序号").Value, "", 1)
                strTmp = IIf(IsNull(gRs!段落字体), "宋体;12;0;0;0", gRs!段落字体)
                j = objDraw.NextTxtIndex
                Select Case IIf(IsNull(gRs!插图位置), 0, gRs!插图位置)
                Case 0
                    Call objDraw.AddPageItemPic(objDraw.NextPicIndex, 0, vNextY, FileName, vTmp1, vTmp, W, H)
                    vWidth = QueryItem.Width - vTmp1 - 120 - 120
                    Call objDraw.AddPageItemTxt(j, vNextX + vTmp1 + 60, vNextY, strText & Chr(13) & Chr(10), strTmp, vWidth, vHeight)
                Case 1
                    Call objDraw.AddPageItemPic(objDraw.NextPicIndex, 1, vNextY, FileName, vWidth, vTmp, W, H)
                    vTmp1 = QueryItem.Width - vWidth - 60 - 90
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
        
    Call objDraw.ResizePage(QueryItem.Width, vNextY)
    QueryItem.Height = QueryItem.FactHeight
    'Call FrameDefault.InitNavigator(FrameDefault.ClientWidth, vNextY)
    
    '获取背景并画出页面背景
    gstrSQL = "select B.类型,B.名称,B.宽度,B.高度 from 咨询页面目录 A,咨询图片元素 B where A.页面背景=B.序号 and A.页面序号=[1]"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo)
    If gRs.BOF = False Then
        Call objDraw.BackPicture(IIf(IsNull(gRs!名称), "", App.Path & "\图形\" & gRs!名称 & IIf(gRs!类型 <> 2, ".pic", ".swf")), IIf(IsNull(gRs!宽度), 0, gRs!宽度) * Screen.TwipsPerPixelX, IIf(IsNull(gRs!高度), 0, gRs!高度) * Screen.TwipsPerPixelY)
    End If
            
    Call objDraw.InitLoad
    objDraw.ClientVisible = True
    
    StopFlatFlash
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub hsb_Change()
    QueryItem.Left = mvarLeftStart - hsb.Value * 600
    If QueryItem.Left + QueryItem.Width < picBack.Left + picBack.Width - vsb.Width Then
        QueryItem.Left = picBack.Left + picBack.Width - QueryItem.Width - vsb.Width
    End If
    If QueryItem.Left > 0 Then QueryItem.Left = 0
    
End Sub

Private Sub hsb_KeyDown(KeyCode As Integer, Shift As Integer)
    Call picBack_KeyDown(KeyCode, Shift)
End Sub

Private Sub picBack_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If vsb.Enabled Then vsb.Value = IIf(vsb.Value < vsb.Max, vsb.Value + 1, vsb.Max)
    End If

    If KeyCode = vbKeyUp Then
        If vsb.Enabled Then vsb.Value = IIf(vsb.Value > 0, vsb.Value - 1, 0)
    End If

    If KeyCode = vbKeyRight Then
        If hsb.Enabled Then hsb.Value = IIf(hsb.Value < hsb.Max, hsb.Value + 1, hsb.Max)
    End If

    If KeyCode = vbKeyLeft Then
        If hsb.Enabled Then hsb.Value = IIf(hsb.Value > 0, hsb.Value - 1, 0)
    End If
End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call InitInternal
End Sub

Private Sub picBack_Paint()
    Call RaisEffect(picBack, -1)
End Sub

Private Sub picBack1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call picBack_KeyDown(KeyCode, Shift)
End Sub

Private Sub picBack1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call InitInternal
End Sub

Private Sub QueryItem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call InitInternal
End Sub

Private Sub vsb_Change()
    QueryItem.Top = 0 - vsb.Value * 600
    If QueryItem.Top + QueryItem.Height < picBack.Top + picBack.Height - hsb.Height Then
        QueryItem.Top = picBack.Top + picBack.Height - hsb.Height - QueryItem.Height
    End If
    If QueryItem.Top > 0 Then QueryItem.Top = 0
    
End Sub

Private Sub CalcVsb()
    vsb.Max = 0 - Int(0 - (QueryItem.Height - picBack.ScaleHeight + hsb.Height + 45) / 600)
    If vsb.Max > 0 Then
        vsb.Enabled = True
        vsb.SmallChange = 1
        vsb.LargeChange = 1
        vsb.Value = 0
    Else
        vsb.Enabled = False
    End If
    
    hsb.Max = 0 - Int(0 - (QueryItem.Width - picBack.ScaleWidth + vsb.Width + 45) / 600)
    If hsb.Max > 0 Then
        hsb.Enabled = True
        hsb.SmallChange = 1
        hsb.LargeChange = 1
        hsb.Value = 0
    Else
        hsb.Enabled = False
    End If
End Sub

Private Sub vsb_KeyDown(KeyCode As Integer, Shift As Integer)
    Call picBack_KeyDown(KeyCode, Shift)
End Sub
