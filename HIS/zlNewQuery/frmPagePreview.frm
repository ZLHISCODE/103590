VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPagePreview 
   Caption         =   "页面预览"
   ClientHeight    =   6165
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   9330
   Icon            =   "frmPagePreview.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBack 
      BackColor       =   &H8000000C&
      Height          =   4275
      Left            =   75
      ScaleHeight     =   4215
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   45
      Width           =   6015
      Begin VB.PictureBox picBack1 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   4125
         ScaleHeight     =   495
         ScaleWidth      =   570
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   3240
         Width           =   570
      End
      Begin VB.HScrollBar hsb 
         Height          =   255
         Left            =   285
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   3870
         Width           =   2010
      End
      Begin VB.VScrollBar vsb 
         Height          =   3990
         Left            =   5670
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   -15
         Width           =   255
      End
      Begin zl9NewQuery.ctlQueryItem QueryItem 
         Height          =   2820
         Left            =   1365
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   285
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   4974
      End
   End
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   6960
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":06EA
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":090A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":0B2A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":0D4A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":0F6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":14C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":1A1E
            Key             =   "View"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":1C3A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":1E5A
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   7545
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":207A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":229A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":24BA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":26DA
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":28FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":2E54
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":33AE
            Key             =   "View"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":35CA
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":37EA
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPagePreview"
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

Private Sub Form_Activate()
    If mblnFist = False Then Exit Sub
    mblnFist = False
            
    DoEvents
    Call LoadPageItemList(mvarPageNo)
        
    Call CalcVsb
End Sub

Private Sub Form_Load()
    mblnFist = True
    RestoreWinState Me, App.ProductName
    
    QueryItem.Height = Screen.Height
End Sub

Private Sub Form_Resize()
    '根据窗体状态,调整窗体中各控件的显示位置

    QueryItem.Width = Screen.Width - 2010 - 45
    Call ResizeControl(picBack, 0, 0, Me.ScaleWidth, Me.ScaleHeight)
    Call ResizeControl(QueryItem, (Me.ScaleWidth - QueryItem.Width) / 2, 45, QueryItem.Width, QueryItem.Height)
    
    If QueryItem.Left < 45 Then QueryItem.Left = 45
    mvarLeftStart = QueryItem.Left
    
    Call ResizeControl(vsb, picBack.ScaleWidth - vsb.Width, 0, vsb.Width, picBack.ScaleHeight - hsb.Height)
    Call ResizeControl(hsb, 0, picBack.ScaleHeight - hsb.Height, picBack.ScaleWidth - vsb.Width, hsb.Height)
    picBack1.Left = vsb.Left
    picBack1.Top = hsb.Top
    
    Call CalcVsb
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Public Function ShowPreview(frmMain As Object, ByVal PageNo As Long)
    mvarPageNo = PageNo
    frmPagePreview.Show 1, frmMain
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
                gstrSQL = "select C.页面名称||decode(B.标题文本,NULL,'','：'||B.标题文本) as 标题文本,A.链接页面,A.页内段号 from 咨询段落链接 A,咨询段落目录 B,咨询页面目录 C Where A.链接页面=C.页面序号 and A.链接页面=B.页面序号(+) and A.页内段号=B.段落序号(+) and A.页面序号 = [1] And A.段落序号 =[2] "
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
                    gstrSQL = "select B.姓名,A.链接页面,A.页内段号 from 咨询段落链接 A,人员表 B Where A.页内段号=B.id And (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null) and A.页面序号 = [1] And A.段落序号 = [2] "
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
    On Error Resume Next
    QueryItem.Left = mvarLeftStart - hsb.Value * 600
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

Private Sub vsb_Change()
    On Error Resume Next
    QueryItem.Top = 45 - vsb.Value * 600
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
