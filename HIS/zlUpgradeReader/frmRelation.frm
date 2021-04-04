VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmRelation 
   Caption         =   "关联问题"
   ClientHeight    =   7560
   ClientLeft      =   165
   ClientTop       =   525
   ClientWidth     =   12060
   Icon            =   "frmRelation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   12060
   StartUpPosition =   2  '屏幕中心
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   5310
      Left            =   330
      TabIndex        =   0
      Top             =   1305
      Width           =   7260
      _Version        =   589884
      _ExtentX        =   12806
      _ExtentY        =   9366
      _StockProps     =   0
      ShowGroupBox    =   -1  'True
      AutoColumnSizing=   0   'False
   End
   Begin VB.PictureBox picRight 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4875
      Left            =   8250
      ScaleHeight     =   4875
      ScaleWidth      =   3075
      TabIndex        =   5
      Top             =   360
      Width           =   3075
      Begin VB.TextBox txt需求 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   75
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   120
         Width           =   2880
      End
   End
   Begin VB.PictureBox pic说明 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2010
      Left            =   7890
      ScaleHeight     =   2010
      ScaleWidth      =   2550
      TabIndex        =   3
      Top             =   2760
      Width           =   2550
      Begin VB.TextBox txt说明 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   60
         Width           =   2310
      End
   End
   Begin VB.PictureBox pic风险评估 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2010
      Left            =   7920
      ScaleHeight     =   2010
      ScaleWidth      =   2550
      TabIndex        =   1
      Top             =   4995
      Width           =   2550
      Begin VB.TextBox txt风险评估 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   60
         Width           =   2310
      End
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   1485
      Top             =   945
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelation.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelation.frx":0924
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelation.frx":0EBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelation.frx":1258
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelation.frx":17F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelation.frx":1D8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelation.frx":416E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelation.frx":6550
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelation.frx":8932
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelation.frx":AD14
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelation.frx":B0AE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmRelation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFileName As String  'excel文件名
Private mstrRelations As String '问题编号，以","分隔

Private Const Dkp_ID_Rept As Integer = 3
Private Const Dkp_ID_Right As Integer = 4
Private Const Dkp_ID_说明 As Integer = 5
Private Const Dkp_ID_风险 As Integer = 6
Private rowLink As ReportRow        '当前超连接焦点行
Private mlngItemID As Long

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.Id = Dkp_ID_Rept Then
        Item.Handle = rptList.hwnd
    ElseIf Item.Id = Dkp_ID_Right Then
        Item.Handle = picRight.hwnd
    ElseIf Item.Id = Dkp_ID_说明 Then
        Item.Handle = pic说明.hwnd
    ElseIf Item.Id = Dkp_ID_风险 Then
        Item.Handle = pic风险评估.hwnd
    End If
End Sub

Private Sub Form_Load()
    Call initDockPane
    Call initRptList(rptList, ImgList, txt需求.Font, False)
End Sub

Private Sub picRight_Resize()
    On Error Resume Next
    With Me.txt需求
        .Left = picRight.ScaleLeft
        .Top = picRight.ScaleTop
        .Width = picRight.ScaleWidth - 45
        .Height = picRight.ScaleHeight - 45
    End With
End Sub

Private Sub pic风险评估_Resize()
    On Error Resume Next
    With Me.txt风险评估
        .Left = pic风险评估.ScaleLeft
        .Top = pic风险评估.ScaleTop
        .Width = pic风险评估.ScaleWidth - 45
        .Height = pic风险评估.ScaleHeight - 45
    End With
End Sub

Private Sub pic说明_Resize()
    On Error Resume Next
    With Me.txt说明
        .Left = pic说明.ScaleLeft
        .Top = pic说明.ScaleTop
        .Width = pic说明.ScaleWidth - 45
        .Height = pic说明.ScaleHeight - 45
    End With
End Sub

Private Sub rptList_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
     Dim RecordItem As ReportRecordItem
    If (Row.Record(mCol.影响评估).Value = 未填写) Then
        For Each RecordItem In Row.Record
            RecordItem.Bold = True
        Next
    Else
        For Each RecordItem In Row.Record
            RecordItem.Bold = False
        Next
    End If
        
    If (Item.Index = mCol.风险) Then
        Select Case Item.Value
            Case 0: Item.Icon = ICON_Unknown    '不确定
            Case 1: Item.Icon = ICON_Low        '低
            Case 2: Item.Icon = ICON_Center     '中
            Case 3: Item.Icon = ICON_High       '高
        End Select
    End If
    
    If (Item.Index = mCol.编号) Then
        If Row.Record(mCol.连接).Value = "是" Then
            Set Metrics.Font = fntUnderLine
            Metrics.ForeColor = vbBlue
        End If
    End If
End Sub

Private Sub rptList_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim strLinkFile As String
    If Button = 1 Then
        
        If (Not rowLink Is Nothing) Then
            If rowLink.Record(mCol.连接).Value = "是" Then
                strLinkFile = Mid(mstrFileName, 1, InStrRev(mstrFileName, "\")) & "Document\" & rowLink.Record(mCol.编号).Value & ".htm"
                If Dir(strLinkFile) <> "" Then
                    Call ShellExecute(Me.hwnd, "open", "file:///" & Replace(strLinkFile, "\", "/"), vbNullString, vbNullString, 1)
                Else
                    MsgBox "未找到对应的html文件，打开文件失败！", vbInformation, gstrSysname
                End If
            End If
        End If
    End If
End Sub

Private Sub rptList_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    
    Dim htInfo As ReportHitTestInfo
    Set htInfo = rptList.HitTest(X, Y)
    
    Dim objRow As ReportRow

    If (Not htInfo.Item Is Nothing) Then
        
        If (htInfo.Item.Index = mCol.编号) Then
            Set objRow = htInfo.Row
        End If
    End If

    If (Not objRow Is rowLink) Then
        If (Not objRow Is Nothing) Then
            If objRow.Record(mCol.连接).Value = "是" Then
                objRow.Record(mCol.编号).BackColor = RGB(255, 238, 99)
            End If
            
        End If
        
        If (Not rowLink Is Nothing) Then
            rowLink.Record(mCol.编号).BackColor = -1
        End If
        
        Set rowLink = objRow
        rptList.Redraw
    End If
    

End Sub

Private Sub rptList_SelectionChanged()
    '#
    If rptList.FocusedRow Is Nothing Then Exit Sub
    If Not rptList.FocusedRow.GroupRow Then
        txt说明 = rptList.FocusedRow.Record(mCol.说明).Value
        txt需求 = rptList.FocusedRow.Record(mCol.需求).Value
        txt风险评估 = rptList.FocusedRow.Record(mCol.风险评估).Value
    Else
        txt说明 = ""
        txt需求 = ""
        txt风险评估 = ""
    End If
    
End Sub

Public Sub ShowRelation(ByVal strFilename As String, strRelations As String)
    '功能：显示关联问题
    '参数:
    '   StrFileName: Excel文件名
    '   strRelations:问题编号，以","分隔
    If strFilename = "" Or strRelations = "" Then Exit Sub
    mstrFileName = strFilename: mstrRelations = strRelations
    Me.Show
    Call LoadSheet(mstrFileName)
End Sub

'-------------
Private Sub initDockPane()
    Dim paneTree As Pane, paneFind As Pane, paneEdit As Pane, paneList As Pane, paneRight As Pane, pane说明 As Pane, pane风险 As Pane
    
   ' Me.dkpMain.SetCommandBars Me.cbsMenu
    Me.dkpMain.Options.UseSplitterTracker = False '实时拖动
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    
    Me.dkpMain.Options.HideClient = True
    
    Set paneList = Me.dkpMain.CreatePane(Dkp_ID_Rept, 900, 700, DockTopOf, Nothing)
    paneList.Title = "问题清单"
    paneList.Handle = Me.rptList.hwnd
    paneList.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set pane说明 = Me.dkpMain.CreatePane(Dkp_ID_说明, 800, 500, DockBottomOf, paneList)
    pane说明.Title = "修改说明"
    pane说明.Handle = Me.pic说明.hwnd
    pane说明.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set paneRight = Me.dkpMain.CreatePane(Dkp_ID_Right, 100, 300, DockBottomOf, pane说明)
    paneRight.Title = "用户需求"
    paneRight.Handle = Me.picRight.hwnd
    paneRight.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set pane风险 = Me.dkpMain.CreatePane(Dkp_ID_风险, 100, 300, DockRightOf, paneRight)
    pane风险.Title = "风险评估"
    pane风险.Handle = Me.pic风险评估.hwnd
    pane风险.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
End Sub

Private Sub LoadSheet(ByVal Filename As String)

    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim rptItem1 As ReportRecordItem
    Dim rptItem2 As ReportRecordItem
    Dim rptRow As ReportRow
    Dim rptColum As ReportColumn
    Dim blnAdd As Boolean
    Dim strSheet As String, varSheet As Variant, i As Integer
    Dim str编号 As String
    Dim rsSheet As ADODB.Recordset
    
    On Error GoTo errHandle
    If Filename = "" Then Exit Sub
    strSheet = OpenExcelFile(Filename)
    
    If strSheet = "" Then Exit Sub
    
    rptList.Records.DeleteAll '清空原列表
    
    If InStr(strSheet, "|") <= 0 Then Exit Sub
    
    varSheet = Split(strSheet, "|")
    For i = LBound(varSheet) To UBound(varSheet)
        Set rsSheet = OpenExcelSheet(varSheet(i))
        Do Until rsSheet.EOF
            
            '加入明细
            With rptList
                str编号 = "" & rsSheet(Excel_Col.问题编号).Value
                
                If InStr("," & mstrRelations & ",", "," & str编号 & ",") > 0 Then '符合条件则加入
                    Set rptRcd = Me.rptList.Records.Add()
                    
                    '已读 = 0: 风险: 培训: 版本: 分类: 编号: 模块: 影响模块: 风险评估: 用户: 需求: 说明: 关联问题: 备注: 影响评估: 连接
                    Set rptItem = rptRcd.AddItem(""): rptItem.Focusable = False
                    If Val("" & rsSheet(Excel_Col.对用户影响评估).Value) = 0 Then
                        rptItem.Icon = ICON_NoRead
                    Else
                        rptItem.Icon = ICON_Read
                    End If
                        
                    If "" & rsSheet(Excel_Col.问题风险).Value = "高" Then
                        Set rptItem1 = rptRcd.AddItem(3)
                    ElseIf "" & rsSheet(Excel_Col.问题风险).Value = "中" Then
                        Set rptItem1 = rptRcd.AddItem(2)
                    ElseIf "" & rsSheet(Excel_Col.问题风险).Value = "低" Then
                        Set rptItem1 = rptRcd.AddItem(1)
                    Else
                        Set rptItem1 = rptRcd.AddItem(0)
                    End If
                    rptItem1.Caption = " ": rptItem1.Focusable = False

                    Set rptItem = rptRcd.AddItem(CStr("" & rsSheet(Excel_Col.发布版本).Value)): rptItem.Focusable = False
                    
                    Set rptItem = rptRcd.AddItem(CStr(Replace(varSheet(i), "$", ""))): rptItem.Focusable = False
                    Set rptItem = rptRcd.AddItem(CStr("" & rsSheet(Excel_Col.问题编号).Value)): rptItem.Focusable = False
                   
                    Set rptItem = rptRcd.AddItem(CStr("" & rsSheet(Excel_Col.登记模块).Value)): rptItem.Focusable = False
                    
                    Set rptItem = rptRcd.AddItem(CStr("" & rsSheet(Excel_Col.影响模块).Value)): rptItem.Focusable = False
                     Set rptItem = rptRcd.AddItem(CStr("" & rsSheet(Excel_Col.影响类型).Value)): rptItem.Focusable = False
                     
                    Set rptItem = rptRcd.AddItem(CStr("" & rsSheet(Excel_Col.风险评估说明).Value)): rptItem.Focusable = False
                    Set rptItem = rptRcd.AddItem(CStr("" & rsSheet(Excel_Col.登记用户).Value)): rptItem.Focusable = False
                    Set rptItem = rptRcd.AddItem(CStr("" & rsSheet(Excel_Col.用户需求).Value)): rptItem.Focusable = False
                    Set rptItem = rptRcd.AddItem(CStr("" & rsSheet(Excel_Col.修改说明).Value)): rptItem.Focusable = False
                    Set rptItem = rptRcd.AddItem(CStr("" & rsSheet(Excel_Col.相关问题).Value)): rptItem.Focusable = False
                    Set rptItem = rptRcd.AddItem(CStr("" & rsSheet(Excel_Col.备注).Value)): rptItem.Focusable = False
                    
                    '---- 用户操作项
'                    Set rptItem2 = rptRcd.AddItem("")
                    If "" & rsSheet(Excel_Col.是否需要培训) = "否" Then
'                         rptItem2.Icon=-1
                        Set rptItem2 = rptRcd.AddItem("") '无需培训
                        
                    Else
                        If "" & rsSheet(Excel_Col.操作培训情况).Value = "" Then
'                            rptItem2.Icon = ICON_UnTrain
                            Set rptItem2 = rptRcd.AddItem("0-未填写")

                        Else
'                            rptItem2.Icon = ICON_Train
                            Set rptItem2 = rptRcd.AddItem("1-已培训")
       
                        End If
                    End If
                    
                    rptRcd.AddItem Val(CStr("" & rsSheet(Excel_Col.对用户影响评估).Value))   '0-未填写 1-正面影响 2-负面影响 3-无影响
                    
                    '--- 超链接,是否修改
                    rptRcd.AddItem CStr("" & rsSheet(Excel_Col.是否有HTML文档).Value)
                    rptRcd.AddItem CStr("0")
                    rptRcd.PreviewText = "" & rsSheet(Excel_Col.用户需求).Value
                End If
                
            End With
            
            rsSheet.MoveNext
        Loop
        
    Next
    Set rsSheet = Nothing
        
    '定位到上次选中行
    If mlngItemID <> 0 Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If Val(rptRow.Record(mCol.编号).Value) = mlngItemID Then
                    Set Me.rptList.FocusedRow = rptRow
                    Exit For
                End If
            End If
        Next
    End If
    
    '展开选中行
    If Me.rptList.FocusedRow Is Nothing And Me.rptList.Rows.Count > 0 Then
        If Me.rptList.Rows(0).GroupRow Then
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0).Childs(0)
        Else
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        End If
    End If
    
    rptList.Populate
    Call rptList_SelectionChanged '触发选择事件
    Exit Sub
errHandle:
    MsgBox Err.Number & " " & Err.Description, vbQuestion, gstrSysname
    
End Sub


