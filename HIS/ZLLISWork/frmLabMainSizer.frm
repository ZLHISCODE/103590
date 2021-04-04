VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "CO70B6~1.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmLabMainSizer 
   BackColor       =   &H00FDD6C6&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   3495
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   1965
      _Version        =   589884
      _ExtentX        =   3466
      _ExtentY        =   6165
      _StockProps     =   0
      ShowHeader      =   0   'False
   End
   Begin XtremeSuiteControls.ShortcutCaption ShortCaption 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2385
      _Version        =   589884
      _ExtentX        =   4207
      _ExtentY        =   503
      _StockProps     =   6
      Caption         =   "筛选条件"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.01
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   1
      Alignment       =   1
   End
End
Attribute VB_Name = "frmLabMainSizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mstrName As String
Private Const con_主界面筛选_检验中 As String = "门诊病人;住院病人;无主标本;已审标本;未审标本;体检病人;紧急医嘱;紧急标本;质控标本;审核已通过;审核未通过;未做完;已做完;仪器审核通过;仪器审核未通过"

Private Const con_主界面筛选_待核收 As String = "门诊病人;住院病人;体检病人"

Private Const con_frmLisStationWrite As String = "查看中文;查看原始结果;查看上次结果;查看标志;查看单位;查看参考;查看酶标;仪器提示;仪器审核标识"
Private Enum mCol
    选择 = 0
    名称
End Enum

Private Sub showData()
    Dim intLoop As Integer
    Dim lngLoop As Long
    Dim astrName() As String
    Dim Record As ReportRecord
    Dim strCheck As Boolean
    
    With Me.rptList.Columns
        
        rptList.AllowColumnRemove = False
        rptList.ShowItemsInGroups = False
        Me.rptList.Records.DeleteAll
        With rptList.PaintManager
            
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
'        rptList.SetImageList Imglist
        Set Column = .Add(mCol.选择, "", 30, False)
        Column.Icon = 0
        Set Column = .Add(mCol.名称, "名称", 120, False)
    End With
    
    Select Case mstrName
        Case "检验中"
            astrName = Split(con_主界面筛选_检验中, ";")
        Case "待核收"
            astrName = Split(con_主界面筛选_待核收, ";")
        Case "frmLisStationWrite"
            astrName = Split(con_frmLisStationWrite, ";")
    End Select
    
'    astrName = Split(mstrName, ";")
    
    For lngLoop = 0 To UBound(astrName)
        Set Record = Me.rptList.Records.Add
        For intLoop = 0 To Me.rptList.Columns.Count - 1
            Record.AddItem ""
        Next
        Record(mCol.选择).HasCheckbox = True
        Record(mCol.名称).Value = astrName(lngLoop)
        strCheck = zlDatabase.GetPara(mstrName & "_" & astrName(lngLoop), 100, 1208)
        Record(mCol.选择).Checked = strCheck
    Next
    Me.rptList.Populate

End Sub

Public Sub ShowME(Objfrm As Object, strName As String, blnShow As Boolean)
    mstrName = strName
    If blnShow = True Then
        Unload Me
    Else
        showData
        Me.Show modal, Objfrm
    End If
End Sub

Private Sub Form_Resize()
    Me.ShortCaption.Top = 50
    Me.ShortCaption.Left = 50
    Me.ShortCaption.Width = Me.ScaleWidth - 100
    Me.rptList.Top = Me.ShortCaption.Top + Me.ShortCaption.Height
    Me.rptList.Left = 50
    Me.rptList.Width = Me.ScaleWidth - 100
    Me.rptList.Height = Me.ScaleHeight - (Me.ShortCaption.Top + Me.ShortCaption.Height) - 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim intLoop As Integer
    '界面过滤列表
    
    For intLoop = 0 To Me.rptList.Rows.Count
        zlDatabase.SetPara mstrName & "_" & Me.rptList.Rows(intLoop).Record(mCol.名称).Value, _
            Me.rptList.Rows(intLoop).Record(mCol.选择).Checked, 100, 1208
    Next
    mstrName = ""
End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Row.Record(mCol.选择).Checked = Not Row.Record(mCol.选择).Checked
    Me.rptList.Populate
End Sub



