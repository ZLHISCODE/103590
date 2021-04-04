VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmLabAnalyseData 
   Caption         =   "分析数据收集"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11760
   Icon            =   "frmLabAnalyseData.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   11760
   StartUpPosition =   1  '所有者中心
   Begin XtremeReportControl.ReportControl rptSource 
      Height          =   2865
      Left            =   60
      TabIndex        =   19
      Top             =   1530
      Width           =   3975
      _Version        =   589884
      _ExtentX        =   7011
      _ExtentY        =   5054
      _StockProps     =   0
      BorderStyle     =   2
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      ShowItemsInGroups=   -1  'True
      AutoColumnSizing=   0   'False
   End
   Begin XtremeReportControl.ReportControl rptAnalyse 
      Height          =   2865
      Left            =   6660
      TabIndex        =   20
      Top             =   1530
      Width           =   3975
      _Version        =   589884
      _ExtentX        =   7011
      _ExtentY        =   5054
      _StockProps     =   0
      BorderStyle     =   2
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      ShowItemsInGroups=   -1  'True
      AutoColumnSizing=   0   'False
   End
   Begin VB.CommandButton cmdRightAll 
      Caption         =   ">>>"
      Height          =   435
      Left            =   5970
      TabIndex        =   18
      Top             =   4530
      Width           =   525
   End
   Begin VB.CommandButton cmdRight 
      Caption         =   "==>"
      Height          =   435
      Left            =   5970
      TabIndex        =   17
      Top             =   3570
      Width           =   525
   End
   Begin VB.CommandButton cmdLeftAll 
      Caption         =   "<<<"
      Height          =   435
      Left            =   5970
      TabIndex        =   16
      Top             =   2790
      Width           =   525
   End
   Begin VB.CommandButton CmdLeft 
      Caption         =   "<=="
      Height          =   435
      Left            =   5970
      TabIndex        =   15
      Top             =   1980
      Width           =   525
   End
   Begin VB.Frame fraFilter 
      Caption         =   "查询条件"
      Height          =   1095
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   12135
      Begin VB.CommandButton cmdFind 
         Caption         =   "查询"
         Height          =   345
         Left            =   9270
         TabIndex        =   12
         Top             =   210
         Width           =   1065
      End
      Begin VB.ComboBox cbo分析目的 
         Height          =   300
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   630
         Width           =   3285
      End
      Begin VB.ComboBox cbo仪器 
         Height          =   300
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   630
         Width           =   3735
      End
      Begin MSComCtl2.DTPicker DTPStart 
         Height          =   285
         Left            =   5640
         TabIndex        =   6
         Top             =   255
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         _Version        =   393216
         Format          =   95748097
         CurrentDate     =   39769
      End
      Begin VB.TextBox txt标本号 
         Height          =   315
         Left            =   2850
         TabIndex        =   4
         Top             =   240
         Width           =   1785
      End
      Begin VB.TextBox txt批号 
         Height          =   315
         Left            =   930
         TabIndex        =   3
         Top             =   240
         Width           =   1245
      End
      Begin MSComCtl2.DTPicker DTPEnd 
         Height          =   285
         Left            =   7440
         TabIndex        =   7
         Top             =   255
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         _Version        =   393216
         Format          =   95748097
         CurrentDate     =   39769
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "分析目的"
         Height          =   180
         Left            =   4800
         TabIndex        =   10
         Top             =   690
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "仪    器"
         Height          =   180
         Left            =   150
         TabIndex        =   8
         Top             =   690
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "核收时间                  ---"
         Height          =   180
         Left            =   4800
         TabIndex        =   5
         Top             =   300
         Width           =   2610
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "标本批号"
         Height          =   180
         Left            =   150
         TabIndex        =   2
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "标本号"
         Height          =   180
         Left            =   2250
         TabIndex        =   1
         Top             =   300
         Width           =   540
      End
   End
   Begin XtremeSuiteControls.ShortcutCaption ShortCaptAnalyse 
      Height          =   315
      Left            =   6690
      TabIndex        =   14
      Top             =   1200
      Width           =   2895
      _Version        =   589884
      _ExtentX        =   5106
      _ExtentY        =   556
      _StockProps     =   6
      Caption         =   "分析数据"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   1
   End
   Begin XtremeSuiteControls.ShortcutCaption ShortCapSource 
      Height          =   315
      Left            =   60
      TabIndex        =   13
      Top             =   1200
      Width           =   2895
      _Version        =   589884
      _ExtentX        =   5106
      _ExtentY        =   556
      _StockProps     =   6
      Caption         =   "原始数据"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   1
   End
End
Attribute VB_Name = "frmLabAnalyseData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum mCol
    标本ID
    标本号
    姓名
    性别
    年龄
    核收时间
    用途
End Enum
Private mlngMachine As Long



Private Sub cmdFind_Click()
    Dim astrItem() As String
    Dim lngLoop As Long
    Dim varItem() As String
    Dim strBegingNO As String, strEndNO As String
    Dim strWhere  As String
    Dim rsTmp As New ADODB.Recordset
    Dim Record As ReportRecord
    Dim strNumber As String
    
    If DateDiff("d", Me.DTPStart, Me.DTPEnd) > 30 Then
        If MsgBox("你所选择的时间段大于30天，可能导致查询数据过多或查询时间过长。" & vbCrLf & "是否继续？", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            Me.DTPStart.SetFocus
            Exit Sub
        End If
    End If
    
    If Me.cbo分析目的.ListCount = 0 Then
        MsgBox "你没有设置分析目的，请到字典管理工具中增加分析目的!", vbInformation, Me.Caption
        Exit Sub
    End If
    
    '==========================================================查找原始数据=======================================================================
    gstrSql = " Select a.ID,a.标本序号, a.姓名, a.年龄, a.性别, a.核收时间, " & vbNewLine & _
                "  Decode(a.仪器id, Null," & vbNewLine & _
                "                 To_Char(Trunc(a.标本序号 / 10000) + 1, '0000') || '-' || To_Char(Mod(a.标本序号, 10000), '0000')," & vbNewLine & _
                "                 a.标本序号) As 标本号显示 " & vbNewLine & _
                " From 检验标本记录 a , 检验分析记录 b " & vbNewLine & _
                " Where a.id = b.标本ID(+) And (仪器id = [1] Or Nvl(仪器id, -1) = [1]) " & vbNewLine & _
                " And 核收时间 between [2] and [3] and b.用途 is null "
    
    
    '标本号
    If Trim(txt标本号) <> "" Then
        txt标本号 = Replace(Replace(txt标本号, "～", "~"), "-", "~")
        varItem = Split(Trim(txt标本号.Text), ",")
        
        For lngLoop = 0 To UBound(varItem)
            astrItem = Split(varItem(lngLoop), "~")
            
            If UBound(astrItem) <= 0 Then
                strBegingNO = TransSampleNO(IIf(Val(Me.txt批号) <> 0, Val(Me.txt批号) & "-" & Val(varItem(lngLoop)), Val(varItem(lngLoop))))
                strEndNO = TransSampleNO(IIf(Val(Me.txt批号) <> 0, Val(Me.txt批号) & "-" & Val(varItem(lngLoop)), Val(varItem(lngLoop))))
            Else
                strBegingNO = TransSampleNO(IIf(Val(Me.txt批号) <> 0, Val(Me.txt批号) & "-" & Val(astrItem(0)), Val(astrItem(0))))
                strEndNO = TransSampleNO(IIf(Val(Me.txt批号) <> 0, Val(Me.txt批号) & "-" & Val(astrItem(1)), Val(astrItem(1))))
            End If
            If lngLoop = 0 Then
                strWhere = strWhere & " and (to_Number(标本序号) between " & Val(strBegingNO) & " and " & Val(strEndNO) & " "
            Else
                strWhere = strWhere & "  or to_Number(标本序号) between " & Val(strBegingNO) & " and " & Val(strEndNO) & " "
            End If
        Next
        If lngLoop >= 0 Then strWhere = strWhere & ")"
    ElseIf Trim(txt批号) <> "" Then
        strWhere = strWhere & " and to_Number(标本序号) between [4] and [5] "
        strBegingNO = TransSampleNO(Val(Me.txt批号) & "-0001")
        strEndNO = TransSampleNO(Val(Me.txt批号) & "-9999")
    End If
    gstrSql = gstrSql & strWhere
    
    Me.rptSource.Records.DeleteAll
    strTmp = Me.cbo分析目的.List(Me.cbo分析目的.ListIndex)
    strTmp = Mid(strTmp, 1, InStr(strTmp, "-") - 1)
    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex)), _
                    CDate(Format(Me.DTPStart, "yyyy-mm-dd 00:00:00")), _
                    CDate(Format(Me.DTPEnd, "yyyy-mm-dd 23:59:59")), Val(strBegingNO), Val(strEndNO))
                        
    Do Until rsTmp.EOF
        
        Set Record = Me.rptSource.Records.Add
            For intLoop = 0 To Me.rptSource.Columns.Count
                Record.AddItem ""
            Next
            Record.Item(mCol.标本ID).Value = Nvl(rsTmp("ID"))
            Record.Item(mCol.标本号).Value = Nvl(rsTmp("标本序号"))
            Record.Item(mCol.标本号).Caption = Nvl(rsTmp("标本号显示"))
            Record.Item(mCol.姓名).Value = Nvl(rsTmp("姓名"))
            Record.Item(mCol.性别).Value = Nvl(rsTmp("性别"))
            Record.Item(mCol.年龄).Value = Nvl(rsTmp("年龄"))
            Record.Item(mCol.核收时间).Value = Nvl(rsTmp("核收时间"))
        rsTmp.MoveNext
    Loop
    Me.rptSource.Populate
    '==========================================================================================================================================
    
    '===========================================================查询分析数据===================================================================
    gstrSql = " Select a.ID,a.标本序号, a.姓名, a.年龄, a.性别, a.核收时间,c.名称, " & vbNewLine & _
                "  Decode(a.仪器id, Null," & vbNewLine & _
                "                 To_Char(Trunc(a.标本序号 / 10000) + 1, '0000') || '-' || To_Char(Mod(a.标本序号, 10000), '0000')," & vbNewLine & _
                "                 a.标本序号) As 标本号显示 " & vbNewLine & _
                " From 检验标本记录 a , 检验分析记录 b, 检验分析用途 c " & vbNewLine & _
                " Where a.id = b.标本ID And b.用途=c.编码 and (仪器id = [1] Or Nvl(仪器id, -1) = [1]) " & vbNewLine & _
                " And 核收时间 between [2] and [3] and b.用途 is not null and c.编码 = [6] "
    
    
    '标本号
    If Trim(txt标本号) <> "" Then
        txt标本号 = Replace(Replace(txt标本号, "～", "~"), "-", "~")
        varItem = Split(Trim(txt标本号.Text), ",")
        
        For lngLoop = 0 To UBound(varItem)
            astrItem = Split(varItem(lngLoop), "~")
            
            If UBound(astrItem) <= 0 Then
                strBegingNO = TransSampleNO(IIf(Val(Me.txt批号) <> 0, Val(Me.txt批号) & "-" & Val(varItem(lngLoop)), Val(varItem(lngLoop))))
                strEndNO = TransSampleNO(IIf(Val(Me.txt批号) <> 0, Val(Me.txt批号) & "-" & Val(varItem(lngLoop)), Val(varItem(lngLoop))))
            Else
                strBegingNO = TransSampleNO(IIf(Val(Me.txt批号) <> 0, Val(Me.txt批号) & "-" & Val(astrItem(0)), Val(astrItem(0))))
                strEndNO = TransSampleNO(IIf(Val(Me.txt批号) <> 0, Val(Me.txt批号) & "-" & Val(astrItem(1)), Val(astrItem(1))))
            End If
            If lngLoop = 0 Then
                strWhere = strWhere & " and (to_Number(标本序号) between " & Val(strBegingNO) & " and " & Val(strEndNO) & " "
            Else
                strWhere = strWhere & "  or to_Number(标本序号) between " & Val(strBegingNO) & " and " & Val(strEndNO) & " "
            End If
        Next
        If lngLoop >= 0 Then strWhere = strWhere & ")"
    ElseIf Trim(txt批号) <> "" Then
        strWhere = strWhere & " and to_Number(标本序号) between [4] and [5] "
        strBegingNO = TransSampleNO(Val(Me.txt批号) & "-0001")
        strEndNO = TransSampleNO(Val(Me.txt批号) & "-9999")
    End If
    gstrSql = gstrSql & strWhere
    Me.rptAnalyse.Records.DeleteAll
    strNumber = Me.cbo分析目的.List(Me.cbo分析目的.ListIndex)
    strNumber = Mid(strNumber, 1, InStr(strNumber, "-") - 1)
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex)), _
                    CDate(Format(Me.DTPStart, "yyyy-mm-dd 00:00:00")), _
                    CDate(Format(Me.DTPEnd, "yyyy-mm-dd 23:59:59")), Val(strBegingNO), Val(strEndNO), strNumber)
                        
    Do Until rsTmp.EOF
        
        Set Record = Me.rptAnalyse.Records.Add
            For intLoop = 0 To Me.rptSource.Columns.Count
                Record.AddItem ""
            Next
            Record.Item(mCol.标本ID).Value = Nvl(rsTmp("ID"))
            Record.Item(mCol.标本号).Value = Nvl(rsTmp("标本序号"))
            Record.Item(mCol.标本号).Caption = Nvl(rsTmp("标本号显示"))
            Record.Item(mCol.姓名).Value = Nvl(rsTmp("姓名"))
            Record.Item(mCol.性别).Value = Nvl(rsTmp("性别"))
            Record.Item(mCol.年龄).Value = Nvl(rsTmp("年龄"))
            Record.Item(mCol.核收时间).Value = Nvl(rsTmp("核收时间"))
            Record.Item(mCol.用途).Value = Nvl(rsTmp("名称"))
        rsTmp.MoveNext
    Loop
    Me.rptAnalyse.Populate
    '==========================================================================================================================================
    
End Sub

Private Sub cmdLeft_Click()
    Call SaveData(1)
End Sub

Private Sub cmdLeftAll_Click()
    Call SaveData(2)
End Sub

Private Sub cmdRight_Click()
    Call SaveData(3)
End Sub

Private Sub cmdRightAll_Click()
    Call SaveData(4)
End Sub

Private Sub Form_Load()
    Dim Column As ReportColumn
    Dim rsTmp As New ADODB.Recordset

    With Me.rptSource.Columns
        rptSource.AllowColumnRemove = False
        rptSource.ShowItemsInGroups = False
        
        With rptSource.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
'        rptListSource.SetImageList ImgList
        Set Column = .Add(mCol.标本ID, "标本ID", 75, True): Column.Visible = False
        Set Column = .Add(mCol.标本号, "标本号", 75, True)
        Set Column = .Add(mCol.姓名, "姓名", 75, True)
        Set Column = .Add(mCol.性别, "性别", 75, True)
        Set Column = .Add(mCol.年龄, "年龄", 75, True)
        Set Column = .Add(mCol.核收时间, "核收时间", 75, True)
    End With
    
    With Me.rptAnalyse.Columns
        rptAnalyse.AllowColumnRemove = False
        rptAnalyse.ShowItemsInGroups = False
        
        With rptAnalyse.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
'        rptListSource.SetImageList ImgList
        Set Column = .Add(mCol.标本ID, "标本ID", 75, True): Column.Visible = False
        Set Column = .Add(mCol.标本号, "标本号", 75, True)
        Set Column = .Add(mCol.姓名, "姓名", 75, True)
        Set Column = .Add(mCol.性别, "性别", 75, True)
        Set Column = .Add(mCol.年龄, "年龄", 75, True)
        Set Column = .Add(mCol.核收时间, "核收时间", 75, True)
        Set Column = .Add(mCol.用途, "用途", 100, True)
    End With
    
    Me.DTPStart = Now
    Me.DTPEnd = Now
    
    gstrSql = "select Id,编码,名称 from 检验仪器 "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With Me.cbo仪器
        .Clear
        .AddItem "[手工]"
        .ItemData(.NewIndex) = -1
        Do While Not rsTmp.EOF
            .AddItem Nvl(rsTmp("编码")) & "-" & Nvl(rsTmp("名称"))
            .ItemData(.NewIndex) = Nvl(rsTmp("ID"))
            If mlngMachine = Nvl(rsTmp("ID")) Then
                .ListIndex = .NewIndex
            End If
            rsTmp.MoveNext
        Loop
        If .ListCount > 0 And .ListIndex = -1 Then
            .ListIndex = 0
        End If
    End With
    
    gstrSql = "select 编码,名称 from 检验分析用途 "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With Me.cbo分析目的
        .Clear
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("编码")) & "-" & Nvl(rsTmp("名称"))
            .ItemData(.NewIndex) = Nvl(rsTmp("编码"))
            rsTmp.MoveNext
        Loop
        If .ListCount > 0 And .ListIndex = -1 Then
            .ListIndex = 0
        End If
    End With
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    With Me.fraFilter
        .Left = 0
        .Width = Me.ScaleWidth
    End With
    
    With Me.ShortCapSource
        .Left = 0
        .Width = (Me.ScaleWidth / 2) - Me.CmdLeft.Width - (50 * 2)
    End With
    
    With Me.rptSource
        .Left = 0
        .Width = Me.ShortCapSource.Width
        .Height = Me.ScaleHeight - .Top
    End With
    
    With Me.ShortCaptAnalyse
        .Left = Me.ShortCapSource.Left + Me.ShortCapSource.Width + Me.CmdLeft.Width + 100
        .Width = Me.ScaleWidth - .Left - 50
    End With
    
    With Me.rptAnalyse
        .Left = Me.ShortCaptAnalyse.Left
        .Width = Me.ShortCaptAnalyse.Width
        .Height = Me.ScaleHeight - .Top
    End With
    
    With Me.CmdLeft
        .Left = Me.rptSource.Left + Me.rptSource.Width + 50
        .Top = (Me.rptSource.Height / 4 / 2 * 1) + (.Height / 2) + Me.rptSource.Top
    End With
    
    With Me.cmdLeftAll
        .Left = Me.CmdLeft.Left
        .Top = (Me.rptSource.Height / 4 / 2 * 2) + (.Height / 2) + Me.rptSource.Top
    End With
    
    With Me.cmdRight
        .Left = Me.CmdLeft.Left
        .Top = (Me.rptSource.Height / 4 / 2 * 3) + (.Height / 2) + Me.rptSource.Top
    End With
    
    With Me.cmdRightAll
        .Left = Me.CmdLeft.Left
        .Top = (Me.rptSource.Height / 4 / 2 * 4) + (.Height / 2) + Me.rptSource.Top
    End With
End Sub
Public Sub ShowMe(Objfrm As Object, lngMachine As Long)
    mlngMachine = lngMachine
    Me.Show vbModal, Objfrm
End Sub
Public Sub SaveData(EditMode As Integer)
    '功能               写入分析或删除分析
    '参数               EditMode
    '                   1=删除当前一条分析数据
    '                   2=删除当前所有分析数据
    '                   3=插入当前一条分析数据
    '                   4=删除当前所有分析数据
    Dim lngLoop As Long
    Dim strAnalyse As String
    Dim intColCount As Integer
    Dim Record As ReportRecord
    
    Select Case EditMode
        Case 1, 2                                                   '删除分析数据记录
            '没有原始数据或没有选分析目的时退出
            If Me.rptAnalyse.Records.Count = 0 Then
                MsgBox "没有数据可以选择，请重新选择条件进行查询!", vbInformation, Me.Caption
                Exit Sub
            End If
            strAnalyse = Me.cbo分析目的.List(Me.cbo分析目的.ListIndex)
            If EditMode = 1 Then
                If Me.rptAnalyse.FocusedRow Is Nothing Then
                    Exit Sub
                End If
                gstrSql = "Zl_检验分析记录_Edit(2," & Me.rptAnalyse.FocusedRow.Record(mCol.标本ID).Value & ",'" & _
                            Mid(strAnalyse, 1, InStr(strAnalyse, "-") - 1) & "')"
                zlDatabase.ExecuteProcedure gstrSql, Me.Caption
                Set Record = Me.rptSource.Records.Add
                For intColCount = 0 To Me.rptSource.Columns.Count - 1
                    Record.AddItem ""
                Next
                Record.Item(mCol.标本ID).Value = Me.rptAnalyse.FocusedRow.Record(mCol.标本ID).Value
                Record.Item(mCol.标本号).Value = Me.rptAnalyse.FocusedRow.Record(mCol.标本号).Value
                Record.Item(mCol.标本号).Caption = Me.rptAnalyse.FocusedRow.Record(mCol.标本号).Caption
                Record.Item(mCol.姓名).Value = Me.rptAnalyse.FocusedRow.Record(mCol.姓名).Value
                Record.Item(mCol.性别).Value = Me.rptAnalyse.FocusedRow.Record(mCol.性别).Value
                Record.Item(mCol.年龄).Value = Me.rptAnalyse.FocusedRow.Record(mCol.年龄).Value
                Record.Item(mCol.核收时间).Value = Me.rptAnalyse.FocusedRow.Record(mCol.核收时间).Value
                
                Me.rptAnalyse.Records.RemoveAt (Me.rptAnalyse.FocusedRow.Index)
            Else
                '删除所有
                If MsgBox("是否确定要删除当前条件下的所有分析数据?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                    Exit Sub
                End If
                For lngLoop = 0 To Me.rptAnalyse.Records.Count - 1
                    gstrSql = "Zl_检验分析记录_Edit(2," & Me.rptAnalyse.Records(lngLoop).Item(mCol.标本ID).Value & ",'" & _
                            Mid(strAnalyse, 1, InStr(strAnalyse, "-") - 1) & "')"
                    zlDatabase.ExecuteProcedure gstrSql, Me.Caption
                    
                    Set Record = Me.rptSource.Records.Add
                    For intColCount = 0 To Me.rptSource.Columns.Count - 1
                        Record.AddItem ""
                    Next
                    Record.Item(mCol.标本ID).Value = Me.rptAnalyse.Records(lngLoop).Item(mCol.标本ID).Value
                    Record.Item(mCol.标本号).Value = Me.rptAnalyse.Records(lngLoop).Item(mCol.标本号).Value
                    Record.Item(mCol.标本号).Caption = Me.rptAnalyse.Records(lngLoop).Item(mCol.标本号).Caption
                    Record.Item(mCol.姓名).Value = Me.rptAnalyse.Records(lngLoop).Item(mCol.姓名).Value
                    Record.Item(mCol.性别).Value = Me.rptAnalyse.Records(lngLoop).Item(mCol.性别).Value
                    Record.Item(mCol.年龄).Value = Me.rptAnalyse.Records(lngLoop).Item(mCol.年龄).Value
                    Record.Item(mCol.核收时间).Value = Me.rptAnalyse.Records(lngLoop).Item(mCol.核收时间).Value
                Next
                Me.rptAnalyse.Records.DeleteAll
            End If
            Me.rptAnalyse.Populate
            Me.rptSource.Populate
        Case 3, 4                                                   '插入分析数据
            '没有原始数据或没有选分析目的时退出
            If Me.rptSource.Records.Count = 0 Then
                MsgBox "没有数据可以选择，请重新选择条件进行查询!", vbInformation, Me.Caption
                Exit Sub
            End If
            If Me.cbo分析目的.ListCount = 0 Then
                MsgBox "请选择一个分析目的!", vbInformation, Me.Caption
                Me.cbo分析目的.SetFocus
                Exit Sub
            End If
            strAnalyse = Me.cbo分析目的.List(Me.cbo分析目的.ListIndex)
            If EditMode = 3 Then
                '单个写入
                If Me.rptSource.FocusedRow Is Nothing Then
                    Exit Sub
                End If
                gstrSql = "Zl_检验分析记录_Edit(1," & Me.rptSource.FocusedRow.Record(mCol.标本ID).Value & ",'" & _
                            Mid(strAnalyse, 1, InStr(strAnalyse, "-") - 1) & "')"
                zlDatabase.ExecuteProcedure gstrSql, Me.Caption
                
                Set Record = Me.rptAnalyse.Records.Add
                For intColCount = 0 To Me.rptAnalyse.Columns.Count - 1
                    Record.AddItem ""
                Next
                Record.Item(mCol.标本ID).Value = Me.rptSource.FocusedRow.Record(mCol.标本ID).Value
                Record.Item(mCol.标本号).Value = Me.rptSource.FocusedRow.Record(mCol.标本号).Value
                Record.Item(mCol.标本号).Caption = Me.rptSource.FocusedRow.Record(mCol.标本号).Caption
                Record.Item(mCol.姓名).Value = Me.rptSource.FocusedRow.Record(mCol.姓名).Value
                Record.Item(mCol.性别).Value = Me.rptSource.FocusedRow.Record(mCol.性别).Value
                Record.Item(mCol.年龄).Value = Me.rptSource.FocusedRow.Record(mCol.年龄).Value
                Record.Item(mCol.核收时间).Value = Me.rptSource.FocusedRow.Record(mCol.核收时间).Value
                Record.Item(mCol.用途).Value = strAnalyse
                Me.rptSource.Records.RemoveAt (Me.rptSource.FocusedRow.Index)
                
            Else
                '写入所有
                For lngLoop = 0 To Me.rptSource.Records.Count - 1
                    gstrSql = "Zl_检验分析记录_Edit(1," & Me.rptSource.Records(lngLoop).Item(mCol.标本ID).Value & ",'" & _
                            Mid(strAnalyse, 1, InStr(strAnalyse, "-") - 1) & "')"
                    zlDatabase.ExecuteProcedure gstrSql, Me.Caption
                    Set Record = Me.rptAnalyse.Records.Add
                    For intColCount = 0 To Me.rptAnalyse.Columns.Count - 1
                        Record.AddItem ""
                    Next
                    Record.Item(mCol.标本ID).Value = Me.rptSource.Records(lngLoop).Item(mCol.标本ID).Value
                    Record.Item(mCol.标本号).Value = Me.rptSource.Records(lngLoop).Item(mCol.标本号).Value
                    Record.Item(mCol.标本号).Caption = Me.rptSource.Records(lngLoop).Item(mCol.标本号).Caption
                    Record.Item(mCol.姓名).Value = Me.rptSource.Records(lngLoop).Item(mCol.姓名).Value
                    Record.Item(mCol.性别).Value = Me.rptSource.Records(lngLoop).Item(mCol.性别).Value
                    Record.Item(mCol.年龄).Value = Me.rptSource.Records(lngLoop).Item(mCol.年龄).Value
                    Record.Item(mCol.核收时间).Value = Me.rptSource.Records(lngLoop).Item(mCol.核收时间).Value
                    Record.Item(mCol.用途).Value = strAnalyse
                Next
                Me.rptSource.Records.DeleteAll
            End If
            Me.rptSource.Populate
            Me.rptAnalyse.Populate
    End Select
End Sub

Private Sub rptAnalyse_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call SaveData(1)
End Sub

Private Sub rptSource_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call SaveData(3)
End Sub
