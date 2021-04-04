VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmLabMainPrintFormat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "打印单据格式设置"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4845
   Icon            =   "frmLabMainPrintFormat.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin XtremeReportControl.ReportControl rptAlist 
      Height          =   2535
      Left            =   30
      TabIndex        =   1
      Top             =   300
      Width           =   4785
      _Version        =   589884
      _ExtentX        =   8440
      _ExtentY        =   4471
      _StockProps     =   0
      BorderStyle     =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "请选择一个单据格式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1890
   End
End
Attribute VB_Name = "frmLabMainPrintFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mstr医嘱ID As String
Dim mstrPrintCode As String
Private Enum mCol
    报表编号
    名称
End Enum
Public Sub ShowMe(Objfrm As Object, str医嘱 As String, strPrintCode As String)
    Dim rsTmp As New ADODB.Recordset
    Dim Record As ReportRecord                      '列表数据集
    Dim Column As ReportColumn
    Dim intLoop As Integer
    
    rptAlist.AllowColumnRemove = False
    rptAlist.ShowItemsInGroups = False
    With rptAlist.PaintManager
        .ColumnStyle = xtpColumnShaded
        .GridLineColor = RGB(225, 225, 225)
        .NoGroupByText = "拖动列标题到这里,按该列分组..."
        .NoItemsText = "没有可显示的项目..."
        .VerticalGridStyle = xtpGridSolid
        .HideSelection = True
    End With
    With Me.rptAlist.Columns
        Set Column = .Add(mCol.报表编号, "报表编号", 120, True)
        Set Column = .Add(mCol.名称, "名称", 120, True)
    End With

    gstrSql = "Select /*+ rule */" & vbNewLine & _
        " Distinct 'ZLCISBILL' || Trim(To_Char(C.编号, '00000')) || '-2' As 报表编号,  A.记录性质,c.名称" & vbNewLine & _
        " From 病人医嘱发送 A, 病历文件列表 C, 病人医嘱记录 D, 病历单据应用 E" & vbNewLine & _
        " Where E.病历文件id = C.ID And D.诊疗项目id = E.诊疗项目id And A.医嘱id = D.ID And " & vbNewLine & _
        " E.应用场合 = Decode(D.病人来源, 2, 2, 4, 4, 1) And" & vbNewLine & _
        " D.相关id In (Select * From Table(Cast(f_Num2list([1]) As ZLTOOLS.t_Numlist)))"

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, str医嘱)
    If rsTmp.RecordCount <= 1 Then
        If rsTmp.EOF = False Then
            strPrintCode = Nvl(rsTmp("报表编号"))
        End If
        Unload Me
        Exit Sub
    End If
    Do While Not rsTmp.EOF
        Set Record = Me.rptAlist.Records.Add
        For intLoop = 0 To Me.rptAlist.Columns.Count - 1
            Record.AddItem ""
        Next
        Record(mCol.报表编号).Value = Nvl(rsTmp("报表编号"))
        Record(mCol.名称).Value = Nvl(rsTmp("名称"))
        rsTmp.MoveNext
    Loop
    Me.rptAlist.Populate
    mstr医嘱ID = str医嘱
    Me.Show vbModal, Objfrm
    strPrintCode = mstrPrintCode
End Sub

Private Sub rptAlist_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    mstrPrintCode = Item.Record(mCol.报表编号).Value
    Unload Me
End Sub
