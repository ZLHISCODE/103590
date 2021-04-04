VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPageMedRecNOSel 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5595
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvwInStationNo 
      Height          =   1995
      Left            =   300
      TabIndex        =   0
      Top             =   450
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   3519
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "住院号"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "姓名"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "入院日期"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "出院日期"
         Object.Width           =   2646
      EndProperty
   End
End
Attribute VB_Name = "frmPageMedRecNOSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mstrID As String                 '住院号_主页id
Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        mstrID = 0
        Unload Me
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mstrID = Me.lvwInStationNo.SelectedItem.Key
        Unload Me
    End If
End Sub

Private Sub Form_Resize()
    Me.lvwInStationNo.Top = 0
    Me.lvwInStationNo.Left = 0
    Me.lvwInStationNo.Width = Me.ScaleWidth
    Me.lvwInStationNo.Height = Me.ScaleHeight
End Sub
Public Function ShowMe(objfrm As Object, rsPatient As ADODB.Recordset, strIf As String) As String
    '---------------------------------------------------------------------------------------------'
    '功能        提供给上级窗体调用
    '参数        objfrm上级窗体对象
    '            rsPatient数据集
    '            Visible是否显示或关闭窗体
    '---------------------------------------------------------------------------------------------'
    On Error GoTo errH
    
    If Me.Visible = True Then
        Unload Me
        ShowMe = "_"
        Exit Function
    End If
    
    Me.lvwInStationNo.ListItems.Clear
    If rsPatient.State = 1 Then
        rsPatient.Filter = strIf
        rsPatient.Sort = "住院次数 Desc"
        If rsPatient.EOF = False Then
            rsPatient.MoveFirst
        End If
    ElseIf rsPatient.State = adStateClosed Then
        Unload Me
        ShowMe = "_"
        Exit Function
    End If
    
    Do Until rsPatient.EOF
        If Check是否存在病案(rsPatient("住院号"), rsPatient("住院次数")) Then
            Set objList = Me.lvwInStationNo.ListItems.Add(, rsPatient("住院号") & "_" & rsPatient("住院次数"), rsPatient("住院号"))
            objList.SubItems(1) = rsPatient("姓名")
            objList.SubItems(2) = Format(rsPatient("入院日期"), "yyyy-mm-dd")
            objList.SubItems(3) = Format(rsPatient("出院日期"), "yyyy-mm-dd")
        End If
        rsPatient.MoveNext
    Loop
    
    If Me.lvwInStationNo.ListItems.Count > 0 Then
        Me.Show vbModal, objfrm
        ShowMe = mstrID
    Else
        MsgBox "没有找到病人信息!", vbInformation, gstrSysName
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub lvwInStationNo_Click()
    mstrID = Me.lvwInStationNo.SelectedItem.Key
End Sub

Private Sub lvwInStationNo_DblClick()
    mstrID = Me.lvwInStationNo.SelectedItem.Key
    Unload Me
End Sub
Private Function Check是否存在病案(str住院号 As String, lng主页ID As Integer) As Boolean
    '----------------------------------------------------------------------------------------------------------
    '功能:检查是否已建立病案
    '参数:  str住院号-住院号
    '       lng主页ID-主页ID
    '返回:True不存在 False存在
    '----------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    gstrSQL = "" & _
        "   Select 1 from 病人信息 a , 病案主页 b " & _
        "   Where a.病人ID = b.病人ID and b.编目日期 is not null and b.出院日期 is not null  and " & _
        "         a.住院号 = [1] and b.主页ID = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str住院号, lng主页ID)
    If rsTmp.EOF Then
        Check是否存在病案 = True
    End If
    rsTmp.Close
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

