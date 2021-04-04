VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEdit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "交接班记录-编辑"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7380
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7380
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraEdit 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.Frame fraSplit1 
         Height          =   30
         Left            =   0
         TabIndex        =   24
         Top             =   4080
         Width           =   6855
      End
      Begin VB.CommandButton cmdIn 
         Caption         =   "…"
         Height          =   290
         Left            =   3225
         TabIndex        =   21
         Top             =   2300
         Width           =   255
      End
      Begin VB.ComboBox cboType 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1082
         Width           =   1935
      End
      Begin VB.ComboBox cboHoldType 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2736
         Width           =   1935
      End
      Begin VB.ComboBox cboDept 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtPer 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         TabIndex        =   2
         Top             =   661
         Width           =   1935
      End
      Begin VB.TextBox txtHold 
         Height          =   300
         Left            =   1560
         TabIndex        =   4
         Top             =   2315
         Width           =   1935
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   4440
         TabIndex        =   6
         Top             =   4320
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   5760
         TabIndex        =   7
         Top             =   4320
         Width           =   1100
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   285
         Left            =   1560
         TabIndex        =   13
         Top             =   1500
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   243662851
         CurrentDate     =   401769
         MaxDate         =   402133
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   285
         Left            =   1560
         TabIndex        =   14
         Top             =   1905
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   243662851
         CurrentDate     =   401769
         MaxDate         =   402133
      End
      Begin MSComCtl2.DTPicker dtpHoldBegin 
         Height          =   285
         Left            =   1560
         TabIndex        =   17
         Top             =   3150
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   243662851
         CurrentDate     =   401769
         MaxDate         =   402133
      End
      Begin MSComCtl2.DTPicker dtpHoldEnd 
         Height          =   285
         Left            =   1560
         TabIndex        =   18
         Top             =   3570
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   243662851
         CurrentDate     =   401769
         MaxDate         =   402133
      End
      Begin MSComctlLib.TreeView tvwDoc 
         Height          =   1575
         Left            =   1080
         TabIndex        =   22
         Top             =   4200
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   2778
         _Version        =   393217
         Indentation     =   353
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imgList"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfPaiType 
         Height          =   3375
         Left            =   3840
         TabIndex        =   25
         Top             =   480
         Width           =   3075
         _cx             =   5433
         _cy             =   5953
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
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
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   11
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEdit.frx":6852
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
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "自动导入以下病人"
         Height          =   180
         Left            =   3840
         TabIndex        =   23
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label lblInEnd 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "接班结束时间"
         Height          =   180
         Left            =   360
         TabIndex        =   20
         Top             =   3600
         Width           =   1080
      End
      Begin VB.Label lblInBegin 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "接班开始时间"
         Height          =   180
         Left            =   360
         TabIndex        =   19
         Top             =   3180
         Width           =   1080
      End
      Begin VB.Label lblEnd 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "交班结束时间"
         Height          =   180
         Left            =   360
         TabIndex        =   16
         Top             =   1950
         Width           =   1080
      End
      Begin VB.Label lblBegin 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "交班开始时间"
         Height          =   180
         Left            =   360
         TabIndex        =   15
         Top             =   1545
         Width           =   1080
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "科室"
         Height          =   180
         Index           =   0
         Left            =   1080
         TabIndex        =   12
         Top             =   315
         Width           =   360
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "接班班次"
         Height          =   180
         Index           =   3
         Left            =   720
         TabIndex        =   11
         Top             =   2775
         Width           =   720
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "接班医生"
         Height          =   180
         Index           =   4
         Left            =   720
         TabIndex        =   10
         Top             =   2370
         Width           =   720
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "交班班次"
         Height          =   180
         Index           =   6
         Left            =   720
         TabIndex        =   9
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "交班医生"
         Height          =   180
         Index           =   7
         Left            =   720
         TabIndex        =   8
         Top             =   720
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   8760
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":69C1
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":6F5B
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":74F5
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":DD57
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":145B9
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":1AE1B
            Key             =   "add"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":2167D
            Key             =   "Person"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":2208F
            Key             =   "Dept"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngId As Long
Private mrsTime As ADODB.Recordset
Private mstrDeptId As String
Private mblnOk As Boolean
Private mrsDoc As ADODB.Recordset
Private mrsPati As ADODB.Recordset
Private mbytType As Byte '0-新增；1-修改
Private mstrDept As String '科室传入
Private mstrOutPer As String '交班人姓名
Private mstrOutTime As String '交班时间范围
Private mstrInPer As String '接班人姓名
Private mstrInTime As String '接班时间范围

Public Function ShowMe(ByVal bytType As Byte, Optional ByVal lngId As Long, Optional strDept As String, Optional strOutPer As String, Optional strOutTime As String, _
                Optional strInPer As String, Optional strInTime As String) As Boolean
'bytType:0-新增交接班记录；1-修改交接班记录
'strDept格式-主界面选择科室索引|下拉框科室1|下拉框科室2...
    mstrDeptId = ""
    mlngId = lngId
    mbytType = bytType
    mstrDept = strDept
    mstrOutPer = strOutPer
    mstrOutTime = strOutTime
    mstrInPer = strInPer
    mstrInTime = strInTime
    
    Me.Show 1
    ShowMe = mblnOk
End Function

Private Sub SetBasic()
    Dim strTemp As String
    Dim varTemp As Variant, varData As Variant
    Dim i As Long, lngTemp As Long
    Dim rsTemp As ADODB.Recordset
    
    Select Case mbytType
        Case 0
            Me.Caption = "交接班记录-新增"
            Me.Width = fraEdit.Width + 100
            Me.Height = fraEdit.Height + 350
            fraEdit.Visible = True
            fraEdit.Move 0, 0
            Set rsTemp = GetPatientType
            With vsfPaiType
                .Rows = 1
                .Rows = rsTemp.RecordCount + 1
                Do While Not rsTemp.EOF
                    .TextMatrix(rsTemp.AbsolutePosition, .ColIndex("简称")) = rsTemp!简称
                    .TextMatrix(rsTemp.AbsolutePosition, .ColIndex("病人类型")) = rsTemp!名称
                    .TextMatrix(rsTemp.AbsolutePosition, .ColIndex("提取SQL")) = rsTemp!提取SQL & ""
                    rsTemp.MoveNext
                Loop
                For i = 0 To .Rows - 1
                    .Cell(flexcpChecked, i, 0) = flexChecked
                Next
                If .Rows > 11 Then
                    .ColWidth(.ColIndex("病人类型")) = 1800
                Else
                    .ColWidth(.ColIndex("病人类型")) = 2055
                End If
            End With
            Call vsfPaiType_AfterRowColChange(1, 1, 0, 1)
        Case 1
            Me.Caption = "交接班记录-修改"
            cboDept.Enabled = False
            lblType.Visible = False
            vsfPaiType.Visible = False
            cboType.Enabled = False
            Me.Width = fraEdit.Width - vsfPaiType.Width
            Me.Height = fraEdit.Height + 350
            cmdCancel.Left = dtpHoldEnd.Left + dtpHoldEnd.Width - cmdCancel.Width
            cmdOK.Left = cmdCancel.Left - 200 - cmdOK.Width
            fraEdit.Visible = True
            fraEdit.Move 0, 0
    End Select
End Sub

Private Sub cboDept_Click()
    Dim strDeptID As Long
    Dim rsTemp As ADODB.Recordset
    
    strDeptID = cboDept.ItemData(cboDept.ListIndex)
    Set rsTemp = GetShiftType(2, strDeptID)
    Set mrsTime = GetShiftType(1, strDeptID)
    cboType.Clear
    cboHoldType.Clear
    Do While Not rsTemp.EOF
        cboType.AddItem rsTemp!班次名称
        cboHoldType.AddItem rsTemp!班次名称
        rsTemp.MoveNext
    Loop
End Sub

Private Sub cboHoldType_Change()
    If cboHoldType.Text = "" Then
        dtpHoldBegin.Value = "3000/1/1"
        dtpHoldEnd.Value = "3000/1/1"
    End If
End Sub

Private Sub cboHoldType_Click()
    Dim objDate As Date
    
    objDate = Format(IIf(cboType.Text = "", zlDatabase.Currentdate, dtpEnd.Value), "yyyy-mm-dd")
    mrsTime.Filter = "班次名称='" & cboHoldType.Text & "'"
    If mrsTime.RecordCount = 1 Then
        dtpHoldBegin.Value = objDate & " " & mrsTime!开始时间
        dtpHoldEnd.Value = IIf(mrsTime!开始时间 >= mrsTime!结束时间, objDate + 1, objDate) & " " & mrsTime!结束时间
    End If
End Sub

Private Sub cboType_Change()
    If cboType = "" Then
        dtpBegin.Enabled = False
        dtpBegin.Value = "3000/1/1"
        dtpEnd.Value = "3000/1/1"
    Else
        dtpBegin.Enabled = True
    End If
End Sub

Private Sub cboType_Click()
'交班班次选择后，自动显示交班开始时间和交班结束时间，交班开始时间可调整
    Dim objDate As Date
    
    objDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    mrsTime.Filter = "班次名称='" & cboType.Text & "'"
    If mrsTime.RecordCount = 1 Then
        If mrsTime!开始时间 >= mrsTime!结束时间 Then
            dtpBegin.Value = objDate - 1 & " " & mrsTime!开始时间
            dtpEnd.Value = objDate & " " & mrsTime!结束时间
        Else
            dtpBegin.Value = objDate & " " & mrsTime!开始时间
            dtpEnd.Value = objDate & " " & mrsTime!结束时间
        End If
        If mrsTime!开始时间 = mrsTime!结束时间 Then
            cboHoldType.Text = cboType.Text
            Call cboHoldType_Click
        Else
            mrsTime.Filter = "开始时间='" & mrsTime!结束时间 & "'"
            If mrsTime.RecordCount > 0 Then
                cboHoldType.Text = mrsTime!班次名称
                Call cboHoldType_Click
            End If
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdIn_Click()

    mrsDoc.Filter = ""
    If mrsDoc.RecordCount = 0 Then Exit Sub
    tvwDoc.Visible = True
    tvwDoc.SetFocus
End Sub

Private Sub ShowDoc()
'加载医生信息的数据
    Dim strDept As String
    Dim objNode As Object
    
    On Error GoTo errH
    If mbytType = 0 Then
        gstrSQL = "Select b.部门id, c.名称, a.Id,a.编号, a.姓名,a.简码 From 人员表 a, 部门人员 b, 部门表 c" & vbNewLine & _
            "Where a.Id = b.人员id And b.缺省 = 1 And b.部门id In(Select * From Table(f_str2list([1]))) And b.部门id = c.Id" & vbNewLine & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) Order By 部门id, Id"
        Set mrsDoc = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrDeptId)
    Else
        gstrSQL = "Select b.部门id, c.名称, a.Id, a.编号, a.姓名,a.简码 " & vbNewLine & _
            "From 人员表 a, 部门人员 b, 部门表 c" & vbNewLine & _
            "Where a.Id = b.人员id And b.缺省 = 1 And" & vbNewLine & _
            "      b.部门id In (Select 部门id From 临床部门 Where 工作性质 In (Select 工作性质 From 临床部门 Where 部门id =[1])) And b.部门id = c.Id" & vbNewLine & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) Order By 部门id, Id"
        Set mrsDoc = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrDeptId))
    End If
    If mrsDoc.RecordCount = 0 Then Exit Sub
    
    With tvwDoc
        .Nodes.Clear
        Do While Not mrsDoc.EOF
            If strDept <> mrsDoc!名称 Then
                '部门和人员的id可能重复，故关键字是id和名称一起
                Set objNode = .Nodes.Add(, , "K" & mrsDoc!部门id & mrsDoc!名称, mrsDoc!名称, "Dept")
                Set objNode = .Nodes.Add("K" & mrsDoc!部门id & mrsDoc!名称, tvwChild, "K" & mrsDoc!id, mrsDoc!姓名, "Person")
                strDept = mrsDoc!名称
            Else
                Set objNode = .Nodes.Add("K" & mrsDoc!部门id & mrsDoc!名称, tvwChild, "K" & mrsDoc!id, mrsDoc!姓名, "Person")
            End If
            mrsDoc.MoveNext
        Loop
    End With
    tvwDoc.Left = txtHold.Left
    tvwDoc.Top = txtHold.Top + txtHold.Height
    tvwDoc.ZOrder 0
    Exit Sub
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdOK_Click()
    Dim arrTemp As Variant, arrSQL As Variant
    Dim i As Long, lngId As Long
    Dim blnBegin As Boolean
        
    If CheckRecordData = False Then Exit Sub
    gstrSQL = ""
    arrTemp = Array()
    arrSQL = Array()
    If mbytType = 0 Then
        '新增，可能一个医生值班多个科室
        If cboDept.Text = "所有科室" Then
            For i = 1 To cboDept.ListCount - 1
                ReDim Preserve arrTemp(UBound(arrTemp) + 1)
                arrTemp(UBound(arrTemp)) = cboDept.ItemData(i) & ",'" & txtPer.Text & "','" & cboType.Text & "'," & _
                zlStr.To_Date(dtpBegin.Value) & "," & zlStr.To_Date(dtpEnd.Value) & ",'" & _
                txtHold.Text & "','" & cboHoldType.Text & "'," & _
                zlStr.To_Date(dtpHoldBegin.Value) & "," & zlStr.To_Date(dtpHoldEnd.Value)
            Next
        Else
            ReDim Preserve arrTemp(UBound(arrTemp) + 1)
            arrTemp(UBound(arrTemp)) = cboDept.ItemData(cboDept.ListIndex) & ",'" & txtPer.Text & "','" & cboType.Text & "'," & _
            zlStr.To_Date(dtpBegin.Value) & "," & zlStr.To_Date(dtpEnd.Value) & ",'" & _
            txtHold.Text & "','" & cboHoldType.Text & "'," & _
            zlStr.To_Date(dtpHoldBegin.Value) & "," & zlStr.To_Date(dtpHoldEnd.Value)
        End If
        Set mrsPati = GetTimeRangePati(dtpBegin.Value, dtpEnd.Value, mstrDeptId)
        For i = LBound(arrTemp) To UBound(arrTemp)
            lngId = GetNextId("医生交接班记录", "记录ID")
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_医生交接班记录_Edit(0," & lngId & "," & arrTemp(i) & ",'" & grsUserInfo!姓名 & "')"
            Call SavePatiData(arrSQL, lngId, Mid(arrTemp(i), 1, InStr(arrTemp(i), ",") - 1))
        Next
    Else '修改
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_医生交接班记录_Edit(1," & mlngId & "," & Val(mstrDeptId) & ",'" & txtPer.Text & "','" & cboType.Text & "'," & _
            zlStr.To_Date(dtpBegin.Value) & "," & zlStr.To_Date(dtpEnd.Value) & ",'" & _
            txtHold.Text & "','" & cboHoldType.Text & "'," & _
            zlStr.To_Date(dtpHoldBegin.Value) & "," & zlStr.To_Date(dtpHoldEnd.Value) & ",'" & grsUserInfo!姓名 & "')"
    End If
    On Error GoTo ErrHand
    gcnOracle.BeginTrans
    blnBegin = True
    For i = LBound(arrSQL) To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans
    mblnOk = True
    Unload Me
    Exit Sub
ErrHand:
    If blnBegin Then gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub SavePatiData(arrSQL As Variant, ByVal lngRecordId As Long, ByVal lngDeptId As Long)
'根据勾选保存交接班内容数据以及汇总数据
    Dim rsTemp As ADODB.Recordset
    Dim strType As String, strTypes As String, strPsiId As String, str诊断 As String, str主诉 As String
    Dim i As Long, lng新入 As Long, lng出院 As Long, lng总人数 As Long
    
    '新增记录时自动加入汇总表数据
    Set rsTemp = GetPatiType
    Do While Not rsTemp.EOF
        mrsPati.Filter = "出院科室id=" & lngDeptId & " And 类型='" & rsTemp!简称 & "'"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_医生交接班汇总_Insert(" & lngRecordId & "," & rsTemp!顺序 & ",'" & rsTemp!简称 & "'," & mrsPati.RecordCount & ")"
        rsTemp.MoveNext
    Loop
    
    '住院总人数
    gstrSQL = "Select Count(*) 人数 From 在院病人 Where 科室id =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngDeptId)
    lng总人数 = rsTemp!人数
    If DateDiff("s", dtpEnd.Value, zlDatabase.Currentdate) <= 0 Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_医生交接班汇总_Insert(" & lngRecordId & ",99,'住院总'," & rsTemp!人数 & ")"
    Else
        gstrSQL = "Select count(*) 人数 " & vbNewLine & _
            "From 病案主页 a" & vbNewLine & _
            "Where a.入院日期 > " & zlStr.To_Date(dtpEnd.Value) & " And" & vbNewLine & _
            "      a.入院日期 <=sysdate And a.出院日期 Is Null and a.出院科室id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngDeptId)
        lng新入 = rsTemp!人数
        gstrSQL = "Select count(*) 人数" & vbNewLine & _
            "From 病案主页 a" & vbNewLine & _
            "Where a.出院日期 > " & zlStr.To_Date(dtpEnd.Value) & " And" & vbNewLine & _
            "      a.出院日期 <=sysdate and a.出院科室id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngDeptId)
        lng出院 = rsTemp!人数
        lng总人数 = lng总人数 - lng新入 + lng出院
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_医生交接班汇总_Insert(" & lngRecordId & ",99,'住院总人数'," & IIf(lng总人数 > 0, lng总人数, 0) & ")"
    End If
    
    mrsPati.Filter = ""
    If mrsPati.RecordCount > 0 Then
        Set rsTemp = zlDatabase.CopyNewRec(mrsPati)
        With vsfPaiType
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, .ColIndex("选择")) = flexChecked Then
                    mrsPati.Filter = "出院科室id=" & lngDeptId & " And 类型='" & .TextMatrix(i, .ColIndex("简称")) & "'"
                    Do While Not mrsPati.EOF
                        '一个病人如果属于多种类型，则病人类型中用字符串拼起来(简称)
                        '一个病人只能加在一种类型中，按照表格顺序排列的
                        strType = mrsPati!类型
                        rsTemp.Filter = "出院科室id=" & lngDeptId & " And 病人id=" & mrsPati!病人ID & " And 类型<>" & "'" & .TextMatrix(i, .ColIndex("简称")) & "'"
                        Do While Not rsTemp.EOF
                            strType = strType & "," & rsTemp!类型
                            rsTemp.MoveNext
                        Loop
                        If InStr(strPsiId & ",", "," & mrsPati!病人ID & ",") = 0 Then
                            strPsiId = strPsiId & "," & mrsPati!病人ID
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "Zl_医生交接班内容_Edit(0,0," & lngRecordId & ",0,'" & strType & "'," & mrsPati!病人ID & "," & NVL(mrsPati!主页ID, 0) & ",'" & _
                                mrsPati!姓名 & "','" & mrsPati!性别 & "','" & mrsPati!年龄 & "','" & mrsPati!床号 & "" & "'," & NVL(mrsPati!标识号, "Null") & _
                                "," & zlStr.To_Date(mrsPati!入院时间) & ",'" & _
                                mrsPati!入院方式 & "')"
                        End If
                        mrsPati.MoveNext
                    Loop
                End If
            Next
        End With
    End If
End Sub

Private Function CheckRecordData() As Boolean
'交接班记录数据的检查
    
    If cboType.Text = "" Then MsgBox "交班班次不能为空，请选择！", vbExclamation, Me.Caption: Call zlControl.ControlSetFocus(txtPer): Exit Function
    If txtHold.Text = "" Then MsgBox "接班医生不能为空，请填写！", vbExclamation, Me.Caption: Call zlControl.ControlSetFocus(txtHold): Exit Function
    If cboHoldType.Text = "" Then MsgBox "接班班次不能为空，请选择！", vbExclamation, Me.Caption: Call zlControl.ControlSetFocus(cboHoldType): Exit Function
    If dtpEnd.Value <> dtpHoldBegin.Value Then MsgBox "交班结束时间与接班开始时间不一致，请检查!", vbExclamation, Me.Caption: Call zlControl.ControlSetFocus(cboHoldType): Exit Function
    mrsDoc.Filter = "姓名='" & txtHold.Text & "'"
    If mrsDoc.RecordCount = 0 Then
        MsgBox "接班医生不属于当前所属科室，请重新选择！", vbExclamation, Me.Caption
        Exit Function
    End If
    CheckRecordData = True
End Function

Private Sub dtpBegin_CloseUp()
    
    mrsTime.Filter = "班次名称='" & cboType.Text & "'"
    If mrsTime.RecordCount = 1 Then
        If mrsTime!开始时间 >= mrsTime!结束时间 Then
            dtpEnd.Value = Format(dtpBegin.Value + 1, "yyyy-mm-dd") & " " & mrsTime!结束时间
        Else
            dtpEnd.Value = Format(dtpBegin.Value, "yyyy-mm-dd") & " " & mrsTime!结束时间
        End If
        Call cboHoldType_Click
    End If
End Sub

Private Sub dtpEnd_CloseUp()
    Call cboHoldType_Click
End Sub

Private Sub Form_Load()
    Dim varTemp As Variant, varData As Variant
    Dim i As Long

    Call SetBasic
    varTemp = Split(mstrDept, "|")
    Select Case mbytType
        Case 0
            '新增时科室与主界面的科室一致
            For i = 1 To UBound(varTemp)
                varData = Split(varTemp(i), ",")
                cboDept.AddItem varData(0)
                cboDept.ItemData(cboDept.NewIndex) = varData(1)
                mstrDeptId = IIf(mstrDeptId = "", "", mstrDeptId & ",") & varData(1)
            Next
            cboDept.ListIndex = IIf(varTemp(0) < 0, 0, varTemp(0))
        Case 1
            cboDept.AddItem varTemp(0)
            cboDept.ListIndex = 0
            mstrDeptId = varTemp(1)
    End Select
    '交班人、交班班次、交班时期
    txtPer.Text = mstrOutPer
    varTemp = Split(mstrOutTime, "|")
    If UBound(varTemp) > 0 Then
        cboType.AddItem varTemp(0)
        cboType.ListIndex = 0
        dtpBegin.Value = varTemp(1)
        dtpEnd.Value = varTemp(2)
    End If
    '接班人、接班班次、接班时期
    txtHold.Text = mstrInPer
    varTemp = Split(mstrInTime, "|")
    If UBound(varTemp) > 0 Then
        cboHoldType.AddItem varTemp(0)
        cboHoldType.ListIndex = 0
        dtpHoldBegin.Value = varTemp(1)
        dtpHoldEnd.Value = varTemp(2)
    End If
    Call ShowDoc
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsTime = Nothing
    Set mrsDoc = Nothing
    Set mrsPati = Nothing
End Sub

Private Sub tvwDoc_DblClick()
    
    If Not tvwDoc.SelectedItem.Parent Is Nothing Then
        txtHold.Text = tvwDoc.SelectedItem.Text
        Call tvwDoc_LostFocus
    End If
End Sub

Private Sub tvwDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Call tvwDoc_LostFocus
    End If
End Sub

Private Sub tvwDoc_LostFocus()
    tvwDoc.Visible = False
End Sub

Private Sub txtHold_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strTemp As String
        
    If KeyCode = vbKeyReturn Then
        strTemp = UCase(txtHold.Text)
        mrsDoc.MoveFirst
        Do While Not mrsDoc.EOF
            If InStr(mrsDoc!姓名, strTemp) > 0 Or InStr(mrsDoc!简码, strTemp) > 0 Then
                txtHold.Text = mrsDoc!姓名
                Exit Do
            End If
            mrsDoc.MoveNext
        Loop
    End If
End Sub

Private Sub txtHold_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Or KeyAscii = Asc("%") Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txtHold_Validate(Cancel As Boolean)

    '点取消也会触发这个事件，故在保存的时候检查
'    If txtHold.Text = "" Then Exit Sub
'    mrsDoc.Filter = "姓名='" & txtHold.Text & "'"
'    If mrsDoc.RecordCount = 0 Then
'        MsgBox "接班医生不属于当前所属科室，请重新选择！"
'    End If
End Sub

Private Sub vsfPaiType_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    
    With vsfPaiType
        If Col = 0 Then
            '实现全选和全消的功能
            If Row = 0 Then
                If .Cell(flexcpChecked, 0, 0) = flexChecked Then
                    .Cell(flexcpChecked, 0, 0) = flexChecked
                    For i = .FixedRows To .Rows - .FixedRows
                        .Cell(flexcpChecked, i, 0) = flexChecked
                    Next
                Else
                    .Cell(flexcpChecked, 0, 0) = flexUnchecked
                    For i = .FixedRows To .Rows - .FixedRows
                        .Cell(flexcpChecked, i, 0) = flexUnchecked
                    Next
                End If
            Else
                If .Cell(flexcpChecked, 0, 0) = flexChecked Then
                    .Cell(flexcpChecked, 0, 0) = flexUnchecked
                End If
                For i = .FixedRows To .Rows - .FixedRows
                    If .Cell(flexcpChecked, i, 0) = flexUnchecked Then: Exit For
                    If i = .Rows - .FixedRows Then
                        .Cell(flexcpChecked, 0, 0) = flexChecked
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Sub vsfPaiType_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Or NewRow < 1 Then Exit Sub
    With vsfPaiType
        If NewRow = 1 Then
            .Cell(flexcpPicture, NewRow, .ColIndex("上移")) = ""
            .Cell(flexcpPicture, NewRow, .ColIndex("下移")) = imgList.ListImages("Down").Picture
        Else
            If NewRow = .Rows - 1 Then
                .Cell(flexcpPicture, NewRow, .ColIndex("下移")) = ""
                .Cell(flexcpPicture, NewRow, .ColIndex("上移")) = imgList.ListImages("Up").Picture
            Else
                .Cell(flexcpPicture, NewRow, .ColIndex("上移")) = imgList.ListImages("Up").Picture
                .Cell(flexcpPicture, NewRow, .ColIndex("下移")) = imgList.ListImages("Down").Picture
            End If
        End If
        If OldRow < .Rows Then
            .Cell(flexcpPicture, OldRow, .ColIndex("上移")) = ""
            .Cell(flexcpPicture, OldRow, .ColIndex("下移")) = ""
        End If
    End With
End Sub

Private Sub vsfPaiType_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfPaiType
        If Col <> .ColIndex("选择") Then Cancel = True
    End With
End Sub

Private Sub vsfPaiType_Click()
    Dim lngCheck As Long, lngNum As Long, lngRow As Long
    Dim strPati As String, strName As String, strSQL As String
    
    With vsfPaiType
        If .Row < 1 Then Exit Sub
        If .Col = .ColIndex("上移") Then
            If Not .Cell(flexcpPicture, .Row, .ColIndex("上移")) Is Nothing Then
                lngRow = .Row - 1
            End If
        ElseIf .Col = .ColIndex("下移") Then
            If Not .Cell(flexcpPicture, .Row, .ColIndex("下移")) Is Nothing Then
                lngRow = .Row + 1
            End If
        End If
        If lngRow = 0 Then Exit Sub
        lngCheck = .Cell(flexcpChecked, .Row, .ColIndex("选择"))
        strPati = .TextMatrix(.Row, .ColIndex("病人类型"))
        strName = .TextMatrix(.Row, .ColIndex("简称"))
        strSQL = .TextMatrix(.Row, .ColIndex("提取SQL"))
        
        .Cell(flexcpChecked, .Row, .ColIndex("选择")) = .Cell(flexcpChecked, lngRow, .ColIndex("选择"))
        .TextMatrix(.Row, .ColIndex("病人类型")) = .TextMatrix(lngRow, .ColIndex("病人类型"))
        .TextMatrix(.Row, .ColIndex("简称")) = .TextMatrix(lngRow, .ColIndex("简称"))
        .TextMatrix(.Row, .ColIndex("提取SQL")) = .TextMatrix(lngRow, .ColIndex("提取SQL"))
        .Cell(flexcpChecked, lngRow, .ColIndex("选择")) = lngCheck
        .TextMatrix(lngRow, .ColIndex("病人类型")) = strPati
        .TextMatrix(lngRow, .ColIndex("简称")) = strName
        .TextMatrix(lngRow, .ColIndex("提取SQL")) = strSQL
        .Row = lngRow
        .ShowCell lngRow, 1
    End With
    
End Sub


