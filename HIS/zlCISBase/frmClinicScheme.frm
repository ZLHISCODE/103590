VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicScheme 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "成套方案"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   Icon            =   "frmClinicScheme.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkAll 
      Caption         =   "调用本方案时全选"
      Height          =   270
      Left            =   4890
      TabIndex        =   47
      ToolTipText     =   "勾选时医嘱下达调用本方案时默认全选所有项目，否则不选任何项目。"
      Top             =   2310
      Width           =   1770
   End
   Begin VB.TextBox txtFind 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   300
      Left            =   5160
      TabIndex        =   26
      Top             =   3225
      Width           =   975
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Enabled         =   0   'False
      Height          =   300
      Left            =   6120
      TabIndex        =   27
      Top             =   3225
      Width           =   855
   End
   Begin VB.TextBox txt建档时间 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   300
      Left            =   4560
      MaxLength       =   13
      TabIndex        =   45
      Top             =   5280
      Width           =   2370
   End
   Begin VB.TextBox txt建档人 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1080
      MaxLength       =   13
      TabIndex        =   43
      Top             =   5280
      Width           =   1890
   End
   Begin VB.Frame fraline 
      Height          =   45
      Index           =   3
      Left            =   0
      TabIndex        =   41
      Top             =   5640
      Width           =   7335
   End
   Begin VB.ComboBox cmbStationNo 
      Height          =   300
      Left            =   825
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   2250
      Visible         =   0   'False
      Width           =   2220
   End
   Begin MSComctlLib.ListView lvw科室 
      Height          =   1380
      Left            =   1125
      TabIndex        =   29
      Top             =   3615
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   2434
      View            =   2
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483633
      Appearance      =   1
      Enabled         =   0   'False
      NumItems        =   0
   End
   Begin VB.CheckBox chk范围 
      Caption         =   "住院使用(&I)"
      Height          =   195
      Index           =   1
      Left            =   5805
      TabIndex        =   22
      Top             =   2880
      Value           =   1  'Checked
      Width           =   1290
   End
   Begin VB.CheckBox chk范围 
      Caption         =   "门诊使用(&C)"
      Height          =   195
      Index           =   0
      Left            =   4470
      TabIndex        =   21
      Top             =   2880
      Value           =   1  'Checked
      Width           =   1290
   End
   Begin VB.Frame fraline 
      Height          =   30
      Index           =   2
      Left            =   0
      TabIndex        =   38
      Top             =   2640
      Width           =   8490
   End
   Begin VB.OptionButton opt范围 
      Caption         =   "个人(&1)"
      Height          =   180
      Index           =   0
      Left            =   1110
      TabIndex        =   18
      Top             =   2880
      Width           =   930
   End
   Begin VB.OptionButton opt范围 
      Caption         =   "全院(&3)"
      Height          =   180
      Index           =   2
      Left            =   3135
      TabIndex        =   20
      Top             =   2880
      Value           =   -1  'True
      Width           =   930
   End
   Begin VB.OptionButton opt范围 
      Caption         =   "科室(&2)"
      Height          =   180
      Index           =   1
      Left            =   2115
      TabIndex        =   19
      Top             =   2880
      Width           =   930
   End
   Begin VB.ComboBox cbo人员 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1125
      TabIndex        =   24
      Top             =   3225
      Width           =   3030
   End
   Begin VB.CommandButton cmdScheme 
      Caption         =   "方案内容(&E)…"
      Height          =   350
      Left            =   1590
      TabIndex        =   30
      Top             =   5775
      Width           =   1590
   End
   Begin VB.TextBox txt说明 
      Height          =   300
      Left            =   825
      MaxLength       =   30
      TabIndex        =   16
      Top             =   1875
      Width           =   5835
   End
   Begin VB.TextBox txt名称 
      Height          =   300
      Index           =   1
      Left            =   825
      MaxLength       =   40
      TabIndex        =   11
      Top             =   1485
      Width           =   2250
   End
   Begin VB.TextBox txt拼音 
      Height          =   300
      Index           =   1
      Left            =   4080
      MaxLength       =   12
      TabIndex        =   13
      Top             =   1485
      Width           =   960
   End
   Begin VB.TextBox txt五笔 
      Height          =   300
      Index           =   1
      Left            =   5700
      MaxLength       =   12
      TabIndex        =   14
      Top             =   1485
      Width           =   960
   End
   Begin VB.Frame fraline 
      Height          =   30
      Index           =   1
      Left            =   -45
      TabIndex        =   37
      Top             =   5160
      Width           =   8490
   End
   Begin VB.TextBox txt五笔 
      Height          =   300
      Index           =   0
      Left            =   5700
      MaxLength       =   12
      TabIndex        =   9
      Top             =   1110
      Width           =   960
   End
   Begin VB.TextBox txt拼音 
      Height          =   300
      Index           =   0
      Left            =   4080
      MaxLength       =   12
      TabIndex        =   8
      Top             =   1110
      Width           =   960
   End
   Begin VB.TextBox txt名称 
      Height          =   300
      Index           =   0
      Left            =   825
      MaxLength       =   40
      TabIndex        =   6
      Top             =   1110
      Width           =   2250
   End
   Begin VB.TextBox txt编码 
      Height          =   300
      Left            =   825
      MaxLength       =   13
      TabIndex        =   1
      Top             =   735
      Width           =   2250
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   3570
      Left            =   180
      TabIndex        =   35
      TabStop         =   0   'False
      Tag             =   "1000"
      Top             =   6345
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   6297
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5865
      TabIndex        =   32
      Top             =   5775
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   420
      Picture         =   "frmClinicScheme.frx":058A
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5775
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4800
      TabIndex        =   31
      Top             =   5775
      Width           =   1100
   End
   Begin VB.TextBox txt分类 
      Height          =   300
      Left            =   4080
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   3
      Top             =   735
      Width           =   2580
   End
   Begin VB.CommandButton cmd分类 
      Caption         =   "&P"
      Height          =   285
      Left            =   6675
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   735
      Width           =   285
   End
   Begin VB.Frame fraline 
      Height          =   60
      Index           =   0
      Left            =   -30
      TabIndex        =   36
      Top             =   540
      Width           =   8490
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   3780
      Top             =   6375
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
            Picture         =   "frmClinicScheme.frx":06D4
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicScheme.frx":0C6E
            Key             =   "expend"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "查找科室："
      Height          =   180
      Left            =   4320
      TabIndex        =   25
      Top             =   3285
      Width           =   900
   End
   Begin VB.Label lbl建档时间 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "建档时间(&T)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3480
      TabIndex        =   46
      Top             =   5340
      Width           =   990
   End
   Begin VB.Label lbl建档人 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "建档人(&P)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   44
      Top             =   5340
      Width           =   810
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   15
      Left            =   480
      TabIndex        =   42
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label lblStationNo 
      AutoSize        =   -1  'True
      Caption         =   "院区(&C)"
      Height          =   180
      Left            =   135
      TabIndex        =   40
      Top             =   2310
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label lbl科室 
      AutoSize        =   -1  'True
      Caption         =   "使用科室："
      Height          =   180
      Left            =   210
      TabIndex        =   28
      Top             =   3690
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "使用范围："
      Height          =   180
      Left            =   210
      TabIndex        =   17
      Top             =   2880
      Width           =   900
   End
   Begin VB.Label lbl人员 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "使用人员："
      Height          =   180
      Left            =   210
      TabIndex        =   23
      Top             =   3285
      Width           =   900
   End
   Begin VB.Label lbl说明 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "说明(&Z)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   135
      TabIndex        =   15
      Top             =   1920
      Width           =   630
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   390
      Picture         =   "frmClinicScheme.frx":1208
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbl名称 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "别名(&A)"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   135
      TabIndex        =   10
      Top             =   1545
      Width           =   630
   End
   Begin VB.Label lbl简码 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "简码(&M)           (拼音)            (五笔)"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   3420
      TabIndex        =   12
      Top             =   1545
      Width           =   3780
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    根据常用的典型医嘱，经过适当筛选，形成成套的医嘱方案，以方便医生快速地下达病人医嘱。"
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   1155
      TabIndex        =   34
      Top             =   105
      Width           =   5925
   End
   Begin VB.Label lbl简码 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "简码(&S)           (拼音)            (五笔)"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   3420
      TabIndex        =   7
      Top             =   1170
      Width           =   3780
   End
   Begin VB.Label lbl名称 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "名称(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   5
      Top             =   1170
      Width           =   630
   End
   Begin VB.Label lbl编码 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "编码(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   135
      TabIndex        =   0
      Top             =   795
      Width           =   630
   End
   Begin VB.Label lbl分类 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "分类(&F)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3420
      TabIndex        =   2
      Top             =   795
      Width           =   630
   End
End
Attribute VB_Name = "frmClinicScheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'说明：
'   1、上级程序通过本窗体ShowMe函数，将父窗体、权限、编辑项目的分类ID、ID,编辑状态等信息传递进入本程序
'   2、编辑状态：由Me.tag存放，分别为"增加"、"修改"、"查阅"，由上级程序通过ShowMe传入
'---------------------------------------------------
Private mint范围 As Integer '1-门诊,2-住院,3-门诊和住院
Private mstrPrivs As String
Private lngClassId As Long       '被编辑的分类ID，上级程序通过ShowMe传递进入
Private lngItemId As Long        '被编辑的项目ID，修改、查阅时由上级程序通过ShowMe传递进入,增加时为0，
Private mblnNoCheck As Boolean
Private mblnFirst As Boolean
Private mstrLike As String
Private mblnChange As Boolean
Private mlngFind As Long

Private rsTemp As New ADODB.Recordset
Private mrsScheme As ADODB.Recordset
Private mblnOK As Boolean

Public Function ShowMe(ByVal frmParent As Object, ByVal strPrivs As String, ByVal byt状态 As Byte, _
    ByVal lng分类id As Long, Optional ByVal lng项目id As Long, Optional ByVal int范围 As Integer = 3) As Boolean
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '---------------------------------------------------
    Dim objNode As Node
    
    mint范围 = int范围
    mstrPrivs = strPrivs
    
    Me.Tag = Switch(byt状态 = 0, "增加", byt状态 = 1, "修改", byt状态 = 2, "查阅")
    Me.Caption = "成套方案" & Me.Tag
    lngClassId = lng分类id: lngItemId = lng项目id
    
    '填写需要选择的数据
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select ID,上级ID,编码,名称,简码" & _
                " From 诊疗分类目录 Where 类型=6 And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Start With 上级ID is Null Connect by Prior ID=上级ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "请首先建立配方诊疗分类项目之后增加配方", vbExclamation, gstrSysName
            Unload Me: Exit Function
        End If
        
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!上级ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !编码 & "]" & !名称, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !上级ID, tvwChild, "_" & !ID, "[" & !编码 & "]" & !名称, "close")
            End If
            objNode.Sorted = True
            objNode.Tag = IIf(IsNull(!简码), "", !简码)
            objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
        Me.tvwClass.Nodes("_" & lng分类id).Selected = True
        Me.txt分类.Text = Me.tvwClass.SelectedItem.Text
        Me.txt分类.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    End With
    
    '显示窗体
    Me.Show 1, frmParent
    ShowMe = mblnOK
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub GetDefineSize()
    '功能：得到数据库的表字段的长度
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset
    gstrSql = "Select A.编码,A.标本部位,B.名称,B.简码 From 诊疗项目目录 A, 诊疗项目别名 B Where A.ID=B.诊疗项目ID and A.ID=0"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSql, Me.Caption)
    
    txt编码.MaxLength = rsTmp.Fields("编码").DefinedSize
    txt名称(0).MaxLength = rsTmp.Fields("名称").DefinedSize
    txt名称(1).MaxLength = rsTmp.Fields("名称").DefinedSize
    txt拼音(0).MaxLength = rsTmp.Fields("简码").DefinedSize
    txt拼音(1).MaxLength = rsTmp.Fields("简码").DefinedSize
    txt五笔(0).MaxLength = rsTmp.Fields("简码").DefinedSize
    txt五笔(1).MaxLength = rsTmp.Fields("简码").DefinedSize
    txt说明.MaxLength = rsTmp.Fields("标本部位").DefinedSize

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo人员_Click()
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long, strSql As String
    Dim intIdx As Integer, i As Long
    Dim vRect As RECT, blnCancel As Boolean
    
    If cbo人员.ListIndex <> -1 Then
        If cbo人员.ItemData(cbo人员.ListIndex) = 0 Then
            strSql = "Select ID,编号,姓名,简码,性别 From 人员表 Where 撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null Order by 编号"
            vRect = zlControl.GetControlRect(cbo人员.hWnd)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSql, 0, "人员", , , , , , True, vRect.Left, vRect.Top, cbo人员.Height, blnCancel, , True)
            If Not rsTmp Is Nothing Then
                intIdx = Cbo.FindIndex(cbo人员, rsTmp!ID)
                If intIdx <> -1 Then
                    cbo人员.ListIndex = intIdx
                Else
                    cbo人员.AddItem rsTmp!编号 & "-" & rsTmp!姓名, 0
                    cbo人员.ItemData(cbo人员.NewIndex) = rsTmp!ID
                    cbo人员.ListIndex = cbo人员.NewIndex
                End If
                mblnChange = True
            Else
                If Not blnCancel Then
                    MsgBox "没有人员数据，请先到人员管理中设置。", vbInformation, gstrSysName
                End If
                Call zlControl.CboSetIndex(cbo人员.hWnd, Val(cbo人员.Tag))
            End If
        Else
            cbo人员.Tag = cbo人员.ListIndex
        End If
    Else
        cbo人员.Tag = cbo人员.ListIndex
    End If
End Sub

Private Sub cbo人员_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo人员.ListIndex = -1 Then
            Call cbo人员_Validate(blnCancel)
        End If
        If Not blnCancel Then Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo人员_Validate(Cancel As Boolean)
'功能：根据输入的内容,自动匹配执行科室
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, intIdx As Long, i As Long
    Dim vRect As RECT, blnCancel As Boolean
    
    If cbo人员.ListIndex <> -1 Then Exit Sub '已选中
    If cbo人员.Text = "" Then Exit Sub '无输入
    
    On Error GoTo errH
    
    strSql = "Select ID,编号,姓名,简码,性别 From 人员表" & _
        " Where Upper(编号) Like '" & UCase(cbo人员.Text) & "%'" & _
        " Or Upper(姓名) Like '" & mstrLike & UCase(cbo人员.Text) & "%'" & _
        " Or Upper(简码) Like '" & mstrLike & UCase(cbo人员.Text) & "%'" & _
        " And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) " & _
        " Order by 编号"
    vRect = zlControl.GetControlRect(cbo人员.hWnd)
    Set rsTmp = zlDatabase.ShowSelect(Me, strSql, 0, "人员", , , , , , True, vRect.Left, vRect.Top, cbo人员.Height, blnCancel, , True)
    If Not rsTmp Is Nothing Then
        intIdx = Cbo.FindIndex(cbo人员, rsTmp!ID)
        If intIdx <> -1 Then
            cbo人员.ListIndex = intIdx
        Else
            cbo人员.AddItem rsTmp!编号 & "-" & rsTmp!姓名, 0
            cbo人员.ItemData(cbo人员.NewIndex) = rsTmp!ID
            cbo人员.ListIndex = cbo人员.NewIndex
        End If
        mblnChange = True
    Else
        If Not blnCancel Then
            MsgBox "未找到对应的人员。", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub chkAll_Click()
    If mblnNoCheck Then Exit Sub
    mblnChange = True
End Sub

Private Sub chk范围_Click(Index As Integer)
    
    If mblnNoCheck Then Exit Sub
    
    If Index = 1 And chk范围((Index + 1) Mod 2).Value = 1 And chk范围(Index).Value = 0 Then
        If Not mrsScheme Is Nothing Then
            mrsScheme.Filter = "期效=0"
            If mrsScheme.RecordCount > 0 Then
                MsgBox "此成套方案中存在长嘱，不能设置为仅使用于门诊！", vbInformation, gstrSysName
                mblnNoCheck = True
                chk范围(Index).Value = 1
                mblnNoCheck = False
                mrsScheme.Filter = "": Exit Sub
            End If
            mrsScheme.Filter = ""
        End If
        
    ElseIf chk范围((Index + 1) Mod 2).Value = 0 And chk范围(Index).Value = 0 Then
        mblnNoCheck = True
        chk范围(Index).Value = 1
        mblnNoCheck = False
        Exit Sub
    End If
    
    If InStr(mstrPrivs, "全院成套方案") > 0 Then
        Call LoadDeptList(True)
    ElseIf InStr(mstrPrivs, "本科成套方案") > 0 Then
        Call LoadDeptList(False)
    Else
    End If
    
    mblnChange = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me: Exit Sub
End Sub

Private Sub cmdFind_Click()
    Dim strFind As String
    Dim i As Long
    Dim blnIsFind As Boolean
    
    strFind = UCase(Trim(txtFind.Text))
    If strFind = "" Then Exit Sub
    For i = mlngFind To Lvw科室.ListItems.Count
        If zlCommFun.SpellCode(Mid(Lvw科室.ListItems(i).Text, InStr(Lvw科室.ListItems(i).Text, "-") + 1)) Like UCase(IIf(gstrMatch <> "", "*", "") & strFind & "*") Or _
                UCase(Lvw科室.ListItems(i).Text) Like UCase(IIf(gstrMatch <> "", "*", "") & strFind & "*") Then
            Lvw科室.ListItems(i).Selected = True
            Lvw科室.ListItems(i).EnsureVisible
            Lvw科室.SetFocus
            blnIsFind = True
            mlngFind = i + 1
            Exit For
        End If
    Next
    If blnIsFind = False Then
        If mlngFind = 1 Then
            MsgBox "没有找到您查找的科室。", vbInformation, Me.Caption
        Else
            MsgBox "已经是最后一个科室了。", vbInformation, Me.Caption
            mlngFind = 1
        End If
    End If
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Function Get服务对象() As Integer
    If chk范围(0).Value = 1 And chk范围(1).Value = 1 Then
        Get服务对象 = 3
    ElseIf chk范围(0).Value = 1 Then
        Get服务对象 = 1
    ElseIf chk范围(1).Value = 1 Then
        Get服务对象 = 2
    End If
End Function

Private Sub cmdOK_Click()
    Dim arrSql() As Variant
    Dim strTmp As String, i As Long
    Dim str科室IDs As String, lng人员ID As Long
    Dim strSql As String
    Dim str站点 As String
    Dim str编码 As String
    
    '重新检查名称，并去掉特殊字符
    strTmp = MoveSpecialChar(txt名称(0).Text)
    If txt名称(0).Text <> strTmp Then
        txt名称(0).Text = strTmp
        Me.txt拼音(0).Text = zlStr.GetCodeByORCL(Me.txt名称(0).Text, False)
        Me.txt五笔(0).Text = zlStr.GetCodeByORCL(Me.txt名称(0).Text, True)
    End If
    strTmp = MoveSpecialChar(txt名称(1).Text)
    If txt名称(1).Text <> strTmp Then
        txt名称(1).Text = strTmp
        Me.txt拼音(1).Text = zlStr.GetCodeByORCL(Me.txt名称(1).Text, False)
        Me.txt五笔(1).Text = zlStr.GetCodeByORCL(Me.txt名称(1).Text, True)
    End If
    
    '一般特性检查
    If Trim(Me.txt编码.Text) = "" Then MsgBox "请输入编码！", vbInformation, gstrSysName: Me.txt编码.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt编码.Text), vbFromUnicode)) > Me.txt编码.MaxLength Then MsgBox "编码的超长（最多" & Me.txt编码.MaxLength & "个字符）！", vbInformation, gstrSysName: Me.txt编码.SetFocus: Exit Sub
    If Trim(Me.txt名称(0).Text) = "" Then MsgBox "请输入名称！", vbInformation, gstrSysName: Me.txt名称(0).SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt名称(0).Text), vbFromUnicode)) > Me.txt名称(0).MaxLength Then
        MsgBox "名称超长（" & Me.txt名称(0).MaxLength & "个字符或" & Me.txt名称(0).MaxLength / 2 & "个汉字）！", vbInformation, gstrSysName: Me.txt名称(0).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt名称(1).Text), vbFromUnicode)) > Me.txt名称(1).MaxLength Then
        MsgBox "别名超长（" & Me.txt名称(1).MaxLength & "个字符或" & Me.txt名称(1).MaxLength / 2 & "个汉字）！", vbInformation, gstrSysName: Me.txt名称(1).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt拼音(0).Text), vbFromUnicode)) > Me.txt拼音(0).MaxLength Then
        MsgBox "名称拼音简码超长（" & Me.txt拼音(0).MaxLength & "个字符）！", vbInformation, gstrSysName: Me.txt拼音(0).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt拼音(1).Text), vbFromUnicode)) > Me.txt拼音(1).MaxLength Then
        MsgBox "别名拼音简码超长（" & Me.txt拼音(1).MaxLength & "个字符）！", vbInformation, gstrSysName: Me.txt拼音(1).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt五笔(0).Text), vbFromUnicode)) > Me.txt五笔(0).MaxLength Then
        MsgBox "名称五笔简码超长（" & Me.txt五笔(0).MaxLength & "个字符）！", vbInformation, gstrSysName: Me.txt五笔(0).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt五笔(1).Text), vbFromUnicode)) > Me.txt五笔(1).MaxLength Then
        MsgBox "别名五笔简码超长（" & Me.txt五笔(1).MaxLength & "个字符）！", vbInformation, gstrSysName: Me.txt五笔(1).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt说明.Text), vbFromUnicode)) > Me.txt说明.MaxLength Then
        MsgBox "说明超长（" & Me.txt说明.MaxLength & "个字符或" & Me.txt说明.MaxLength / 2 & "个汉字）！", vbInformation, gstrSysName: Me.txt说明.SetFocus: Exit Sub
    End If
    
    '新增项目时，保证不出现重复编码，如果有重复自动在原编码基础上加1，直到不重复
    str编码 = Trim(txt编码.Text)
    If Me.Tag = "增加" Then
        Do While True
            gstrSql = "select a.编码 from 诊疗项目目录 a,诊疗项目类别 b where a.编码=[1] and a.类别=b.编码"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "编码是否重复", str编码)
            If rsTemp.RecordCount <> 0 Then
                str编码 = zlCommFun.IncStr(str编码)
            Else
                Exit Do
            End If
        Loop
    End If
    
    '使用范围检查
    If opt范围(0).Value Then
        If cbo人员.ListIndex = -1 Then
            MsgBox "请指定成套方案的使用人员。", vbInformation, gstrSysName
            cbo人员.SetFocus: Exit Sub
        End If
        lng人员ID = cbo人员.ItemData(cbo人员.ListIndex)
    ElseIf opt范围(1).Value Then
        For i = 1 To Lvw科室.ListItems.Count
            If Lvw科室.ListItems(i).Checked Then
                str科室IDs = str科室IDs & "," & Mid(Lvw科室.ListItems(i).Key, 2)
            End If
        Next
        If str科室IDs = "" Then
            MsgBox "请指定成套方案的使用科室。", vbInformation, gstrSysName
            Lvw科室.SetFocus: Exit Sub
        End If
        str科室IDs = Mid(str科室IDs, 2)
    End If
    
    '内容检查
    If mrsScheme Is Nothing Then
        MsgBox "成套方案中没有内容，请先录入成套方案内容！", vbInformation, gstrSysName
        cmdScheme.SetFocus: Exit Sub
    ElseIf mrsScheme.RecordCount = 0 Then
        MsgBox "成套方案中没有内容，请先录入成套方案内容！", vbInformation, gstrSysName
        cmdScheme.SetFocus: Exit Sub
    End If
    
    If cmbStationNo.Text = "" Then
        str站点 = "Null"
    Else
        str站点 = Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1)
    End If
    
    '数据保存
    arrSql = Array()
    If Me.Tag = "增加" Then
        lngItemId = zlDatabase.GetNextId("诊疗项目目录")
    Else
        If zlClinicCodeRepeat(str编码, lngItemId) = True Then Exit Sub
    End If
    
    ReDim Preserve arrSql(UBound(arrSql) + 1)
    arrSql(UBound(arrSql)) = "ZL_成套方案项目_Update(" & _
        lngItemId & "," & Val(Me.txt分类.Tag) & ",'" & str编码 & "'," & _
        "'" & Trim(Me.txt名称(0).Text) & "','" & Trim(Me.txt拼音(0).Text) & "','" & Trim(Me.txt五笔(0).Text) & "'," & _
        "'" & Trim(Me.txt名称(1).Text) & "','" & Trim(Me.txt拼音(1).Text) & "','" & Trim(Me.txt五笔(1).Text) & "'," & _
        "'" & Trim(Me.txt说明.Text) & "'," & IIf(opt范围(0).Value, lng人员ID, "Null") & "," & _
        IIf(opt范围(1).Value, "'" & str科室IDs & "'", "Null") & "," & Get服务对象 & "," & IIf(cmbStationNo.Visible = False Or cmbStationNo.Text = "", "Null", str站点) & ",'" & UserInfo.姓名 & "'," & chkAll.Value & ")"
    
    If mrsScheme.RecordCount > 0 Then mrsScheme.MoveFirst
    Do While Not mrsScheme.EOF
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = "ZL_成套方案内容_Insert(" & _
            lngItemId & "," & mrsScheme!序号 & "," & ZVal(Nvl(mrsScheme!相关序号, 0)) & "," & _
            mrsScheme!期效 & "," & ZVal(Nvl(mrsScheme!诊疗项目id, 0)) & "," & _
            IIf(IsNull(mrsScheme!诊疗项目id), "'" & Nvl(mrsScheme!医嘱内容) & "',", "NULL,") & _
            ZVal(Nvl(mrsScheme!天数, 0)) & "," & ZVal(Nvl(mrsScheme!单次用量, 0)) & "," & ZVal(Nvl(mrsScheme!总给予量, 0)) & "," & _
            ZVal(Nvl(mrsScheme!收费细目ID, 0)) & ",'" & Nvl(mrsScheme!标本部位) & "'," & _
            "'" & Nvl(mrsScheme!执行频次) & "'," & ZVal(Nvl(mrsScheme!频率次数, 0)) & "," & _
            ZVal(Nvl(mrsScheme!频率间隔, 0)) & ",'" & Nvl(mrsScheme!间隔单位) & "'," & _
            "'" & Nvl(mrsScheme!医生嘱托) & "'," & Nvl(mrsScheme!执行性质, 0) & "," & _
            ZVal(Nvl(mrsScheme!执行科室ID, 0)) & ",'" & Nvl(mrsScheme!时间方案) & "'," & _
            "'" & Nvl(mrsScheme!检查方法) & "'," & ZVal(Val(mrsScheme!配方ID & "")) & "," & _
            ZVal(Val(mrsScheme!组合项目ID & "")) & "," & Val(mrsScheme!执行标记 & "") & ")"
        mrsScheme.MoveNext
    Loop

    On Error GoTo ErrHand
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    If Me.Tag = "增加" Then
        If Val(zlDatabase.GetPara("诊疗项目连续增加", glngSys, 1054, 0)) = 1 Then
            lngItemId = 0: mblnFirst = True
            Call Form_Activate
            Me.txt编码.SetFocus
            Exit Sub
        End If
    End If
    mblnChange = False
    mblnOK = True
    Unload Me
    Exit Sub
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdScheme_Click()
    Dim rsTmp As ADODB.Recordset
    Dim str使用科室 As String, i As Long
    
    If Lvw科室.Enabled Then
        For i = 1 To Lvw科室.ListItems.Count
            If Lvw科室.ListItems(i).Checked Then str使用科室 = str使用科室 & "," & Mid(Lvw科室.ListItems(i).Key, 2)
        Next
    End If
    
'    '测试代码
'    Dim mobjCISKernel As New clsCISKernel
'    Call mobjCISKernel.InitCISKernel(gcnOracle, Me, glngSys, mstrPrivs)
'    Set rsTmp = mobjCISKernel.ShowSchemeEdit(Me, Get服务对象, mrsScheme, Me.Tag = "查阅", , Mid(str使用科室, 2))

    Call gobjKernel.InitCISKernel(gcnOracle, Me, glngSys, mstrPrivs)
    Set rsTmp = gobjKernel.ShowSchemeEdit(Me, Get服务对象, mrsScheme, Me.Tag = "查阅", , Mid(str使用科室, 2))
    If Not rsTmp Is Nothing Then
        Set mrsScheme = rsTmp
        mblnChange = True
    End If
End Sub

Private Sub cmd分类_Click()
    With Me.tvwClass
        .Left = Me.txt分类.Left
        .Top = Me.txt分类.Top + Me.txt分类.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
End Sub

Private Sub Form_Activate()
    Dim strTemp As String
    Dim bln修改 As Boolean
    
    
    If Not mblnFirst Then Exit Sub
    mblnFirst = False
    bln修改 = True
    
    '查阅时先设置界面不可编辑(OptionButton在Enabled时值会变化)
    '-------------------------------------------------
    If Me.Tag = "查阅" Then
        Me.cmdOk.Visible = False
        Me.cmdCancel.Caption = "关闭(&C)"
        Me.txt分类.Enabled = False: Me.cmd分类.Enabled = False
        Me.txt编码.Enabled = False
        Me.txt名称(0).Enabled = False: Me.txt拼音(0).Enabled = False: Me.txt五笔(0).Enabled = False
        Me.txt名称(1).Enabled = False: Me.txt拼音(1).Enabled = False: Me.txt五笔(1).Enabled = False
        Me.txt说明.Enabled = False
        
        opt范围(0).Enabled = False: opt范围(1).Enabled = False: opt范围(2).Enabled = False
        chk范围(0).Enabled = False: chk范围(1).Enabled = False
        cbo人员.Enabled = False: cbo人员.BackColor = vbButtonFace
        Lvw科室.Enabled = False: Lvw科室.BackColor = vbButtonFace
    End If
    
    Err = 0: On Error GoTo ErrHand
    
    '提取执行项目的信息
    '-------------------------------------------------
    If Me.Tag = "增加" Then
        
        lngItemId = 0
        Set mrsScheme = Nothing '连续增加有用

        If Val(zlDatabase.GetPara(61, glngSys)) = 0 Then '诊疗项目编码递增模式
            gstrSql = "Select Nvl(Max(编码),'0000000') as 编码 From 诊疗项目目录"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
            Me.txt编码.Text = Right(String(10, "0") & Val(rsTemp!编码) + 1, Len(rsTemp!编码))
        Else
            strTemp = Mid(Me.txt分类.Text, 2, InStr(1, Me.txt分类.Text, "]") - 2)
            gstrSql = "Select Nvl(Max(编码),'0000000') as 编码" & _
                    " From 诊疗项目目录" & _
                    " Where 编码 like [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, "9" & strTemp & "%")
            
            Err = 0: On Error Resume Next
            Me.txt编码.Text = "9" & strTemp & Right(String(10, "0") & Val(rsTemp!编码) + 1, Len(rsTemp!编码) - 1 - Len(strTemp))
        End If

        Me.txt名称(0).Text = "": Me.txt名称(1).Text = ""
        Me.txt拼音(0).Text = "": Me.txt拼音(1).Text = ""
        Me.txt五笔(0).Text = "": Me.txt五笔(1).Text = ""
        Me.txt说明.Text = ""
        Me.txt建档人.Text = UserInfo.姓名
        Me.txt建档时间.Text = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    Else
        '显示基本信息
        gstrSql = "Select A.编码,A.名称,A.标本部位 as 说明,A.服务对象,A.人员ID,B.编号,B.姓名,A.站点,A.建档人,A.建档时间,a.执行分类 From 诊疗项目目录 A,人员表 B Where A.人员ID=B.ID(+) And A.ID=[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
        With rsTemp
            Me.txt编码.MaxLength = .Fields("编码").DefinedSize
            If .RecordCount > 0 Then
                Me.txt编码.Text = !编码
                Me.txt名称(0).Text = !名称
                Me.txt说明.Text = Nvl(!说明)
                Me.txt建档人.Text = Nvl(!建档人)
                Me.txt建档时间.Text = IIf(Nvl(!建档时间) = "", "", Format(!建档时间, "yyyy-mm-dd"))
                SetStationNo IIf(IsNull(!站点), "", !站点)
                mblnNoCheck = True
                If Nvl(!服务对象, 0) = 3 Then
                    chk范围(0).Value = 1
                    chk范围(1).Value = 1
                ElseIf Nvl(!服务对象, 0) = 1 Then
                    chk范围(0).Value = 1
                    chk范围(1).Value = 0
                ElseIf Nvl(!服务对象, 0) = 2 Then
                    chk范围(0).Value = 0
                    chk范围(1).Value = 1
                End If
                If Nvl(!执行分类, 0) = 1 Then
                    chkAll.Value = 1
                Else
                    chkAll.Value = 0
                End If
                mblnNoCheck = False
                
                If Nvl(!人员ID, 0) <> 0 Then
                    Me.cbo人员.AddItem Nvl(!编号) & "-" & Nvl(!姓名), 0
                    Me.cbo人员.ItemData(Me.cbo人员.NewIndex) = Nvl(!人员ID, 0)
                    Me.cbo人员.ListIndex = Me.cbo人员.NewIndex
                Else
                    Me.cbo人员.Text = ""
                    Me.cbo人员.ListIndex = -1
                End If
            End If
        End With
        
        '显示别名
        gstrSql = "Select 名称,性质,简码,码类 From 诊疗项目别名 Where 诊疗项目ID=[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
        With rsTemp
            Do While Not .EOF
                If !性质 = 1 And !码类 = 1 Then Me.txt拼音(0).Text = !简码
                If !性质 = 1 And !码类 = 2 Then Me.txt五笔(0).Text = !简码
                If !性质 = 9 Then Me.txt名称(1).Text = !名称
                If !性质 = 9 And !码类 = 1 Then Me.txt拼音(1).Text = !简码
                If !性质 = 9 And !码类 = 2 Then Me.txt五笔(1).Text = !简码
                .MoveNext
            Loop
        End With
        
        '提取方案内容
        Call LoadScheme(lngItemId)
    
        '确定项目使用范围
        If cbo人员.Text <> "" Then
            opt范围(0).Value = True
        Else
            gstrSql = "Select B.ID,B.名称 From 诊疗适用科室 A,部门表 B Where A.科室ID=B.ID And A.项目ID=[1] Order by B.编码"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
            If Not rsTemp.EOF Then opt范围(1).Value = True
            '原样加入主要供查阅用,也作为后面保持选择的基础
            Do While Not rsTemp.EOF
                Lvw科室.ListItems.Add(, "_" & rsTemp!ID, rsTemp!名称).Checked = True
                rsTemp.MoveNext
            Loop
        End If
    End If
    
    '根据权限设置控件可用性
    '-------------------------------------------------
    If Me.Tag <> "查阅" Then
        '根据权限限制使用范围
        If InStr(mstrPrivs, "全院成套方案") > 0 Then
            '有全院成套方案权限时，无限制
            Call LoadDeptList(True)
        ElseIf InStr(mstrPrivs, "本科成套方案") > 0 Then
            '只有本科成套方案权限时限制于本科内或自已的
            opt范围(2).Enabled = False
            If opt范围(2).Value Then opt范围(1).Value = True
            Call LoadDeptList(False)
        Else
            '都没有则只能看自已的
            opt范围(1).Enabled = False
            opt范围(2).Enabled = False
            If opt范围(1).Value Or opt范围(2).Value Then opt范围(0).Value = True
        End If
        If InStr(mstrPrivs, "全院成套方案") = 0 Then
            cbo人员.Locked = True
            If cbo人员.Text = "" Then '备用于新增或修改时选择为本人使用，当前不一定选择了本人使用
                Me.cbo人员.AddItem UserInfo.编号 & "-" & UserInfo.姓名, 0
                Me.cbo人员.ItemData(Me.cbo人员.NewIndex) = UserInfo.ID
                Me.cbo人员.ListIndex = Me.cbo人员.NewIndex
            End If
        End If
        
        mblnNoCheck = True
        If mint范围 = 1 Then
            '固定在门诊使用
            chk范围(0).Value = 1: chk范围(0).Visible = False
            chk范围(1).Value = 0: chk范围(1).Visible = False
        ElseIf mint范围 = 2 Then
            '固定在住院使用
            chk范围(0).Value = 0: chk范围(0).Visible = False
            chk范围(1).Value = 1: chk范围(1).Visible = False
        Else
            '保持现有值,可以选择在哪里使用
        End If
        mblnNoCheck = False
    Else
        On Error Resume Next
        cmdCancel.SetFocus
    End If
    
    If Me.Tag = "修改" Then
        If opt范围(0).Value = True And InStr(mstrPrivs, "修改个人成套方案") < 1 Then
            bln修改 = False
        ElseIf opt范围(1).Value = True And InStr(mstrPrivs, "修改科室成套方案") < 1 Then
            bln修改 = False
        ElseIf opt范围(2).Value = True And InStr(mstrPrivs, "修改全院成套方案") < 1 Then
            bln修改 = False
        End If
    End If
    
    
     
    '查阅时先设置界面不可编辑(OptionButton在Enabled时值会变化)
    '-------------------------------------------------
    If Me.Tag = "查阅" Or bln修改 = False Then
        Me.cmdOk.Visible = False
        Me.cmdCancel.Caption = "关闭(&C)"
        Me.txt分类.Enabled = False: Me.cmd分类.Enabled = False
        Me.txt编码.Enabled = False
        Me.txt名称(0).Enabled = False: Me.txt拼音(0).Enabled = False: Me.txt五笔(0).Enabled = False
        Me.txt名称(1).Enabled = False: Me.txt拼音(1).Enabled = False: Me.txt五笔(1).Enabled = False
        Me.txt说明.Enabled = False
        
        opt范围(0).Enabled = False: opt范围(1).Enabled = False: opt范围(2).Enabled = False
        chk范围(0).Enabled = False: chk范围(1).Enabled = False
        cbo人员.Enabled = False: cbo人员.BackColor = vbButtonFace
        Lvw科室.Enabled = False: Lvw科室.BackColor = vbButtonFace
        cmbStationNo.Enabled = False
    End If
    
    mblnChange = False
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadDeptList(ByVal BlnAll As Boolean)
'功能：根据权限读取可以使用的科室列表
'参数：blnAll=是否读取所有的科室，否则只读取自已的科室
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, strTmp As String
    Dim objItem As ListItem, i As Long
    
    On Error GoTo errH
    
    strTmp = IIf(chk范围(0).Value = 1, ",1", "") & IIf(chk范围(1).Value = 1, ",2", "") & ",3,"
    If BlnAll Then
        '可以指定的全院科室
        strSql = "Select Distinct A.ID,A.名称,A.编码 From 部门表 A,部门性质说明 B" & _
            " Where A.ID=B.部门ID And Instr([1],B.服务对象)>0" & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is Null)" & _
            " And B.工作性质 IN('临床','护理','检查','检验','手术','治疗','营养')" & _
            " Order by A.编码"
    Else
        '只能指定自已的科室
        strSql = "Select Distinct A.ID,A.名称,A.编码 From 部门表 A,部门性质说明 B,部门人员 C" & _
            " Where A.ID=B.部门ID And Instr([1],B.服务对象)>0 And A.ID=C.部门ID And C.人员ID=[2]" & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is Null)" & _
            " And B.工作性质 IN('临床','护理','检查','检验','手术','治疗','营养')" & _
            " Order by A.编码"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strTmp, UserInfo.ID)
    
    strTmp = ""
    For i = 1 To Lvw科室.ListItems.Count
        If Lvw科室.ListItems(i).Checked Then
            strTmp = strTmp & "," & Mid(Lvw科室.ListItems(i).Key, 2)
        End If
    Next
    If strTmp <> "" Then strTmp = strTmp & ","
    Lvw科室.ListItems.Clear
    
    i = 0
    Do While Not rsTmp.EOF
        Set objItem = Lvw科室.ListItems.Add(, "_" & rsTmp!ID, rsTmp!名称)
        If InStr(strTmp, "," & rsTmp!ID & ",") > 0 Then '保持原先的选择
            objItem.Checked = True
            objItem.ForeColor = vbBlue
            If i = 0 Then
                objItem.Selected = True
                objItem.EnsureVisible
            End If
            i = i + 1
        End If
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If Me.tvwClass.Visible Then
            Me.tvwClass.Visible = False: Me.txt分类.SetFocus
        Else
            Call cmdCancel_Click
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnFirst = True
    mblnOK = False
    mstrLike = IIf(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    
    Me.cbo人员.AddItem "[选择人员...]"
    cbo人员.Tag = cbo人员.ListIndex
    mlngFind = 1
    
    Call GetDefineSize
    Call IniStationNo
End Sub

Private Sub SetStationNo(ByVal strNo As String)
    Dim n As Integer
    
'    If gstrNodeNo = "-" Then Exit Sub
    
    If strNo = "" Then
        cmbStationNo.ListIndex = 0
    Else
        For n = 1 To cmbStationNo.ListCount - 1
            If Mid(cmbStationNo.List(n), 1, InStr(1, cmbStationNo.List(n), "-") - 1) = strNo Then
                cmbStationNo.ListIndex = n
            End If
        Next
    End If
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnChange Then
        If MsgBox("你已经对数据作了更改，确实要放弃更改退出吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsScheme = Nothing
End Sub

Private Sub IniStationNo()
    Dim dblHeight As Double
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    
'    lblStationNo.Visible = False
'    cmbStationNo.Visible = False
'
'    If gstrNodeNo <> "-" Then
        lblStationNo.Visible = True
        cmbStationNo.Visible = True
        
        strSql = "select 编号,名称 from zlnodelist"
        Set rsRecord = zlDatabase.OpenSQLRecord(strSql, "站点查询")
        With cmbStationNo
            .AddItem ""
            Do While Not rsRecord.EOF
                .AddItem rsRecord!编号 & "-" & rsRecord!名称
                rsRecord.MoveNext
            Loop
        End With
        
'        With cmbStationNo
'            .Clear
'            .AddItem ""
'            .AddItem "0"
'            .AddItem "1"
'            .AddItem "2"
'            .AddItem "3"
'            .AddItem "4"
'            .AddItem "5"
'            .AddItem "6"
'            .AddItem "7"
'            .AddItem "8"
'            .AddItem "9"
'
'            .ListIndex = 0
'        End With
'    Else
'        dblHeight = cmbStationNo.Height
'
'        fraLine(1).Top = fraLine(1).Top - dblHeight
'        fraLine(2).Top = fraLine(2).Top - dblHeight
'        Label1.Top = Label1.Top - dblHeight
'        opt范围(0).Top = opt范围(0).Top - dblHeight
'        opt范围(1).Top = opt范围(1).Top - dblHeight
'        opt范围(2).Top = opt范围(2).Top - dblHeight
'        chk范围(0).Top = chk范围(0).Top - dblHeight
'        chk范围(1).Top = chk范围(1).Top - dblHeight
'        lbl人员.Top = lbl人员.Top - dblHeight
'        cbo人员.Top = cbo人员.Top - dblHeight
'        lbl科室.Top = lbl科室.Top - dblHeight
'        lvw科室.Top = lvw科室.Top - dblHeight
'        cmdHelp.Top = cmdHelp.Top - dblHeight
'        cmdScheme.Top = cmdScheme.Top - dblHeight
'        cmdOK.Top = cmdOK.Top - dblHeight
'        cmdCancel.Top = cmdCancel.Top - dblHeight
'        Me.Height = Me.Height - dblHeight
'    End If
End Sub


Private Sub lvw科室_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        Item.ForeColor = vbBlue
    Else
        Item.ForeColor = Lvw科室.ForeColor
    End If
    mlngFind = Item.Index + 1
    mblnChange = True
End Sub

Private Sub lvw科室_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then Call cmdFind_Click
End Sub

Private Sub opt范围_Click(Index As Integer)
    If Me.Tag = "查阅" Then Exit Sub
    
    cbo人员.Enabled = Index = 0
    Lvw科室.Enabled = Index = 1
    txtFind.Enabled = Lvw科室.Enabled
    cmdFind.Enabled = Lvw科室.Enabled
    
    If cbo人员.Enabled Then
        cbo人员.BackColor = vbWindowBackground
    Else
        cbo人员.BackColor = vbButtonFace
    End If
    If Lvw科室.Enabled Then
        Lvw科室.BackColor = vbWindowBackground
        txtFind.BackColor = vbWindowBackground
    Else
        Lvw科室.BackColor = vbButtonFace
        txtFind.BackColor = vbButtonFace
    End If
    
    If cbo人员.Enabled Then
        If Trim(cbo人员.Text) = "" And cbo人员.ListCount = 1 Then
            If cbo人员.List(0) = "[选择人员...]" Then
                cbo人员.AddItem UserInfo.编号 & "-" & UserInfo.姓名, 0
                cbo人员.ItemData(cbo人员.NewIndex) = UserInfo.ID
                cbo人员.ListIndex = cbo人员.NewIndex
                cbo人员.Tag = cbo人员.ListIndex
            End If
        End If
    End If
    
    mblnChange = True
End Sub

Private Sub tvwClass_DblClick()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Me.txt分类.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    Me.txt分类.Text = Me.tvwClass.SelectedItem.Text
    Me.txt分类.SetFocus
End Sub

Private Sub tvwClass_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
        If Me.tvwClass.SelectedItem.Children > 0 Then Exit Sub
        Call tvwClass_DblClick
    Case vbKeySpace
        Call tvwClass_DblClick
    Case vbKeyEscape
        Call tvwClass_LostFocus
    End Select
End Sub

Private Sub tvwClass_LostFocus()
    If Me.cmd分类 Is ActiveControl Then Exit Sub
    Me.tvwClass.Visible = False
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtFind.Text <> "" Then Call cmdFind_Click
End Sub

Private Sub txt编码_Change()
    mblnChange = True
End Sub

Private Sub txt编码_GotFocus()
    Me.txt编码.SelStart = 0: Me.txt编码.SelLength = 100
End Sub

Private Sub txt编码_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt分类_Change()
    mblnChange = True
End Sub

Private Sub txt分类_GotFocus()
    Me.txt分类.SelStart = 0: Me.txt分类.SelLength = 100
End Sub

Private Sub txt分类_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt名称_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txt名称_GotFocus(Index As Integer)
    Me.txt名称(Index).SelStart = 0: Me.txt名称(Index).SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt名称_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txt名称(Index).Text = MoveSpecialChar(txt名称(Index).Text)
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
'    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
             
End Sub

Private Sub txt名称_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Me.txt拼音(Index).Text = zlStr.GetCodeByORCL(Me.txt名称(Index).Text, False, txt拼音(Index).MaxLength)
    Me.txt五笔(Index).Text = zlStr.GetCodeByORCL(Me.txt名称(Index).Text, True, txt五笔(Index).MaxLength)
End Sub

Private Sub txt名称_LostFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt拼音_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txt拼音_GotFocus(Index As Integer)
    Me.txt拼音(Index).SelStart = 0: Me.txt拼音(Index).SelLength = 100
End Sub

Private Sub txt拼音_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub cbo人员_GotFocus()
    Call zlControl.TxtSelAll(cbo人员)
End Sub

Private Sub txt说明_Change()
    mblnChange = True
End Sub

Private Sub txt说明_GotFocus()
    Me.txt说明.SelStart = 0: Me.txt说明.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt说明_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt说明_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt五笔_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txt五笔_GotFocus(Index As Integer)
    Me.txt五笔(Index).SelStart = 0: Me.txt五笔(Index).SelLength = 100
End Sub

Private Sub txt五笔_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Function LoadScheme(ByVal lng成套ID As Long) As Boolean
'功能：读取并显示数据库中的成套方案内容
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select 序号,相关序号,期效,诊疗项目ID,收费细目ID,医嘱内容,天数,单次用量,总给予量," & _
        " 医生嘱托,执行频次,频率次数,频率间隔,间隔单位,时间方案,执行科室ID,标本部位,检查方法,执行性质,执行标记,配方ID,组合项目ID" & _
        " From 诊疗项目组合 Where 诊疗组合ID=[1] Order by 序号"
    Set mrsScheme = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng成套ID)
    LoadScheme = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
