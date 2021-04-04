VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmNurseFileEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "文件编辑"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4155
   Icon            =   "frmNurseFileEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCanCel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2640
      TabIndex        =   8
      Top             =   1920
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1410
      TabIndex        =   7
      Top             =   1920
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -120
      TabIndex        =   6
      Top             =   1710
      Width           =   4545
   End
   Begin MSMask.MaskEdBox mskEdit 
      Height          =   285
      Left            =   1350
      TabIndex        =   5
      Top             =   1110
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   19
      Format          =   "yyyy-MM-dd HH:mm:ss"
      Mask            =   "####-##-## ##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txt文件名称 
      Height          =   285
      Left            =   1350
      MaxLength       =   50
      TabIndex        =   3
      Top             =   720
      Width           =   2415
   End
   Begin VB.ComboBox cbo格式来源 
      Height          =   300
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   330
      Width           =   2415
   End
   Begin VB.Label lbl开始时间 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "开始时间"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   570
      TabIndex        =   4
      Top             =   1155
      Width           =   720
   End
   Begin VB.Label lbl文件名称 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "文件名称"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   570
      TabIndex        =   2
      Top             =   765
      Width           =   720
   End
   Begin VB.Label lbl格式来源 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "格式来源"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   570
      TabIndex        =   0
      Top             =   390
      Width           =   720
   End
End
Attribute VB_Name = "frmNurseFileEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstr入院时间 As String
Private mlngFile As Long        '文件ID,传入0表示新增,否则表示修改(修改时不允许修改文件来源)
Private mlng病人id As Long
Private mlng主页id As Long
Private mint婴儿 As Long
Private mlng科室ID As Long
Private mstrDept As String      '当前科室
Private mblnOK As Boolean       '是否保存成功
Private mblnExist体温单 As Boolean
Private mblnExist记录单 As Boolean
Private mblnOnly As Boolean     '住院病人同一时点只记录一份护理文件

Public Function ShowEditor(ByVal lng病人id As Long, ByVal lng主页id As Long, ByVal int婴儿 As Integer, ByVal lng科室ID As Long, _
    ByVal str科室 As String, Optional ByVal lngFile As Long) As Boolean
    mblnOK = False
    mlng病人id = lng病人id
    mlng主页id = lng主页id
    mint婴儿 = int婴儿
    mlng科室ID = lng科室ID
    mlngFile = lngFile
    mstrDept = str科室
    Me.Show 1
    ShowEditor = mblnOK
End Function

Private Sub cbo格式来源_Click()
    Dim bln体温单 As Boolean
    txt文件名称.Text = Split(cbo格式来源.Text, "-")(1)
    If cbo格式来源.ListIndex <> Val(cbo格式来源.Tag) Or cbo格式来源.Tag = "" Then '记录单时如下处理
        txt文件名称.Text = "[" & mstrDept & "]" & txt文件名称.Text
    End If
    
    '新增:如果不存护理文件则缺省时间为入院时间,否则为当前时间
    '修改:护理文件的开始时间不能小于入院时间,不能大于数据发生时间,如无数据则不能大于当前时间
    bln体温单 = (cbo格式来源.Tag <> "" And cbo格式来源.ListIndex = Val(cbo格式来源.Tag))
    If mlngFile = 0 Then
        If (Not mblnExist记录单 And Not bln体温单) Or (Not mblnExist体温单 And bln体温单) Then
            mskEdit.Text = mstr入院时间
        Else
            mskEdit.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngID As Long
    On Error GoTo errHand
    If txt文件名称.Text = "" Then
        MsgBox "请输入文件名称！", vbInformation, gstrSysName
        txt文件名称.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(txt文件名称.Text, vbFromUnicode)) > 50 Then
        MsgBox "文件名称超长！（最多50个字符或25个汉字）", vbInformation, gstrSysName
        txt文件名称.SetFocus
        Exit Sub
    End If
    
    If mlngFile = 0 Then
        lngID = zlDatabase.GetNextId("病人护理文件")
    Else
        lngID = mlngFile
    End If
    
    gstrSQL = "ZL_病人护理文件_UPDATE(" & lngID & "," & mlng科室ID & "," & mlng病人id & "," & mlng主页id & "," & mint婴儿 & "," & _
              cbo格式来源.ItemData(cbo格式来源.ListIndex) & ",'" & txt文件名称.Text & "',to_date('" & mskEdit.Text & "','yyyy-MM-dd hh24:mi:ss')," & IIf(mlngFile = 0 And mblnOnly, 1, 0) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新数据")
    
    mblnOK = True
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim blnSeek As Boolean
    Dim lng格式 As Long, str文件名称 As String, str开始时间 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '提取当前病人入院时间
    gstrSQL = " Select 入院日期 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取当前病人入院时间", mlng病人id, mlng主页id)
    mstr入院时间 = Format(rsTemp!入院日期, "yyyy-MM-dd HH:mm:ss")
    
    '检查是否已设定体温单,如已存在则不允许再次添加体温单
    gstrSQL = " Select B.保留" & _
              " From 病人护理文件 A,病历文件列表 B" & _
              " Where A.格式ID=B.ID And A.病人ID=[1] And A.主页ID=[2] And A.婴儿=[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否已定义体温单", mlng病人id, mlng主页id, mint婴儿)
    rsTemp.Filter = "保留=-1"
    mblnExist体温单 = rsTemp.RecordCount
    rsTemp.Filter = "保留<>-1"
    mblnExist记录单 = rsTemp.RecordCount
    rsTemp.Filter = 0
    
    '读取文件设置
    gstrSQL = " Select A.科室ID,B.名称 AS 科室,A.格式ID,A.文件名称,A.开始时间 " & _
              " From 病人护理文件 A,部门表 B" & _
              " Where A.科室ID=B.ID And A.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取文件设置", mlngFile)
    If rsTemp.RecordCount <> 0 Then
        mlng科室ID = rsTemp!科室ID
        mstrDept = rsTemp!科室
        lng格式 = rsTemp!格式ID
        str文件名称 = rsTemp!文件名称
        str开始时间 = Format(rsTemp!开始时间, "yyyy-MM-dd HH:mm:ss")
    End If
    
    '读取护理相关的病历文件并加载
    gstrSQL = "Select ID,保留,编号||'-'||名称 AS 格式 From 病历文件列表 Where 种类=3 And 通用 > 0  Order by 保留,编号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取护理相关的病历文件")
    With rsTemp
        Me.cbo格式来源.Clear
        Do While Not .EOF
            If (((!保留 = -1 And Not mblnExist体温单) Or !保留 <> -1) And mlngFile = 0) Or mlngFile <> 0 Then
                Me.cbo格式来源.AddItem !格式
                Me.cbo格式来源.ItemData(Me.cbo格式来源.NewIndex) = !ID
                If !保留 = -1 Then Me.cbo格式来源.Tag = .AbsolutePosition - 1
                If !ID = lng格式 Then
                    Me.cbo格式来源.ListIndex = .AbsolutePosition - 1
                    blnSeek = True
                End If
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 And Not blnSeek Then Me.cbo格式来源.ListIndex = 0
    End With
    
    If mlngFile <> 0 Then
        Me.txt文件名称.Text = str文件名称
        If str开始时间 <> "" Then Me.mskEdit.Text = str开始时间
    Else
        mblnOnly = (Val(zlDatabase.GetPara("对应多份护理文件", glngSys, 1255, 0)) = 0)
    End If
    Me.cbo格式来源.Enabled = (mlngFile = 0)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mskEdit_GotFocus()
    Call zlControl.TxtSelAll(mskEdit)
End Sub

Private Sub txt文件名称_GotFocus()
    Call zlControl.TxtSelAll(txt文件名称)
End Sub
