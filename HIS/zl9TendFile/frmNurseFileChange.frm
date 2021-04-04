VERSION 5.00
Begin VB.Form frmNurseFileChange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "文件格式变更"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4620
   Icon            =   "frmNurseFileChange.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCanCel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3000
      TabIndex        =   10
      Top             =   1935
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   45
      TabIndex        =   8
      Top             =   1740
      Width           =   4545
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1725
      TabIndex        =   9
      Top             =   1935
      Width           =   1155
   End
   Begin VB.TextBox txtOldName 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1275
      MaxLength       =   50
      TabIndex        =   3
      Top             =   570
      Width           =   2895
   End
   Begin VB.TextBox txtOldFormat 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1275
      MaxLength       =   50
      TabIndex        =   1
      Top             =   225
      Width           =   2895
   End
   Begin VB.ComboBox cboFormat 
      Height          =   300
      Left            =   1275
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   930
      Width           =   2895
   End
   Begin VB.TextBox txtNewName 
      Height          =   285
      Left            =   1275
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1305
      Width           =   2895
   End
   Begin VB.Label lblOldName 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "旧文件名称"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   315
      TabIndex        =   2
      Top             =   615
      Width           =   900
   End
   Begin VB.Label lblOldFormat 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "旧格式"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   675
      TabIndex        =   0
      Top             =   270
      Width           =   540
   End
   Begin VB.Label lblNewForamt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "新格式"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   675
      TabIndex        =   4
      Top             =   990
      Width           =   540
   End
   Begin VB.Label lblNewName 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "新文件名称"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   315
      TabIndex        =   6
      Top             =   1350
      Width           =   900
   End
End
Attribute VB_Name = "frmNurseFileChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOK As Boolean
Private mlngFileID As Long
Private mlngFormatID As Long
Private mblnWave As Boolean 'TRUE：体温单 FALSE：记录单
Private mstrDept As String '科室名称

Public Function ShowEditor(ByVal mfrmParent As Form, ByVal lngFileID As Long) As Boolean
    mblnOK = False: mblnWave = False
    mlngFileID = lngFileID
    Me.Show 1, mfrmParent
    ShowEditor = mblnOK
End Function

Private Sub cboFormat_Click()
    txtNewName.Text = Split(cboFormat.Text, "-")(1)
    If mstrDept <> "" Then txtNewName.Text = "[" & mstrDept & "]" & txtNewName.Text
End Sub

Private Sub cmdCanCel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
'功能：完成文件格式变更
    Dim blnShow As Boolean
    On Error GoTo ErrHand
    
    If mlngFormatID = cboFormat.ItemData(cboFormat.ListIndex) Then
        MsgBox "替换文件的格式不能和之前的格式相同，请重新选择！", vbInformation, gstrSysName
        cboFormat.SetFocus
        Exit Sub
    End If
    If txtNewName.Text = "" Then
        MsgBox "请输入文件名称！", vbInformation, gstrSysName
        txtNewName.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(txtNewName.Text, vbFromUnicode)) > 50 Then
        MsgBox "文件名称超长！（最多50个字符或25个汉字）", vbInformation, gstrSysName
        txtNewName.SetFocus
        Exit Sub
    End If
    If MsgBox("该操作可能需要一段时间，请问您是否继续！", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
    Screen.MousePointer = 11
    zlCommFun.ShowFlash "正在更新文件格式，请您耐心等待....", Me
    blnShow = True
    gstrSQL = "Zl_病人护理文件_Repalce(" & mlngFileID & "," & Val(cboFormat.ItemData(cboFormat.ListIndex)) & ",'" & Trim(txtNewName.Text) & "'," & IIf(mblnWave = True, 0, 1) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "Zl_病人护理文件_Repalce")
    zlCommFun.StopFlash
    Screen.MousePointer = 0
    
    mblnOK = True
    Unload Me
    Exit Sub
ErrHand:
    Screen.MousePointer = 0
    If blnShow = True Then zlCommFun.StopFlash
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 39 Then KeyCode = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    '提取当前文件信息
    gstrSQL = " Select B.名称 格式名称,B.编号,B.保留,B.ID 格式ID,A.文件名称,A.病人ID,A.主页ID,A.婴儿,C.名称 科室名称" & _
          " From 病人护理文件 A,病历文件列表 B,部门表 C" & _
          " Where A.格式ID=B.ID And A.ID=[1] And B.种类=3 And A.科室ID=C.ID(+)"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取当前文件信息", mlngFileID)
    txtOldFormat.Text = rsTemp!编号 & "-" & rsTemp!格式名称
    txtOldName.Text = rsTemp!文件名称
    mlngFormatID = rsTemp!格式ID
    mblnWave = (Val(NVL(rsTemp!保留, 0)) = -1)
    mstrDept = NVL(rsTemp!科室名称)
    '提取适应于当前病区的记录单或体温单
    gstrSQL = _
        " Select ID, 保留, 编号, 编号 || '-' || 名称 As 格式" & vbNewLine & _
        " From 病历文件列表" & vbNewLine & _
        " Where 种类 = 3  " & IIf(mblnWave = True, " And 保留 =-1", " And 保留 <> 1 And  保留 <> -1") & " And (通用 = 1 Or (通用 = 2 And ID In (Select 文件id From 病历应用科室 Where 科室id = [1])))" & vbNewLine & _
        " Order By 编号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取文件列表", glng病区ID)
    With rsTemp
        cboFormat.Clear
        Do While Not .EOF
            If !ID <> mlngFormatID Then
                cboFormat.AddItem !格式
                cboFormat.ItemData(cboFormat.NewIndex) = !ID
            End If
        .MoveNext
        Loop
    End With
    If cboFormat.ListCount > 0 Then
        cboFormat.ListIndex = 0
    Else '没有可以选择的记录单或体温单
        On Error Resume Next
        MsgBox "在[病历文件列表]中没有找到适用于本病区的其他格式" & IIf(mblnWave, "", "") & "文件！", vbInformation, gstrSysName
        Unload Me
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

