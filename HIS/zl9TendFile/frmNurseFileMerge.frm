VERSION 5.00
Begin VB.Form frmNurseFileMerge 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "合并打印"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4155
   Icon            =   "frmNurseFileMerge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1410
      TabIndex        =   4
      Top             =   1440
      Width           =   1155
   End
   Begin VB.CommandButton cmdCanCel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2640
      TabIndex        =   5
      Top             =   1440
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   -120
      TabIndex        =   3
      Top             =   1110
      Width           =   4365
   End
   Begin VB.ComboBox cbo续打文件 
      Height          =   300
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   630
      Width           =   2625
   End
   Begin VB.Label lbl续打文件 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "续打文件"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   1
      Top             =   690
      Width           =   720
   End
   Begin VB.Label lbl当前文件 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "当前文件:"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   330
      Width           =   810
   End
End
Attribute VB_Name = "frmNurseFileMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngFile As Long
Private mblnOK As Boolean

Public Function ShowEditor(ByVal lngFile As Long) As Boolean
    On Error Resume Next
    mlngFile = lngFile
    mblnOK = False
    Me.Show 1
    ShowEditor = mblnOK
End Function

Private Sub cmdCanCel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim blnTrans As Boolean
    On Error GoTo ErrHand
    gcnOracle.BeginTrans
    blnTrans = True
    gstrSQL = "ZL_病人护理文件_STATE(" & mlngFile & ",2,NULL," & Me.cbo续打文件.ItemData(Me.cbo续打文件.ListIndex) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "合并打印")
    gstrSQL = "Zl_病人护理打印_Batchretrypage(" & Me.cbo续打文件.ItemData(Me.cbo续打文件.ListIndex) & ",'1;0')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "页码重整")
    gcnOracle.CommitTrans
    blnTrans = False
    
    mblnOK = True
    Unload Me
    Exit Sub
ErrHand:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Dim str开始时间 As String, str创建时间 As String
    Dim lngPati As Long, lngPage As Long, lngBaby As Long, lngFormat As Long
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    '提取该病人与指定文件格式相同的文件,设定合并打印(只能按时间的先后顺序进行设定)
    
    gstrSQL = " Select 病人ID,主页ID,nvl(婴儿,0) 婴儿,格式ID,文件名称,开始时间,创建时间 From 病人护理文件 Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取文件属性", mlngFile)
    lngPati = NVL(rsTemp!病人ID, 0)
    lngPage = NVL(rsTemp!主页ID, 0)
    lngBaby = NVL(rsTemp!婴儿, 0)
    lngFormat = NVL(rsTemp!格式ID, 0)
    str开始时间 = Format(rsTemp!开始时间, "yyyy-MM-dd HH:mm:ss")
    str创建时间 = Format(rsTemp!创建时间, "yyyy-MM-dd HH:mm:ss")
    Me.lbl当前文件.Caption = "当前文件：" & rsTemp!文件名称
    
    gstrSQL = _
            "  Select ID,文件名称" & vbNewLine & _
            "  From (With 病人护理文件_F1 As" & vbNewLine & _
            "   (Select Id, 续打id, 文件名称, 开始时间,创建时间" & vbNewLine & _
            "   From 病人护理文件" & vbNewLine & _
            "   Where 病人id = [2] And 主页id = [3] And Nvl(婴儿, 0) = [4] And 格式id = [5])" & vbNewLine & _
            "   Select Id, 文件名称" & vbNewLine & _
            "   From 病人护理文件_F1 a" & vbNewLine & _
            "   Where a.Id <> [1] And (a.开始时间>[6] OR (a.开始时间=[6] And a.创建时间>[7]))" & vbNewLine & _
            "   And Not Exists (Select Id From 病人护理文件_F1 Where a.Id = 续打id)" & vbNewLine & _
            "   Order By a.开始时间)"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取该病人与指定文件格式相同的文件,设定合并打印", mlngFile, lngPati, lngPage, lngBaby, lngFormat, CDate(str开始时间), CDate(str创建时间))
    With rsTemp
        Me.cbo续打文件.Clear
        Do While Not .EOF
            cbo续打文件.AddItem !文件名称
            cbo续打文件.ItemData(cbo续打文件.NewIndex) = !ID
            .MoveNext
        Loop
    End With
    If cbo续打文件.ListCount = 0 Then
        MsgBox "当前文件之后不存在同格式的文件，因此不需要为当前文件指定合并打印！", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    cbo续打文件.ListIndex = 0
    
End Sub

