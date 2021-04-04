VERSION 5.00
Begin VB.Form frmAltAdviceSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "备选医嘱"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11730
   Icon            =   "frmAltAdviceSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   11730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   -120
      TabIndex        =   3
      Top             =   5760
      Width           =   11895
   End
   Begin VB.CommandButton cmdCancle 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   10440
      TabIndex        =   2
      Top             =   5925
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   9120
      TabIndex        =   1
      Top             =   5925
      Width           =   1100
   End
   Begin zlCISPath.UCAdviceList UCAdvice 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   -50
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10186
   End
End
Attribute VB_Name = "frmAltAdviceSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng路径项目ID As Long
Private mstrSelectedIds As String
Private mblnOK As Boolean
Private mrsAdvice As Recordset
Private mint婴儿 As Integer
Private mintFunc As Integer '0-缺省为住院临床路径;1-门诊临床路径

Public Function ShowSelect(ByVal frmParent As Object, ByVal lng路径项目ID As Long, Optional ByVal strSelectedIDs As String, _
    Optional ByVal int婴儿 As Integer, Optional ByVal intFunc As Integer) As String
'功能：调用选择界面，返回选择后的医嘱IDs
    mlng路径项目ID = lng路径项目ID
    mstrSelectedIds = strSelectedIDs
    mint婴儿 = int婴儿
    mintFunc = intFunc
    Me.Show 1, frmParent
    ShowSelect = IIf(mblnOK, mstrSelectedIds, strSelectedIDs)
End Function

Private Sub ShowAltAdvice()
'功能：显示备选医嘱
    Dim strSQL As String, rstmp As Recordset
    Dim i As Long
    
    On Error GoTo errH
    If mintFunc = 0 Then
        strSql = _
            "Select Distinct (a.Id * 10 +" & mint婴儿 & ") as id, Decode(a.相关ID,NULL,NULL,(a.相关id * 10 + " & mint婴儿 & ")) AS 相关id, a.序号, a.期效, a.诊疗项目id, a.收费细目id, a.医嘱内容, a.单次用量, a.总给予量, a.标本部位, a.检查方法, a.医生嘱托, a.执行频次," & vbNewLine & _
                    "                a.频率次数, a.频率间隔, a.间隔单位, a.执行性质,a.执行标记,a.组合项目ID, a.执行科室id, a.时间方案, a.是否缺省, Decode(instr(',' || [4] || ',',',' || (a.ID *10 + " & mint婴儿 & ") || ','), 0, 0, 1) As 是否备选" & vbNewLine & _
                    "From 路径医嘱内容 A, 临床路径医嘱 B, 临床路径项目 C" & vbNewLine & _
                    "Where a.Id = b.医嘱内容id And b.路径项目id = c.Id And c.id=[3]" & vbNewLine & _
                    "Order By a.序号, a.Id"
    Else
        strSql = _
            "Select Distinct (a.Id * 10 +" & mint婴儿 & ") as id, Decode(a.相关ID,NULL,NULL,(a.相关id * 10 + " & mint婴儿 & ")) AS 相关id, a.序号, a.期效, a.诊疗项目id, a.收费细目id, a.医嘱内容, a.单次用量, a.总给予量, a.标本部位, a.检查方法, a.医生嘱托, a.执行频次," & vbNewLine & _
                    "                a.频率次数, a.频率间隔, a.间隔单位, a.执行性质,a.执行标记,a.组合项目ID, a.执行科室id, a.时间方案, a.是否缺省, Decode(instr(',' || [4] || ',',',' || (a.ID *10 + " & mint婴儿 & ") || ','), 0, 0, 1) As 是否备选" & vbNewLine & _
                    "From 门诊路径医嘱内容 A, 门诊路径医嘱 B, 门诊路径项目 C" & vbNewLine & _
                    "Where a.Id = b.医嘱内容id And b.路径项目id = c.Id And c.id=[3]" & vbNewLine & _
                    "Order By a.序号, a.Id"
    End If
    Call UCAdvice.ShowAdvice(3, strSql, 0, 0, , mlng路径项目ID, mstrSelectedIds)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mstrSelectedIds = UCAdvice.GetAdviceIDSelected(1)
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    Call ShowAltAdvice
End Sub
