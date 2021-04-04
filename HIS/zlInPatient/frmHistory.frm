VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmHistory 
   AutoRedraw      =   -1  'True
   Caption         =   "病人变动记录"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10605
   Icon            =   "frmHistory.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   10605
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   10605
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3300
      Width           =   10605
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "退出(&X)"
         Height          =   350
         Left            =   8280
         TabIndex        =   2
         Top             =   105
         Width           =   1575
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh 
      Height          =   3240
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   5715
      _Version        =   393216
      BackColor       =   16777215
      FixedCols       =   0
      RowHeightMin    =   250
      BackColorBkg    =   16777215
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mlng病人ID As Long
Public mlng主页ID As Long

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    Call RestoreWinState(Me, App.ProductName)
    
    msh.RowHeight(0) = 320
    
    On Error GoTo errH
    
    strSQL = _
        " Select B.名称 as 病区,C.名称 as 科室,A.床号,D.名称 as 床位等级," & _
        " E.名称 as 护理等级,A.责任护士 as 护士,A.经治医师 as 住院医师,A.主治医师 as 主治医生,A.主任医师 as 主任医生,A.病情 as 当前病况,A.操作员姓名 as 开始操作员," & _
        " Decode(A.开始原因,1,'入院',2,'入住',3,'转科',4,'换床',5,'调整床位等级',6,'调整护理等级',7,'调整住院医师',8,'调整护士',9,'留观转住院',10,'预出院',11,'调整主治医师',12,'调整主任医师',13,'调整病况',14 ,'调整医疗小组', 15, '调整病区') as 开始原因," & _
        " To_Char(A.开始时间,'YYYY-MM-DD HH24:MI:SS') as 开始时间,A.终止人员 as 终止操作员," & _
        " Decode(A.终止原因,1,'出院',2,'入住',3,'转科',4,'换床',5,'调整床位等级',6,'调整护理等级',7,'调整住院医师',8,'调整护士',9,'留观转住院',10,'预出院',11,'调整主治医师',12,'调整主任医师',13,'调整病况',14 ,'调整医疗小组', 15, '调整病区') as 终止原因," & _
        " To_Char(A.终止时间,'YYYY-MM-DD HH24:MI:SS') as 终止时间" & _
        " " & _
        " From 病人变动记录 A,部门表 B,部门表 C,收费项目目录 D,收费项目目录 E" & _
        " Where A.病区ID=B.ID And A.科室ID=C.ID" & _
        " And A.床位等级ID=D.ID(+) And A.护理等级ID=E.ID(+)" & _
        " And A.病人ID=[1] And A.主页ID=[2]" & _
        " And A.开始时间 is Not NULL" & _
        " Order by A.终止时间,A.开始时间,A.床号"

    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    
    If Not rsTmp.EOF Then
        Set msh.DataSource = rsTmp
        
        For i = 0 To msh.Cols - 1
            If InStr(1, ",床号,护士,住院医师,主治医师,主任医师,开始操作员,开始时间,终止操作员,终止时间,", "," & msh.TextMatrix(0, i) & ",") = 0 Then
                msh.colAlignment(i) = 1
            Else
                msh.colAlignment(i) = 4
            End If
        Next
    End If
    Call SetGridWidth(msh, Me)
    
    RestoreFlexState msh, App.ProductName & "\" & Me.Name
    For i = 0 To msh.Cols - 1
        msh.ColAlignmentFixed(i) = 4
    Next
    msh.Row = 1: msh.Col = 0: msh.ColSel = msh.Cols - 1
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    msh.Left = 0
    msh.Top = 0
    msh.width = Me.ScaleWidth
    msh.Height = Me.ScaleHeight - picCmd.Height
    cmdExit.Left = Me.ScaleWidth - cmdExit.width * 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub
