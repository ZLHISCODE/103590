VERSION 5.00
Begin VB.Form frmTendBodyDateSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "体温单开始日期设定"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4665
   Icon            =   "frmTendBodyDateSet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1920
      TabIndex        =   4
      Top             =   1710
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3180
      TabIndex        =   5
      Top             =   1710
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   -60
      TabIndex        =   3
      Top             =   1380
      Width           =   5175
   End
   Begin VB.ComboBox cboDate 
      Height          =   300
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   900
      Width           =   3105
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTendBodyDateSet.frx":000C
      ForeColor       =   &H00C00000&
      Height          =   585
      Left            =   300
      TabIndex        =   2
      Top             =   150
      Width           =   4035
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   315
      TabIndex        =   0
      Top             =   960
      Width           =   720
   End
End
Attribute VB_Name = "frmTendBodyDateSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mblnReturn As Boolean

Public Function ShowMe(ByVal lng病人id As Long, ByVal lng主页id As Long) As Boolean
    If lng病人id = 0 Then Exit Function
    
    mblnReturn = False
    mlng病人ID = lng病人id
    mlng主页ID = lng主页id
    Me.Show 1
    ShowMe = mblnReturn
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '首先检查在设置的时间前是否已经发生数据了
    gstrSQL = " Select 1 From 病人护理记录 A,病人护理内容 B " & _
        "   Where B.记录ID=A.ID And A.病人ID=[1] And A.主页ID=[2] And A.发生时间<[3] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng病人ID, mlng主页ID, CDate(Mid(Me.cboDate.Text, 2, 19)))
    If rsTemp.RecordCount > 0 Then
        MsgBox "该段时间前已经发生了相应的数据，不允许修改体温单开始时间！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gcnOracle.Execute "ZL_体温单开始日期_UPDATE(" & mlng病人ID & "," & mlng主页ID & ",'" & Mid(Me.cboDate.Text, 2, 19) & "')", , adCmdStoredProc
    
    mblnReturn = True
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Dim str开始日期 As String
    Dim intStart As Integer, intEnd As Integer
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = " SELECT '于'||TO_CHAR(开始时间,'YYYY-MM-DD HH24:MI:SS')||DECODE(开始原因,1,'入院',2,'入科-'||B.名称,'转入-'||B.名称) AS 内容" & _
              " FROM 病人变动记录 A,部门表 B" & _
              " WHERE A.科室ID=B.ID AND A.开始原因 IN (1,2,3) AND A.病人ID=[1] AND A.主页ID=[2]" & _
              " ORDER BY A.病人ID,A.主页ID,A.开始原因,A.开始时间"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    With rsTemp
        Me.cboDate.Clear
        Do While Not .EOF
            Me.cboDate.AddItem !内容
            .MoveNext
        Loop
        Me.cboDate.ListIndex = 0
    End With
    
    '提取病案主页从表中的体温单开始日期
    gstrSQL = " Select 信息值 From 病案主页从表 Where 病人ID=[1] And 主页ID=[2] And 信息名='体温单开始日期'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If rsTemp.RecordCount <> 0 Then
        str开始日期 = Format(rsTemp!信息值, "yyyy-MM-dd HH:mm:ss")
    End If
    
    '定位
    If str开始日期 <> "" Then
        intEnd = Me.cboDate.ListCount
        For intStart = 1 To intEnd
            If InStr(1, Me.cboDate.List(intStart - 1), str开始日期) <> 0 Then
                Me.cboDate.ListIndex = intStart - 1
                Exit For
            End If
        Next
    End If
    
    cmdOK.Enabled = True
End Sub


