VERSION 5.00
Begin VB.Form frmSetStock 
   Caption         =   "系统设置"
   ClientHeight    =   2160
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5505
   Icon            =   "frmSetStock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5505
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox cboCardType 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   780
      Width           =   3255
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "下一步"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   5685
   End
   Begin VB.ComboBox cboStock 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Width           =   3255
   End
   Begin VB.Label lblCardType 
      AutoSize        =   -1  'True
      Caption         =   "刷卡类别"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1140
   End
   Begin VB.Label lblStock 
      AutoSize        =   -1  'True
      Caption         =   "当前药房"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1140
   End
End
Attribute VB_Name = "frmSetStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub GetCardType()
    '取各种医疗卡列表
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSql = "Select ID, 编码, 名称 From 医疗卡类别 Order By 编码"
    Set rsData = zldatabase.OpenSQLRecord(gstrSql, "取医疗卡列表")
    
    If rsData.RecordCount = 0 Then
        MsgBox "未找到医疗卡信息，请先设置医疗卡！", vbInformation, "自助签到系统"
        Unload Me
        Exit Sub
    End If
    
    With cboCardType
        .Clear
        Do While Not rsData.EOF
            .AddItem rsData!编码 & "-" & rsData!名称
            .ItemData(.NewIndex) = rsData!Id
            rsData.MoveNext
        Loop
        .ListIndex = 0
        glngCardTypeID = Val(.ItemData(.ListIndex))
    End With
    Exit Sub
errHandle:
    If errcenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub GetStock()
    '取药房列表
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSql = " Select Distinct a.Id, a.编码, a.名称 " & _
        " From 部门表 A, 部门性质说明 B " & _
        " Where a.Id = b.部门id And a.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd') And b.工作性质 Like '%药房' " & _
        " Order By a.编码 "
    Set rsData = zldatabase.OpenSQLRecord(gstrSql, "取药房列表")
    
    If rsData.RecordCount = 0 Then
        MsgBox "未找到药房信息，请在部门管理中设置！", vbInformation, "自助签到系统"
        Unload Me
        Exit Sub
    End If
    
    With cboStock
        .Clear
        Do While Not rsData.EOF
            .AddItem rsData!编码 & "-" & rsData!名称
            .ItemData(.NewIndex) = rsData!Id
            rsData.MoveNext
        Loop
        .ListIndex = 0
        glngStock = Val(.ItemData(.ListIndex))
    End With
    Exit Sub
errHandle:
    If errcenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdNext_Click()
    glngStock = Val(cboStock.ItemData(cboStock.ListIndex))
    gstrStockName = Mid(cboStock.Text, InStr(1, cboStock.Text, "-") + 1)
    
    glngCardTypeID = Val(cboCardType.ItemData(cboCardType.ListIndex))
    
    Call frmCheckIn.ShowMe(glngStock, Mid(cboCardType.Text, InStr(1, cboCardType.Text, "-") + 1))
    
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Load()
    Call GetStock
    Call GetCardType
End Sub
