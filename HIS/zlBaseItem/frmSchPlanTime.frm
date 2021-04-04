VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSchPlanTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "检查预约--时间计划设置"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4740
   Icon            =   "frmSchPlanTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4740
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2760
      TabIndex        =   5
      Top             =   2760
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   840
      TabIndex        =   4
      Top             =   2760
      Width           =   1100
   End
   Begin VB.TextBox txtCapacity 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "30"
      Top             =   2100
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dpTimeStart 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "HH:mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   885
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   46006274
      CurrentDate     =   .333333333333333
   End
   Begin VB.ComboBox cboSchExamTimeCalcType 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   322
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dpTimeEnd 
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   1485
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   46006274
      CurrentDate     =   .5
   End
   Begin VB.Label Label4 
      Caption         =   "预约容量"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "结束时间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "开始时间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "检查时长计算方法"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmSchPlanTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mlngTimeProjectID As Long
Dim mlngPlanID As Long

Public Sub zlShowMe(frmParent As Form, lngTimeProjectID As Long, lngPlanID As Long)
'------------------------------------------------
'功能：装载时间表的表格格式和基础内容
'参数： frmParent -- 父窗体
'       lngTimeProjectID -- 时间计划ID，如果是新增，则=0
'       lngPlanID -- 预约方案ID
'返回：无
'------------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    mlngTimeProjectID = lngTimeProjectID
    mlngPlanID = lngPlanID
    
    If mlngTimeProjectID <> 0 Then
        '从数据库读取当前的时间计划
        strSql = "select 开始时间,结束时间,预约容量,计算方法,预约方案ID from 影像预约时间计划 where id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "影像预约时间设置", mlngTimeProjectID)
        If rsTemp.EOF = False Then
            cboSchExamTimeCalcType.ListIndex = IIF(NVL(rsTemp!计算方法, 1) = 1, 0, 1)
            dpTimeStart = Format(rsTemp!开始时间, "hh:mm:ss")
            dpTimeEnd = Format(rsTemp!结束时间, "hh:mm:ss")
            txtCapacity = NVL(rsTemp!预约容量)
        End If
    End If
    
    Me.Show 1, frmParent
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '保存时间计划设置
    If saveTimeProject = True Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    cboSchExamTimeCalcType.AddItem "按人次平均"
    cboSchExamTimeCalcType.AddItem "项目时长"
    cboSchExamTimeCalcType.ListIndex = 0
End Sub

Private Function saveTimeProject() As Boolean
'------------------------------------------------
'功能：保存时间计划设置
'参数：
'返回：True -- 成功； False -- 失败
'------------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strS1 As String
    Dim strS2 As String
    Dim strE1 As String
    Dim strE2 As String
    
    On Error GoTo err
    '先检查输入数据的有效性
    If dpTimeEnd.value <= dpTimeStart.value Then
        MsgBox "请重新输入开始时间和结束时间，开始时间应该小于结束时间。", vbOKOnly, "检查预约提示"
        dpTimeEnd.SetFocus
        Exit Function
    End If
    
    If Val(txtCapacity.Text) = 0 Then
        MsgBox "请重新输入预约容量。", vbOKOnly, "检查预约提示"
        txtCapacity.SetFocus
        Exit Function
    End If
    
    '如果平均检查时长小于2分钟，给出提示
    If DateDiff("n", dpTimeStart.value, dpTimeEnd.value) / Val(txtCapacity.Text) < 2 Then
        MsgBox "请检查预约容量是否输入错误，这个时间段内平均检查时长小于2分钟。", vbOKOnly, "检查预约提示"
        txtCapacity.SetFocus
        Exit Function
    End If
    
    '判断时间计划是否存在重复的时间
    strSql = "select ID,开始时间,结束时间 from 影像预约时间计划 where 预约方案ID=[1] order by 开始时间"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "查询重复的时间计划", mlngPlanID)
    strS1 = Format(dpTimeStart.value, "HH:MM")
    strE1 = Format(dpTimeEnd.value, "HH:MM")
    
    While rsTemp.EOF = False
        If rsTemp!ID <> mlngTimeProjectID Then
            strS2 = Format(NVL(rsTemp!开始时间), "HH:MM")
            strE2 = Format(NVL(rsTemp!结束时间), "HH:MM")
            If (strS1 <= strS2 And strS2 < strE1) _
                Or (strS1 < strE2 And strE2 <= strE1) _
                Or (strS2 <= strS1 And strS1 < strE2) Then
            
                MsgBox "请重新输入开始时间和结束时间，这个计划和其他计划存在时间重复。", vbOKOnly, "检查预约提示"
                dpTimeStart.SetFocus
                Exit Function
            End If
        End If
        rsTemp.MoveNext
    Wend
    
    strSql = "Zl_影像预约时间计划_更新(" & mlngTimeProjectID & "," & mlngPlanID & "," _
            & zlStr.To_Date(CDate(dpTimeStart.value)) & "," & zlStr.To_Date(CDate(dpTimeEnd.value)) _
            & "," & Val(txtCapacity.Text) & "," & IIF(cboSchExamTimeCalcType.ListIndex = 0, 1, 2) & ")"
    zlDatabase.ExecuteProcedure strSql, "保存检查预约时间计划"
    
    saveTimeProject = True
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub txtCapacity_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
