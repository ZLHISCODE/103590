VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPACSFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤条件"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPACSFilter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox PicButton 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   6345
      TabIndex        =   22
      Top             =   4740
      Width           =   6345
      Begin VB.ComboBox cboSchemaName 
         Height          =   330
         Left            =   1080
         TabIndex        =   61
         Top             =   120
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmdDelSchema 
         Caption         =   "删除方案(&D)"
         Height          =   375
         Left            =   3480
         TabIndex        =   60
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdSaveSchema 
         Caption         =   "保存方案(&S)"
         Height          =   375
         Left            =   4920
         TabIndex        =   59
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkmore 
         Caption         =   "更多条件(&M)"
         Height          =   375
         Left            =   1440
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   1305
      End
      Begin VB.CommandButton cmdDefault 
         Caption         =   "缺省(&F)"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   375
         Left            =   5025
         TabIndex        =   15
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   375
         Left            =   3660
         TabIndex        =   14
         ToolTipText     =   "F2"
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label labSchema 
         Caption         =   "查询方案："
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picmore 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   0
      ScaleHeight     =   2745
      ScaleWidth      =   6345
      TabIndex        =   28
      Top             =   4920
      Visible         =   0   'False
      Width           =   6345
      Begin VB.TextBox txt报告内容 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   38
         Top             =   405
         Width           =   4410
      End
      Begin VB.TextBox Txt影像诊断 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   10
         Top             =   30
         Width           =   4410
      End
      Begin VB.Frame Frame1 
         Caption         =   "PACS报告查询"
         Height          =   1455
         Left            =   240
         TabIndex        =   29
         Top             =   1200
         Width           =   5850
         Begin VB.TextBox txtPacsRpt 
            Height          =   300
            Index           =   0
            Left            =   1440
            TabIndex        =   12
            Top             =   240
            Width           =   4215
         End
         Begin VB.TextBox txtPacsRpt 
            Height          =   300
            Index           =   1
            Left            =   1440
            TabIndex        =   16
            Top             =   600
            Width           =   4215
         End
         Begin VB.TextBox txtPacsRpt 
            Height          =   300
            Index           =   2
            Left            =   1440
            TabIndex        =   13
            Top             =   960
            Width           =   4215
         End
         Begin VB.Label lblPacsRpt 
            Caption         =   "检查所见"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   32
            Top             =   390
            Width           =   975
         End
         Begin VB.Label lblPacsRpt 
            Caption         =   "或 诊断意见"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   31
            Top             =   750
            Width           =   1335
         End
         Begin VB.Label lblPacsRpt 
            Caption         =   "或 建议"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   30
            Top             =   1110
            Width           =   1095
         End
      End
      Begin VB.TextBox txt随访 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   11
         Top             =   795
         Width           =   4410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "报告内容"
         Height          =   210
         Left            =   240
         TabIndex        =   39
         Top             =   465
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "疾病诊断"
         Height          =   210
         Left            =   240
         TabIndex        =   34
         Top             =   90
         Width           =   840
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "随访"
         Height          =   210
         Left            =   240
         TabIndex        =   33
         Top             =   855
         Width           =   420
      End
   End
   Begin VB.Frame Frabase 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4860
      Left            =   -120
      TabIndex        =   17
      Top             =   0
      Width           =   7005
      Begin VB.ComboBox cboAgeType 
         Height          =   330
         ItemData        =   "frmPACSFilter.frx":000C
         Left            =   4800
         List            =   "frmPACSFilter.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   2640
         Width           =   1425
      End
      Begin VB.TextBox txtEndAge 
         Height          =   330
         Left            =   3840
         TabIndex        =   57
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox txtBeginAge 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1275
         TabIndex        =   56
         Top             =   2640
         Width           =   855
      End
      Begin VB.ComboBox cboSex 
         Height          =   330
         ItemData        =   "frmPACSFilter.frx":0038
         Left            =   1275
         List            =   "frmPACSFilter.frx":003A
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   2280
         Width           =   1665
      End
      Begin VB.ComboBox cboAgeWhere 
         Height          =   330
         ItemData        =   "frmPACSFilter.frx":003C
         Left            =   2280
         List            =   "frmPACSFilter.frx":0052
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   2640
         Width           =   1425
      End
      Begin VB.CommandButton cmdDayCfg 
         Caption         =   "近半年(&Y)"
         Height          =   375
         Index           =   7
         Left            =   5040
         TabIndex        =   51
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdDayCfg 
         Caption         =   "近三月(&H)"
         Height          =   375
         Index           =   6
         Left            =   3480
         TabIndex        =   50
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdDayCfg 
         Caption         =   "近二月(&U)"
         Height          =   375
         Index           =   5
         Left            =   1920
         TabIndex        =   49
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdDayCfg 
         Caption         =   "近一月(&N)"
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   48
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton cmdDayCfg 
         Caption         =   "近二周(&K)"
         Height          =   375
         Index           =   3
         Left            =   5040
         TabIndex        =   47
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdDayCfg 
         Caption         =   "近一周(&W)"
         Height          =   375
         Index           =   2
         Left            =   3480
         TabIndex        =   46
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdDayCfg 
         Caption         =   "近两天(&A)"
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   45
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdDayCfg 
         Caption         =   "今天(&T)"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   44
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ComboBox cboYinYangXing 
         Height          =   330
         ItemData        =   "frmPACSFilter.frx":006E
         Left            =   4365
         List            =   "frmPACSFilter.frx":007B
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   4320
         Width           =   1890
      End
      Begin VB.OptionButton optFindType 
         Caption         =   "按采图时间查找"
         Height          =   300
         Index           =   3
         Left            =   4560
         TabIndex        =   37
         Top             =   240
         Width           =   1800
      End
      Begin VB.ComboBox cbodiagdoc 
         Height          =   330
         Left            =   4365
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3960
         Width           =   1890
      End
      Begin VB.ComboBox cbo检查技师 
         Height          =   330
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3975
         Width           =   1905
      End
      Begin VB.ComboBox cboAuditing 
         Height          =   330
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   4350
         Width           =   1905
      End
      Begin VB.ComboBox cboModality 
         Height          =   330
         ItemData        =   "frmPACSFilter.frx":0095
         Left            =   4365
         List            =   "frmPACSFilter.frx":0097
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3615
         Width           =   1890
      End
      Begin VB.ComboBox cbo质量 
         Height          =   330
         ItemData        =   "frmPACSFilter.frx":0099
         Left            =   1275
         List            =   "frmPACSFilter.frx":00A6
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3615
         Width           =   1905
      End
      Begin VB.OptionButton optFindType 
         Caption         =   "按报到时间查找"
         Height          =   300
         Index           =   2
         Left            =   2520
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1995
      End
      Begin VB.ComboBox cboPart 
         Height          =   330
         Left            =   4365
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3240
         Width           =   1890
      End
      Begin VB.OptionButton optFindType 
         Caption         =   "按申请时间查找"
         Height          =   300
         Index           =   1
         Left            =   315
         TabIndex        =   18
         Top             =   240
         Width           =   1800
      End
      Begin VB.ComboBox cboDept 
         Height          =   330
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3240
         Width           =   1905
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   4080
         TabIndex        =   2
         Top             =   570
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   529
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
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   89260035
         CurrentDate     =   38082
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1440
         TabIndex        =   1
         Top             =   600
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
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
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   89260035
         CurrentDate     =   38082
      End
      Begin MSComctlLib.Slider sldDays 
         Height          =   300
         Left            =   1320
         TabIndex        =   43
         Top             =   960
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   529
         _Version        =   393216
         Max             =   180
         TickFrequency   =   7
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人性别"
         Height          =   210
         Left            =   315
         TabIndex        =   55
         Top             =   2340
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "病人年龄"
         Height          =   210
         Left            =   315
         TabIndex        =   54
         Top             =   2760
         Width           =   840
      End
      Begin VB.Label Label 
         Caption         =   "时间范围"
         Height          =   255
         Index           =   0
         Left            =   315
         TabIndex        =   42
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "阴 阳 性"
         Height          =   210
         Left            =   3435
         TabIndex        =   40
         Top             =   4440
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "诊断医生"
         Height          =   210
         Left            =   3435
         TabIndex        =   36
         Top             =   4035
         Width           =   840
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检查技师"
         Height          =   210
         Left            =   315
         TabIndex        =   35
         Top             =   4035
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核医生"
         Height          =   210
         Left            =   315
         TabIndex        =   27
         Top             =   4410
         Width           =   840
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "影像类别"
         Height          =   210
         Left            =   3435
         TabIndex        =   26
         Top             =   3675
         Width           =   840
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "影像质量"
         Height          =   210
         Left            =   315
         TabIndex        =   25
         Top             =   3675
         Width           =   840
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "检查部位"
         Height          =   210
         Left            =   3435
         TabIndex        =   21
         Top             =   3300
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人科室"
         Height          =   210
         Left            =   315
         TabIndex        =   20
         Top             =   3300
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "～"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3720
         TabIndex        =   19
         Top             =   600
         Width           =   180
      End
   End
End
Attribute VB_Name = "frmPACSFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mlngModul As Long      '模块号调用
Public mBeforeDays As Integer '查询天数
Public mDept As Long
Public mblnOK As Boolean '确定退出


Private Sub LoadSex()
'********************************************
'
'读取病人性别
'
'********************************************
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    strSQL = "Select 名称 From 性别"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取病人性别")
    cboSex.Clear
    cboSex.AddItem "全部"
        
    With Me.cboSex
        Do While Not rsTmp.EOF
            .AddItem zlCommFun.SpellCode(Nvl(rsTmp("名称"))) & "-" & Nvl(rsTmp("名称"))
            rsTmp.MoveNext
        Loop
    End With
    
    cboSex.ListIndex = 0
 
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub cboAgeWhere_Click()
    txtBeginAge.Enabled = IIf(cboAgeWhere.Text = "~ ", True, False)
    If txtBeginAge.Enabled Then
        txtBeginAge.BackColor = &HFFFFFF
    Else
        txtBeginAge.Text = ""
        txtBeginAge.BackColor = &HE0E0E0
    End If
End Sub

Private Sub chkmore_Click()
    If chkmore.Value = 1 Then
        Me.Height = Picmore.Top + Picmore.Height + PicButton.Height + 500
        Picmore.Visible = True
    Else
        Me.Height = Frabase.Top + Frabase.Height + PicButton.Height + 500
        Picmore.Visible = False
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Me.Hide
End Sub

Private Sub cmdDayCfg_Click(Index As Integer)
    Select Case Index
        Case 0
            dtpBegin.Value = CDate(Now - Time)
            sldDays.Value = 0
        Case 1
            dtpBegin.Value = CDate(dtpEnd.Value - 2)
            sldDays.Value = 2
        Case 2
            dtpBegin.Value = CDate(dtpEnd.Value - 7)
            sldDays.Value = 7
        Case 3
            dtpBegin.Value = CDate(dtpEnd.Value - 14)
            sldDays.Value = 14
        Case 4
            dtpBegin.Value = CDate(dtpEnd.Value - 30)
            sldDays.Value = 30
        Case 5
            dtpBegin.Value = CDate(dtpEnd.Value - 60)
            sldDays.Value = 60
        Case 6
            dtpBegin.Value = CDate(dtpEnd.Value - 90)
            sldDays.Value = 90
        Case 7
            dtpBegin.Value = CDate(dtpEnd.Value - 180)
            sldDays.Value = 180
    End Select
End Sub

Private Sub cmdDefault_Click()
    Call Form_Load
    Call Form_Activate
End Sub

Private Sub cmdOK_Click()

    If dtpEnd.Value < dtpBegin.Value Then
        MsgBox "开始时间不能大于截止时间，请检查！", vbInformation, gstrSysName
        dtpEnd.SetFocus
        Exit Sub
    End If
    
    mblnOK = True
    Me.Hide
End Sub

Private Sub GetQuerySchemaCfg(ByRef strSchemaFormat As String, ByRef strSchemaField As String)
    Dim strSchema As String
    Dim intBeginAge As Integer
    Dim intEndAge As Integer
    
    If Not dtpBegin Is Nothing Then
        If optFindType(1).Value Then strSchema = strSchema & GetFieldFormatStr("发送时间", "病人医嘱发送", ">=", dtpBegin.Value, "And", "")
        If optFindType(2).Value Then strSchema = strSchema & GetFieldFormatStr("首次时间", "病人医嘱发送", ">=", dtpBegin.Value, "And", "")
        If optFindType(3).Value Then strSchema = strSchema & GetFieldFormatStr("接收日期", "影像检查记录", ">=", dtpBegin.Value, "And", "")
    End If
        
    If Not dtpEnd Is Nothing Then
        If optFindType(1).Value Then strSchema = strSchema & GetFieldFormatStr("发送时间", "病人医嘱发送", "<=", dtpEnd.Value, "And", "")
        If optFindType(2).Value Then strSchema = strSchema & GetFieldFormatStr("首次时间", "病人医嘱发送", "<=", dtpEnd.Value, "And", "")
        If optFindType(3).Value Then strSchema = strSchema & GetFieldFormatStr("接收日期", "影像检查记录", "<=", dtpEnd.Value, "And", "")
    End If
    
    If cboSex.ListIndex <> 0 Then strSchema = strSchema & GetFieldFormatStr("性别", "病人信息", "<=", NeedName(cboSex.Text), "And", "")
    
    
    Select Case NeedName(cboAgeType.Text)
        Case "岁"
            intBeginAge = Val(txtBeginAge.Text) * 365
            intEndAge = Val(txtEndAge.Text) * 365
        Case "月"
            intBeginAge = Val(txtBeginAge.Text) * 30
            intEndAge = Val(txtEndAge.Text) * 30
        Case "周"
            intBeginAge = Val(txtBeginAge.Text) * 7
            intEndAge = Val(txtEndAge.Text) * 7
        Case "天"
            intBeginAge = Val(txtBeginAge.Text) * 1
            intEndAge = Val(txtEndAge.Text) * 1
    End Select
        
    If Trim(txtBeginAge.Text) <> "" Then
        If Trim(cboAgeWhere.Text) = "~" Then
            strSchema = strSchema & GetFieldFormatStr("年龄", "病人信息", ">=", CStr(intBeginAge), "And", "")
        End If
    End If
    
    If Trim(txtEndAge.Text) <> "" Then
        If Trim(cboAgeWhere.Text) = "~" Then
            strSchema = strSchema & GetFieldFormatStr("年龄", "病人信息", "<=", CStr(intEndAge), "And", "")
        Else
            strSchema = strSchema & GetFieldFormatStr("年龄", "病人信息", Trim(cboAgeWhere.Text), CStr(intEndAge), "And", "")
        End If
    End If
    
    If cboDept.ListIndex <> 0 Then
        strSchema = strSchema & GetFieldFormatStr("病人科室ID+0", "病人医嘱记录", "=", cboDept.ItemData(cboDept.ListIndex), "And", "")
    End If
    
    If cboPart.ListIndex <> 0 Then
        strSchema = strSchema & GetFieldFormatStr("病人科室ID+0", "病人医嘱记录", "=", cboDept.ItemData(cboDept.ListIndex), "And", "")
    End If
        
End Sub

Private Function GetFieldFormatStr(strFieldName As String, strTabName As String, _
    strWhere As String, strData As String, strLink As String, strQueryType As String, Optional strBracket As String) As String
    Dim strResult As String
            
    On Error GoTo errHandle
        strResult = "<" & strFieldName & ">#B=" & strBracket & "#F=" & strTabName & "#W=" & strWhere & "#D=" & strData & "#L=" & strLink & "#T=" & strQueryType & "</" & strFieldName & ">"
        
        GetFieldFormatStr = strResult & vbNewLine
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub GetQueryTimeType()
    'If optFindType(1).value Then GetQueryTimeType = "发送时间"
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = vbKeyF2 Then cmdOK_Click
End Sub
Private Sub Form_Activate()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim pReport_CheckViewName As String
    Dim pReport_ResultName As String
    Dim pReport_AdviceName As String
    
    Txt影像诊断.Text = ""
    txt报告内容.Text = ""
    'cboYinYangXing.ListIndex = 0
    'cbo质量.ListIndex = 0
    'cboSex.ListIndex = 0
    'cboAgeWhere.ListIndex = 4
    'cboAgeType.ListIndex = 0
    txt随访.Text = ""
    
    
        
        pReport_CheckViewName = "检查所见"
        pReport_ResultName = "诊断意见"
        pReport_AdviceName = "建议"

    txtPacsRpt(0).Text = ""
    txtPacsRpt(1).Text = ""
    txtPacsRpt(2).Text = ""
    lblPacsRpt(0).Caption = pReport_CheckViewName
    lblPacsRpt(1).Caption = "或 " & pReport_ResultName
    lblPacsRpt(2).Caption = "或 " & pReport_AdviceName
    
    If mlngModul = 1290 Then    '影像医技工作站
        Label15.Visible = True
        cboModality.Visible = True
    ElseIf mlngModul = 1291 Then    '影像采集工作站
        Label15.Visible = False
        cboModality.Visible = False
    ElseIf mlngModul = 1293 Then        '影像病理工作站
        Label15.Visible = False
        cboModality.Visible = False
    End If
        
    dtpBegin.SetFocus
End Sub
Private Sub LoadDept()
'功能：根据病人来源读取病人科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    strSQL = "Select Distinct A.ID,A.编码,A.名称,B.服务对象" & _
        " From 部门表 A,部门性质说明 B" & _
        " Where A.ID=B.部门ID And B.工作性质 IN('临床','手术')" & _
        " And (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by A.编码"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    cboDept.Clear
    cboDept.AddItem "所有科室"
    cboDept.ListIndex = 0
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Next

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function InitPart() As Boolean
    Dim rsTmp As ADODB.Recordset
    gstrSQL = "Select Distinct 名称 From 诊疗检查部位"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取部位")
    cboPart.Clear
    cboPart.AddItem "所有部位"
    cboPart.ListIndex = 0
    With Me.cboPart
        Do While Not rsTmp.EOF
            .AddItem zlCommFun.SpellCode(Nvl(rsTmp("名称"))) & "-" & Nvl(rsTmp("名称"))
            rsTmp.MoveNext
        Loop
    End With
 
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub InitDoc()
Dim rsTmp As ADODB.Recordset
    cbodiagdoc.Clear: cboAuditing.Clear: cbo检查技师.Clear
    cbodiagdoc.AddItem "所有医生": cboAuditing.AddItem "所有医生": cbo检查技师.AddItem "所有医生"
    cbodiagdoc.ListIndex = 0: cboAuditing.ListIndex = 0: cbo检查技师.ListIndex = 0
    On Error GoTo errH
    gstrSQL = "select distinct A.简码,A.姓名 from 人员表 A,部门人员 B where B.部门ID=[1] AND A.ID=B.人员ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取本科室医生", mDept)
    If rsTmp Is Nothing Then Exit Sub
    Do While Not rsTmp.EOF
        cbodiagdoc.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        cboAuditing.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        cbo检查技师.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optAdviceTime_Click()
    Me.dtpBegin.SetFocus
End Sub

Private Sub optCheckTime_Click()
    Me.dtpBegin.SetFocus
End Sub

Private Sub Form_Load()
Dim curDate As Date
Dim int时间类型 As Integer

    curDate = zlDatabase.Currentdate
    dtpEnd.Value = Format(curDate, "yyyy-MM-dd 23:59")
    dtpEnd.Tag = dtpEnd.Value
    dtpBegin.Value = Format(dtpEnd.Value - mBeforeDays, "yyyy-MM-dd 00:00")
    
    int时间类型 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "过滤时间类型", 1))
    If int时间类型 = 1 Then
        optFindType(1).Value = True
    ElseIf int时间类型 = 2 Then
        optFindType(2).Value = True
    Else
        optFindType(3).Value = True
    End If
    
    LoadDept
    LoadSex
    InitPart
    InitDoc
    InitModality
    
    cboYinYangXing.ListIndex = 0
    cbo质量.ListIndex = 0
    cboSex.ListIndex = 0
    cboAgeWhere.ListIndex = 0
    cboAgeType.ListIndex = 0
End Sub
Private Sub InitModality()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    strSQL = "select 编码,名称 from 影像检查类别"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "影像检查类别")
    
    cboModality.Clear
    cboModality.AddItem "全部"
    
    Do Until rsTemp.EOF
        cboModality.AddItem rsTemp!名称 & "--" & rsTemp!编码
        rsTemp.MoveNext
    Loop
    cboModality.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim int时间类型 As Integer
    '保存过滤时间类型
    If optFindType(1).Value = True Then
        int时间类型 = 1
    ElseIf optFindType(2).Value = True Then
        int时间类型 = 2
    Else
        int时间类型 = 3
    End If
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "过滤时间类型", int时间类型
End Sub

Private Sub sldDays_Scroll()
    '设置查询时间范围
    dtpBegin.Value = CDate(dtpEnd.Value - sldDays.Value)
End Sub
