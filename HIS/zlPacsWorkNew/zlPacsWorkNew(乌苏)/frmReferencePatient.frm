VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReferencePatient 
   Caption         =   "关联病人"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   Icon            =   "frmReferencePatient.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkUseOldCheckNo 
      Caption         =   "应用关联病人的检查号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   27
      Top             =   8880
      Width           =   2895
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9240
      TabIndex        =   6
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton cmdDisRelating 
      Caption         =   "取消关联"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6840
      TabIndex        =   5
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton cmdRelating 
      Caption         =   "关联"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5400
      TabIndex        =   4
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查询"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   3
      Top             =   8760
      Width           =   1100
   End
   Begin VB.Frame Frame3 
      Caption         =   "查询条件"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   6360
      Width           =   10335
      Begin VB.OptionButton optFilter 
         Caption         =   "检查号 ="
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtFilter 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   6
         Left            =   1800
         TabIndex        =   25
         Top             =   772
         Width           =   3075
      End
      Begin VB.TextBox txtFilter 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   2
         Left            =   1800
         TabIndex        =   18
         Top             =   1725
         Width           =   3075
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "IC卡号   ="
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtFilter 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   5
         Left            =   7080
         TabIndex        =   16
         Top             =   1252
         Width           =   3075
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "身份证号 ="
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   5520
         TabIndex        =   15
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtFilter 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   1800
         TabIndex        =   14
         Top             =   1252
         Width           =   3075
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "就诊卡号 ="
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtFilter 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   4
         Left            =   7080
         TabIndex        =   12
         Top             =   772
         Width           =   3075
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "住院号   ="
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   5520
         TabIndex        =   11
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtFilter 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   3
         Left            =   7080
         TabIndex        =   10
         Top             =   285
         Width           =   3075
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "门诊号   ="
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   5520
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtFilter 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   1800
         TabIndex        =   8
         Top             =   292
         Width           =   3075
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "姓名     ="
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame frmRelated 
      Caption         =   "已关联"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   10335
      Begin MSComctlLib.ListView lvwRelated 
         Height          =   1215
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   2143
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "等待关联"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   10335
      Begin MSComctlLib.ListView lvwToBeRelate 
         Height          =   1335
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   2355
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwStudies 
         Height          =   1335
         Left            =   120
         TabIndex        =   24
         Top             =   1800
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   2355
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Label lblPatientInfo 
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
      Index           =   2
      Left            =   240
      TabIndex        =   23
      Top             =   840
      Width           =   10095
   End
   Begin VB.Label lblPatientInfo 
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
      Index           =   1
      Left            =   240
      TabIndex        =   22
      Top             =   480
      Width           =   10095
   End
   Begin VB.Label lblPatientInfo 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   21
      Top             =   120
      Width           =   10095
   End
End
Attribute VB_Name = "frmReferencePatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstr姓名 As String      '病人姓名，窗口显示时的默认查询条件
Private mlngOrderID As Long     '病人本次检查的医嘱ID
Private mlngPatietID As Long    '病人ID
Private mlngRelatingID As Long  '当前的关联ID
Private mblnInCheck As Boolean  '是否在程序控制更改Check的状态中，不再触发Check状态
Private mfrmParent As Form      '父窗体
Private mlngDetpID As Long      '当前科室ID
Private mlngStudyNoBuildType As Long        '检查号生成方式,0-按类别递增 1-按科室递增
Private mstrModality As String              '影像类别


Public Sub zlShowMe(lngOrderID As Long, str姓名 As String, frmParent As Form, blnShow As Boolean, lngDetpID As Long)
'显示关联病人的窗口
'参数： lngOrderID --- 医嘱ID
'       str姓名 --- 病人姓名
'       frmParent --- 父窗体
'       blnShow --- 没有可关联的病人是是否显示窗体，True-显示；False-不显示
'       mlngDetpID --- 执行科室ID，用来提取可是流程参数

    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim rsToBeRelate As ADODB.Recordset
    
    mstr姓名 = str姓名
    mlngOrderID = lngOrderID
    Set mfrmParent = frmParent
    mlngRelatingID = 0
    mlngDetpID = lngDetpID
    mlngStudyNoBuildType = Val(GetDeptPara(mlngDetpID, "检查号生成方式", 0))
    
    
    On Error GoTo err
    '查询并记录当前的病人ID
    strSql = "Select a.病人id,b.姓名, b.性别, b.年龄,to_char(b.出生日期,'yyyy-mm-dd') 出生日期, " & _
             " b.门诊号,b.住院号,b.就诊卡号, " & _
             " b.身份证号,b.职业,b.民族,b.婚姻状况,nvl(b.家庭地址,b.工作单位) 地址,nvl(b.家庭电话,b.联系人电话) 电话 " & _
             " From 病人医嘱记录 a ,病人信息 b Where a.病人id=b.病人id and id= [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取病人ID", mlngOrderID)
    If rsTemp.EOF = True Then Exit Sub
    
    mlngPatietID = rsTemp!病人ID
    lblPatientInfo(0).Caption = " 姓名：" & Nvl(rsTemp!姓名) & " 性别：" & Nvl(rsTemp!性别) & _
            " 年龄：" & Nvl(rsTemp!年龄) & " 出生日期：" & Nvl(rsTemp!出生日期)
    lblPatientInfo(1).Caption = " 门诊号：" & Nvl(rsTemp!门诊号) & " 住院号：" & Nvl(rsTemp!住院号) & _
            " 就诊卡号：" & Nvl(rsTemp!就诊卡号) & " 身份证号：" & Nvl(rsTemp!身份证号)
    lblPatientInfo(2).Caption = " 民族：" & Nvl(rsTemp!民族) & " 电话：" & Nvl(rsTemp!电话) & " 地址：" & Nvl(rsTemp!地址)
    
    
    strSql = "Select 影像类别 From 影像检查记录 Where 医嘱ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "查询影像类别", mlngOrderID)
    If rsTemp.EOF = False Then
        mstrModality = Nvl(rsTemp!影像类别)
    End If
        
    '查询是否有同名而且没有关联的病人
    strSql = "Select Distinct a.病人id, b.检查号, b.关联id, a.姓名, a.性别, a.年龄, a.出生日期, a.门诊号, a.住院号, a.就诊卡号, a.费别," & _
             " a.医疗付款方式 , a.身份证号, a.职业, a.民族, a.婚姻状况, a.地址, a.电话,b.影像类别,c.执行科室ID " & _
             " From (Select 病人id, 姓名, 性别, 年龄, To_Char(出生日期, 'yyyy-mm-dd') As 出生日期, 门诊号, 住院号, 就诊卡号, 费别, " & _
             "        医疗付款方式, 身份证号, 职业, 民族, 婚姻状况, Nvl(家庭地址, 工作单位) As 地址, Nvl(家庭电话, 联系人电话) 电话 " & _
             "       From 病人信息 Where 姓名 = [1] And 病人id <> [2]) a, 影像检查记录 b, 病人医嘱记录 c " & _
             " Where c.病人id = a.病人id And c.ID = b.医嘱ID And b.关联id Is Null Order By a.病人id "
    
    Set rsToBeRelate = zlDatabase.OpenSQLRecord(strSql, "提取关联病人", mstr姓名, mlngPatietID)
    
    If mlngStudyNoBuildType = 1 Then
        '检查号生成方式=1，则只查询本科室的检查
        rsToBeRelate.Filter = "执行科室ID = " & mlngDetpID
    Else
        '检查号生成方式=0，则只查询本影像类别的检查
        rsToBeRelate.Filter = "影像类别 = '" & mstrModality & "'"
    End If
    
    '如果没有关联的病人，且不显示窗体，则退出
    If rsToBeRelate.EOF = True And blnShow = False Then
        Exit Sub
    Else
        '初始化窗体
        Call InitLists
        
        Call FillToBeRelateList(rsToBeRelate)
        
        '再填充已经关联的列表
        strSql = "Select Distinct b.病人ID,a.检查号,a.关联ID,b.姓名,b.性别, b.年龄,to_char(b.出生日期,'yyyy-mm-dd') 出生日期," & _
             " b.门诊号,b.住院号,b.就诊卡号,b.费别,b.医疗付款方式,b.病人ID, " & _
             " b.身份证号,b.职业,b.民族,b.婚姻状况,nvl(b.家庭地址,b.工作单位) 地址,nvl(b.家庭电话,b.联系人电话) 电话 " & _
             " From (Select 医嘱id,检查号,关联ID From 影像检查记录 Where 关联id =(Select 关联ID From 影像检查记录 Where 医嘱id=[1])) a, " & _
             " 病人信息 b, 病人医嘱记录 c " & _
             " Where c.病人id = b.病人id And a.医嘱id = c.Id and b.病人ID <> [2] "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取关联病人", mlngOrderID, mlngPatietID)
        Call FillRelatedList(rsTemp)
        
        '显示窗体
        Me.Show 1, mfrmParent
    End If
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitLists()
'初始化“已关联”和“等待关联”列表
    
    On Error GoTo err
    '初始化“已关联”列表
    With lvwRelated.ColumnHeaders
        .Clear
        .Add , , "姓名", 1000
        .Add , , "性别", 800
        .Add , , "年龄", 800
        .Add , , "出生日期", 1200
        .Add , , "门诊号", 1000
        .Add , , "住院号", 1000
        .Add , , "检查号", 1000
        .Add , , "就诊卡号", 1000
        .Add , , "身份证号", 1400
        .Add , , "民族", 800
        .Add , , "电话", 1000
        .Add , , "地址", 2000
        .Add , , "关联ID", 0
    End With
    lvwRelated.ListItems.Add , , "Temp"
    
    '初始化“等待关联”列表
    With lvwToBeRelate.ColumnHeaders
        .Clear
        .Add , , "姓名", 1000
        .Add , , "性别", 800
        .Add , , "年龄", 800
        .Add , , "出生日期", 1200
        .Add , , "门诊号", 1000
        .Add , , "住院号", 1000
        .Add , , "检查号", 1000
        .Add , , "就诊卡号", 1000
        .Add , , "身份证号", 1400
        .Add , , "民族", 800
        .Add , , "电话", 1000
        .Add , , "地址", 2000
        .Add , , "关联ID", 0
    End With
    lvwToBeRelate.ListItems.Add , , "Temp"
    
    '初始化“检查”列表
    With lvwStudies.ColumnHeaders
        .Clear
        .Add , , "序号", 800
        .Add , , "影像类别", 1000
        .Add , , "检查号", 2000
        .Add , , "采图时间", 2000
        .Add , , "医嘱内容", 4000
    End With
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub FillToBeRelateList(rsToBeRelate As ADODB.Recordset)
'填充“等待关联”列表
'参数： rsToBeRelate --- 等待关联列表

    Dim tmpItem As MSComctlLib.ListItem
    Dim strPatientID As String
    Dim strRelatedPatientID As String   '记录已关联的病人ID
    Dim i As Integer
    
    On Error GoTo err
    
    '记录已关联的病人ID
    strRelatedPatientID = ""
    If lvwRelated.ListItems.Count >= 1 Then
        For i = 1 To lvwRelated.ListItems.Count
            strRelatedPatientID = strRelatedPatientID & "," & Mid(lvwRelated.ListItems(i).Key, 2)
        Next i
    End If
    
    '填充等待关联的列表
    strPatientID = ""
    lvwToBeRelate.ListItems.Clear
    
    While rsToBeRelate.EOF = False
        If InStr(strPatientID, rsToBeRelate("病人ID")) = 0 And InStr(strRelatedPatientID, rsToBeRelate("病人ID")) = 0 Then
            strPatientID = strPatientID & "," & rsToBeRelate("病人ID")
            Set tmpItem = lvwToBeRelate.ListItems.Add(, "_" & rsToBeRelate("病人ID"), rsToBeRelate("姓名"))
            tmpItem.SubItems(1) = Nvl(rsToBeRelate("性别"))
            tmpItem.SubItems(2) = Nvl(rsToBeRelate("年龄"))
            tmpItem.SubItems(3) = Nvl(rsToBeRelate("出生日期"))
            tmpItem.SubItems(4) = Nvl(rsToBeRelate("门诊号"))
            tmpItem.SubItems(5) = Nvl(rsToBeRelate("住院号"))
            tmpItem.SubItems(6) = Nvl(rsToBeRelate("检查号"))
            tmpItem.SubItems(7) = Nvl(rsToBeRelate("就诊卡号"))
            tmpItem.SubItems(8) = Nvl(rsToBeRelate("身份证号"))
            tmpItem.SubItems(9) = Nvl(rsToBeRelate("民族"))
            tmpItem.SubItems(10) = Nvl(rsToBeRelate("电话"))
            tmpItem.SubItems(11) = Nvl(rsToBeRelate("地址"))
            tmpItem.SubItems(12) = Nvl(rsToBeRelate("关联ID"))
        End If
        rsToBeRelate.MoveNext
    Wend
    
    '填写检查列表
    If lvwToBeRelate.ListItems.Count >= 1 Then
        Call lvwToBeRelate_ItemClick(lvwToBeRelate.ListItems(1))
    Else
        lvwStudies.ListItems.Clear
    End If
        
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub FillRelatedList(rsRelated As ADODB.Recordset)
'填充“已关联”列表
'参数： rsRelated --- 已关联的数据集

    Dim tmpItem As MSComctlLib.ListItem
    Dim strPatientID As String
    
    On Error GoTo err
    
    strPatientID = ""
    lvwRelated.ListItems.Clear
    
    While rsRelated.EOF = False
        If InStr(strPatientID, rsRelated("病人ID")) = 0 Then
            strPatientID = strPatientID & "," & rsRelated("病人ID")
            Set tmpItem = lvwRelated.ListItems.Add(, "_" & rsRelated("病人ID"), rsRelated("姓名"))
            tmpItem.SubItems(1) = Nvl(rsRelated("性别"))
            tmpItem.SubItems(2) = Nvl(rsRelated("年龄"))
            tmpItem.SubItems(3) = Nvl(rsRelated("出生日期"))
            tmpItem.SubItems(4) = Nvl(rsRelated("门诊号"))
            tmpItem.SubItems(5) = Nvl(rsRelated("住院号"))
            tmpItem.SubItems(6) = Nvl(rsRelated("检查号"))
            tmpItem.SubItems(7) = Nvl(rsRelated("就诊卡号"))
            tmpItem.SubItems(8) = Nvl(rsRelated("身份证号"))
            tmpItem.SubItems(9) = Nvl(rsRelated("民族"))
            tmpItem.SubItems(10) = Nvl(rsRelated("电话"))
            tmpItem.SubItems(11) = Nvl(rsRelated("地址"))
            tmpItem.SubItems(12) = Nvl(rsRelated("关联ID"))
            mlngRelatingID = Nvl(rsRelated("关联ID"))
        End If
        rsRelated.MoveNext
    Wend
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdDisRelating_Click()
'把“已关联”列表中被选中的项目设置成“等待关联”
    Dim ItemSelected As MSComctlLib.ListItem
    Dim ItemDeleteRelate As MSComctlLib.ListItem
    Dim lngPatientID As Long
    Dim i As Integer
    
    On Error GoTo err
    
    '判断当前“已关联”列表中是否有被选中的项目，循环被选中的项目，取消关联
    For Each ItemSelected In lvwRelated.ListItems
        lngPatientID = Mid(ItemSelected.Key, 2)
        If ItemSelected.Checked = True Then
            '取消数据库的关联
            Call DeleteRelating(lngPatientID)
            '把当前“已关联”列表中选中的项目移动到“等待关联”列表中
            Set ItemDeleteRelate = lvwToBeRelate.ListItems.Add(, ItemSelected.Key, ItemSelected.Text)
            ItemDeleteRelate.SubItems(1) = ItemSelected.SubItems(1)
            ItemDeleteRelate.SubItems(2) = ItemSelected.SubItems(2)
            ItemDeleteRelate.SubItems(3) = ItemSelected.SubItems(3)
            ItemDeleteRelate.SubItems(4) = ItemSelected.SubItems(4)
            ItemDeleteRelate.SubItems(5) = ItemSelected.SubItems(5)
            ItemDeleteRelate.SubItems(6) = ItemSelected.SubItems(6)
            ItemDeleteRelate.SubItems(7) = ItemSelected.SubItems(7)
            ItemDeleteRelate.SubItems(8) = ItemSelected.SubItems(8)
            ItemDeleteRelate.SubItems(9) = ItemSelected.SubItems(9)
            ItemDeleteRelate.SubItems(10) = ItemSelected.SubItems(10)
            ItemDeleteRelate.SubItems(11) = ItemSelected.SubItems(11)
            ItemDeleteRelate.SubItems(12) = ItemSelected.SubItems(12)
        End If
    Next
    
    '删除列表中被选中的项目
    For i = lvwRelated.ListItems.Count To 1 Step -1
        If lvwRelated.ListItems(i).Checked = True Then
            lvwRelated.ListItems.Remove i
        End If
    Next i
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdQuery_Click()
'查询等待关联的病人信息
    Dim i As Integer
    Dim blnQuery As Boolean
    Dim intFilterIndex As Integer
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strFilter As String
    
    intFilterIndex = -1
    On Error GoTo err
        For i = 0 To 6
            If optFilter(i).value = True Then
                blnQuery = True
                intFilterIndex = i
                Exit For
            End If
        Next i
        
        If blnQuery = False Then
            MsgBoxD mfrmParent, "请选择条件后再查询", vbOKOnly, "提示信息"
        Else
            strSql = "Select Distinct b.病人ID,a.检查号,a.关联ID,nvl(b.姓名,a.姓名) 姓名,nvl(b.性别,a.性别) 性别, nvl(b.年龄,a.年龄) 年龄,to_char(b.出生日期,'yyyy-mm-dd') 出生日期," & _
             " b.门诊号,b.住院号,b.就诊卡号,b.费别,b.医疗付款方式,b.病人ID,a.影像类别,c.执行科室ID, " & _
             " b.身份证号,b.职业,b.民族,b.婚姻状况,nvl(b.家庭地址,b.工作单位) 地址,nvl(b.家庭电话,b.联系人电话) 电话 " & _
             " From 影像检查记录 a, 病人信息 b, 病人医嘱记录 c " & _
             " Where c.病人id = b.病人id And a.医嘱id = c.Id and b.病人ID <> [1] "
            If mlngRelatingID <> 0 Then
                strFilter = strFilter & " and (a.关联ID <> [2] Or a.关联id Is Null) "
            End If
            Select Case intFilterIndex
            Case 0  '姓名
                strFilter = strFilter & " and b.姓名 = [3] "
            Case 1  '就诊卡
                strFilter = strFilter & " and b.就诊卡号 = [4] "
            Case 2  'IC卡
                strFilter = strFilter & " and b.IC卡号 = [5] "
            Case 3  '门诊号
                strFilter = strFilter & " and b.门诊号 = [6] "
            Case 4  '住院号
                strFilter = strFilter & " and b.住院号 = [7] "
            Case 5  '身份证号
                strFilter = strFilter & " and b.身份证号 = [8] "
            Case 6  '检查号
                strFilter = strFilter & " and a.检查号 = [9] "
            End Select
            
            strSql = strSql & strFilter
            
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取关联病人", mlngPatietID, mlngRelatingID, CStr(txtFilter(0).Text), _
                        CStr(txtFilter(1).Text), CStr(txtFilter(2).Text), CLng(Val(txtFilter(3).Text)), CLng(Val(txtFilter(4).Text)), _
                        CStr(txtFilter(5).Text), CStr(txtFilter(6).Text))
            
            If mlngStudyNoBuildType = 1 Then
                '检查号生成方式=1，则只查询本科室的检查
                rsTemp.Filter = "执行科室ID = " & mlngDetpID
            Else
                '检查号生成方式=0，则只查询本影像类别的检查
                rsTemp.Filter = "影像类别 = '" & mstrModality & "'"
            End If
    
            Call FillToBeRelateList(rsTemp)
        End If
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdRelating_Click()
'把“等待关联列表”中选中的项目设置成跟当前病人关联
    Dim ItemSelected As MSComctlLib.ListItem
    Dim ItemRelated As MSComctlLib.ListItem
    Dim lngPatientID As Long
    Dim i As Integer
    Dim str检查号  As String
    Dim arr检查号() As String
    Dim strReturn As String
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    str检查号 = "||"
    
    '判断当前“等待关联列表”是否有被选中的项目,循环被选中的项目，设置关联
    For Each ItemSelected In lvwToBeRelate.ListItems
        lngPatientID = Mid(ItemSelected.Key, 2)
        If ItemSelected.Checked = True Then
            '在数据库中设置关联
            Call SetRelating(lngPatientID)
            '把当前“等待关联”列表中选中的项目移动到“已关联”列表中
            Set ItemRelated = lvwRelated.ListItems.Add(, ItemSelected.Key, ItemSelected.Text)
            ItemRelated.SubItems(1) = ItemSelected.SubItems(1)
            ItemRelated.SubItems(2) = ItemSelected.SubItems(2)
            ItemRelated.SubItems(3) = ItemSelected.SubItems(3)
            ItemRelated.SubItems(4) = ItemSelected.SubItems(4)
            ItemRelated.SubItems(5) = ItemSelected.SubItems(5)
            ItemRelated.SubItems(6) = ItemSelected.SubItems(6)
            ItemRelated.SubItems(7) = ItemSelected.SubItems(7)
            ItemRelated.SubItems(8) = ItemSelected.SubItems(8)
            ItemRelated.SubItems(9) = ItemSelected.SubItems(9)
            ItemRelated.SubItems(10) = ItemSelected.SubItems(10)
            ItemRelated.SubItems(11) = ItemSelected.SubItems(11)
            ItemRelated.SubItems(12) = ItemSelected.SubItems(12)
            
            '查询当前要关联的病人已经有的检查号
            strSql = "Select 影像类别,检查号,a.执行科室ID From 影像检查记录 a,病人医嘱记录 b Where a.医嘱ID=b.Id And  b.病人ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取病人检查号", lngPatientID)
            
            If mlngStudyNoBuildType = 1 Then
                '检查号生成方式=1，则只查询本科室的检查
                rsTemp.Filter = "执行科室ID = " & mlngDetpID
            Else
                '检查号生成方式=0，则只查询本影像类别的检查
                rsTemp.Filter = "影像类别 = '" & mstrModality & "'"
            End If
            
            If rsTemp.RecordCount > 1 Then
                For i = 1 To rsTemp.RecordCount
                    If InStr(str检查号, "||" & rsTemp("检查号") & "||") = 0 Then
                        str检查号 = str检查号 & rsTemp("检查号") & "||" '确保每一个检查号都是由一对||包围
                    End If
                    rsTemp.MoveNext
                Next i
            Else
                If InStr(str检查号, "||" & ItemSelected.SubItems(6) & "||") = 0 Then
                    str检查号 = str检查号 & ItemSelected.SubItems(6) & "||" '确保每一个检查号都是由一对||包围
                End If
            End If
             
            
        End If
    Next
    
    '删除列表中被选中的项目
    For i = lvwToBeRelate.ListItems.Count To 1 Step -1
        If lvwToBeRelate.ListItems(i).Checked = True Then
            lvwToBeRelate.ListItems.Remove i
        End If
    Next i
    
    '处理是否自动应用检查号
    If chkUseOldCheckNo.value = 1 And str检查号 <> "||" Then
        '是否有多个检查号，如果有多个不同的检查号，则提示用户选择
        arr检查号 = Split(str检查号, "||")
        If UBound(arr检查号) > 2 Then
            '有多个检查号，提示用户自己选择
            For i = 1 To UBound(arr检查号) - 1
                strReturn = strReturn & i & "----" & arr检查号(i) & vbCrLf
            Next i
            strReturn = InputBox("本次关联中使用了多个检查号" & vbCrLf & "请输入编号选择其中一个检查号" & vbCrLf & vbCrLf _
                        & strReturn & vbCrLf & "如果输入无效编号，表示不应用任何一个检查号。", "选择检查号", "1")
            If Val(strReturn) >= 1 And Val(strReturn) <= UBound(arr检查号) - 1 Then
                strReturn = arr检查号(Val(strReturn))
                Call subSetCheckNo(mlngOrderID, strReturn)
            End If
        Else
            '只有一个检查号，直接修改
            strReturn = arr检查号(1)
            Call subSetCheckNo(mlngOrderID, strReturn)
        End If
    End If
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetRelating(lngPatientID As Long)
'设置关联
'参数： lngPatientID -- 需要被关联的病人ID
    Dim strSql As String
    
    On Error GoTo err
    
    strSql = "ZL_影像关联病人(" & mlngOrderID & "," & mlngPatietID & "," & lngPatientID & ")"
    zlDatabase.ExecuteProcedure strSql, "关联病人"
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub DeleteRelating(lngPatientID As Long)
'取消关联
'参数：lngPatientID -- 需要被取消关联的病人ID
    Dim strSql As String

    On Error GoTo err
        
    strSql = "ZL_影像取消关联病人(" & mlngOrderID & "," & lngPatientID & ")"
    zlDatabase.ExecuteProcedure strSql, "关联病人"
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Dim strRegPath As String
    
    strRegPath = "公共模块\" & App.ProductName & "\frmReferencePatient"
    
    chkUseOldCheckNo.value = Val(GetSetting("ZLSOFT", strRegPath, "应用关联病人的检查号", 0))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strRegPath As String
    
    strRegPath = "公共模块\" & App.ProductName & "\frmReferencePatient"
    Call SaveSetting("ZLSOFT", strRegPath, "应用关联病人的检查号", chkUseOldCheckNo.value)
End Sub

Private Sub lvwToBeRelate_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    '设置关联,检查是否还有跟他关联的项目，同时也选择上
    Dim i As Integer
    
    If mblnInCheck = True Then Exit Sub
    
    mblnInCheck = True
    For i = 1 To lvwToBeRelate.ListItems.Count
        If lvwToBeRelate.ListItems(i).SubItems(12) <> "" And lvwToBeRelate.ListItems(i).SubItems(12) = Item.SubItems(12) Then
            lvwToBeRelate.ListItems(i).Checked = Item.Checked
        End If
    Next i
    mblnInCheck = False
End Sub

Private Sub lvwToBeRelate_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '显示当前病人对应的所有检查
    Dim strSql As String
    Dim rsStudies As ADODB.Recordset
    Dim lngPatientID As Long
    
    On Error GoTo err
    
    lngPatientID = Mid(Item.Key, 2)
    strSql = "Select 影像类别,医嘱ID,接收日期,医嘱内容,检查号,b.执行科室ID From 影像检查记录 a,病人医嘱记录 b Where a.医嘱ID=b.Id And b.病人id=[1]  And b.相关id Is Null order by 接收日期"
    Set rsStudies = zlDatabase.OpenSQLRecord(strSql, "提取病人检查信息", lngPatientID)
    
    If mlngStudyNoBuildType = 1 Then
        '检查号生成方式=1，则只查询本科室的检查
        rsStudies.Filter = "执行科室ID = " & mlngDetpID
    Else
        '检查号生成方式=0，则只查询本影像类别的检查
        rsStudies.Filter = "影像类别 = '" & mstrModality & "'"
    End If
    
    Call FillStudies(rsStudies)
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub txtFilter_GotFocus(Index As Integer)
    optFilter(Index).value = True
End Sub

Private Sub txtFilter_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdQuery_Click
    End If
End Sub

Private Sub FillStudies(rsStudies As ADODB.Recordset)
    '填充检查列表
    Dim tmpItem As MSComctlLib.ListItem
    Dim i As Integer
    
    On Error GoTo err
    lvwStudies.ListItems.Clear
    i = 1
    
    While rsStudies.EOF = False
        Set tmpItem = lvwStudies.ListItems.Add(, "_" & rsStudies("医嘱ID"), i)
        tmpItem.SubItems(1) = Nvl(rsStudies("影像类别"))
        tmpItem.SubItems(2) = Nvl(rsStudies("检查号"))
        tmpItem.SubItems(3) = Nvl(rsStudies("接收日期"))
        tmpItem.SubItems(4) = Nvl(rsStudies("医嘱内容"))
        rsStudies.MoveNext
        i = i + 1
    Wend
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub subSetCheckNo(lng医嘱ID As Long, str检查号 As String)
'------------------------------------------------
'功能：设置新的检查号
'参数： lng医嘱ID--医嘱ID
'       str检查号 -- 新的检查号
'返回：无
'----------------------------------------------
    Dim strSql As String
    
    On Error GoTo err
    
    strSql = "Zl_影像检查号_更新( " & lng医嘱ID & "," & str检查号 & ")"
    zlDatabase.ExecuteProcedure strSql, "设置新的检查号"
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
