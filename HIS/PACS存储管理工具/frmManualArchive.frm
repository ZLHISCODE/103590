VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmManualArchive 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "手动归档"
   ClientHeight    =   3630
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab sstabManualArchive 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   6376
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "第一步"
      TabPicture(0)   =   "frmManualArchive.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdStep1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "第二步"
      TabPicture(1)   =   "frmManualArchive.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdStep2"
      Tab(1).Control(1)=   "cobManualMoveDelete"
      Tab(1).Control(2)=   "Label4(2)"
      Tab(1).Control(3)=   "Label4(1)"
      Tab(1).Control(4)=   "Label4(0)"
      Tab(1).Control(5)=   "Label3"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "第三步"
      TabPicture(2)   =   "frmManualArchive.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblDetail"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdStep3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.CommandButton cmdStep2 
         Caption         =   "下一步"
         Height          =   350
         Left            =   -71520
         TabIndex        =   12
         Top             =   3100
         Width           =   1100
      End
      Begin VB.ComboBox cobManualMoveDelete 
         Height          =   300
         ItemData        =   "frmManualArchive.frx":0054
         Left            =   -73920
         List            =   "frmManualArchive.frx":0061
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2520
         Width           =   2415
      End
      Begin VB.CommandButton cmdStep3 
         Caption         =   "开始归档"
         Height          =   350
         Left            =   3360
         TabIndex        =   7
         Top             =   3120
         Width           =   1100
      End
      Begin VB.CommandButton cmdStep1 
         Caption         =   "下一步"
         Height          =   350
         Left            =   -71500
         TabIndex        =   6
         Top             =   3100
         Width           =   1100
      End
      Begin VB.Frame Frame1 
         Caption         =   "选择存储设备"
         Height          =   1815
         Left            =   -74760
         TabIndex        =   2
         Top             =   1200
         Width           =   4335
         Begin VB.ComboBox cobDevice 
            Height          =   315
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1200
            Width           =   2520
         End
         Begin VB.OptionButton optManualSelectDevice 
            Caption         =   "手工指定存储设备"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   4
            Top             =   720
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.OptionButton optManualSelectDevice 
            Caption         =   "自动选择存储设备"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   3
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Label Label4 
         Caption         =   "3、只删除：不移动数据，仅仅删除源设备中的数据"
         Height          =   495
         Index           =   2
         Left            =   -74760
         TabIndex        =   15
         Top             =   1800
         Width           =   4215
      End
      Begin VB.Label Label4 
         Caption         =   "2、归档且删除：将数据从源设备移动到目的设备，然后删除源设备的数据"
         Height          =   495
         Index           =   1
         Left            =   -74760
         TabIndex        =   14
         Top             =   1320
         Width           =   4215
      End
      Begin VB.Label Label4 
         Caption         =   "1、只归档：将数据从源设备移动到目的设备，同时保留源设备的数据"
         Height          =   495
         Index           =   0
         Left            =   -74760
         TabIndex        =   13
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label Label3 
         Caption         =   "归档方式选择："
         Height          =   375
         Left            =   -74760
         TabIndex        =   10
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   "确认归档设置："
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblDetail 
         Caption         =   "源设备："
         Height          =   1335
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "选择存储设备：在默认情况下，系统将自动选择一个最佳归档目的设备"
         Height          =   360
         Left            =   -74760
         TabIndex        =   1
         Top             =   600
         Width           =   4260
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmManualArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bArchive As Boolean          '标识归档/反归档 .1---归档;0---反归档

Private Sub Command1_Click()
    
End Sub

Private Sub cmdStep1_Click()
    
    Me.sstabManualArchive.Tab = 1
End Sub

Private Sub Command3_Click()

End Sub

Private Sub cmdStep2_Click()
    Me.sstabManualArchive.Tab = 2
End Sub

Private Sub cmdStep3_Click()

    On Error GoTo errH
    '开始归档
    If Me.cmdStep3.Caption = "完成" Then
        Unload Me
    Else
        '向数据库归档作业表添加一条归档作业记录
        Dim strSQL As String
        Dim tmpset As ADODB.Recordset
        Dim lngJobNum As Long
        
        If Me.cobDevice.ListCount <= 0 Then
            MsgBox "没有目的设备，请检查设备配置。"
            Exit Sub
        End If
        Dim strDeviceNo As String
        strDeviceNo = Left(Me.cobDevice.List(Me.cobDevice.ListIndex), InStr(Me.cobDevice.List(Me.cobDevice.ListIndex), "-") - 1)
        
        strSQL = "select 影像归档作业_ID.nextval as JobID from dual"
        Set tmpset = gcnOracle.Execute(strSQL)
        lngJobNum = tmpset!JobID
        strSQL = "Insert into 影像归档作业 (编码,名称,执行时间,源设备,目的设备,指定设备,是否迁移,是否删除,自动备份,执行过程) values (" & _
                 lngJobNum & ",'手动归档" & lngJobNum & "',to_date('" & Date & " " & Time & "','yyyy-mm-dd hh24:mi:ss') " & _
                 IIf(bArchive, ",'1','2','", ",'2','1','") & IIf(Me.optManualSelectDevice(1).Value = True, "", strDeviceNo) & "' ," & _
                 IIf(Me.cobManualMoveDelete.ListIndex <> 2, 1, 0) & "," & _
                 IIf(Me.cobManualMoveDelete.ListIndex <> 0, 1, 0) & ",0,0)"
        gcnOracle.Execute (strSQL)
        
        zl9comlib.ZlCommFun.ShowFlash
        '执行归档作业
        frmMain.funcdoArchiveJob lngJobNum
        zl9comlib.ZlCommFun.StopFlash
        Me.cmdStep3.Caption = "完成"
        frmMain.ShowChkRecord
    End If
    Exit Sub
errH:
    zl9comlib.ZlCommFun.StopFlash
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optManualSelectDevice_Click(Index As Integer)
    If Index = 1 Then       '自动选择备份设备
        Me.cobDevice.Enabled = False
    Else                    '手工指定备份设备
        Me.cobDevice.Enabled = True
    End If
End Sub

Private Sub sstabManualArchive_Click(PreviousTab As Integer)
    If Me.sstabManualArchive.Tab = 2 Then              '完成页面
        Me.lblDetail.Caption = "源设备：" & IIf(bArchive = True, "主存储设备", "辅助存储设备") & vbCrLf & vbCrLf & _
                               "目的设备：" & IIf(Me.optManualSelectDevice(1).Value = True, "自动选择", Me.cobDevice.Text) & vbCrLf & vbCrLf & _
                               "归档：" & IIf(Me.cobManualMoveDelete.ListIndex <> 2, "是", "否") & vbCrLf & vbCrLf & _
                               "删除：" & IIf(Me.cobManualMoveDelete.ListIndex <> 0, "是", "否")
        
    End If
End Sub

 
