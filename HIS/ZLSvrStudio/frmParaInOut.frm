VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmParaInOut 
   Caption         =   "参数导入导出选项"
   ClientHeight    =   3630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6060
   Icon            =   "frmParaInOut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6060
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraSplit 
      BackColor       =   &H80000012&
      Height          =   30
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   3000
      Width           =   6700
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   6060
      TabIndex        =   0
      Top             =   3015
      Width           =   6060
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   4515
         TabIndex        =   2
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   3360
         TabIndex        =   1
         Top             =   120
         Width           =   1100
      End
   End
   Begin VB.PictureBox picSet 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2925
      Left            =   0
      ScaleHeight     =   2925
      ScaleWidth      =   6075
      TabIndex        =   3
      Top             =   0
      Width           =   6075
      Begin VB.CommandButton cmdFile 
         Caption         =   "…"
         Height          =   255
         Left            =   5355
         TabIndex        =   11
         Top             =   203
         Width           =   300
      End
      Begin VB.Frame fraHos 
         Caption         =   "医院选项"
         Height          =   1215
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   5535
         Begin VB.OptionButton optCurHos 
            Caption         =   "本院"
            Height          =   255
            Left            =   360
            TabIndex        =   9
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optOtherHos 
            Caption         =   "他院"
            Height          =   180
            Left            =   1560
            TabIndex        =   8
            Top             =   390
            Width           =   855
         End
         Begin VB.Label lblInfo 
            Caption         =   "导出参数清单、本机私有参数设置、部门参数设置。"
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   360
            TabIndex        =   14
            Top             =   720
            Width           =   4935
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame fraSys 
         Caption         =   "系统选择"
         Height          =   735
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   5535
         Begin VB.OptionButton optAllSys 
            Caption         =   "所有系统"
            Height          =   180
            Left            =   1560
            TabIndex        =   6
            Top             =   390
            Width           =   1095
         End
         Begin VB.OptionButton optCurSys 
            Caption         =   "当前系统"
            Height          =   255
            Left            =   360
            TabIndex        =   5
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.TextBox txtFile 
         Height          =   300
         Left            =   1080
         MaxLength       =   256
         TabIndex        =   12
         Top             =   180
         Width           =   4575
      End
      Begin VB.Label lblPath 
         AutoSize        =   -1  'True
         Caption         =   "导入路径"
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   720
      End
   End
   Begin MSComDlg.CommonDialog cmmFile 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmParaInOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mpstType As ParaSetType '0-导入。1-导出，2-Excel转换XMl,3-XML合并
Public Enum ParaSetType
    PST_Imp = 0
    PST_Exp = 1
End Enum
Private mstrReturn As String
Private mlngSys As Long '主界面当前的系统

Public Function ShowMe(ByVal pstType As ParaSetType, Optional ByVal lngSys As Long) As String
    mpstType = pstType
    mlngSys = lngSys
    mstrReturn = ""
    Me.Show vbModal
    ShowMe = mstrReturn
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFile_Click()
    cmmFile.FileName = txtFile.Text
    If mpstType = PST_Imp Then
        cmmFile.Filter = "导入文件(*.xml)|*.xml"
        cmmFile.ShowOpen
    Else
        cmmFile.Filter = "导出文件(*.xml)|*.xml"
        cmmFile.ShowSave
    End If
    If cmmFile.FileName <> "" Then
        If mpstType = PST_Imp Then
            If CheckImpFile(cmmFile.FileName, True) Then
                txtFile.Text = cmmFile.FileName
            End If
        Else
            txtFile.Text = cmmFile.FileName
        End If
    End If
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    mstrReturn = txtFile.Text & "|" & IIf(optCurSys.value, 0, 1) & "|" & IIf(optCurHos.value, 0, 1)
    If txtFile.Text <> "" Then
        Call SaveSetting("ZLSOFT", "用户设置", "参数导入导出", txtFile.Text)
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strPath As String
    picSet.Visible = False
    Me.Width = 6300: Me.Height = 4200
    Select Case mpstType
        Case PST_Exp
            Me.Caption = "导出设置": lblPath.Caption = "导出文件"
            picSet.Visible = True
            If gblnInIDE Then
                txtFile.Text = gobjFile.GetFile("C:\APPSOFT\zlSvrStudio.exe").ParentFolder & "\ZLParasInfo.xml"
            Else
                txtFile.Text = GetSetting("ZLSOFT", "用户设置", "参数导入导出", App.Path & "\ZLParasInfo.xml")
            End If
        Case PST_Imp
            Me.Caption = "导入设置": lblPath.Caption = "导入文件"
            picSet.Visible = True
            If gblnInIDE Then
                txtFile.Text = gobjFile.GetFile("C:\APPSOFT\zlSvrStudio.exe").ParentFolder & "\ZLParasInfo.xml"
            Else
                txtFile.Text = GetSetting("ZLSOFT", "用户设置", "参数导入导出", App.Path & "\ZLParasInfo.xml")
            End If
            If Not CheckImpFile(txtFile.Text) Then
                txtFile.Text = ""
            End If
    End Select
    Call optCurHos_Click
End Sub

Private Sub Form_Resize()
    Me.Width = 6300
    Me.Height = 4200
End Sub

Private Sub optCurHos_Click()
    If mpstType = PST_Exp Then
        lblInfo.Caption = "导出参数清单、本机私有参数设置、部门参数设置。"
    Else
        lblInfo.Caption = "导入参数清单、本机私有参数设置、部门参数设置。"
    End If
End Sub

Private Sub optOtherHos_Click()
    If mpstType = PST_Exp Then
        lblInfo.Caption = "导出参数清单"
    Else
        lblInfo.Caption = "导入参数清单"
    End If
End Sub

Private Sub picBottom_Resize()
    cmdCancel.Left = picBottom.ScaleWidth - 120 - cmdCancel.Width
    cmdOK.Left = cmdCancel.Left - 60 - cmdOK.Width
End Sub

Private Function CheckImpFile(ByVal strFile As String, Optional blnMsg As Boolean) As Boolean
'功能：检查导入文件
'参数：strFile=参数文件
'          blnMsg=是否弹出消息提示
'返回：是否检查通过
    Dim rsParas As ADODB.Recordset, rsDBSys As ADODB.Recordset, rsComInfo As ADODB.Recordset
    Dim lngSys As Long, blnDetial As Boolean
    
    On Error GoTo errH
    If Dir(strFile) = "" Then Exit Function
    Set rsParas = New ADODB.Recordset
    '获取所有数据
    rsParas.Open strFile, , adOpenStatic, adLockOptimistic, adCmdFile
    
    optOtherHos.Enabled = True: optCurHos.Enabled = True
    optCurSys.Enabled = True: optAllSys.Enabled = True
    rsParas.Filter = "类型 = -99" '配置信息
    If rsParas.EOF Then
        If blnMsg Then MsgBox "该参数文件中未发现有效的配置信息，无法导入！", vbInformation, gstrSysName
        Exit Function
    End If
    blnDetial = Val(rsParas!私有) <> 0: lngSys = Val(rsParas!参数号)
    '判断是否有可导入的系统
    rsParas.Filter = "类型=-9"
    Set rsDBSys = GetALLPars(-9)
    '名称 参数名, 版本号 参数值, User 缺省值,To_Char(Sysdate, 'yyyy-mm-dd HH24:mi:ss') 影响控制说明
    Set rsComInfo = GetCompareRec(rsDBSys, rsParas, "系统", "参数值")
    rsComInfo.Filter = "State=0 OR State=2"
    If rsComInfo.RecordCount = 0 Then '没有可以导入的系统
        If blnMsg Then MsgBox "该参数文件中未发现可导入的系统！", vbInformation, gstrSysName
        Exit Function
    End If
    If Not blnDetial Then '该文件未包含部门参数详情，本机私有参数详情，只能选择他院导入
        optOtherHos.value = True
        optOtherHos.Enabled = False: optCurHos.Enabled = False
    End If
    If lngSys <> -1 Then
        If lngSys <> mlngSys Then '是其他系统的参数文件
            optAllSys.value = True
        Else '是当前系统的参数文件
            optCurSys.value = True
        End If
        optOtherHos.Enabled = False: optCurHos.Enabled = False
    Else '多个系统的参数文件
        rsComInfo.Filter = "State<>-1 And MainKey='" & mlngSys & "'" '查看导入文件中是否有当前系统的参数
        If rsComInfo.EOF Then '不存在当前系统的参数，不能选择当前系统
            optAllSys.value = True
            optOtherHos.Enabled = False: optCurHos.Enabled = False
        End If
    End If
    CheckImpFile = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Function
