VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDescribe 
   BackColor       =   &H80000005&
   Caption         =   "装卸管理"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5445
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmDescribe.frx":0000
   ScaleHeight     =   4275
   ScaleWidth      =   5445
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ils32 
      Left            =   270
      Top             =   2970
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDescribe.frx":04F9
            Key             =   "K01"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDescribe.frx":114B
            Key             =   "K02"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDescribe.frx":1A25
            Key             =   "K03"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDescribe.frx":22FF
            Key             =   "K04"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDescribe.frx":2BD9
            Key             =   "K05"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "装卸管理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   200
      TabIndex        =   1
      Top             =   100
      Width           =   960
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  'Transparent
      Height          =   3735
      Left            =   930
      TabIndex        =   0
      Top             =   645
      Width           =   4140
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   240
      Picture         =   "frmDescribe.frx":34B3
      Top             =   645
      Width           =   480
   End
End
Attribute VB_Name = "frmDescribe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstr编号 As String    '窗体编号

Private Sub Form_Load()
    Select Case mstr编号
        Case "01"
            lblTitle.Caption = "装卸管理"
            lblMain.Caption = "安装或拆卸各个应用系统的数据服务器。" & _
                vbCrLf & vbCrLf & "由于需要对所有者授予一些特殊权限（如Select on sys.v_$session、Select on sys.v_$parameter、Select on sys.dba_role_privs、Execute on sys.dbms_sql），因此只有具有这些权限及其GRANT选项的DBA用户才能运行装卸管理。" & _
                vbCrLf & vbCrLf & "由于数据库系统的物理设计对系统的运行效率有较大的影响，因此在系统安装之前，请参照相关资料文件，并根据主机系统的具体情况正确规划。" & _
                vbCrLf & vbCrLf & "由于装卸过程涉及数据文件的增删，因此请保证在没有其他用户使用的情况下进行。"
        Case "02"
            lblTitle.Caption = "数据管理"
            lblMain.Caption = "对指定应用系统的数据存储调整、逻辑备份与恢复、大数据装载文本生成与装载等操作。" & _
                vbCrLf & vbCrLf & "由于数据相关数据管理的操作，需要消耗较大的系统资源，因此以上操作尽量安排在系统较为空闲的时间进行，以降低对相关联机事务操作的影响；" & _
                vbCrLf & vbCrLf & "部分数据管理操作(数据导入、数据清除)将进行系统数据的彻底清除，请务必在此之前保证已经存在安全的数据备份。" & _
                vbCrLf & vbCrLf & "由于导出文件(Export)、数据导入(Import)、数据装载(Load)等是使用数据库自身功能，请保证数据库相关命令能正确执行；" & _
                vbCrLf & vbCrLf & "由于数据导入与装载未检查版本的合法性，请确认与实际情况符合，以保证正常的操作。"
        Case "03"
            lblTitle.Caption = "运行管理"
            lblMain.Caption = "完成系统现行状态的监控、历史日志的查看和未来运行状态的设置。" & _
                vbCrLf & vbCrLf & "后台作业是利用数据库作业功能实现的自动化措施，如需要使用后台作业，请调整数据库的相关init参数(如JOB_QUEUE_PROCESSES、JOB_QUEUE_INTERVAL)；" & _
                vbCrLf & vbCrLf & "同样，也建议在系统较空闲时执行后台作业，以减少和其他任务的资源竞争。" & _
                vbCrLf & vbCrLf & "错误日志和运行日志的记录，需要占用一定的数据空间和资源，请注意经常清理历史日志数据；如果系统已经较长时间稳定运行，可以通过选项关闭对日志的记录。"
        Case "04"
            lblTitle.Caption = "权限管理"
            lblMain.Caption = "建立系统的功能角色，创建、授予用户权限并指定用户身份，调整重组菜单。" & _
                vbCrLf & vbCrLf & "系统角色受数据库系统init参数max_enabled_roles的限制，如果需要划分建立较多的角色，请修改该数据参数后重新启动数据库服务生效；" & _
                vbCrLf & vbCrLf & "为了更好地控制权限，本系统的用户也就是数据库系统的用户，但密码经过了转换，请不要试图使用本系统建立的用户直接使用数据库的工具联接进入数据库；" & _
                vbCrLf & vbCrLf & "自己建立的角色缺省自己也具有该角色，但一旦你不需要该角色而取消后，你将再也看不到该角色的存在，除非是DBA用户。"
        Case "05"
            lblTitle.Caption = "专项工具"
            lblMain.Caption = "使用报表工具可完成系统各种票据格式与输出内容的定义修改。" & _
                vbCrLf & vbCrLf & "该工具采用面向对象的策略，独特的图元定制方式（图形元素点选描绘），精确定制票据与报表，可随心所欲地调整票据的纸张特性(大小、类型)、输出格式（字体、颜色、排列），并可立即预览打印。" & _
                vbCrLf & vbCrLf & vbCrLf & vbCrLf & "使用函数工具完成各系统数据传递函数的管理，包括函数文本及其参数向导的定义、修改与设置。" & _
                vbCrLf & vbCrLf & "数据传递函数是本软件各应用系统间相互抽选传递数据的重要方式，使整个应用成为一个完整的整体；较多地应用于财务总帐自动凭证、成本效益核算和报表分析提取各应用系统的发生数据。"
    End Select
    Me.Caption = lblTitle.Caption
    imgMain.Picture = ils32.ListImages("K" & mstr编号).Picture
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With imgMain
        .Top = 700
        .Left = ScaleLeft + 200
    End With
    
    With lblMain
        .Top = imgMain.Top
        .Height = ScaleHeight - .Top * 2
        .Left = imgMain.Left * 2 + imgMain.Width
        .Width = ScaleWidth - .Left - imgMain.Left
    End With
End Sub

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口

End Sub
