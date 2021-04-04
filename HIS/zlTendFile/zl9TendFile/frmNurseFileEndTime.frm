VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmNurseFileEndTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "结束时间"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4155
   Icon            =   "frmNurseFileEndTime.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSMask.MaskEdBox txt结束时间 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy-MM-dd HH:mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   0
      EndProperty
      Height          =   300
      Left            =   1350
      TabIndex        =   2
      Top             =   630
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   19
      Format          =   "yyyy-MM-dd HH:mm:ss"
      Mask            =   "####-##-## ##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1410
      TabIndex        =   4
      Top             =   1440
      Width           =   1155
   End
   Begin VB.CommandButton cmdCanCel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2640
      TabIndex        =   5
      Top             =   1440
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   -120
      TabIndex        =   3
      Top             =   1110
      Width           =   4365
   End
   Begin VB.Label lbl结束时间 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "结束时间"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   1
      Top             =   690
      Width           =   720
   End
   Begin VB.Label lbl当前文件 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "当前文件:"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   330
      Width           =   810
   End
End
Attribute VB_Name = "frmNurseFileEndTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngFile As Long
Private mblnOK As Boolean

Public Function ShowEditor(ByVal lngFile As Long) As Boolean
    On Error Resume Next
    mlngFile = lngFile
    mblnOK = False
    Me.Show 1
    ShowEditor = mblnOK
End Function

Private Sub cmdCanCel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHand
    
    If Not IsDate(txt结束时间.Text) Then
        MsgBox "请输入合法的时点！", vbInformation, gstrSysName
        txt结束时间.SetFocus
        Exit Sub
    End If
    If txt结束时间.Text < txt结束时间.Tag Then
        MsgBox "文件的结束时间不能小于护理数据的最后发生时间[" & txt结束时间.Tag & "]！", vbInformation, gstrSysName
        txt结束时间.SetFocus
        Exit Sub
    End If
    
    gstrSQL = "ZL_病人护理文件_ENDTIME(" & mlngFile & ",to_date('" & txt结束时间.Text & "','yyyy-MM-dd hh24:mi:ss'))"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "设定当前文件的结束时间")
    
    mblnOK = True
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Dim str结束时间 As String
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    '提取该病人与指定文件格式相同的文件,设定合并打印(只能按时间的先后顺序进行设定)
    
    gstrSQL = " Select 文件名称,结束时间 From 病人护理文件 Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取文件属性", mlngFile)
    txt结束时间.Text = Format(rsTemp!结束时间, "yyyy-MM-dd HH:mm:ss")
    Me.lbl当前文件.Caption = "当前文件：" & rsTemp!文件名称
    
    gstrSQL = " Select max(发生时间) AS 发生时间 from 病人护理数据 B Where B.文件ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取最后修改时间", mlngFile)
    txt结束时间.Tag = Format(rsTemp!发生时间, "yyyy-MM-dd HH:mm:ss")
    
    cmdOK.Enabled = (txt结束时间.Tag <> "")
End Sub

Private Sub txt结束时间_GotFocus()
    txt结束时间.SelStart = 0
    txt结束时间.SelLength = 20
End Sub
