VERSION 5.00
Begin VB.Form frmDiseaseStationSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "疾病报告范围设置"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   30
      TabIndex        =   4
      Top             =   525
      Width           =   5730
   End
   Begin VB.ListBox lstFiles 
      Height          =   1320
      Left            =   1080
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   960
      Width           =   3210
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   2385
      Width           =   5730
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2565
      TabIndex        =   1
      Top             =   2565
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3675
      TabIndex        =   0
      Top             =   2565
      Width           =   1100
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   195
      Picture         =   "frmDiseaseStationSet.frx":0000
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "设置本工作站管理的疾病报告文件。"
      Height          =   180
      Left            =   720
      TabIndex        =   6
      Top             =   150
      Width           =   3975
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblFiles 
      AutoSize        =   -1  'True
      Caption         =   "本工作站可管理文件(&F):"
      Height          =   180
      Left            =   90
      TabIndex        =   5
      Top             =   690
      Width           =   1980
   End
End
Attribute VB_Name = "frmDiseaseStationSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOk As Boolean

Public Function ShowMe(ByVal frmParent As Object, ByVal blnFiles As Boolean, ByRef strFiles As String) As Boolean
'功能：显示本窗体并提供用户设置
'参数： blnFiles,   是否允许设置文件
'       strFiles,   目前可管理的文件id列表
    Dim strSetFiles As String, strReturn As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngCount As Long
    Dim strSQL As String

    strSetFiles = Trim(gobjComlib.zlDatabase.GetPara("本工作站可管理文件", glngSys, 1278))
  On Error GoTo errHand
    strSQL = "Select Id, 编号, 名称 From 病历文件列表 Where 种类 = 5  Order By 编号"
    Set rsTemp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With rsTemp
        Me.lstFiles.Clear
        Do While Not .EOF
            '为支持新病，空格看不见，同时用于作分隔符
            Me.lstFiles.AddItem !编号 & "-" & !名称 & "                                   " & !ID
            Me.lstFiles.ItemData(Me.lstFiles.NewIndex) = !ID
            If InStr(1, "," & strSetFiles & ",", "," & !ID & ",") > 0 Then
                Me.lstFiles.Selected(Me.lstFiles.NewIndex) = True
            End If
            .MoveNext
        Loop
    End With
    
    Me.lstFiles.Enabled = blnFiles
    
    '显示窗体
    Me.Show vbModal, frmParent
    
    '返回处理
    If mblnOk Then
        If Me.lstFiles.Enabled Then
            strFiles = ""
            For lngCount = 0 To Me.lstFiles.ListCount - 1
                If Me.lstFiles.Selected(lngCount) Then
                    If IsNumeric(Split(lstFiles.List(lngCount), "                                   ")(1)) Then
                        strFiles = strFiles & "," & Me.lstFiles.ItemData(lngCount)
                    End If
                End If
            Next
            If strFiles <> "" Then strFiles = Mid(strFiles, 2)
            Call gobjComlib.zlDatabase.SetPara("本工作站可管理文件", strFiles, glngSys, 1278)
        End If
    End If
    ShowMe = mblnOk
    Unload Me
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Unload Me
End Function

Private Sub cmdCancel_Click()
    mblnOk = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim blnSelected As Boolean
    Dim lngCount As Long
   
    If Me.lstFiles.Enabled Then
        For lngCount = 0 To Me.lstFiles.ListCount - 1
            If Me.lstFiles.Selected(lngCount) Then
                blnSelected = True
                Exit For
            End If
        Next
        If Not blnSelected Then
            MsgBox "没有设置本工作站可管理的疾病报告文件！", vbExclamation, gstrSysName: Me.lstFiles.SetFocus: Exit Sub
        End If
    End If
    mblnOk = True
    Me.Hide
End Sub

Private Sub lstFiles_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call gobjComlib.zlCommFun.PressKey(vbKeyTab)
End Sub
