VERSION 5.00
Begin VB.Form frmPacsInterfaceVBSTest 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "功能验证"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txt 
      Height          =   270
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&S)"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "参数值"
      Height          =   180
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "参数名"
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   540
   End
   Begin VB.Label lab 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   540
   End
End
Attribute VB_Name = "frmPacsInterfaceVBSTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintParsCount As Integer '不重复的参数数量
Private mstrVBSOld As String
Private mstrVBSTest As String
Private mobjOwner As frmPacsInterfaceCfg
Private mintVBSTest As Integer
Public Function zlShowMe(ByVal strVBS As String, objOwner As frmPacsInterfaceCfg) As Integer
    Set mobjOwner = objOwner
    mstrVBSOld = strVBS
    
    Call LoadControlAndLayout
    
    Call Me.Show(1, objOwner)
    zlShowMe = mintVBSTest
End Function
Private Sub LoadControlAndLayout()
'根据mstrVBSOld中的预定义参数，动态生成控件并且布局
    Dim strTmp As String
    Dim intTMP As Integer
    Dim lngL As Long, lngT As Long, lngW As Long, lngH As Long
    Dim strParName As String
    Dim blHaveSamePar As Boolean '是否已经存在一样的参数
    Dim i As Integer
    
    mintParsCount = 0
    strTmp = mstrVBSOld
    While InStr(strTmp, "[[")
        
        '功能：提取参数名称
        strTmp = Mid(strTmp, InStr(strTmp, "[[") + 2)
        strParName = Mid(strTmp, 1, InStr(strTmp, "]]") - 1)
        blHaveSamePar = False
        
        For i = 0 To mintParsCount - 1
            If lab(i).Caption = strParName Then
                blHaveSamePar = True
                Exit For
            End If
        Next
        
        If mintParsCount > 0 Then
            '首先排除重复情况
            If Not blHaveSamePar Then
                mintParsCount = mintParsCount + 1
                
                Load lab(mintParsCount - 1)
                lab(mintParsCount - 1).Caption = strParName
                lab(mintParsCount - 1).AutoSize = True
                Call lab(mintParsCount - 1).Move(240, 240 + 360 * mintParsCount)
                lab(mintParsCount - 1).Visible = True
        
                Load txt(mintParsCount - 1)
                txt(mintParsCount - 1).Text = ""
                txt(mintParsCount - 1).tag = strParName
                Call txt(mintParsCount - 1).Move(1560, 240 + 360 * mintParsCount, 1215, 270)
                txt(mintParsCount - 1).Visible = True
                
            Call SetDefaltValue(strParName, False)
            End If
        Else
            lab(0).Caption = strParName
            txt(0).tag = strParName

            Call SetDefaltValue(strParName, True)
            mintParsCount = mintParsCount + 1
        End If
        
        strTmp = Mid(strTmp, InStr(strTmp, "]]") + 2)
    Wend
    
    Call cmdOK.Move(1080, lab(mintParsCount - 1).Top + lab(mintParsCount - 1).Height + 360)
    Call cmdCancel.Move(2160, lab(mintParsCount - 1).Top + lab(mintParsCount - 1).Height + 360)
    
    lngW = 3315
    lngL = mobjOwner.Left + (mobjOwner.Width - lngW) / 2
    lngH = cmdOK.Top + cmdOK.Height + 600
    lngT = mobjOwner.Top + (mobjOwner.Height - lngH) / 2
    
    Call Me.Move(lngL, lngT, lngW, lngH)
    
End Sub

Private Sub ReplacePars(ByVal strParName As String)
'替换参数内容
On Error GoTo ErrorHnad
    Dim strValue As String
    Dim i As Integer

    If strParName = "当前窗口句柄" Then
        mstrVBSTest = Replace(mstrVBSTest, "[[" & strParName & "]]", mobjOwner.hWnd)
    Else
        For i = 0 To mintParsCount - 1
            If txt(i).tag = strParName Then
                strValue = txt(i).Text
                Exit For
            End If
        Next
        mstrVBSTest = Replace(mstrVBSTest, "[[" & strParName & "]]", strValue)
    End If
    
    Exit Sub
ErrorHnad:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub cmdCancel_Click()
    mintVBSTest = 未测试
    Unload Me
End Sub

Private Sub CmdOK_Click()
'首先替换与定义参数，然后测试
On Error GoTo ErrorHnad
    mstrVBSTest = mstrVBSOld
    Call ReplacePars("用户名")
    Call ReplacePars("账号名")
    Call ReplacePars("系统号")
    Call ReplacePars("模块号")
    Call ReplacePars("科室ID")
    Call ReplacePars("病人ID")
    Call ReplacePars("医嘱ID")
    Call ReplacePars("检查号")
    Call ReplacePars("门诊号")
    Call ReplacePars("住院号")
    Call ReplacePars("身份证号")
    Call ReplacePars("影像类别")
    Call ReplacePars("当前窗口句柄")
    
    Call TestExecuteSub(mstrVBSTest)
    Exit Sub
ErrorHnad:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub


Private Function TestExecuteSub(ByVal strVBS As String) As Boolean
'调用vbs脚本实现功能
On Error GoTo ErrorHnad
    Dim objCall As Object
    Dim strTempVBS As String
    
    '创建脚本执行对象
    Set objCall = CreateObject("ScriptControl")
    objCall.Timeout = 60000
    objCall.Language = "vbscript"
    
    Call objCall.AddCode(strVBS)
    
    Call objCall.Run(Trim("ExcuteSub"))
    
    TestExecuteSub = True
    mintVBSTest = 通过
    Unload Me
    Exit Function
ErrorHnad:
    If err.Description <> "" Then MsgBox err.Description, vbExclamation, gstrSysName
End Function

Private Sub Form_Terminate()
    Set mobjOwner = Nothing
End Sub

Private Sub SetDefaltValue(ByVal strName As String, ByVal blnFirst As Boolean)
'填充参数初始值的处理
    Dim lngIndex As Long '(控件数组)
    
    If blnFirst Then
        lngIndex = 0
    Else
        lngIndex = mintParsCount - 1
    End If
    
    Select Case strName
        Case "用户名"
            txt(lngIndex).Text = UserInfo.姓名
             
        Case "账号名"
            txt(lngIndex).Text = UserInfo.用户名
                                
        Case "系统号"
            txt(lngIndex).Text = "100"
            
        Case "模块号"
            txt(lngIndex).Text = "1291"
            
        Case "科室ID"
            txt(lngIndex).Text = UserInfo.部门ID
        
        Case "病人ID"
            txt(lngIndex).Text = "1"
            
        Case "医嘱ID"
            txt(lngIndex).Text = "101"
            
        Case "检查号"
            txt(lngIndex).Text = "110"
            
        Case "门诊号"
            txt(lngIndex).Text = "1"
        
        Case "住院号"
            txt(lngIndex).Text = "110"
            
        Case "身份证号"
            txt(lngIndex).Text = "500105190001010000"
            
        Case "影像类别"
            txt(lngIndex).Text = "CT"
                                
        Case "当前窗口句柄"
            txt(lngIndex).Text = Me.hWnd
            txt(lngIndex).Enabled = False
    End Select

End Sub
