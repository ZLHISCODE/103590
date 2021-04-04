VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7065
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   8550
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdGetAllDev 
      Caption         =   "取检验仪器列表"
      Height          =   360
      Left            =   6675
      TabIndex        =   17
      Top             =   270
      Width           =   1650
   End
   Begin VB.CommandButton cmdUnAudit 
      Caption         =   "取消已审报告"
      Height          =   360
      Left            =   5235
      TabIndex        =   16
      Top             =   4200
      Width           =   1605
   End
   Begin VB.CommandButton cmdWritLIS 
      Caption         =   "写入已审报告"
      Height          =   360
      Left            =   3285
      TabIndex        =   15
      Top             =   4230
      Width           =   1605
   End
   Begin VB.CommandButton cmd取消核收 
      Caption         =   "取消核收"
      Height          =   360
      Left            =   6555
      TabIndex        =   14
      Top             =   3435
      Width           =   990
   End
   Begin VB.TextBox txt仪器ID 
      Height          =   270
      Left            =   5100
      TabIndex        =   12
      Text            =   "41"
      Top             =   3525
      Width           =   525
   End
   Begin VB.TextBox txt标本号 
      Height          =   270
      Left            =   5835
      TabIndex        =   11
      Text            =   "1"
      Top             =   3510
      Width           =   645
   End
   Begin VB.CommandButton cmdReg 
      Caption         =   "核收标本到LIS"
      Height          =   480
      Index           =   1
      Left            =   3285
      TabIndex        =   10
      Top             =   3360
      Width           =   1620
   End
   Begin VB.CommandButton cmdReg 
      Caption         =   "标为不可退费"
      Height          =   480
      Index           =   0
      Left            =   345
      TabIndex        =   9
      Top             =   3315
      Width           =   1620
   End
   Begin VB.CommandButton cmdItem 
      Caption         =   "3、提取申请明细"
      Height          =   480
      Left            =   300
      TabIndex        =   8
      Top             =   1410
      Width           =   1620
   End
   Begin VB.TextBox txtItem 
      Height          =   315
      Left            =   2085
      TabIndex        =   7
      Top             =   1485
      Width           =   1980
   End
   Begin VB.TextBox txtClinic 
      Height          =   315
      Left            =   2085
      TabIndex        =   6
      Top             =   930
      Width           =   1980
   End
   Begin VB.CommandButton cmdGetClinic 
      Caption         =   "2、提取申请内容"
      Height          =   480
      Left            =   300
      TabIndex        =   5
      Top             =   855
      Width           =   1620
   End
   Begin VB.TextBox txtInput 
      Height          =   315
      Left            =   2085
      TabIndex        =   4
      Top             =   390
      Width           =   1980
   End
   Begin VB.CommandButton cmdGetA 
      Caption         =   "1、提取检验申请"
      Height          =   480
      Left            =   300
      TabIndex        =   3
      Top             =   300
      Width           =   1620
   End
   Begin VB.CommandButton cmdSaveTiJian 
      Caption         =   "保存体检结果"
      Height          =   480
      Left            =   225
      TabIndex        =   2
      Top             =   4125
      Width           =   1620
   End
   Begin VB.CommandButton cmdDelRtf 
      Caption         =   "删除Rtf报告"
      Height          =   480
      Left            =   4035
      TabIndex        =   1
      Top             =   5565
      Width           =   1620
   End
   Begin VB.CommandButton cmdInsRtf 
      Caption         =   "保存Rtf报告"
      Height          =   480
      Left            =   1860
      TabIndex        =   0
      Top             =   5565
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "输入住院号等，"
      Height          =   180
      Left            =   4170
      TabIndex        =   18
      Top             =   390
      Width           =   1260
   End
   Begin VB.Label lblIDID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "仪器ID  标本号"
      Height          =   180
      Left            =   5115
      TabIndex        =   13
      Top             =   3225
      Width           =   1260
   End
   Begin VB.Line Line10 
      X1              =   2760
      X2              =   2505
      Y1              =   2655
      Y2              =   2895
   End
   Begin VB.Line Line9 
      X1              =   2130
      X2              =   2430
      Y1              =   2625
      Y2              =   2895
   End
   Begin VB.Line Line8 
      X1              =   2760
      X2              =   2610
      Y1              =   5400
      Y2              =   5490
   End
   Begin VB.Line Line7 
      X1              =   2430
      X2              =   2580
      Y1              =   5370
      Y2              =   5490
   End
   Begin VB.Line Line6 
      X1              =   2580
      X2              =   2580
      Y1              =   4980
      Y2              =   5490
   End
   Begin VB.Line Line4 
      Index           =   3
      X1              =   3810
      X2              =   3810
      Y1              =   4590
      Y2              =   4965
   End
   Begin VB.Line Line4 
      Index           =   2
      X1              =   1305
      X2              =   1305
      Y1              =   4575
      Y2              =   4950
   End
   Begin VB.Line Line5 
      X1              =   1320
      X2              =   3840
      Y1              =   4995
      Y2              =   4995
   End
   Begin VB.Line Line4 
      Index           =   1
      X1              =   3795
      X2              =   3795
      Y1              =   3825
      Y2              =   4200
   End
   Begin VB.Line Line4 
      Index           =   0
      X1              =   1350
      X2              =   1350
      Y1              =   3825
      Y2              =   4200
   End
   Begin VB.Line Line3 
      Index           =   1
      X1              =   3840
      X2              =   3840
      Y1              =   2940
      Y2              =   3315
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   1335
      X2              =   1335
      Y1              =   2940
      Y2              =   3285
   End
   Begin VB.Line Line2 
      X1              =   1335
      X2              =   3870
      Y1              =   2925
      Y2              =   2925
   End
   Begin VB.Line Line1 
      X1              =   2460
      X2              =   2460
      Y1              =   1905
      Y2              =   2925
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lis As New ZLLISInterface.clsLisInterface
Private mblnReg As Boolean

Private Sub cmdGetAllDev_Click()
    Dim strReturn As String, strErr As String
    strReturn = Lis.gGetAllDevice(strErr)
    If strReturn = "" Then
        MsgBox strErr
    Else
        MsgBox strReturn
    End If
End Sub

Private Sub cmdGetClinic_Click()
    Dim lngID As Long, strReturn As String
    lngID = Val(txtClinic)
    If lngID > 0 Then
        strReturn = Lis.gGetClinicItem(lngID)
        MsgBox strReturn
    End If
End Sub

Private Sub cmdGetA_Click()
    '提取病人申请
    Dim strInput As String, strReturn As String
    strInput = txtInput
    strReturn = Lis.gGetApplication(strInput)
    MsgBox strReturn
    
End Sub

Private Sub cmdInsRtf_Click()
    Dim strTmp As String
    Dim astrTmp() As String
    
    Call Lis.gInsertReport(CLng(txtClinic.Text), "c:\Temp\test.rtf", strTmp)

    MsgBox strTmp
End Sub

Private Sub cmdDelRtf_Click()
    Dim strTmp As String
    
    rsTmp = Lis.gDeleteReport(1380633)
    MsgBox rsTmp
End Sub

Private Sub cmdItem_Click()
    Dim lngID As Long, strReturn As String
    lngID = Val(txtItem)
    If lngID > 0 Then
        strReturn = Lis.gGetItemList(lngID)
        MsgBox strReturn
    End If
    
End Sub

Private Sub cmdReg_Click(Index As Integer)
    Dim lng仪器id As Long, lngID As Long, strInfo As String, str标本号 As String
    lng仪器id = Val(txt仪器ID)
    lngID = Val(txtClinic)
    str标本号 = Val(txt标本号)
    
    If Lis.gzlLisRegister(lng仪器id, lngID, str标本号, strInfo) = True Then
        MsgBox "核收成功！" & vbNewLine & strInfo
    Else
        MsgBox "核收失败！" & vbNewLine & strInfo
    End If
End Sub

Private Sub cmdSaveLIS_Click()

End Sub

Private Sub cmdSaveTiJian_Click()
       '诊治项目id;检验结果1;单位1;结果参1考;结果标志1|诊治项目id;检验结果2;单位2;结果参考2;结果标志2
    Dim strResult As String, varTmp As Variant
    Dim strErr As String, strInfo As String
    strResult = Lis.gTestResults(8252908, "检验技师", Format(Now, "yyyy-MM-dd HH:mm:ss"), _
                "548;3.72;X10E+12/L;3.5～5.5;|" & _
                "443;10.10;10E+9/L;4～10;偏高|" & _
                "721;107.00;g/L;110～160;偏低|" & _
                "554;0.37;%;.37～.49;偏低|" & _
                "648;98.1;fL;80～100;|" & _
                "550;28.8;pg;27～31;|" & _
                "551;293.20;g/L;320～360;偏低|" & _
                "735;216.00;10E+9/L;100～450;|" & _
                "610;16.1;%;20～40;偏低|" & _
                "500;6.9;％;3～8;|" & _
                "792;74.9;％;50～70;偏高|" & _
                "673;2.0;％;.5～5;|" & _
                "670;0.1;％;0～1;|" & _
                "611;1.63;10E+9/L;.8～4;|" & _
                "501;0.70;10E+9/L;.12～.8;|" & _
                "793;7.56;10E+9/L;2～7.5;偏高|" & _
                "674;0.20;10E+9/L;0～.5;|" & _
                "671;0.01;10E+9/L;0～.1;|" & _
                "544;0.15;％;.11～.16;|" & _
                "545;53.1;fL;37～54;|" & _
                "738;12.3;fL;9～17;|" & _
                "649;10.1;fL;8～13;|" & _
                "497;0.26;％;.13～.43;|" & _
                "734;0.22;％;.18～.22;")
    If strResult <> "" Then
        varTmp = Split(strResult, vbNewLine)
        For i = LBound(varTmp) To UBound(varTmp)
            If varTmp(i) Like "0|*" Then
                strErr = strErr & Mid(varTmp(i), 3) & vbNewLine
            ElseIf varTmp(i) Like "1|*" Then
                strInfo = strInfo & Mid(varTmp(i), 3) & vbNewLine
            End If
        Next
        MsgBox IIf(strErr <> "", "错误提示：" & vbNewLine & strErr, "") & IIf(strInfo <> "", "提示信息：" & vbNewLine & strInfo, "")
    Else
        MsgBox "OK"
    End If
End Sub

Private Sub cmdUnAudit_Click()
    Dim strErrinfo As String
    If Lis.gzlLisUnAudit(Val(txtClinic), strErrinfo) Then
        MsgBox "Ok" & strErrinfo
    Else
        MsgBox "失败" & strErrinfo
    End If
End Sub

Private Sub cmdWritLIS_Click()
    Dim lngID As Long, strItem As String, strErrinfo As String, blnOK As Boolean
    Dim varItem As Variant
    lngID = Val("" & txtClinic)
    If Val(txtItem) = 0 Then
        MsgBox "请输入组合项目ID！"
        txtItem.SetFocus
        Exit Sub
    End If
    strItem = Lis.gGetItemList(Val(txtItem))
    varItem = Split(strItem, "|")
    strItem = ""
    '这里是模拟数据，检验结果为序号，实际接口中需要改为具体的检验结果
    For i = LBound(varItem) To UBound(varItem)
        strItem = strItem & "|" & Split(varItem(i), "^")(0) & "^" & i
    Next
    blnOK = Lis.gZLLisInsterReport(lngID, strItem, strErrinfo)
    If blnOK Then
        MsgBox "完成" & strErrinfo
    Else
        MsgBox "失败" & vbNewLine & strErrinfo
    End If
    
End Sub

Private Sub cmd取消核收_Click()
    Dim strErr As String
    If Lis.gzlLisUnRegister(Val(txt采集ID), strErr) Then
        MsgBox "已取消" & strErr
    Else
        MsgBox "失败" & vbNewLine & strErr
    End If
End Sub

Private Sub Form_Activate()
    Dim ctl As Object
    For Each ctl In Me.Controls
        If TypeName(ctl) = "CommandButton" Then
            ctl.Enabled = mblnReg
        End If
    Next
End Sub

Private Sub Form_Load()
  If Lis.gOpenDataBase("txyy118", "zlhis", "aqa") = False Then
    mblnReg = False
  Else
    mblnReg = True
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Lis.gOraDataClose
End Sub
