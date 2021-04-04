VERSION 5.00
Begin VB.Form frm补录病情_重庆渝北 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病种选择"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   Icon            =   "frm补录病情_重庆渝北.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra1 
      Caption         =   "入院病种信息"
      Height          =   1395
      Left            =   60
      TabIndex        =   7
      Top             =   1065
      Width           =   7980
      Begin VB.TextBox Txt疾病 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1095
         TabIndex        =   2
         Top             =   240
         Width           =   6780
      End
      Begin VB.TextBox Txt疾病 
         Height          =   300
         Index           =   1
         Left            =   1095
         TabIndex        =   4
         Top             =   615
         Width           =   6780
      End
      Begin VB.TextBox Txt疾病 
         Height          =   300
         Index           =   2
         Left            =   1095
         TabIndex        =   6
         Top             =   975
         Width           =   6780
      End
      Begin VB.Label lbl病情 
         AutoSize        =   -1  'True
         Caption         =   "病种(&1)"
         Height          =   180
         Index           =   0
         Left            =   420
         TabIndex        =   1
         Top             =   300
         Width           =   630
      End
      Begin VB.Label lbl病情 
         AutoSize        =   -1  'True
         Caption         =   "病种一(&2)"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   690
         Width           =   810
      End
      Begin VB.Label lbl病情 
         AutoSize        =   -1  'True
         Caption         =   "病种二(&3)"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   810
      End
   End
   Begin VB.Frame fra2 
      Caption         =   "出院病种信息"
      Height          =   1395
      Left            =   60
      TabIndex        =   8
      Top             =   2595
      Width           =   7965
      Begin VB.TextBox Txt疾病 
         Height          =   300
         Index           =   3
         Left            =   1080
         TabIndex        =   10
         Top             =   270
         Width           =   6780
      End
      Begin VB.TextBox Txt疾病 
         Height          =   300
         Index           =   4
         Left            =   1080
         TabIndex        =   12
         Top             =   645
         Width           =   6780
      End
      Begin VB.TextBox Txt疾病 
         Height          =   300
         Index           =   5
         Left            =   1080
         TabIndex        =   14
         Top             =   1005
         Width           =   6780
      End
      Begin VB.Label lbl病情 
         AutoSize        =   -1  'True
         Caption         =   "病种(&4)"
         Height          =   180
         Index           =   3
         Left            =   405
         TabIndex        =   9
         Top             =   330
         Width           =   630
      End
      Begin VB.Label lbl病情 
         AutoSize        =   -1  'True
         Caption         =   "病种一(&5)"
         Height          =   180
         Index           =   4
         Left            =   225
         TabIndex        =   11
         Top             =   720
         Width           =   810
      End
      Begin VB.Label lbl病情 
         AutoSize        =   -1  'True
         Caption         =   "病种二(&6)"
         Height          =   180
         Index           =   5
         Left            =   225
         TabIndex        =   13
         Top             =   1110
         Width           =   810
      End
   End
   Begin VB.Frame fra 
      Height          =   105
      Index           =   0
      Left            =   15
      TabIndex        =   18
      Top             =   720
      Width           =   8475
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5790
      TabIndex        =   16
      Top             =   4335
      Width           =   1100
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7050
      TabIndex        =   17
      Top             =   4335
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   105
      Left            =   -150
      TabIndex        =   15
      Top             =   4110
      Width           =   8715
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   165
      Picture         =   "frm补录病情_重庆渝北.frx":000C
      Stretch         =   -1  'True
      Top             =   210
      Width           =   510
   End
   Begin VB.Label lblPatient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "姓名：渝北医保    卡号：01234567    "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   210
      Left            =   960
      TabIndex        =   0
      Top             =   450
      Width           =   7275
   End
End
Attribute VB_Name = "frm补录病情_重庆渝北"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mblnStart As Boolean
Private mintInsure As Integer
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mbln补录 As Boolean
Private Sub Txt疾病_Change(Index As Integer)
    Txt疾病(Index).Tag = ""
End Sub
Private Sub Txt疾病_GotFocus(Index As Integer)
        '医保要求门诊诊断必须输入
        zlControl.TxtSelAll Txt疾病(Index)
End Sub

Private Sub Txt疾病_KeyPress(Index As Integer, KeyAscii As Integer)
  Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim strLike As String, str性别 As String
    Dim StrInput As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Txt疾病(Index).Text = "" Or Txt疾病(Index).Tag <> "" Then
            Call zlCommFun.PressKey(vbKeyTab) '允许不输入
        Else
            StrInput = UCase(Txt疾病(Index).Text)
            gstrSQL = "" & _
            "   Select id, 编码, 名称, 支付类别, 助记码, 病种结算办法, 经办构构代码 " & _
            "   From 医保病种目录" & _
            "   Where 性质=2 and (" & zlCommFun.GetLike("", "编码", StrInput) & " Or " & _
                        zlCommFun.GetLike("", "名称", StrInput) & " Or " & _
                        zlCommFun.GetLike("", "助记码", StrInput) & ") "
            
            Dim sngLeft As Single, sngTop As Single
            
            If Index >= 3 Then
                sngLeft = Txt疾病(Index).Left + Me.Left + fra2.Left
                sngTop = Txt疾病(Index).Top + Me.Top + fra2.Top
            Else
                sngLeft = Txt疾病(Index).Left + Me.Left + fra1.Left
                sngTop = Txt疾病(Index).Top + Me.Top + fra1.Top
            End If
            
            Set rsTmp = zlDatabase.ShowSelect(Me, gstrSQL, 0, "病种编码选择", , , , , , True, _
                sngLeft, sngTop, Txt疾病(Index).Height, blnCancel, , True)
            If Not rsTmp Is Nothing Then
                Txt疾病(Index).Text = "(" & rsTmp!编码 & ")" & rsTmp!名称
                Txt疾病(Index).Tag = Nvl(rsTmp!ID)
                lbl病情(Index).Tag = Nvl(rsTmp!编码)
                If Index < 5 Then
                   If Txt疾病(Index + 1).Enabled Then Txt疾病(Index + 1).SetFocus
                Else
                   If cmdOK.Enabled Then cmdOK.SetFocus
                End If
            Else
                If Not blnCancel Then
                    MsgBox "没有找到匹配的病种编码。", vbInformation, gstrSysName
                End If
                Call Txt疾病_GotFocus(Index)
                Txt疾病(Index).SetFocus
            End If
        End If
    Else
        zlControl.TxtCheckKeyPress Txt疾病(Index), KeyAscii, m文本式
    End If
End Sub
Public Sub Load历史看病信息()
    '功能:加载历史医保病人的看病信息
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0: On Error GoTo errHand:
    
    gstrSQL = "Select 性质,序号,病情ID,病情编码,病情 from 病人诊断情况_91 where  病人id=" & mlng病人ID & IIf(mlng主页ID = 0, " and 主页id is null ", " and 主页id=" & mlng主页ID) & " and 性质 IN (1,2)"
    Call OpenRecordset_OtherBase(rsTemp, "获取诊断情况", gstrSQL, gcnOracle_CQYB)
    
    With rsTemp
        Do While Not .EOF
            If Val(Nvl(!性质)) = 1 Then
                Select Case Nvl(!序号, 0)
                Case 1
                    Txt疾病(0).Text = IIf(IsNull(!病情编码), "", "(" & Nvl(!病情编码) & ")") & Nvl(!病情)
                    Txt疾病(0).Tag = Val(Nvl(!病情ID))
                    lbl病情(0).Tag = Nvl(!病情编码)
                Case 2
                    Txt疾病(1).Text = IIf(IsNull(!病情编码), "", "(" & Nvl(!病情编码) & ")") & Nvl(!病情)
                    Txt疾病(1).Tag = Val(Nvl(!病情ID))
                    lbl病情(1).Tag = Nvl(!病情编码)
                Case 3
                    Txt疾病(2).Text = IIf(IsNull(!病情编码), "", "(" & Nvl(!病情编码) & ")") & Nvl(!病情)
                    Txt疾病(2).Tag = Val(Nvl(!病情ID))
                    lbl病情(2).Tag = Nvl(!病情编码)
                    
                End Select
            Else
                Select Case Nvl(!序号, 0)
                Case 1
                    Txt疾病(3).Text = IIf(IsNull(!病情编码), "", "(" & Nvl(!病情编码) & ")") & Nvl(!病情)
                    Txt疾病(3).Tag = Val(Nvl(!病情ID))
                    lbl病情(3).Tag = Nvl(!病情编码)
                    
                Case 2
                    Txt疾病(4).Text = IIf(IsNull(!病情编码), "", "(" & Nvl(!病情编码) & ")") & Nvl(!病情)
                    Txt疾病(4).Tag = Val(Nvl(!病情ID))
                    lbl病情(4).Tag = Nvl(!病情编码)
                Case 3
                    Txt疾病(5).Text = IIf(IsNull(!病情编码), "", "(" & Nvl(!病情编码) & ")") & Nvl(!病情)
                    Txt疾病(5).Tag = Val(Nvl(!病情ID))
                    lbl病情(5).Tag = Nvl(!病情编码)
                End Select
            End If
           .MoveNext
        Loop
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub






Private Sub cmdOK_Click()
    
     '必须选择病情信息
    If Trim(Txt疾病(0).Text) = "" Then
        MsgBox "请为该参保病人选择入院病情！", vbInformation, gstrSysName
        If Txt疾病(0).Enabled Then Txt疾病(0).SetFocus
        Exit Sub
    End If
    If Trim(Txt疾病(3).Text) = "" Then
        MsgBox "请为该参保病人选择出院病情！", vbInformation, gstrSysName
        If Txt疾病(3).Enabled Then Txt疾病(3).SetFocus
        Exit Sub
    End If
    
    '保存病种
    Err = 0: On Error GoTo errHand
    'gcnOracle.BeginTrans
    gcnOracle_CQYB.BeginTrans
    '保存入院病情
    Call Get病情信息(False)
    Call Save病情信息(mlng病人ID, mlng主页ID, 1)
    
    Call Get病情信息(True)
    Call Save病情信息(mlng病人ID, mlng主页ID, 2)
    
    '1.保存到保险帐户
    gstrSQL = "ZL_保险帐户_更新信息(" & mlng病人ID & "," & mintInsure & ",'病情ID','''" & g病人身份_重庆渝北.病情ID & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "疾病ID")
    gstrSQL = "ZL_保险帐户_更新信息(" & mlng病人ID & "," & mintInsure & ",'病情1ID','" & IIf(g病人身份_重庆渝北.病情1ID = 0, "NULL", g病人身份_重庆渝北.病情1ID) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "疾病ID")
    gstrSQL = "ZL_保险帐户_更新信息(" & mlng病人ID & "," & mintInsure & ",'病情2ID','" & IIf(g病人身份_重庆渝北.病情2ID = 0, "NULL", g病人身份_重庆渝北.病情2ID) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "疾病ID")
    gstrSQL = "ZL_保险帐户_更新信息(" & mlng病人ID & "," & mintInsure & ",'病情1','''" & g病人身份_重庆渝北.病情名称 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "疾病ID")
    gstrSQL = "ZL_保险帐户_更新信息(" & mlng病人ID & "," & mintInsure & ",'病情2','''" & g病人身份_重庆渝北.病情名称1 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "疾病ID")
    gstrSQL = "ZL_保险帐户_更新信息(" & mlng病人ID & "," & mintInsure & ",'病情3','''" & g病人身份_重庆渝北.病情名称2 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "疾病ID")
   ' gcnOracle.CommitTrans
    gcnOracle_CQYB.CommitTrans
    mblnOK = True
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
  '  gcnOracle.RollbackTrans
    gcnOracle_CQYB.CRollbackTrans
End Sub

Private Sub Get病情信息(ByVal bln出院 As Boolean)
    '功能:获取病情信息
    Dim i As Integer
    i = IIf(bln出院, 3, 0)
    
    g病人身份_重庆渝北.病情ID = Val(Txt疾病(i).Tag)
    If g病人身份_重庆渝北.病情ID <> 0 Then
        g病人身份_重庆渝北.病情编码 = Trim(lbl病情(i).Tag)
        g病人身份_重庆渝北.病情名称 = Replace(Trim(Txt疾病(i).Text), "(" & Trim(lbl病情(i).Tag) & ")", "", 1, 1)
    Else
        g病人身份_重庆渝北.病情编码 = ""
        g病人身份_重庆渝北.病情名称 = Trim(Txt疾病(i).Text)
    End If
    i = i + 1
    g病人身份_重庆渝北.病情1ID = Val(Txt疾病(i).Tag)
    If g病人身份_重庆渝北.病情1ID <> 0 Then
        g病人身份_重庆渝北.病情编码1 = Trim(lbl病情(i).Tag)
        g病人身份_重庆渝北.病情名称1 = Replace(Trim(Txt疾病(i).Text), "(" & Trim(lbl病情(i).Tag) & ")", "", 1, 1)
    Else
        g病人身份_重庆渝北.病情编码1 = ""
        g病人身份_重庆渝北.病情名称1 = Trim(Txt疾病(i).Text)
    End If

    i = i + 1
    g病人身份_重庆渝北.病情2ID = Val(Txt疾病(i).Tag)
    If g病人身份_重庆渝北.病情2ID <> 0 Then
        g病人身份_重庆渝北.病情编码2 = Trim(lbl病情(i).Tag)
        g病人身份_重庆渝北.病情名称2 = Replace(Trim(Txt疾病(i).Text), "(" & Trim(lbl病情(i).Tag) & ")", "", 1, 1)
    Else
        g病人身份_重庆渝北.病情编码2 = ""
        g病人身份_重庆渝北.病情名称2 = Trim(Txt疾病(i).Text)
    End If
End Sub

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If Not mblnStart Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = " Select B.姓名,A.卡号,A.医保号 " & _
              " From 保险帐户 A,病人信息 B " & _
              " Where A.病人ID=B.病人ID And A.病人ID=[1] And A.险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取该病人的基本信息", mlng病人ID, mintInsure)
    If rsTemp.EOF Then
        mblnStart = False
        MsgBox "医保病人不存在!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    lblPatient.Caption = "姓名：" & Nvl(rsTemp!姓名) & Space(4) & "卡号：" & Nvl(rsTemp!卡号) & Space(4) & "个人编号：" & Nvl(rsTemp!医保号)
    
    Call Load历史看病信息
    
    Err = 0: On Error Resume Next
    fra1.Enabled = mbln补录
    Txt疾病(0).Enabled = mbln补录
    Txt疾病(1).Enabled = mbln补录
    Txt疾病(2).Enabled = mbln补录
    
    mblnStart = True
End Sub

Public Function ShowSelect(ByVal intinsure As Integer, _
    ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional bln补录 As Boolean = False) As Boolean
    '选择病人的入院病情及出院病情，同时将病人本次住院的相关信息显示出来
    '更新保险帐户的病情ID（入院病情）及出院病情，并将入院病情及出院病情编码返回给调用模块
    
    mblnOK = False
    mintInsure = intinsure
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mbln补录 = bln补录
    Me.Show 1
    ShowSelect = mblnOK
End Function


