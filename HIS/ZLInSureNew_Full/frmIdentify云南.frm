VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIdentify云南 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保病人就诊类型选择"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIdentify云南.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fra就诊类型 
      Caption         =   "就诊类型"
      Height          =   5145
      Left            =   210
      TabIndex        =   3
      Top             =   1440
      Width           =   4305
      Begin MSComctlLib.ListView lvw疾病 
         Height          =   3975
         Left            =   270
         TabIndex        =   6
         Top             =   1020
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   7011
         View            =   1
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         _Version        =   393217
         Icons           =   "ils32"
         SmallIcons      =   "ils32"
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "编码"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.OptionButton opt疾病 
         Caption         =   "特殊(&B)"
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   5
         Top             =   540
         Width           =   1215
      End
      Begin VB.OptionButton opt疾病 
         Caption         =   "普通(&A)"
         Height          =   285
         Index           =   0
         Left            =   1380
         TabIndex        =   4
         Top             =   540
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   1
         Left            =   270
         Picture         =   "frmIdentify云南.frx":000C
         Top             =   360
         Width           =   240
      End
   End
   Begin VB.Frame fra卡类型 
      Caption         =   "医保卡类型"
      Height          =   1125
      Left            =   210
      TabIndex        =   0
      Top             =   150
      Width           =   4305
      Begin VB.OptionButton opt医保卡 
         Caption         =   "磁卡(&2)"
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   1
         Left            =   2910
         TabIndex        =   2
         Top             =   570
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton opt医保卡 
         Caption         =   "IC卡(&1)"
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   0
         Left            =   1410
         TabIndex        =   1
         Top             =   570
         Width           =   1215
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   0
         Left            =   300
         Picture         =   "frmIdentify云南.frx":0E4E
         Top             =   390
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   405
      Left            =   4770
      TabIndex        =   8
      Top             =   870
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   405
      Left            =   4770
      TabIndex        =   7
      Top             =   270
      Width           =   1305
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   4980
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIdentify云南.frx":1C90
            Key             =   "Disease"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmIdentify云南"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum card类型
    cardIC = 0
    card磁卡 = 1
End Enum

Dim mint险类 As Integer
Dim mstr卡类型 As String
Dim mstr疾病编码 As String
Dim mlng疾病ID As Long
Dim mblnOK As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If opt疾病(1).Value = True Then
        If lvw疾病.ListItems.Count = 0 Then
            MsgBox "请在医保病种管理程序中增加中心认可的病种。", vbInformation, gstrSysName
            Exit Sub
        End If
    
        If lvw疾病.SelectedItem Is Nothing Then
            MsgBox "请选择疾病类型。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    mstr卡类型 = IIf(opt医保卡(cardIC).Value = True, cardIC, card磁卡)
    If opt疾病(0).Value = True Then
        mlng疾病ID = 0
        mstr疾病编码 = ""
    Else
        mlng疾病ID = Mid(lvw疾病.SelectedItem.Key, 2)
        mstr疾病编码 = lvw疾病.SelectedItem.SubItems(1)
    End If
    
    '保存使用的卡类型
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "卡类型", mstr卡类型
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKey1 Or KeyCode = vbKeyF1 Or KeyCode = vbKeyNumpad1 Then
        opt医保卡(cardIC) = 0
    ElseIf KeyCode = vbKey2 Or KeyCode = vbKeyF2 Or KeyCode = vbKeyNumpad2 Then
        opt医保卡(card磁卡) = 1
    End If
End Sub

Public Function GetIdentifyMode(ByVal intInsure As Integer, ByVal bytType As Byte, str卡类型 As String, lng疾病ID As Long, str疾病编码 As String) As Boolean
'功能：获得身份验证的模式
'参数：bytType     0-门诊，1-住院，2-仅身份验证
'      str卡类型   0-IC卡,1-磁卡
'      lng疾病ID   0-普通就诊,否则为疾病编码
'返回：成功为True
    Dim bln特殊就诊 As Boolean
    Dim rsTemp As New ADODB.Recordset, lst As ListItem
    
    mblnOK = False
    mint险类 = intInsure
    mstr卡类型 = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "卡类型", "1") '缺省为磁卡
    mlng疾病ID = 0
    mstr疾病编码 = ""
    
    '根据注册表信息设置前一次使用的卡类型
    opt医保卡(IIf(mstr卡类型 = "0", 0, 1)).Value = True
    
    '根据身份验证类型，显示疾病信息
'    If bytType = 0 Or bytType = 1 Then
        '门诊与住院都可能要选择疾病
'        gstrSQL = "select 参数值 from 保险参数 where 险类=" & mint险类 & " and 中心=0 and 参数名='支持慢性病、特种病'"
'        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
'        If rsTemp.EOF = False Then
'            If rsTemp("参数值") = "1" Then
                '要进行特殊病选择
                bln特殊就诊 = True
'            End If
'        End If
'    End If
    
    If bln特殊就诊 = False Then
        '不允许使用特殊就诊，目前只有普通帐户验证时不需读疾病
        opt疾病(1).Enabled = False
    Else
        '本院支持使用特殊病种
        If bytType = 0 Then
            '门诊，使用慢性病与特种病
            gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
                    " From 保险病种 A where A.类别 in (0,1,2) and A.险类=" & mint险类
        Else
            '住院，使用普通病
            'Modified by ZYB 2004-10-12 昆明
            '-------------------------------
            gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
                    " From 保险病种 A where A.类别 in (0,1,2) and A.险类=" & mint险类 & IIf(mint险类 = TYPE_云南建水, "", " And A.编码 IN ('0094','0093')")
        End If
        
        Call OpenRecordset(rsTemp, "医保身份验证")
        Do Until rsTemp.EOF
            Set lst = lvw疾病.ListItems.Add(, "K" & rsTemp("ID"), rsTemp("名称"), "Disease", "Disease")
            lst.SubItems(1) = rsTemp("编码")
            
            rsTemp.MoveNext
        Loop
    End If
    
    frmIdentify云南.Show vbModal
    GetIdentifyMode = mblnOK
    If mblnOK = True Then
        str卡类型 = mstr卡类型
        If mint险类 = TYPE_云南建水 Then str卡类型 = 3
        lng疾病ID = mlng疾病ID
        str疾病编码 = mstr疾病编码
    End If
End Function

Private Sub Form_Load()
    If mint险类 = TYPE_云南建水 Then
        opt医保卡(1).Visible = False
        opt医保卡(0).Caption = "建水卡"
    End If
    opt医保卡(0).Value = True
End Sub

Private Sub fra卡类型_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub lvw疾病_DblClick()
'担心操作员出现误操作
'    Call cmdOK_Click
End Sub

Private Sub lvw疾病_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    lvw疾病.Drag 0
End Sub

Private Sub lvw疾病_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If Not lvw疾病.HitTest(x, y) Is Nothing Then
            lvw疾病.Drag 1
        End If
    End If
End Sub

Private Sub opt疾病_Click(Index As Integer)
    lvw疾病.Enabled = (opt疾病(1).Value = True)
    
    If lvw疾病.Enabled = False Then
        lvw疾病.BackColor = &H8000000F '按钮表面
    Else
        lvw疾病.BackColor = &H80000005 '窗口背景
    End If
End Sub

Private Sub opt医保卡_Click(Index As Integer)

End Sub
