VERSION 5.00
Begin VB.UserControl UsrHome 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   6780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9315
   KeyPreview      =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   9315
   Begin zl9NewQuery.ctlPicture picHome 
      Height          =   3825
      Left            =   1080
      TabIndex        =   4
      Top             =   315
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   6747
      Border          =   0
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "联系方式:"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   570
      TabIndex        =   3
      Top             =   5370
      Width           =   1080
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "医院等级:"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   585
      TabIndex        =   2
      Top             =   4995
      Width           =   1080
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "医院地址:"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   570
      TabIndex        =   1
      Top             =   4635
      Width           =   1080
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "医院名称:"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   585
      TabIndex        =   0
      Top             =   4275
      Width           =   1080
   End
End
Attribute VB_Name = "UsrHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)

Public Sub InitLoad()
    Dim W As Single
    Dim H As Single
    Dim strFileName As String
    Dim vCountX As Long
    Dim vCountY As Long
    Dim i As Long
    Dim j As Long
    Dim X1 As Single
    Dim Y1 As Single
    Dim picObj As StdPicture
    
    Call InitCommon(gcnOracle)
            
'    picHome.Visible = False
    On Error GoTo errHand
    
    lbl(0).Visible = True
    lbl(1).Visible = True
    lbl(2).Visible = True
    lbl(3).Visible = True
            

    If Val(GetPara("关闭主页上的医院信息显示", "0")) = 1 Then
        lbl(0).Visible = False
        lbl(1).Visible = False
        lbl(2).Visible = False
        lbl(3).Visible = False
    End If

    

    '读取主页设置信息,背景、图片和宣传标语
    gstrSQL = "select 插图序号 from 咨询段落目录 B where B.段落序号=1 and B.页面序号=0"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "主页面")
    picHome.Border = 0
    If gRs.BOF = False Then
        picHome.Tag = GetFileName(IIf(IsNull(gRs!插图序号), 0, gRs!插图序号), W, H)
        W = IIf(W > 7695, 7695, W)
        H = IIf(H > 5220, 5220, H)
        Call picHome.ShowPictureByFile(picHome.Tag, True, W, H)
        Call UserControl_Resize
    End If
    
    gstrSQL = "Select B.类型,B.名称,B.宽度,B.高度 from 咨询页面目录 A,咨询图片元素 B where A.页面背景=B.序号 and A.页面序号=0"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "主页面")
    If gRs.BOF = False Then
        strFileName = IIf(IsNull(gRs!名称), "", App.Path & "\图形\" & gRs!名称 & IIf(gRs!类型 <> 2, ".pic", ".swf"))
        If Dir(strFileName) <> "" Then

            W = IIf(IsNull(gRs!宽度), 0, gRs!宽度) * Screen.TwipsPerPixelX
            H = IIf(IsNull(gRs!高度), 0, gRs!高度) * Screen.TwipsPerPixelY
                        
            vCountX = Int(UserControl.Width / W) + 1
            vCountY = Int(UserControl.Height / H) + 1

            Set picObj = VB.LoadPicture(strFileName)
            If Not picObj Is Nothing Then
                Select Case GetPara("背景显示模式", "平铺")
                    Case "拉伸"
                        UserControl.PaintPicture picObj, X1, Y1, UserControl.ScaleWidth, UserControl.ScaleHeight
                    Case "居中"
                        X1 = (UserControl.ScaleWidth - W) / 2
                        Y1 = (UserControl.ScaleHeight - H) / 2
                        UserControl.PaintPicture picObj, X1, Y1, W, H
                    Case Else
                        For j = 1 To vCountY
                            For i = 1 To vCountX
                                X1 = (i - 1) * W
                                Y1 = (j - 1) * H
                                UserControl.PaintPicture picObj, X1, Y1, W, H
                            Next
                        Next
                End Select
            End If
        End If
    End If
'    picHome.Visible = True
'    DoEvents
    
    If lbl(0).Visible Then
        lbl(0).Caption = "医院名称:" & GetUnitName
        lbl(1).Caption = "医院地址:" & GetUnitInfo("地址")
        lbl(2).Caption = "医院等级:" & GetUnitInfo("医院等级")
        lbl(3).Caption = "联系方式:"
        If GetUnitInfo("电话") <> "" Then lbl(3).Caption = lbl(3).Caption & "  电话 " & GetUnitInfo("电话")
        If GetUnitInfo("联系人") <> "" Then lbl(3).Caption = lbl(3).Caption & "  联系人 " & GetUnitInfo("联系人")
    End If
    
    Exit Sub
errHand:
    If ErrCenter() = -1 Then Resume
    Call SaveErrLog
End Sub

Private Sub picHome_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    picHome.Left = (UserControl.Width - picHome.Width) / 2
    picHome.Top = (UserControl.Height - picHome.Height - (lbl(0).Height + 120) * 4 - 300) / 2
    picHome.Top = IIf(picHome.Top < 300, 300, picHome.Top)
    
    lbl(0).Left = 300
    lbl(0).Top = picHome.Top + picHome.Height + 300 + 120
    lbl(0).Width = UserControl.Width - lbl(0).Left
    
    lbl(1).Left = 300
    lbl(1).Top = lbl(0).Top + lbl(0).Height + 120
    lbl(1).Width = UserControl.Width - lbl(1).Left
    
    lbl(2).Left = 300
    lbl(2).Top = lbl(1).Top + lbl(1).Height + 120
    lbl(2).Width = UserControl.Width - lbl(2).Left
    
    lbl(3).Left = 300
    lbl(3).Top = lbl(2).Top + lbl(2).Height + 120
    lbl(3).Width = UserControl.Width - lbl(3).Left
    
    
End Sub

Public Property Let Enabled(ByVal vData As Boolean)
    UserControl.Enabled = vData
End Property
