VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4395
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   6480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   4395
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picHos 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2750
      Left            =   1650
      ScaleHeight     =   2745
      ScaleWidth      =   4830
      TabIndex        =   10
      Top             =   0
      Width           =   4835
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   1005
      TabIndex        =   8
      Top             =   3720
      Width           =   6135
   End
   Begin VB.Image imgPic 
      Height          =   2745
      Left            =   150
      Picture         =   "frmSplash.frx":5D0A2
      Top             =   420
      Width           =   1260
   End
   Begin VB.Label lblTag 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   4785
      TabIndex        =   9
      Top             =   1815
      Width           =   615
   End
   Begin VB.Label LblProductName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   1515
      TabIndex        =   7
      Top             =   1350
      Width           =   4650
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "使用权属于："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1650
      TabIndex        =   6
      Top             =   2430
      Width           =   1080
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "产品开发商："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1650
      TabIndex        =   5
      Top             =   3255
      Width           =   1080
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "技术支持商："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1650
      TabIndex        =   4
      Top             =   2835
      Width           =   1080
   End
   Begin VB.Label lblGrant 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2745
      TabIndex        =   3
      Top             =   2430
      Width           =   90
   End
   Begin VB.Label lbl技术支持商 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2745
      TabIndex        =   2
      Top             =   2835
      Width           =   90
   End
   Begin VB.Label lbl开发商 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2745
      TabIndex        =   1
      Top             =   3255
      Width           =   90
   End
   Begin VB.Image ImgIndicate 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   165
      Picture         =   "frmSplash.frx":5D923
      Stretch         =   -1  'True
      Top             =   3390
      Width           =   720
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "警告：本软件受软件保护法和软件使用许可证保护。未经授权许可，任何人不得复制、销售及解密此软件，否则将承担全部法律责任。"
      Height          =   465
      Left            =   1095
      TabIndex        =   0
      Top             =   3825
      Width           =   5550
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowSplash()
    Dim StrUnitName As String
    Dim intCount As Integer
    Dim objPic As IPictureDisp
    
    Load frmSplash '强制装入，以方便后面的程序直接卸载
    '由注册表中获取用户注册相关信息,如果用户单位名称不为空,则显示闪现窗体
    StrUnitName = GetSetting("ZLSOFT", "注册信息", "单位名称", "")
    If StrUnitName <> "" Then
        lblGrant = StrUnitName
        LblProductName = GetSetting("ZLSOFT", "注册信息", "产品全称", "")
        If Len(LblProductName.Caption) > 10 Then
            LblProductName.FontSize = 15.75 '三号
        Else
            LblProductName.FontSize = 21.75  '二号
        End If
        gstrSysName = GetSetting("ZLSOFT", "注册信息", "产品名称", "") & "软件"
        lbl技术支持商 = GetSetting("ZLSOFT", "注册信息", "技术支持商", "")
        lbl开发商 = ""
        StrUnitName = GetSetting("ZLSOFT", "注册信息", "开发商", "")
    
        gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "")
        gstrUltimatetag = GetSetting("ZLSOFT", "注册信息", "产品系列", "")
        lblTag = gstrUltimatetag
        
        Call ApplyOEM_Picture(ImgIndicate, "Picture")
        Call ApplyOEM_Picture(imgPic, "PictureB")
        If gobjFile Is Nothing Then Set gobjFile = New FileSystemObject
        If gstrAppsoft = "" Then
            gstrAppsoft = App.Path
            If gblnInIDE Then
                gstrAppsoft = "C:\APPSOFT"
            End If
        End If

        If gobjFile.FileExists(gstrAppsoft & "\附加文件\logo_login.jpg") Then
            Set objPic = LoadPicture(gstrAppsoft & "\附加文件\logo_login.jpg")
            picHos.Visible = True
            picHos.Height = IIf(objPic.Height < 2745, objPic.Height, 2745) '183像素
            picHos.Width = IIf(objPic.Width < 4845, objPic.Width, 4845) '323像素
            picHos.PaintPicture objPic, 0, 0, picHos.Width, picHos.Height
        Else
            picHos.Visible = False
        End If
        If Trim$(lbl技术支持商.Caption) = "" Then
            Label1.Visible = False
            lbl技术支持商.Visible = False
        Else
            Label1.Visible = True
            lbl技术支持商.Visible = True
        End If
        
        If Trim(StrUnitName) = "" Then
            Label3.Visible = False
            lbl开发商.Visible = False
        Else
            Label3.Visible = True
            lbl开发商.Visible = True
            lbl开发商.Caption = ""
            For intCount = 0 To UBound(Split(StrUnitName, ";"))
                lbl开发商.Caption = lbl开发商.Caption & Split(StrUnitName, ";")(intCount) & vbCrLf
            Next
        End If
        frmSplash.Show
        DoEvents
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    gdtStart = 0
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    gdtStart = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strCode As String
       
    If gcnOracle.State = adStateOpen Then
        If gblnCreate Then
            gstrUltimatetag = ""
            gstrProductTitle = gobjRegister.zlRegInfo("产品标题")
            If gstrProductTitle <> "" Then
                If InStr(gstrProductTitle, "-") > 0 Then
                    If Split(gstrProductTitle, "-")(1) = "Ultimate" Then
                        gstrUltimatetag = "旗舰版"
                    ElseIf Split(gstrProductTitle, "-")(1) = "Professional" Then
                        gstrUltimatetag = "专业版"
                    End If
                End If
            End If
            gstrProductTitle = Split(gstrProductTitle, "-")(0)
            
            gstrProductName = gobjRegister.zlRegInfo("产品简名")
            gstrDevelopers = gobjRegister.zlRegInfo("产品开发商", , -1)
            gstrSustainer = gobjRegister.zlRegInfo("技术支持商", , -1)
            gstrWebSustainer = gobjRegister.zlRegInfo("支持商简名")
            gstrWebURL = gobjRegister.zlRegInfo("支持商URL")
            gstrWebEmail = gobjRegister.zlRegInfo("支持商MAIL")
            
            '将用户注册相关信息写入注册表,供下次启动时显示
            SaveSetting "ZLSOFT", "注册信息", "产品全称", gstrProductTitle
            SaveSetting "ZLSOFT", "注册信息", "产品名称", gstrProductName
            SaveSetting "ZLSOFT", "注册信息", "技术支持商", gstrSustainer
            SaveSetting "ZLSOFT", "注册信息", "开发商", gstrDevelopers
            SaveSetting "ZLSOFT", "注册信息", "WEB支持商简名", gstrWebSustainer
            SaveSetting "ZLSOFT", "注册信息", "WEB支持EMAIL", gstrWebEmail
            SaveSetting "ZLSOFT", "注册信息", "WEB支持URL", gstrWebURL
            SaveSetting "ZLSOFT", "注册信息", "单位名称", gobjRegister.zlRegInfo("单位名称", , -1)
            SaveSetting "ZLSOFT", "注册信息", "产品系列", gstrUltimatetag
            
        End If
    End If
End Sub

'立即退出
Private Sub ImgIndicate_Click()
    gdtStart = 0
End Sub

Private Sub LblProductName_Click()
    gdtStart = 0
End Sub

Private Sub Label1_Click()
    gdtStart = 0
End Sub

Private Sub Label2_Click()
    gdtStart = 0
End Sub

Private Sub Label3_Click()
    gdtStart = 0
End Sub

Private Sub lblWarning_Click()
    gdtStart = 0
End Sub
