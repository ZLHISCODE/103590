VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMarkMapEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病历标记图形编辑"
   ClientHeight    =   3300
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7305
   Icon            =   "frmFigureEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkAspectRatio 
      Caption         =   "锁定纵横比(&M)"
      Height          =   255
      Left            =   1305
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2430
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.TextBox txtHeight 
      Height          =   300
      Left            =   2340
      TabIndex        =   5
      Top             =   2070
      Width           =   795
   End
   Begin VB.TextBox txtWidth 
      Height          =   300
      Left            =   1290
      TabIndex        =   4
      Top             =   2070
      Width           =   795
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&P"
      Height          =   300
      Left            =   3870
      TabIndex        =   3
      Top             =   1650
      Width           =   375
   End
   Begin VB.CheckBox chkFitMode 
      Alignment       =   1  'Right Justify
      Caption         =   "适合大小(&F)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5985
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2970
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   1260
      Top             =   2745
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   $"frmFigureEdit.frx":058A
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2025
      TabIndex        =   7
      Top             =   2850
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3150
      TabIndex        =   8
      Top             =   2850
      Width           =   1100
   End
   Begin VB.TextBox txt简码 
      Height          =   300
      Left            =   1290
      TabIndex        =   2
      Top             =   1650
      Width           =   1350
   End
   Begin VB.TextBox txt名称 
      Height          =   300
      Left            =   1290
      TabIndex        =   1
      Top             =   1230
      Width           =   2955
   End
   Begin VB.TextBox txt编码 
      Height          =   300
      Left            =   1290
      TabIndex        =   0
      Top             =   825
      Width           =   795
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   570
      TabIndex        =   13
      Top             =   600
      Width           =   3735
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -45
      TabIndex        =   12
      Top             =   2745
      Width           =   4320
   End
   Begin zlRichEPR.ucCanvas Canvas 
      Height          =   2805
      Left            =   4410
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   90
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   4948
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "象素"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3195
      TabIndex        =   21
      Top             =   2130
      Width           =   360
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "×"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2115
      TabIndex        =   20
      Top             =   2130
      Width           =   180
   End
   Begin VB.Label lblSize 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "大小(&Z)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   585
      TabIndex        =   19
      Top             =   2130
      Width           =   630
   End
   Begin VB.Label lblPic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "图片:"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3330
      TabIndex        =   18
      Top             =   1710
      Width           =   450
   End
   Begin VB.Label lblColor 
      Caption         =   "颜色深度:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4410
      TabIndex        =   11
      Top             =   2970
      Width           =   2190
   End
   Begin VB.Label lbl简码 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "简码(&S)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   585
      TabIndex        =   17
      Top             =   1710
      Width           =   630
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   45
      Picture         =   "frmFigureEdit.frx":0647
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lblNote 
      Caption         =   "编辑将应用的统一的病历标记图形资源，设置其编码和命名，以便后续使用。"
      Height          =   345
      Left            =   585
      TabIndex        =   16
      Top             =   135
      Width           =   3660
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl名称 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "名称(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   585
      TabIndex        =   15
      Top             =   1290
      Width           =   630
   End
   Begin VB.Label lbl编码 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "编码(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   585
      TabIndex        =   14
      Top             =   885
      Width           =   630
   End
End
Attribute VB_Name = "frmMarkMapEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'简码：
'   1、上级程序通过本窗体ShowMe函数，将父窗体、编辑单据ID,编辑状态等信息传递进入本程序
'   2、编辑状态：由Me.tag存放，分别为"新增"、"修改"，由上级程序通过ShowMe传入
'---------------------------------------------------
Private mstrItemCode As String      '被编辑的项目编码，修改、查阅时由上级程序通过ShowMe传递进入,新增时为0，
Private mblnOK As Boolean           '是否完成编辑退出

'临时变量
Dim rsTemp As New ADODB.Recordset

'################################################################################################################
'-- 位图控制
Private WithEvents DIBFilter As cDIBFilter      ' DIB 滤镜对象(24 bpp)
Attribute DIBFilter.VB_VarHelpID = -1
Private WithEvents DIBDither As cDIBDither      ' DIB 抖动对象(1, 4, 8 bpp)
Attribute DIBDither.VB_VarHelpID = -1
Private DIBPal               As New cDIBPal     ' DIB 调色板对象 (1, 4, 8 bpp)
Private DIBSave              As New cDIBSave    ' Save 对象 (BMP)  (1, 4, 8, 24 bpp)
Private DIBbpp               As Byte            ' 当前颜色深度
Private WithEvents cPicEditor As cPictureEditor     ' 图片编辑对象
Attribute cPicEditor.VB_VarHelpID = -1
Private m_LastFilename As String                    ' 最后打开的图片位置
Private m_Temp As String                            ' 临时文件路径
Private m_AppID As Long
'-- GDI+
Private m_GDIpToken         As Long         ' 用于关闭 GDI+
Private mblnAdd As Boolean                  ' 是否是新增
Private Const MAX_PIXELS_SIZE As Long = 4000000
Private W As Long, chgW As Boolean
Private H As Long, chgH As Boolean
Private mfrmParent As Object

Public Function ShowMe(ByRef frmParent As Object, ByVal blnAdd As Boolean, Optional ByVal strItemCode As String, _
    Optional oDIB As cDIB) As String
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '返回：确定返回新增或修改的编码；取消返回""
    '---------------------------------------------------
    mblnAdd = blnAdd
    If mblnAdd Then
        Me.Tag = "新增"
    Else
        Me.Tag = "修改"
    End If
    
    Set mfrmParent = frmParent
    
    If mblnAdd = False Then
        Set Canvas.DIB = oDIB
        Me.Canvas.Resize
        If Me.Canvas.DIB.hDIB <> 0 Then
            W = oDIB.Width
            H = oDIB.Height
            txtWidth = W
            txtHeight = H
            txtWidth.Enabled = True
            txtHeight.Enabled = True
            chkAspectRatio.Enabled = True
            chkFitMode.Enabled = True
            lblSize.Enabled = True
            lblColor.Enabled = True
            lblColor = "颜色深度:24 位"
        Else
            txtWidth.Enabled = False
            txtHeight.Enabled = False
            chkAspectRatio.Enabled = False
            chkFitMode.Enabled = False
            lblSize.Enabled = False
            lblColor.Enabled = False
        End If
    End If
    
    mstrItemCode = strItemCode
    
    '提取信息
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select 编码,名称,简码 From 病历标记图形 Where 编码=[1]"
    Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, mstrItemCode)
    With rsTemp
        If .RecordCount > 0 Then
            Me.txt编码.Text = !编码: Me.txt名称.Text = !名称
            Me.txt简码.Text = IIf(IsNull(!简码), "", !简码)
        End If
        Me.txt编码.MaxLength = .Fields("编码").DefinedSize
        Me.txt名称.MaxLength = .Fields("名称").DefinedSize
        Me.txt简码.MaxLength = .Fields("简码").DefinedSize
    End With
    If Me.Tag = "新增" Then
        gstrSQL = "Select nvl(max(编码),'" & String(Me.txt编码.MaxLength, "0") & "') as 编码 From 病历标记图形"
        Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption)
        Me.txt编码.Text = Format(Val(rsTemp!编码) + 1, String(Me.txt编码.MaxLength, "0"))
    End If
    
    txtWidth.Enabled = (Me.Canvas.DIB.hDIB <> 0)
    txtHeight.Enabled = (Me.Canvas.DIB.hDIB <> 0)
    '显示窗体
    Me.Show vbModal, frmParent
    If mblnOK Then
        ShowMe = Trim(Me.txt编码.Text)
        Set frmParent.Canvas.DIB = Me.Canvas.DIB
        frmParent.Canvas.Resize
    Else
        ShowMe = ""
    End If
    Unload Me
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ShowMe = ""
End Function

Private Sub chkFitMode_Click()
    Canvas.FitMode = CBool(chkFitMode)
    Call Canvas.Resize
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False: Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim arySql() As String, lngSql As Long

    If Trim(Me.txt编码.Text) = "" Then MsgBox "请输入编码！", vbInformation, gstrSysName: Me.txt编码.SetFocus: Exit Sub
    If Len(Me.txt编码.Text) < Me.txt编码.MaxLength Then MsgBox "编码长度不足！", vbInformation, gstrSysName: Me.txt编码.SetFocus: Exit Sub
    If Trim(Me.txt名称.Text) = "" Then MsgBox "请输入名称！", vbInformation, gstrSysName: Me.txt名称.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt名称.Text), vbFromUnicode)) > Me.txt名称.MaxLength Then
        MsgBox "名称超长（最多" & Me.txt名称.MaxLength & "个字符或等长的汉字）！", vbInformation, gstrSysName: Me.txt名称.SetFocus: Exit Sub
    End If
    
    '数据保存
    gstrSQL = "'" & Trim(Me.txt编码.Text) & "','" & Trim(Me.txt名称.Text) & "','" & Trim(Me.txt简码.Text) & "'"
    If Me.Tag = "新增" Then
        If Me.Canvas.DIB.hDIB = 0 Then
            MsgBox "必须选择一幅图片！", vbOKOnly + vbInformation
            If cmdSelect.Visible And cmdSelect.Enabled Then cmdSelect.SetFocus
            Exit Sub
        Else
            gstrSQL = "ZL_病历标记图形_INSERT(" & gstrSQL & ")"
        End If
    Else
        '修改模式，要确保该病历图形还存在
        gstrSQL = "select count(*) from 病历标记图形 where 编码 =[1]"
        Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, mstrItemCode)
        If rsTemp(0) = 0 Then
            MsgBox "该图片已经被其他用户删除，保存失败！", vbOKOnly + vbInformation, "保存失败"
            Unload Me
            mfrmParent.zlRefLists
            Exit Sub
        End If
        rsTemp.Close
        gstrSQL = "ZL_病历标记图形_UPDATE('" & mstrItemCode & "'," & gstrSQL & ")"
    End If
        
    ReDim Preserve arySql(0 To 0)
    arySql(0) = gstrSQL

    '设置图片新尺寸
    If Me.Canvas.DIB.hDIB <> 0 Then
        If (txtWidth * txtHeight > MAX_PIXELS_SIZE) Then
            Call MsgBox(vbCrLf & _
                "图片大小超过最大允许范围(4M 象素)" & vbCrLf & vbCrLf & _
                "请减小图片尺寸！", vbExclamation)
            txtWidth.SelStart = 0: txtWidth.SelLength = Len(txtWidth): txtWidth.SetFocus
            Exit Sub
        End If
    
        Dim lngR As Long, strMsg As String
        If (txtWidth <> Me.Canvas.DIB.Width) Or (txtHeight <> Me.Canvas.DIB.Height) Then
            If txtWidth < 10 Or txtHeight < 10 Then
                lngR = MsgBox("注意：图片尺寸过小，一般象素低于10×10的图片将无法有效利用！" & vbCrLf & _
                    "是否继续？ 选“是”继续，选“否”取消。", vbYesNo + vbQuestion, gstrSysName)
            ElseIf txtWidth / Me.Canvas.DIB.Width < 0.5 And txtHeight / Me.Canvas.DIB.Height < 0.5 Then
                lngR = MsgBox("注意：图片尺寸小于原始尺寸的一半，这将导致图片质量损失，而且不可恢复！" & vbCrLf & _
                    "是否继续？ 选“是”继续，选“否”取消。", vbYesNo + vbQuestion, gstrSysName)
            ElseIf txtWidth / Me.Canvas.DIB.Width < 0.5 Then
                lngR = MsgBox("注意：图片宽度小于原始尺寸的一半，这将导致图片质量损失，而且不可恢复！" & vbCrLf & _
                    "是否继续？ 选“是”继续，选“否”取消。", vbYesNo + vbQuestion, gstrSysName)
            ElseIf txtHeight / Me.Canvas.DIB.Height < 0.5 Then
                lngR = MsgBox("注意：图片高度小于原始尺寸的一半，这将导致图片质量损失，而且不可恢复！" & vbCrLf & _
                    "是否继续？ 选“是”继续，选“否”取消。", vbYesNo + vbQuestion, gstrSysName)
            Else
                lngR = MsgBox("注意：改变图像尺寸将会损失图片质量，而且不可恢复！" & vbCrLf & _
                    "是否继续？ 选“是”继续，选“否”取消。", vbYesNo + vbQuestion, gstrSysName)
            End If
            If lngR = vbYes Then
                Screen.MousePointer = vbHourglass
                Call mGdIpEx.ScaleDIB(Me.Canvas.DIB, txtWidth, txtHeight, True)
                Call Me.Canvas.RemoveCropRectangle
                Call Me.Canvas.Resize
                Screen.MousePointer = vbNormal
            Else
                txtWidth = Me.Canvas.DIB.Width
                txtHeight = Me.Canvas.DIB.Height
                Exit Sub
            End If
        End If
    
        '同时保存位图
        Dim strFileName As String
        Screen.MousePointer = vbHourglass
        strFileName = m_Temp & "\R" & m_AppID & ".jpg"
        Call mGdIpEx.SaveDIB(Me.Canvas.DIB, strFileName, [ImageJPEG], 90)         '90%的图片质量，部分压缩
        
        If gobjFSO.FileExists(strFileName) Then
            If zlBlobSql(0, Trim(Me.txt编码.Text), strFileName, arySql) = False Then
                Screen.MousePointer = vbNormal
                MsgBox "标记图形保存失败", vbExclamation, gstrSysName
                Exit Sub
            End If
            gobjFSO.DeleteFile strFileName  '删除临时文件
        End If
        Screen.MousePointer = vbNormal
    End If
    
    '执行保存
    Err = 0: On Error GoTo errHand
    gcnOracle.BeginTrans
    For lngSql = LBound(arySql) To UBound(arySql)
        Call SQLTest(App.ProductName, Me.Caption, arySql(lngSql))
        gcnOracle.Execute arySql(lngSql), , adCmdStoredProc
        Call SQLTest
    Next
    gcnOracle.CommitTrans
    
    mblnOK = True: Me.Hide
    Exit Sub

errHand:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSelect_Click()
    Dim strFileName As String, bSuccess As Boolean, strTmp As String
    dlgThis.InitDir = m_LastFilename
    dlgThis.CancelError = True
    On Error GoTo LL
    dlgThis.ShowOpen
    
    strFileName = dlgThis.Filename
    If Trim(strFileName) <> "" Then
        '-- Create DIB
'        DoEvents
        Call pvSetDIBPicture(pvGetStdPicture(strFileName, bSuccess))
        
        If (bSuccess) Then
            m_LastFilename = strFileName
            W = Me.Canvas.DIB.Width
            H = Me.Canvas.DIB.Height
            txtWidth = W
            txtHeight = H
            lblColor = "颜色深度:" & DIBbpp & " 位"
            txtWidth.Enabled = True
            txtHeight.Enabled = True
            chkAspectRatio.Enabled = True
            chkFitMode.Enabled = True
            lblSize.Enabled = True
            lblColor.Enabled = True
        End If
    End If
    txtWidth.Enabled = (Me.Canvas.DIB.hDIB <> 0)
    txtHeight.Enabled = (Me.Canvas.DIB.hDIB <> 0)
LL:
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

'################################################################################################################
'## 功能：  画布相关函数
'################################################################################################################
Private Function pvGetStdPicture(ByVal sFileName As String, bSuccess As Boolean) As StdPicture
    On Error Resume Next
    If (pvGetExt(sFileName) = "png" Or pvGetExt(sFileName) = "tif") Then
        '-- Use GDI+ loading
        Set pvGetStdPicture = mGdIpEx.LoadPictureEx(sFileName)
      Else
        '-- Use VB LoadPicture
        Set pvGetStdPicture = LoadPicture(sFileName)
    End If
    
    '-- Is there an image ?
    bSuccess = Not (pvGetStdPicture Is Nothing)
    
    If (bSuccess = False) Then
        '-- Nothing loaded
        Call MsgBox("调入图片时发生意外错误！", vbExclamation)
    End If

    On Error GoTo 0
End Function
    
Private Sub pvSetDIBPicture(Image As StdPicture)
  Static lstW As Long
  Static lstH As Long

    If (Not Picture Is Nothing) Then

        '-- Save last DIB dimensions
        lstW = Canvas.DIB.Width
        lstH = Canvas.DIB.Height
        
        '-- Clear palette
        Call DIBPal.Clear
        
        DIBbpp = Canvas.DIB.CreateFromStdPicture(Image, DIBPal, DIBDither)
        
        '-- Select current depth mode
        Call pvSetPalMode(DIBbpp)
        
        '-- Remove Crop rectangle and resize canvas
        Call Canvas.RemoveCropRectangle
        With Canvas.DIB
            If (lstW <> .Width Or lstH <> .Height) Then
                Call Canvas.Resize
              Else
                Call Canvas.Repaint
            End If
        End With
    End If
End Sub

Private Sub pvSetPalMode(ByVal bpp As Long)
  Dim lIdxNew As Long
  Dim lIdxOld As Long
    
    Select Case bpp
        Case 1  '-- 2 colors / Black and White
            lIdxNew = IIf(DIBPal.IsGreyScale, 0, 4)
        Case 4  '-- 16 colors / 16 greys
            lIdxNew = IIf(DIBPal.IsGreyScale, 1, 5)
        Case 8  '-- 256 colors / 256 greys
            lIdxNew = IIf(DIBPal.IsGreyScale, 2, 6)
        Case 24 '-- True color
            lIdxNew = 8
        Case Else
            Exit Sub
    End Select
End Sub

Private Function pvGetExt(ByVal sFileName As String) As String
    pvGetExt = Mid(sFileName, Len(sFileName) - 2)
End Function

Private Sub Form_Load()
    '-----------------------------------------------------
    m_LastFilename = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "LastFilename", App.Path)
    Dim GpInput As GdiplusStartupInput
    '-- 调入 GDI+ Dll
    GpInput.GdiplusVersion = 1
    If (mGdIpEx.GdiplusStartup(m_GDIpToken, GpInput) <> [OK]) Then
        Call MsgBox("调入 GDI+ 出错，无法进行图片插入！请检查 GDI+ DLL 是否存在或者损坏！", vbInformation + vbOKOnly)
        Call Unload(Me)
        Exit Sub
    End If
    
    m_Temp = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
    m_AppID = Me.hWnd
    Set DIBFilter = New cDIBFilter
    Set DIBDither = New cDIBDither
    Set cPicEditor = New cPictureEditor
    
    Canvas.FitMode = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "LastFilename", m_LastFilename
    If mblnAdd Then Me.Canvas.DIB.Destroy    '修改模式是不能删除该DIB的！
    
    LockWindowUpdate 0
    UpdateWindow Me.hWnd
    ' Unload the GDI+ Dll
    Call mGdIpEx.GdiplusShutdown(m_GDIpToken)

    '-- Free objects
    Set DIBFilter = Nothing
    Set DIBDither = Nothing
    Set DIBPal = Nothing
    Set DIBSave = Nothing
    Set cPicEditor = Nothing
End Sub

Private Sub txtHeight_Change()
    txtHeight = Val(txtHeight)
    If (Val(txtHeight) = 0) Then
        If Me.Canvas.DIB.hDIB <> 0 Then
            txtHeight = "1"
        Else
            txtHeight = "0"
        End If
        txtHeight.SelLength = 1
    End If
    If (chkAspectRatio) Then
        If (Not chgW) Then
            chgH = True
            If Me.Canvas.DIB.hDIB <> 0 Then
                txtWidth = CInt(txtHeight / H * W)
            Else
                txtWidth = "0"
            End If
            chgH = False
        End If
    End If
End Sub

Private Sub txtHeight_GotFocus()
    txtHeight.SelStart = Len(txtHeight)
End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtWidth_Change()
    txtWidth = Val(txtWidth)
    If (Val(txtWidth) = 0) Then
        If Me.Canvas.DIB.hDIB <> 0 Then
            txtWidth = "1"
        Else
            txtWidth = "0"
        End If
        txtWidth.SelLength = 1
    End If
    If (chkAspectRatio) Then
        If (Not chgH) Then
            chgW = True
            If Me.Canvas.DIB.hDIB <> 0 Then
                txtHeight = CInt(txtWidth / W * H)
            Else
                txtHeight = "0"
            End If
            chgW = False
        End If
    End If
End Sub

Private Sub txtWidth_GotFocus()
    txtWidth.SelStart = Len(txtWidth)
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt编码_Change()
    ValidControlText txt编码
End Sub

Private Sub txt编码_GotFocus()
    Me.txt编码.SelStart = 0: Me.txt编码.SelLength = 100
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt编码_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt简码_Change()
    ValidControlText txt简码
End Sub

Private Sub txt简码_GotFocus()
    Me.txt简码.SelStart = 0: Me.txt简码.SelLength = 100
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt简码_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
        If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt名称_Change()
    ValidControlText txt名称
    If Me.Tag = "新增" Then
        Me.txt简码.Text = zlGetSymbol(Me.txt名称.Text, 0)
    End If
End Sub

Private Sub txt名称_GotFocus()
    Me.txt名称.SelStart = 0: Me.txt名称.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Me.txt简码.Text = zlGetSymbol(Me.txt名称.Text, 0)
        Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
