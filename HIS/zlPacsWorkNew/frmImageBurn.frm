VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\Azl9PacsControl\zl9PacsControl.vbp"
Begin VB.Form frmImageBurn 
   Caption         =   "图像刻录"
   ClientHeight    =   7980
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13200
   Icon            =   "frmImageBurn.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   13200
   StartUpPosition =   3  '窗口缺省
   Begin zl9PacsControl.ZLScrollBar pbState 
      Height          =   255
      Left            =   8180
      TabIndex        =   29
      Top             =   7680
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   450
      Appearance      =   0
      AutoRedraw      =   -1  'True
      ScaleHeight     =   255
      ScaleWidth      =   3480
      ScaleLeft       =   0
      ScaleTop        =   0
      ScaleMode       =   1
      BackColor       =   14737632
      Hwnd            =   1839266
      EndColor        =   65280
      ShpMoveVisible  =   0   'False
      AllowMouseChange=   0
      AutoShowBlock   =   0   'False
   End
   Begin VB.Frame framBurn 
      Height          =   777
      Left            =   120
      TabIndex        =   20
      Top             =   6600
      Width           =   12975
      Begin VB.CommandButton cmdExit 
         Caption         =   "退 出(&E)"
         Height          =   400
         Left            =   11760
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdBurn 
         Caption         =   "刻 录&B)"
         Height          =   400
         Left            =   10560
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chkContainBurnStation 
         Caption         =   "包含CD观片工具"
         Height          =   255
         Left            =   8880
         TabIndex        =   25
         Top             =   278
         Width           =   1575
      End
      Begin VB.TextBox txtVolumeName 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   7080
         TabIndex        =   8
         Top             =   240
         Width           =   1665
      End
      Begin VB.ComboBox cbxBurnSpeed 
         Height          =   300
         ItemData        =   "frmImageBurn.frx":076A
         Left            =   4080
         List            =   "frmImageBurn.frx":076C
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   1905
      End
      Begin VB.ComboBox cbxDeviceName 
         Height          =   300
         ItemData        =   "frmImageBurn.frx":076E
         Left            =   1080
         List            =   "frmImageBurn.frx":0770
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   1905
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "卷标名称："
         Height          =   180
         Left            =   6120
         TabIndex        =   24
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "刻录速度："
         Height          =   180
         Left            =   3120
         TabIndex        =   23
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "刻录设备："
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   300
         Width           =   900
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   120
      ScaleHeight     =   5535
      ScaleWidth      =   12975
      TabIndex        =   18
      Top             =   960
      Width           =   12975
      Begin zl9PacsControl.ucSplitter ucSplitter1 
         Height          =   5535
         Left            =   3735
         TabIndex        =   21
         Top             =   0
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   9763
         DBClickType     =   2
         SplitLevel      =   3
         Control1Name    =   "ufgBurnData"
         Control2Name    =   "picImage"
      End
      Begin VB.PictureBox picImage 
         Height          =   5535
         Left            =   3870
         ScaleHeight     =   5475
         ScaleWidth      =   9045
         TabIndex        =   26
         Top             =   0
         Width           =   9105
         Begin zl9PacsControl.ucSplitPage ucPage 
            Height          =   330
            Left            =   0
            TabIndex        =   27
            Top             =   5160
            Width           =   5955
            _ExtentX        =   10504
            _ExtentY        =   582
            PageCount       =   0
            PageRecord      =   9
         End
         Begin DicomObjects.DicomViewer DViewer 
            Height          =   5055
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   9075
            _Version        =   262147
            _ExtentX        =   16007
            _ExtentY        =   8916
            _StockProps     =   35
            BackColor       =   0
         End
      End
      Begin zl9PACSWork.ucFlexGrid ufgBurnData 
         Height          =   5535
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   9763
         DefaultCols     =   ""
         BackColor       =   12648447
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
      End
   End
   Begin VB.Frame framQuery 
      Height          =   700
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   12975
      Begin VB.CommandButton cmdCustomQuery 
         Caption         =   "自定义查询"
         Height          =   375
         Left            =   11640
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "使用自定义查询"
         Top             =   200
         Width           =   1215
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查  询(&Q)"
         Height          =   375
         Left            =   10560
         TabIndex        =   32
         Top             =   200
         Width           =   1000
      End
      Begin VB.ComboBox cbxDeviceType 
         Height          =   300
         ItemData        =   "frmImageBurn.frx":0772
         Left            =   9360
         List            =   "frmImageBurn.frx":0774
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   7080
         TabIndex        =   2
         Top             =   240
         Width           =   1160
      End
      Begin VB.TextBox txtEndStudyNum 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   5880
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtStartStudyNum 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   5040
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   2640
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   62259203
         CurrentDate     =   38082.9993055556
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   960
         TabIndex        =   0
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   62259203
         CurrentDate     =   38082
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "设备类型："
         Height          =   180
         Left            =   8400
         TabIndex        =   17
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "姓名："
         Height          =   180
         Left            =   6600
         TabIndex        =   16
         Top             =   285
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "～"
         Height          =   180
         Left            =   5685
         TabIndex        =   15
         Top             =   315
         Width           =   180
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "检 查 号："
         Height          =   180
         Left            =   4200
         TabIndex        =   14
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "～"
         Height          =   180
         Left            =   2440
         TabIndex        =   13
         Top             =   315
         Width           =   180
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "检查日期："
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Width           =   900
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   31
      Top             =   7620
      Width           =   13200
      _ExtentX        =   23283
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   1764
            Picture         =   "frmImageBurn.frx":0776
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3704
            MinWidth        =   3704
            Text            =   "光盘可用大小："
            TextSave        =   "光盘可用大小："
            Key             =   "AvailableCapacity"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3704
            MinWidth        =   3704
            Text            =   "光盘预留大小："
            TextSave        =   "光盘预留大小："
            Key             =   "ReserveCapacity"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4409
            MinWidth        =   4409
            Text            =   "待刻录文件大小："
            TextSave        =   "待刻录文件大小："
            Key             =   "FileCapacity"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6326
            MinWidth        =   2
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label labState 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   135
      TabIndex        =   30
      Top             =   7440
      Width           =   12825
   End
End
Attribute VB_Name = "frmImageBurn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


#Const DebugState = False



Private Const C_STR_BURN_COLS As String = "|检查UID,hide,key|检查号,txtright,rowcheck|姓名,read|性别,read|年龄,read|影像类别,read|检查部位,read|检查时间,read,w1900|"
Private Const C_STR_BURN_DATA_CONVERT As String = ""

Private Const STR_ATTACHED_FILE_PATH = "PACSLIST"

Private mlngCurAdviceId As Long
Private mblnMoved As Long

Private mMultiCols As Long
Private mMultiRows As Long

Private mstrReadyUID As String

Private WithEvents mObjBurn As clsImapi2Burn
Attribute mObjBurn.VB_VarHelpID = -1

Private mdsetDicomDir As DicomDataSet
Private mobjFile As New Scripting.FileSystemObject
Private mobjRegBurnFileList As New Collection
Private mstrBurnRoot As String
Private mstrBurnDicomDir As String

Private mrsCustomQuery As ADODB.Recordset
Private mlngCurDeptId As Long
Private mlngModule As Long

'显示刻录窗口
Public Sub ShowBurn(ByVal lngModule As Long, ByVal lngCurDeptId As Long, ByVal lngCurAdviceId As Long, ByVal blnMoved As Boolean, owner As Object)
On Error GoTo ErrHandle
    mlngModule = lngModule
    mlngCurDeptId = lngCurDeptId
    mlngCurAdviceId = lngCurAdviceId
    mblnMoved = blnMoved
    
    Me.Show 1, owner
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


'读取当前检查信息到列表
Private Sub ReadCurStudyInfToList()
    Dim strSQL As String
    
    strSQL = "select a.检查UID, a.检查号,a.姓名,a.性别,a.年龄, a.影像类别,a.接收日期 as 检查时间,b.医嘱内容 as 检查部位 from 影像检查记录 a, 病人医嘱记录 B where a.医嘱ID=b.ID and a.医嘱ID=[1]"
    Set ufgBurnData.AdoData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngCurAdviceId)
    
    Call ufgBurnData.RefreshData
    
    If ufgBurnData.ShowingDataRowCount > 0 Then
        ufgBurnData.HeadCheckValue = True
'        Call ufgBurnData.SetRowChecked(1, True, csCustom)
    End If
End Sub


'查询检查信息到列表
Private Sub QueryStudyInfToList()
    Dim strSQL As String
    Dim strDeviceType As String
    
    Dim strFilter As String
    
    If txtStartStudyNum.Text = "" And txtEndStudyNum.Text = "" Then
        strFilter = " a.接收日期 between [1] and [2]"
    Else
        If txtStartStudyNum.Text <> "" And txtEndStudyNum.Text <> "" Then
            strFilter = " a.检查号 between [3] and [4]"
        Else
            strFilter = " a.检查号 =" & IIf(txtStartStudyNum.Text <> "", "[3]", "[4]")
        End If
    End If
    
    If txtName.Text <> "" Then
        If strFilter <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & " a.姓名 like [5] || '%'"
    End If
    
    If cbxDeviceType.Text <> "" Then
        If strFilter <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & " a.影像类别=upper([6])"
    End If
    
    If strFilter <> "" Then strFilter = strFilter & " and "
    strFilter = strFilter & " a.检查UID is not null"
    
    strSQL = "select a.检查UID, a.检查号,a.姓名,a.性别,a.年龄, a.影像类别,a.接收日期 as 检查时间,b.医嘱内容 as 检查部位 from 影像检查记录 a, 病人医嘱记录 B where a.医嘱ID=b.ID and " & strFilter
    
    strDeviceType = Split(cbxDeviceType.Text & "--#", "--")(1)
    
    Set ufgBurnData.AdoData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
                                                        dtpBegin.value, _
                                                        dtpEnd.value, _
                                                        txtStartStudyNum.Text, _
                                                        txtEndStudyNum.Text, _
                                                        txtName.Text, _
                                                        strDeviceType)
    Call ufgBurnData.RefreshData
    
    ufgBurnData.HeadCheckValue = False
End Sub


Private Function GetImageViewData(ByVal strStudyUID As String, ByVal lngCurPage As Long, _
    ByVal lngPageRecord As Long, Optional blnIsAllData As Boolean = False) As ADODB.Recordset
'获取预览图像数据


    Dim strSQL As String
    Dim lngStartRecord As Long
    Dim lngEndRecord As Long
    
        
    strSQL = "Select rownum as 顺序号, A.图像号,D.FTP用户名 As User1,D.FTP密码 As Pwd1," & _
        "D.IP地址 As Host1,'/'||D.Ftp目录||'/' As Root1," & _
        "Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID||'/'||A.图像UID As URL,d.设备号 as 设备号1, " & _
        "E.FTP用户名 As User2,E.FTP密码 As Pwd2," & _
        "E.IP地址 As Host2,'/'||E.Ftp目录||'/' As Root2," & _
        "e.设备号 as 设备号2, A.图像UID,C.检查UID,B.序列UID,A.动态图,A.编码名称,A.采集时间, A.录制长度 " & _
        "From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像设备目录 D,影像设备目录 E " & _
        "Where A.序列UID=B.序列UID And B.检查UID=C.检查UID And C.位置一=D.设备号(+) And C.位置二=E.设备号(+) and C.检查UID=[1]"
        
    If mblnMoved Then
        strSQL = Replace(strSQL, "影像检查图象", "H影像检查图象")
        strSQL = Replace(strSQL, "影像检查序列", "H影像检查序列")
        strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
    End If
    
    lngStartRecord = (lngCurPage - 1) * lngPageRecord + 1
    lngEndRecord = lngCurPage * lngPageRecord
    
    strSQL = "select * from (" & strSQL & " order by b.序列UID, a.图像号) " & IIf(blnIsAllData, "", " where 顺序号>=" & lngStartRecord & " and 顺序号<=" & lngEndRecord)
    
    Set GetImageViewData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strStudyUID)
End Function


Private Sub LoadViewImageToFace(rsCurImageData As ADODB.Recordset)
'加载预览图像到界面
    Dim strTmpFile As String
    Dim strCachePath As String
    
    Dim curImage As DicomImage
    
    Dim objFile As New Scripting.FileSystemObject
    
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    
    Dim iCols As Integer, iRows As Integer
    
    
    
    DViewer.Images.Clear
    
    If rsCurImageData.RecordCount > 0 Then
        '计算图像显示数量
        ResizeRegion rsCurImageData.RecordCount, DViewer.Width, DViewer.Height, iRows, iCols
        
        mMultiCols = iCols
        mMultiRows = iRows

        DViewer.MultiColumns = iCols
        DViewer.MultiRows = iRows
        
        '创建本地目录
        strCachePath = IIf(Len(App.Path) > 3, App.Path & "\TmpImage\", App.Path & "TmpImage\")
        MkLocalDir strCachePath & objFile.GetParentFolderName(Nvl(rsCurImageData("URL")))
        
        Do While Not rsCurImageData.EOF
            '循环加载图像到DicomViewer中
            strTmpFile = strCachePath & Nvl(rsCurImageData("URL"))
            
            If Nvl(rsCurImageData("动态图"), IMGTAG) = VIDEOTAG Then
                strTmpFile = IIf(Len(App.Path) > 3, App.Path & "\..\附加文件\Avi.bmp", App.Path & "..\附加文件\Avi.bmp")
            ElseIf Nvl(rsCurImageData("动态图"), IMGTAG) = AUDIOTAG Then
                strTmpFile = IIf(Len(App.Path) > 3, App.Path & "\..\附加文件\wav.bmp", App.Path & "..\附加文件\wav.bmp")
            End If
            
            If mobjFile.FileExists(strTmpFile) = False Then
                '本地缓存图像不存在，则读取FTP图像
                
                '建立FTP连接
                If Nvl(rsCurImageData("设备号1")) <> vbNullString And Inet1.hConnection = 0 Then
                    If Inet1.FuncFtpConnect(Nvl(rsCurImageData("Host1")), Nvl(rsCurImageData("User1")), Nvl(rsCurImageData("Pwd1"))) = 0 Then
                        If Nvl(rsCurImageData("设备号2")) <> vbNullString Then
                            If Inet2.FuncFtpConnect(Nvl(rsCurImageData("Host2")), Nvl(rsCurImageData("User2")), Nvl(rsCurImageData("Pwd2"))) = 0 Then
                                MsgBoxD Me, "FTP不能正常连接，请检查网络设置。"
                                Exit Sub
                            End If
                        Else
                            MsgBoxD Me, "FTP不能正常连接，请检查网络设置。"
                            Exit Sub
                        End If
                    End If
                End If
                
                If Inet1.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsCurImageData("Root1")) & rsCurImageData("URL")), strTmpFile, objFile.GetFileName(rsCurImageData("URL"))) <> 0 Then
                    '从设备号1提取图像失败，则从设备号2提取图像
                    If Nvl(rsCurImageData("设备号2")) <> vbNullString Then
                        If Inet2.hConnection = 0 Then Inet2.FuncFtpConnect Nvl(rsCurImageData("Host2")), Nvl(rsCurImageData("User2")), Nvl(rsCurImageData("Pwd2"))
                        Call Inet2.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsCurImageData("Root2")) & rsCurImageData("URL")), strTmpFile, objFile.GetFileName(rsCurImageData("URL")))
                    End If
                End If
            End If
  
            If mobjFile.FileExists(strTmpFile) Then
               If Nvl(rsCurImageData("动态图"), IMGTAG) <> VIDEOTAG And Nvl(rsCurImageData("动态图"), IMGTAG) <> AUDIOTAG Then
                    Set curImage = DViewer.Images.ReadFile(strTmpFile)
                    
                    With curImage
                        .BorderStyle = 6
                        .BorderWidth = 1
                        .BorderColour = vbWhite
                    End With
                Else
                    Set curImage = New DicomImage
                    
                    On Error GoTo continue
                        Call curImage.FileImport(strTmpFile, "DIB/BMP")
continue:
                    
                    Call AddVideoLabelToDicomImage(curImage, _
                        "采集时间：" & Nvl(rsCurImageData("采集时间")), _
                        "录制长度：" & Nvl(rsCurImageData("录制长度"), "0") & " 秒", _
                        "编码名称：" & Nvl(rsCurImageData("编码名称")))
                    
                    With curImage
                        .BorderStyle = 6
                        .BorderWidth = 1
                        .BorderColour = vbWhite
                    End With
                    
                    Call DViewer.Images.Add(curImage)
                End If
                
                
                '取消自动剪影,因为DicomObjects控件本身对处理剪影有BUG，存在（0028，6100）时，会自动对图像进行剪影，
                '导致晋煤的DSA图像不能正常显示
                '虽然设置图像的mask=0 ,可以取消剪影，但是每次图像被添加到新的Dicomimages之后，自动又将mask设置成1了，
                '这样在程序中无法很好的控制，因此直接去掉（0028，6100）这个属性。
                If Not IsNull(curImage.Attributes(&H28, &H6100).value) Then
                    curImage.Attributes.Remove &H28, &H6100
                End If
            End If
            
            rsCurImageData.MoveNext
        Loop
        
        
        Inet1.FuncFtpDisConnect
        Inet2.FuncFtpDisConnect
    Else
        DViewer.MultiColumns = 1
        DViewer.MultiRows = 1
    End If
End Sub


Private Sub AdjustFace()
    framQuery.Left = 120
    framQuery.Width = Me.ScaleWidth - 240
    framQuery.Top = 0
    
    picPane.Left = 120
    picPane.Width = framQuery.Width
    picPane.Top = framQuery.Top + framQuery.Height + 120
    picPane.Height = Me.ScaleHeight - framQuery.Height - stbThis.Height - labState.Height - framBurn.Height - 320
    
    framBurn.Left = 120
    framBurn.Width = framQuery.Width
    framBurn.Top = picPane.Top + picPane.Height + 120
    
    labState.Top = framBurn.Top + framBurn.Height + 60
    
    If Me.ScaleWidth - stbThis.Panels.Item(1).Width - stbThis.Panels.Item(2).Width - stbThis.Panels.Item(3).Width - stbThis.Panels.Item(4).Width - stbThis.Panels.Item(6).Width - stbThis.Panels.Item(7).Width - 430 <= 0 Then
        pbState.Visible = False
    Else
        pbState.Visible = True
        pbState.Width = Me.ScaleWidth - stbThis.Panels.Item(1).Width - stbThis.Panels.Item(2).Width - stbThis.Panels.Item(3).Width - stbThis.Panels.Item(4).Width - stbThis.Panels.Item(6).Width - stbThis.Panels.Item(7).Width - 460
    End If
    
    pbState.Top = Me.ScaleHeight - stbThis.Height + 60

    
    Call ucSplitter1.RePaint(False)
    
End Sub


Private Sub InitBurnList()
    '设置行数
    ufgBurnData.GridRows = glngStandardRowCount
    '设置行高
    ufgBurnData.RowHeightMin = glngStandardRowHeight
    
    ufgBurnData.IsKeepRows = False
    ufgBurnData.DefaultColNames = C_STR_BURN_COLS
    ufgBurnData.ColNames = C_STR_BURN_COLS
    ufgBurnData.ColConvertFormat = C_STR_BURN_DATA_CONVERT
End Sub



Private Sub LoadDriverSpeed()
    Dim i As Long
    
    cbxBurnSpeed.Clear
    For i = 0 To mObjBurn.GetCurSupportedSpeedCount - 1
        cbxBurnSpeed.AddItem mObjBurn.GetCurSupportedSpeed(i)
    Next i
    
    If cbxBurnSpeed.ListCount > 0 Then cbxBurnSpeed.ListIndex = 0
End Sub


Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
'自定义查询
    Dim strSQL As String
    Dim strReturn As String
    Dim strPars As Variant
    
On Error GoTo ErrHandle

    strReturn = frmCustomQueryCall.ShowCustomQuery(control.ID, mlngCurDeptId, mlngModule, strPars, Me)
    
    If strReturn = "" Then Exit Sub
    
    strSQL = "select a.检查UID, a.检查号,a.姓名,a.性别,a.年龄, a.影像类别,a.接收日期 as 检查时间,b.医嘱内容 as 检查部位 " & _
             "from 影像检查记录 a, 病人医嘱记录 B where a.医嘱ID=b.ID and a.检查UID is not null and b.id in (" & strReturn & ")"
    
    Set ufgBurnData.AdoData = GetDataToLocal(strSQL, "自定义查询", strPars(1), strPars(2), strPars(3), strPars(4), strPars(5), strPars(6), strPars(7), strPars(8), strPars(9), strPars(10), _
                                            strPars(11), strPars(12), strPars(13), strPars(14), strPars(15), strPars(16), strPars(17), strPars(18), strPars(19), strPars(20))
                                            
    Call ufgBurnData.RefreshData
    
    ufgBurnData.HeadCheckValue = False

Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbxDeviceName_Click()
On Error Resume Next
    chkContainBurnStation.Enabled = False
    cmdBurn.Enabled = False
        
    If Not mObjBurn.CheckingDeviceIsBurn(cbxDeviceName.Text) Then
            
        stbThis.Panels.Item(2).Text = "光盘可用大小：0 M"
        stbThis.Panels.Item(3).Text = "光盘预留大小：0 M"
        stbThis.Panels.Item(4).Text = "待刻录文件大小：0 M"

        txtVolumeName.Text = ""
        
        Call cbxBurnSpeed.Clear
        
        MsgBoxD Me, "当前设备不支持刻录操作。", vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    Do
        If Not mObjBurn.CheckingDeviceIsReady(cbxDeviceName.Text) Then
            If MsgBoxD(Me, "请先插入光盘。", vbOKCancel, Me.Caption) = vbCancel Then Exit Sub
        Else
            Exit Do
        End If
    Loop While True
    
    mObjBurn.CurBurnDevice = cbxDeviceName.Text
    chkContainBurnStation.Enabled = True
    cmdBurn.Enabled = True
    
    Call LoadDriverSpeed
    
    txtVolumeName.Text = mObjBurn.GetDiscName(cbxDeviceName.Text)
    If txtVolumeName.Text = "" Then
        txtVolumeName.Text = Format(zlDatabase.Currentdate, "yyyymmddhhmmss")
    End If
    
        stbThis.Panels.Item(2).Text = "光盘可用大小：" & Format(mObjBurn.GetDiscFreeSize / 1024 / 1024, "0.00") & " M"
        stbThis.Panels.Item(3).Text = "光盘预留大小：" & Format(mObjBurn.ReserveKBSize / 1024, "0.00") & " M"


End Sub



Private Sub ConfigAppBurnDir(ByVal blnIsClearDir As Boolean)
    mstrBurnRoot = IIf(Len(App.Path) > 3, App.Path & "\CreateCDTmp", App.Path & "CreateCDTmp")
    mstrBurnDicomDir = mstrBurnRoot & "\DICOM"
    
    If blnIsClearDir And mobjFile.FolderExists(mstrBurnDicomDir) Then Call mobjFile.DeleteFolder(mstrBurnDicomDir, blnIsClearDir)
    
    If mobjFile.FolderExists(mstrBurnRoot) = False Then Call MkDir(mstrBurnRoot)
    If mobjFile.FolderExists(mstrBurnDicomDir) = False Then Call MkDir(mstrBurnDicomDir)
End Sub


'创建DICOM目录
Private Sub CreateDicomDir()
    Dim img As DicomImage
    Dim imgs As New DicomImages
    Dim strTransfersyntax As String
    Dim rsTemp As ADODB.Recordset
    Dim strStudyUID As String
    Dim strBufferDir As String
    Dim strTmpFile As String
    Dim strMiddlePath As String
    Dim objFtp As clsFtp
    Dim i As Long
    
   
    
    strTransfersyntax = "1.2.840.10008.1.2.1"
    
    If mdsetDicomDir Is Nothing Then
        Set mdsetDicomDir = New DicomDataSet
        mdsetDicomDir.Name = "ZLPACS"
    End If
    
    strBufferDir = App.Path & "\"
    Set objFtp = New clsFtp
    
On Error GoTo errDisFtpConnect
    pbState.Min = 1
    pbState.Max = ufgBurnData.GridRows - 1
    
    For i = 1 To ufgBurnData.GridRows - 1
        strStudyUID = ufgBurnData.Text(i, "检查UID")
        
        If Trim(strStudyUID) <> "" Then
            pbState.Position = i
            
            labState.Caption = "正在处理 [" & ufgBurnData.Text(i, "姓名") & "] 的检查数据..."
            labState.Refresh
            
            If ufgBurnData.GetRowCheck(i) = True And InStr(mstrReadyUID, strStudyUID & ";") <= 0 Then
                            
                Set rsTemp = GetImageViewData(strStudyUID, 0, 0, True)
                
                If objFtp.FuncFtpConnect(Nvl(rsTemp!Host1), Nvl(rsTemp!User1), Nvl(rsTemp!Pwd1)) = 0 Then
                    If Nvl(rsTemp("设备号2")) <> vbNullString Then
                        If objFtp.FuncFtpConnect(Nvl(rsTemp!Host2), Nvl(rsTemp!User2), Nvl(rsTemp!Pwd2)) = 0 Then
                            MsgBoxD Me, "FTP不能正常连接，不能获取刻录文件，请检查网络设置。"
                            Exit Sub
                        End If
                    Else
                        MsgBoxD Me, "FTP不能正常连接，不能获取刻录文件，请检查网络设置。"
                        Exit Sub
                    End If
                End If
                
                mstrReadyUID = mstrReadyUID & strStudyUID & ";"
                
                
                While Not rsTemp.EOF
                    '判断是否存在需要刻录的本地文件，如果不存在，则需从ftp中下载
                    strBufferDir = IIf(Len(App.Path) > 3, App.Path & "\TmpImage\", App.Path & "TmpImage\")
                    strTmpFile = strBufferDir & Nvl(rsTemp("URL"))
                    
                    If mobjFile.FileExists(strTmpFile) = False Then
                
                        If objFtp.FuncDownloadFile(mobjFile.GetParentFolderName(Nvl(rsTemp("Root1")) & Nvl(rsTemp("URL"))), strTmpFile, mobjFile.GetFileName(Nvl(rsTemp("URL")))) <> 0 Then
                            '从设备号1提取图像失败，则从设备号2提取图像
                            If Nvl(rsTemp("设备号2")) <> vbNullString Then
                                If objFtp.hConnection = 0 Then objFtp.FuncFtpConnect Nvl(rsTemp("Host2")), Nvl(rsTemp("User2")), Nvl(rsTemp("Pwd2"))
                                Call objFtp.FuncDownloadFile(mobjFile.GetParentFolderName(Nvl(rsTemp("Root2")) & Nvl(rsTemp("URL"))), strTmpFile, mobjFile.GetFileName(rsTemp("URL")))
                            End If
                        End If
                        
                    End If
                    
                    '先不处理视频和音频
                    If mobjFile.FileExists(strTmpFile) And Nvl(rsTemp("动态图"), IMGTAG) <> VIDEOTAG And Nvl(rsTemp("动态图"), IMGTAG) <> AUDIOTAG Then
                        
                        Set img = imgs.ReadFile(strTmpFile)
                        img.StudyUID = strStudyUID
                        
                        '如果目录不存在，则创建目录
                        strMiddlePath = "IMAGES"
                        If mobjFile.FolderExists(mstrBurnDicomDir & "\" & strMiddlePath) = False Then
                            MkDir (mstrBurnDicomDir & "\" & strMiddlePath)
                        End If
                        
                        strMiddlePath = strMiddlePath & "\" & ChkDir(img.Name & "(" & img.PatientID & ")")
                        If mobjFile.FolderExists(mstrBurnDicomDir & "\" & strMiddlePath) = False Then
                            MkDir (mstrBurnDicomDir & "\" & strMiddlePath)
                        End If
                        
                        strMiddlePath = strMiddlePath & "\" & img.StudyUID
                        If mobjFile.FolderExists(mstrBurnDicomDir & "\" & strMiddlePath) = False Then
                            MkDir (mstrBurnDicomDir & "\" & strMiddlePath)
                        End If
                        
                        Call img.WriteFile(mstrBurnDicomDir & "\" & strMiddlePath & "\" & img.InstanceUID & ".DCM", True, strTransfersyntax)
                        Call mdsetDicomDir.AddToDirectory(img, strMiddlePath & "\" & img.InstanceUID & ".DCM", strTransfersyntax, 0)
                        
                        Call imgs.Clear
                    End If
                
                    Call rsTemp.MoveNext
                Wend
                
                Call objFtp.FuncFtpDisConnect
                
                Call mobjRegBurnFileList.Add(mstrBurnDicomDir & "\" & strMiddlePath, strStudyUID)
                
            End If
        End If
        
        DoEvents
    Next i
    
    If mdsetDicomDir.Children.Count > 0 Then
        mdsetDicomDir.WriteDirectory mstrBurnDicomDir & "\DICOMDIR"
    Else
        Call ConfigAppBurnDir(True)
    End If
    
    labState.Caption = "检查数据处理完毕..."
    
    Exit Sub
errDisFtpConnect:
    Call objFtp.FuncFtpDisConnect
    Call err.Raise(err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext)
End Sub

Private Function ChkDir(StrDirectory As String) As String
    '检查目录是否有不正常的字符，并修正
    ChkDir = Replace(StrDirectory, "/", "")
    ChkDir = Replace(StrDirectory, "\", "")
    ChkDir = Replace(StrDirectory, ":", "")
    ChkDir = Replace(StrDirectory, "*", "")
    ChkDir = Replace(StrDirectory, "?", "")
    ChkDir = Replace(StrDirectory, """", "")
    ChkDir = Replace(StrDirectory, "<", "")
    ChkDir = Replace(StrDirectory, ">", "")
    ChkDir = Replace(StrDirectory, "|", "")
End Function



Private Sub chkContainBurnStation_Click()
On Error GoTo ErrHandle
    Dim strAppPath As String
    Dim objFSO As New FileSystemObject
    
    strAppPath = objFSO.GetParentFolderName(App.Path) & IIf(Len(App.Path) > 3, "\", "") & STR_ATTACHED_FILE_PATH
    
    '保存“独立CD观片站”
    If chkContainBurnStation.value <> 0 Then
        
        If mobjFile.FolderExists(strAppPath) Then
'            mobjFile.CopyFile strAppPath & "\*.*", mstrBurnRoot
            Call mObjBurn.AddBurnDirTree(strAppPath)
        Else
            MsgBoxD Me, "没有找到复制文件路径！", vbOKOnly, Me.Caption
        End If
    Else
        '移除刻录文件
        Call mObjBurn.RemoveBurnDirTree(strAppPath)
    End If

    stbThis.Panels.Item(4).Text = "待刻录文件大小：" & Format(mObjBurn.GetBurnResourceTotalSize() / 1024 / 1024, "0.00") & " M"
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdBurn_Click()
On Error GoTo ErrHandle
    Do
        If Not mObjBurn.CheckingDeviceIsReady(cbxDeviceName.Text) Then
            If MsgBoxD(Me, "请先插入光盘。", vbOKCancel, Me.Caption) = vbCancel Then Exit Sub
            Call cbxDeviceName_Click
        Else
            Exit Do
        End If
    Loop While True
    
    If mObjBurn.GetBurnResourceTotalSize() > mObjBurn.GetDiscFreeSize Then
        Call MsgBoxD(Me, "待刻录文件数量大于光盘可用容量，不能进行刻录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If txtVolumeName.Text = "" Then
        Call MsgBoxD(Me, "卷标名称不能为空。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If mObjBurn.GetBurnResourceTotalSize <= 0 Then
        Call MsgBoxD(Me, "当前没有发现需要刻录的数据。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    mObjBurn.BurnVolumeName = txtVolumeName.Text
    mObjBurn.WriteSpeed = cbxBurnSpeed.Text
    mObjBurn.CurBurnDevice = cbxDeviceName.Text
    
    Call mObjBurn.StartBurn
    
    pbState.Position = pbState.Max
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdExit_Click()
On Error Resume Next
    Call Me.Hide
    err.Clear
End Sub

Private Sub cmdQuery_Click()
On Error GoTo ErrHandle
    '如果重新查询了数据，则删除原有的刻录数据
    Call DelDicomDir
    
    Call QueryStudyInfToList
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Dim curDate As Date
    
    #If DebugState = True Then
        mlngCurAdviceId = 302
        
        Call InitDebugObject(1290, Me, "zlhis", "HIS")
    #End If
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call ConfigAppBurnDir(True)
    
    Set mObjBurn = New clsImapi2Burn
    
    cmdBurn.Enabled = mObjBurn.HasBurnDeviceInSystem
    
    curDate = zlDatabase.Currentdate
    
    dtpBegin.value = Format(curDate, "yyyy-mm-dd 00:00:00")
    dtpEnd.value = Format(curDate, "yyyy-mm-dd 23:59:59")
    
    Call InitCustomQueryType
    
    Call InitBurnObj
    
    Call InitBurnList
    
    Call InitModality
    
    Call LoadDriverWithImapi2
    
    If mlngCurAdviceId <> 0 Then Call ReadCurStudyInfToList

    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub InitCustomQueryType()
'获取自定义查询方案
    Dim strSQL As String

On Error GoTo ErrHandle

    strSQL = "select Id, 方案名称, 是否默认, 查询语句 from 影像查询方案 where 使用状态=1"
    Set mrsCustomQuery = zlDatabase.OpenSQLRecord(strSQL, "影像查询方案")
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdCustomQuery_Click()
'配置自定义查询菜单
    Dim objPopup As CommandBar
    Dim objControl As CommandBarControl
On Error GoTo ErrHandle

    Set objPopup = cbrMain.Add("自定义查询菜单", xtpBarPopup)
    
    With objPopup.Controls
        If mrsCustomQuery.RecordCount <= 0 Then Exit Sub

        mrsCustomQuery.MoveFirst
        Do While Not mrsCustomQuery.EOF
            Set objControl = .Add(xtpControlButton, Nvl(mrsCustomQuery!ID), Nvl(mrsCustomQuery!方案名称))
            mrsCustomQuery.MoveNext
        Loop
    End With
    
    objPopup.ShowPopup

Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub InitModality()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    strSQL = "select 编码,名称 from 影像检查类别"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "影像检查类别")
    
    cbxDeviceType.Clear
    cbxDeviceType.AddItem ""
    cbxDeviceType.ListIndex = 0
    
    Do Until rsTemp.EOF
        cbxDeviceType.AddItem Nvl(rsTemp!名称) & "--" & Nvl(rsTemp!编码)
        rsTemp.MoveNext
    Loop
    
End Sub


'初始化刻录对象
Private Sub InitBurnObj()
    If mObjBurn Is Nothing Then Exit Sub
    
    mObjBurn.IsOverWirte = False
    mObjBurn.IsIncludeBaseDir = False
    mObjBurn.VerificationLevel = ivlQuick
    mObjBurn.OnceMedia = False

    mObjBurn.ReserveKBSize = 20 * 1024
End Sub

Private Sub LoadDriverWithImapi2()
    Dim i As Long
    Dim strDeviceName As String
    Dim lngBurnIndex As Long
        
    lngBurnIndex = 0
    For i = 0 To mObjBurn.DeviceCount - 1
        strDeviceName = mObjBurn.DeviceName(i)
        
        cbxDeviceName.AddItem strDeviceName
        
        If mObjBurn.CheckingDeviceIsBurn(strDeviceName) Then
            lngBurnIndex = i
        End If
    Next i
    
    If cbxDeviceName.ListCount > 0 Then cbxDeviceName.ListIndex = lngBurnIndex
End Sub


Private Sub Form_Resize()
On Error Resume Next
    Call AdjustFace
    Call err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
    If mobjFile.FolderExists(mstrBurnDicomDir) Then Call mobjFile.DeleteFolder(mstrBurnDicomDir, True)
    
    Set mObjBurn = Nothing
    Set mobjFile = Nothing
    
    
    
    err.Clear
End Sub







Private Sub mObjBurn_OnBurnEvent(ByVal strCurState As String, args As clsImapi2BurnArgs)
On Error Resume Next
    pbState.Visible = True

    pbState.Min = 0
    pbState.Max = args.TotalTime
    pbState.Position = args.ElapsedTime
'    pbState.Orientation = args.ElapsedTime
'    pbState.Refresh
    
    
    labState.Caption = strCurState & "    当前时间：" & args.ElapsedTime & "/预计时间：" & args.TotalTime
    labState.Refresh
    
    err.Clear
End Sub

Private Sub mObjBurn_OnBurnProcedureEvent(ByVal strState As String)
On Error Resume Next
    labState.Caption = strState
    labState.Refresh
    
    err.Clear
End Sub





Private Sub picImage_Resize()
On Error Resume Next
    DViewer.Left = 0
    DViewer.Top = 0
    DViewer.Height = picImage.Height - ucPage.Height - 120
    DViewer.Width = picImage.Width
    
    ucPage.Left = 60
    ucPage.Top = DViewer.Top + DViewer.Height + 60
    
'    labPageRecordCount.Left = ucPage.Left + ucPage.Width + 120
'    labPageRecordCount.Top = ucPage.Top + 60
'
'    txtPageRecordCount.Left = labPageRecordCount.Left + labPageRecordCount.Width
'    txtPageRecordCount.Top = ucPage.Top
'
'    labTotal.Left = txtPageRecordCount.Left + txtPageRecordCount.Width + 60
'    labTotal.Top = labPageRecordCount.Top
    
    err.Clear
End Sub



'Private Sub txtPageRecordCount_Change()
'On Error GoTo errHandle
'    If Not ufgBurnData.Visible Then Exit Sub
'
'    If Val(txtPageRecordCount.Text) = 0 Then Exit Sub
'
'    Call ufgBurnData_OnSelChange
'Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'End Sub

Private Sub ucPage_OnPageChange(ByVal lngPageIndex As Long, ByVal lngPageCount As Long)
On Error GoTo ErrHandle
    Dim rsData As ADODB.Recordset
    Dim strStudyUID As String
    
    If Not ufgBurnData.Visible Then Exit Sub
    If Not ufgBurnData.IsSelectionRow Then Exit Sub
    
    
    strStudyUID = ufgBurnData.Text(ufgBurnData.SelectionRow, "检查UID")
    
    Set rsData = GetImageViewData(strStudyUID, lngPageIndex, lngPageCount)
    Call LoadViewImageToFace(rsData)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

'删除dicomdir目录
Private Sub DelDicomDir()
    Dim strStudyUID As String
    Dim i As Long
    
    For i = 1 To ufgBurnData.GridRows - 1
        strStudyUID = ufgBurnData.Text(i, "检查UID")
        Call RemoveStudyBurnFile(strStudyUID)
    Next i
     
     Call ConfigAppBurnDir(True)
     
     stbThis.Panels.Item(4).Text = "待刻录文件大小：" & Format(mObjBurn.GetBurnResourceTotalSize() / 1024 / 1024, "0.00") & " M"
End Sub

Private Sub ufgBurnData_OnCheckAllChanged()
On Error GoTo ErrHandle
    Dim strBurnPath As String
    Dim strStudyUID As String
    Dim strPath As String
    Dim i As Long
    
    For i = 1 To ufgBurnData.GridRows - 1
        If Not ufgBurnData.GetRowCheck(i) Then
            strStudyUID = ufgBurnData.Text(i, "检查UID")
            Call RemoveStudyBurnFile(strStudyUID)
        End If
    Next i
    
    Call CreateDicomDir
    
    Call mObjBurn.AddBurnDirTree(mstrBurnRoot)
    
    
    stbThis.Panels.Item(4).Text = "待刻录文件大小：" & Format(mObjBurn.GetBurnResourceTotalSize() / 1024 / 1024, "0.00") & " M"
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgBurnData_OnCheckChanged(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrHandle
    Dim strBurnPath As String
    Dim strStudyUID As String
    Dim strPath As String
    
    
    If Not ufgBurnData.GetRowCheck(Row) Then
        strStudyUID = ufgBurnData.Text(Row, "检查UID")
        Call RemoveStudyBurnFile(strStudyUID)
    End If
    
    Call CreateDicomDir
    
    Call mObjBurn.AddBurnDirTree(mstrBurnRoot)
    
    stbThis.Panels.Item(4).Text = "待刻录文件大小：" & Format(mObjBurn.GetBurnResourceTotalSize() / 1024 / 1024, "0.00") & " M"
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub RemoveStudyBurnFile(ByVal strStudyUID As String)
    Dim strPath As String
    Dim i As Long
    Dim j As Long
    
    On Error GoTo continue
    If mobjRegBurnFileList.Count > 0 Then strPath = mobjRegBurnFileList.Item(strStudyUID)
continue:
    
    If Not mdsetDicomDir Is Nothing Then
        '从dicomdir中移除数据
        For i = mdsetDicomDir.Children.Count To 1 Step -1
            '移除包含的检查数据集
            For j = mdsetDicomDir.Children(i).Children.Count To 1 Step -1
                If mdsetDicomDir.Children(i).Children(j).StudyUID = strStudyUID Then
                    Call mdsetDicomDir.Children(i).Children.Remove(j)
                    
                    '删除dicomdir对应文件
                    If mobjFile.FolderExists(strPath) Then
                        Call mobjFile.DeleteFolder(strPath, True)
                    End If
                    
                    Exit For
                End If
            Next j
            
            '移除当前病人数据集
            If mdsetDicomDir.Children(i).Children.Count <= 0 Then
                mdsetDicomDir.Children.Remove (i)
                
                '移除检查的父目录
                If mobjFile.FolderExists(mobjFile.GetParentFolderName(strPath)) Then
                    Call mobjFile.DeleteFolder(mobjFile.GetParentFolderName(strPath))
                End If
            End If
        Next i
        
        If mdsetDicomDir.Children.Count <= 0 Then Call ConfigAppBurnDir(True)
    End If
    
    mstrReadyUID = Replace(mstrReadyUID, strStudyUID & ";", "")
    
    On Error GoTo continue1
    '从注册刻录列表中移除
    If mobjRegBurnFileList.Count > 0 Then Call mobjRegBurnFileList.Remove(strStudyUID)
continue1:
End Sub


Private Sub ufgBurnData_OnSelChange()
On Error GoTo ErrHandle
    Dim strStudyUID As String
    Dim rsData As ADODB.Recordset
    
    
    If Not ufgBurnData.IsSelectionRow Then Exit Sub
    
    strStudyUID = ufgBurnData.Text(ufgBurnData.SelectionRow, "检查UID")
    Call InitPageControl(strStudyUID)
    
    Set rsData = GetImageViewData(strStudyUID, 1, ucPage.PageRecord)
    
    Call LoadViewImageToFace(rsData)
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub InitPageControl(ByVal strStudyUID As String)
'初始化分页控件
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngRecordCount As Long
    

    strSQL = "select count(1)  as 返回值 from 影像检查图象 a, 影像检查序列 b where a.序列UID=b.序列UID and b.检查UID=[1]"
       
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strStudyUID)
    If rsData.RecordCount > 0 Then
        lngRecordCount = Nvl(rsData!返回值)
    Else
        lngRecordCount = 0
    End If
 
    
'    ucPage.PageRecord = Val(txtPageRecordCount.Text)
    ucPage.RecordCount = lngRecordCount
    
'    labTotal.Caption = "总数：" & lngRecordCount
End Sub

