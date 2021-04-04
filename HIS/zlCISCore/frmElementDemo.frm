VERSION 5.00
Object = "{B0475000-7740-11D1-BDC3-0020AF9F8E6E}#6.1#0"; "TTF16.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmElementDemo 
   Caption         =   "病历元素示范"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
   Icon            =   "frmElementDemo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   6345
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame FraButtom 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      TabIndex        =   20
      Top             =   3720
      Width           =   6375
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   3975
         TabIndex        =   14
         Top             =   270
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   150
         Picture         =   "frmElementDemo.frx":058A
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   270
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   5100
         TabIndex        =   15
         Top             =   270
         Width           =   1100
      End
      Begin VB.Frame fraLine 
         Height          =   30
         Index           =   1
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   6555
      End
      Begin MSComDlg.CommonDialog cdgThis 
         Left            =   1500
         Top             =   165
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox picProgBar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   0
         ScaleHeight     =   210
         ScaleWidth      =   6255
         TabIndex        =   24
         Top             =   50
         Visible         =   0   'False
         Width           =   6255
         Begin MSComctlLib.ProgressBar prbRefresh 
            Height          =   195
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   344
            _Version        =   393216
            Appearance      =   0
         End
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   720
      TabIndex        =   18
      Top             =   600
      Width           =   5600
      Begin VB.CommandButton CmdEdit 
         Caption         =   "编辑(&E)"
         Height          =   350
         Left            =   4200
         TabIndex        =   9
         Top             =   1440
         Width           =   1100
      End
      Begin VB.TextBox txt说明 
         Height          =   555
         Left            =   720
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   525
         Width           =   4575
      End
      Begin VB.TextBox txt名称 
         Height          =   300
         Left            =   2595
         MaxLength       =   20
         TabIndex        =   3
         Top             =   135
         Width           =   2685
      End
      Begin VB.TextBox txt编码 
         Height          =   300
         Left            =   720
         MaxLength       =   5
         TabIndex        =   1
         Top             =   135
         Width           =   795
      End
      Begin VB.Frame fraLine 
         Height          =   30
         Index           =   0
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   6105
      End
      Begin VB.ComboBox cbo科室 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1155
         Width           =   1935
      End
      Begin VB.PictureBox picSample 
         BackColor       =   &H80000005&
         Height          =   1215
         Left            =   15
         ScaleHeight     =   1155
         ScaleWidth      =   5235
         TabIndex        =   22
         Top             =   1800
         Width           =   5295
         Begin zl9CISCore.VisItem VisItem 
            Height          =   225
            Index           =   0
            Left            =   840
            TabIndex        =   23
            Top             =   600
            Visible         =   0   'False
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   397
            MousePointer    =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            AllowEdit       =   -1  'True
         End
         Begin zl9CISCore.ctrlVisForm VisForm 
            Height          =   735
            Left            =   2520
            TabIndex        =   13
            Top             =   360
            Visible         =   0   'False
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   1296
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
         Begin VB.PictureBox PicFlag 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   240
            ScaleHeight     =   25
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   25
            TabIndex        =   10
            Top             =   0
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtBox 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   270
            Left            =   2400
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   0
            Visible         =   0   'False
            Width           =   975
         End
         Begin TTF160Ctl.F1Book grdTable 
            Height          =   1335
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Visible         =   0   'False
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   2355
            _0              =   $"frmElementDemo.frx":06D4
            _1              =   $"frmElementDemo.frx":0ADD
            _2              =   $"frmElementDemo.frx":0EE6
            _3              =   $"frmElementDemo.frx":12EF
            _4              =   $"frmElementDemo.frx":16F8
            _5              =   $"frmElementDemo.frx":1B01
            _6              =   $"frmElementDemo.frx":1F0A
            _7              =   $"frmElementDemo.frx":2313
            _8              =   $"frmElementDemo.frx":271C
            _count          =   9
            _ver            =   2
         End
      End
      Begin VB.Label lbl说明 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "说明(&M)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   15
         TabIndex        =   4
         Top             =   540
         Width           =   630
      End
      Begin VB.Label lbl名称 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "名称(&N)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1890
         TabIndex        =   2
         Top             =   195
         Width           =   630
      End
      Begin VB.Label lbl编号 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "编号(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   15
         TabIndex        =   0
         Top             =   195
         Width           =   630
      End
      Begin VB.Label lbl科室 
         AutoSize        =   -1  'True
         Caption         =   "适用(&A)"
         Height          =   180
         Left            =   30
         TabIndex        =   6
         Top             =   1230
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "示范(&S)"
         Height          =   255
         Left            =   15
         TabIndex        =   8
         Top             =   1560
         Width           =   735
      End
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "各科室通用的元素，可指定当前示范适用的科室，否则各科室都可使用该元素示范。"
      Height          =   360
      Left            =   720
      TabIndex        =   17
      Top             =   105
      Width           =   5415
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   105
      Picture         =   "frmElementDemo.frx":28F5
      Top             =   150
      Width           =   480
   End
End
Attribute VB_Name = "frmElementDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'说明：
'   1、上级程序通过本窗体ShowMe函数，将父窗体、编辑单据ID,编辑状态等信息传递进入本程序
'   2、编辑状态：由Me.tag存放，分别为"增加"、"修改"，由上级程序通过ShowMe传入
'---------------------------------------------------
Private lngElementID As Long, lngSampleID As Long '元素ID，元素示范内容ID
Private lngItemID As Long        '被编辑的项目ID，修改、查阅时由上级程序通过ShowMe传递进入,增加时为0，

Private Const MINWIDTH As Long = 6400
Private Const MINHEIGHT As Long = 4700

Private rsTemp As New ADODB.Recordset
Private objItem As ListItem
Private strTemp As String, aryTemp() As String
Private intCount As Integer
Private iElementType As Integer, sElementCode As String, lngDepartID As Long
Private strFont As String
Private aPicFlag As MapItems '标记图编辑返回值
Private objParent As Object
Private bNotRunSelChange As Boolean

Private WithEvents SpecPaper As VBControlExtender
Attribute SpecPaper.VB_VarHelpID = -1

Public Sub ShowMe(ByVal frmParent As Object, ByVal blnAdd As Boolean, ByVal lng元素Id As Long, Optional ByVal lngDemoID As Long = 0)
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '---------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    
    lngElementID = lng元素Id
    lngItemID = lngDemoID
    If blnAdd Then
        Me.Tag = "增加"
        
        '查询元素属性
        lngSampleID = 0
        
        zlDatabase.OpenRecordset rsTmp, "Select 类型,编码,科室ID,字体,名称 From 病历元素目录 Where ID=" & lngElementID, Me.Caption
        If Not rsTmp.EOF Then
            iElementType = rsTmp(0)
            sElementCode = IIf(IsNull(rsTmp(1)), "", rsTmp(1))
            lngDepartID = IIf(IsNull(rsTmp(2)), 0, rsTmp(2))
            strFont = IIf(IsNull(rsTmp(3)), "宋体,9", rsTmp(3))
            
            Me.Caption = "病历元素示范" + IIf(IsNull(rsTmp(4)), "", "-" + rsTmp(4))
        Else
            iElementType = 0
            sElementCode = ""
            lngDepartID = 0
            strFont = "宋体,9"
        End If
    Else
        Me.Tag = "修改"
        '查询元素属性
        zlDatabase.OpenRecordset rsTmp, "Select b.类型,编码,b.科室ID,b.字体,c.ID,b.名称 From 病历示范目录 a,病历元素目录 b,病人病历内容 c Where a.元素ID=b.ID And a.ID=c.病历示范ID And a.ID=" & lngDemoID, Me.Caption
        If Not rsTmp.EOF Then
            iElementType = rsTmp(0)
            sElementCode = IIf(IsNull(rsTmp(1)), "", rsTmp(1))
            lngDepartID = IIf(IsNull(rsTmp(2)), 0, rsTmp(2))
            strFont = IIf(IsNull(rsTmp(3)), "宋体,9", rsTmp(3))
            lngSampleID = rsTmp(4)
            
            Me.Caption = "病历元素示范" + IIf(IsNull(rsTmp(5)), "", "-" + rsTmp(5))
        Else
            iElementType = 0
            sElementCode = ""
            lngDepartID = 0
            strFont = "宋体,9"
            lngSampleID = 0
        End If
    End If
    
    '填写需要选择的数据
    Me.cbo科室.Clear
    Me.cbo科室.AddItem "---公用---": Me.cbo科室.ItemData(Me.cbo科室.NewIndex) = 0: Me.cbo科室.ListIndex = 0
    Err = 0: On Error GoTo errHand
    With rsTemp
        gstrSql = "select distinct D.ID,D.编码,D.名称" & _
                " from 部门表 D,部门性质说明 K" & _
                " where D.ID=K.部门ID and (K.工作性质='临床' or K.工作性质='护理')" & _
                " order by D.编码"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Do While Not .EOF
            Me.cbo科室.AddItem !名称: Me.cbo科室.ItemData(Me.cbo科室.NewIndex) = !ID
            .MoveNext
        Loop
    End With
    
    Me.Tag = "加载" & Me.Tag
    
    Set objParent = frmParent
    '显示窗体
    Me.Show 1, frmParent
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo科室_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub CmdEdit_Click()
    Dim aMapFlags As Variant
    
    Set aMapFlags = EditFlag(Me, lngElementID, aPicFlag)
    If Not aMapFlags Is Nothing Then
        Set aPicFlag = aMapFlags
        ShowFlagInOjbect PicFlag, lngElementID, aPicFlag
    End If
End Sub

Private Sub CmdEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdHelp_Click()
'    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    If Trim(Me.txt编码.Text) = "" Then MsgBox "请输入编码！", vbInformation, gstrSysName: Me.txt编码.SetFocus: Exit Sub
    If Len(Me.txt编码.Text) < Me.txt编码.MaxLength Then MsgBox "编码长度不足！", vbInformation, gstrSysName: Me.txt编码.SetFocus: Exit Sub
    If Trim(Me.txt名称.Text) = "" Then MsgBox "请输入名称！", vbInformation, gstrSysName: Me.txt名称.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt名称.Text), vbFromUnicode)) > 20 Then
        MsgBox "名称超长（最多20个字符或10个汉字）！", vbInformation, gstrSysName: Me.txt名称.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt说明.Text), vbFromUnicode)) > 50 Then
        MsgBox "说明超长（最多50个字符或25个汉字）！", vbInformation, gstrSysName: Me.txt说明.SetFocus: Exit Sub
    End If
    
    Err = 0: On Error GoTo errHand
    '数据保存
    SaveDemo
        
    Unload Me
    Exit Sub
errHand:
    If Err.Number = vbObjectError + 1 Then
        MsgBox Err.Description, vbInformation, gstrSysName
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Sub

Private Sub SaveDemo()
    Dim iContID As Long
    
    Dim strSQL1 As String, strSQLDelete As String, strSQL2 As String, strSQLSaveCont As String
    Dim ErrorNumber As Long, ErrorMsg As String
    
    '数据保存
    gcnOracle.BeginTrans
    Err = 0: On Error GoTo errHand
    
    strSQL1 = "'" & Trim(Me.txt编码.Text) & "'," & _
            "'" & Trim(Me.txt名称.Text) & "'," & _
            "'" & Trim(Me.txt说明.Text) & "'," & _
            Me.cbo科室.ItemData(Me.cbo科室.ListIndex)
    If Me.Tag = "增加" Then
        lngItemID = zlDatabase.GetNextId("病历示范目录")
        strSQL1 = "ZL_病历示范目录_INSERT(" & lngItemID & ",2," & lngElementID & "," & strSQL1 & ",'" & UserInfo.姓名 & "')"
        
        iContID = zlDatabase.GetNextId("病人病历内容")
        strSQLDelete = ""
        strSQL2 = "ZL_病人病历内容_INSERT(" & iContID & "," & lngItemID & ",'',1," & iElementType & ",'" & sElementCode & "'," + _
            "0,'',0,'',0,0,'',0,0,2)"
    Else
        strSQL1 = "ZL_病历示范目录_UPDATE(" & lngItemID & "," & strSQL1 & ")"
        
        iContID = lngSampleID
        
        strSQLDelete = "ZL_病人病历内容_DELETE(" & iContID & ")"
        strSQL2 = "ZL_病人病历内容_INSERT(" & iContID & "," & lngItemID & ",'',1," & iElementType & ",'" & sElementCode & "'," + _
            "0,'',0,'',0,0,'',0,0,2)"
    End If
    
    Call SQLTest(App.ProductName, Me.Caption, strSQL1)
    gcnOracle.Execute strSQL1, , adCmdStoredProc: Call SQLTest
    
    If Len(strSQLDelete) > 0 Then
        Call SQLTest(App.ProductName, Me.Caption, strSQLDelete)
        gcnOracle.Execute strSQLDelete, , adCmdStoredProc: Call SQLTest
    End If
    
    If Len(strSQL2) > 0 Then
        Call SQLTest(App.ProductName, Me.Caption, strSQL2)
        gcnOracle.Execute strSQL2, , adCmdStoredProc: Call SQLTest
    End If
    
    Select Case iElementType
        Case 0
            strSQLSaveCont = "ZL_病人病历文本段_SAVE(" & iContID & ",1,'" & Replace(txtBox, "'", "''") & "')"
                
            Call SQLTest(App.ProductName, Me.Caption, strSQLSaveCont)
            gcnOracle.Execute strSQLSaveCont, , adCmdStoredProc: Call SQLTest
        Case 1
            Me.MousePointer = vbHourglass
            BeginShowProgress
            SaveTable_Patient CStr(iContID), grdTable, gcnOracle, , , Me.prbRefresh
            Me.picProgBar.Visible = False
            Me.MousePointer = vbDefault
        Case 2
            Me.MousePointer = vbHourglass
            BeginShowProgress
            VisForm.SaveForm CStr(iContID), gcnOracle, ErrorNumber, ErrorMsg, Me.prbRefresh
            Me.picProgBar.Visible = False
            Me.MousePointer = vbDefault
            If ErrorNumber <> 0 Then
                Err.Description = ErrorMsg
                Err.Raise ErrorNumber, "元素示范"
            End If
        Case 3
            SaveFlag iContID, aPicFlag, gcnOracle
        Case 4
            If Not SpecPaper.SaveData(0, 0, iContID, strSQLSaveCont, ErrorMsg) Then
                Err.Description = ErrorMsg
                If Err.Number = 0 Then
                    Err.Raise vbObjectError + 1, "元素示范"
                Else
                    Err.Raise Err.Number, "元素示范"
                End If
            Else
                gcnOracle.Execute strSQLSaveCont, , adCmdStoredProc
            End If
    End Select

    gcnOracle.CommitTrans
    Exit Sub
errHand:
    gcnOracle.RollbackTrans
    Err.Raise Err.Number, "元素示范"
End Sub

Private Sub initForm()
    Dim aFont() As String, Font As New StdFont
    Dim i As Long, iNum As Long
    Dim rsTmp As New ADODB.Recordset, sTmpFile As String, FileObj As New Scripting.FileSystemObject
    Dim iTabLeft As Long, iTabTop As Long, iTabWidth As Long, iTabHeight As Long, iShown As Integer
    Dim strTxtBox As String
    Dim iOldCtrlWidth As Long, iOldCtrlHeight As Long, iNewCtrlWidth As Long, iNewCtrlHeight As Long
    Dim rsElement As New ADODB.Recordset
    
    Form_Resize
    '提取执行项目的信息
    Err = 0: On Error GoTo errHand
    With rsTemp
        If Me.Tag = "增加" Then
            gstrSql = "select nvl(max(编号),'00000') as 编码 From 病历示范目录 Where 文件ID Is Null And 元素ID=" & lngElementID
            If .State = adStateOpen Then .Close
            Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
            Me.txt编码.Text = Format(Val(!编码) + 1, "00000")
            If lngDepartID > 0 Then
                For intCount = 0 To Me.cbo科室.ListCount - 1
                    If Me.cbo科室.ItemData(intCount) = lngDepartID Then
                        Me.cbo科室.ListIndex = intCount: Exit For
                    End If
                Next
            End If
        Else
            gstrSql = "select ID,编号,名称,说明,科室id" & _
                    " from 病历示范目录" & _
                    " where ID=" & lngItemID
            If .State = adStateOpen Then .Close
            Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
            If .RecordCount > 0 Then
                Me.txt编码.Text = Format(!编号, "00000"): Me.txt名称.Text = !名称
                Me.txt说明.Text = IIf(IsNull(!说明), "", !说明)
                For intCount = 0 To Me.cbo科室.ListCount - 1
                    If Me.cbo科室.ItemData(intCount) = IIf(IsNull(!科室id), 0, !科室id) Then
                        Me.cbo科室.ListIndex = intCount: Exit For
                    End If
                Next
            End If
        End If
    End With
    
    If lngDepartID > 0 Then Me.cbo科室.Enabled = False
    If iElementType = 3 Then
        Me.CmdEdit.Visible = True
    Else
        Me.CmdEdit.Visible = False
    End If
    
    zlDatabase.OpenRecordset rsElement, "Select * From 病历元素目录 Where ID=" & lngElementID, ""
    Select Case iElementType
        Case 0
            Me.txtBox.Visible = True
            Me.txtBox.TabIndex = CmdEdit.TabIndex + 1
            
            strTxtBox = ""
            If lngItemID <> 0 Then '读取病历内容
                zlDatabase.OpenRecordset rsTmp, "Select * From 病人病历文本段 Where 病历ID=" & lngSampleID, ""
                If Not rsTmp.EOF Then strTxtBox = IIf(IsNull(rsTmp("内容")), "", rsTmp("内容"))
            End If
            
            On Error Resume Next
            aFont = Split(strFont, ",")
            
            With txtBox
                .FontName = aFont(0)
                .FontSize = aFont(1)
                .FontBold = aFont(2)
                .FontItalic = aFont(3)
                .FontUnderline = aFont(4)
                .FontStrikethru = aFont(5)
                
                .Enabled = True
                .Text = strTxtBox
            End With
            Err = 0: On Error GoTo errHand
        Case 1
            Me.grdTable.Visible = True
            Me.grdTable.TabIndex = CmdEdit.TabIndex + 1
        
            On Error Resume Next
            aFont = Split(strFont, ",")
            
            With grdTable
                InitTable grdTable
                
                .DefaultFontName = aFont(0)
                .DefaultFontSize = -1 * (aFont(1) * 1440 / 72) '将磅转为缇
                
                iOldCtrlWidth = .Width: iOldCtrlHeight = .Height
                
                Me.MousePointer = vbHourglass
                BeginShowProgress
                
                If lngItemID <> 0 Then '读取病历内容
                    ReadTable_Patient grdTable, lngSampleID, , Me.prbRefresh
                Else
                    ReadTable grdTable, lngElementID, , Me.prbRefresh
                End If
                .SetSelection 1, 1, .MaxRow, .MaxCol
                .WordWrap = True
                .SetSelection 1, 1, 1, 1
                
                .EnableProtection = True
                
                .RangeToTwips 1, 1, .MaxRow, .MaxCol, iTabLeft, iTabTop, iTabWidth, iTabHeight, iShown
                .Width = iTabWidth + 15
                .Height = iTabHeight + 15
                
                iNewCtrlWidth = .Width: iNewCtrlHeight = .Height
                If iNewCtrlWidth > iOldCtrlWidth Then Me.Width = Me.Width + iNewCtrlWidth - iOldCtrlWidth
                If iNewCtrlHeight > iOldCtrlHeight Then Me.Height = Me.Height + iNewCtrlHeight - iOldCtrlHeight
                
                Me.picProgBar.Visible = False
                Me.MousePointer = vbDefault
                
                .Enabled = True
            End With
            Err = 0: On Error GoTo errHand
        Case 2
            Me.VisForm.Visible = True
            Me.VisForm.TabIndex = CmdEdit.TabIndex + 1
        
            On Error Resume Next
            aFont = Split(strFont, ",")
                
            Me.MousePointer = vbHourglass
            BeginShowProgress
            
            With VisForm
                Font.Name = aFont(0)
                Font.Size = aFont(1)
                Font.Bold = aFont(2)
                Font.Italic = aFont(3)
                Font.Underline = aFont(4)
                Font.Strikethrough = aFont(5)
                
                Set .Font = Font
                
                iOldCtrlWidth = .Width: iOldCtrlHeight = .Height
                If lngItemID <> 0 Then '读取病历内容
                    .ReadForm lngSampleID, False, , , , Me.prbRefresh
                Else
                    .ReadForm lngElementID, , , , , Me.prbRefresh
                End If
                iNewCtrlWidth = .Width: iNewCtrlHeight = .Height
                If iNewCtrlWidth > iOldCtrlWidth Then Me.Width = Me.Width + iNewCtrlWidth - iOldCtrlWidth
                If iNewCtrlHeight > iOldCtrlHeight Then Me.Height = Me.Height + iNewCtrlHeight - iOldCtrlHeight
                
                Me.picProgBar.Visible = False
                Me.MousePointer = vbDefault
                
                .Enabled = True
            End With
            Err = 0: On Error GoTo errHand
        Case 3
            Me.PicFlag.Visible = True
            Me.PicFlag.Enabled = True
            If lngItemID <> 0 Then '读取病历内容
                Set aPicFlag = GetMapItems(lngSampleID)
            Else
                Set aPicFlag = New MapItems
            End If
            With PicFlag
                iOldCtrlWidth = .Width: iOldCtrlHeight = .Height
                
                Set .Picture = ReadCaseMap(lngElementID)
                .Width = .ScaleX(.Picture.Width, vbHimetric, vbTwips): .Height = .ScaleY(.Picture.Height, vbHimetric, vbTwips)
                .Width = IIf(.Width > 3000, 3000, .Width): .Height = .Height * .Width / .ScaleX(.Picture.Width, vbHimetric, vbTwips)
                .Cls: Set .Picture = Nothing
                
                iNewCtrlWidth = .Width: iNewCtrlHeight = .Height
                If iNewCtrlWidth > iOldCtrlWidth Then Me.Width = Me.Width + iNewCtrlWidth - iOldCtrlWidth
                If iNewCtrlHeight > iOldCtrlHeight Then Me.Height = Me.Height + iNewCtrlHeight - iOldCtrlHeight
            End With
            ShowFlagInOjbect PicFlag, lngElementID, aPicFlag
        Case 4 '专用纸
            If Not rsElement.EOF Then
                On Error Resume Next
                Licenses.Add rsElement("部件")
                Err = 0: On Error GoTo errHand
                iOldCtrlWidth = picSample.Width - 30: iOldCtrlHeight = picSample.Height - 30
                Set SpecPaper = Me.Controls.Add(rsElement("部件"), "SpecPaper")
                With SpecPaper
                    Set .Container = picSample
                    .ID病人病历 = lngSampleID
                    
                    .DispMode = False
                    .TabIndex = CmdEdit.TabIndex + 1
                    .Visible = True
                
                    iNewCtrlWidth = .Width: iNewCtrlHeight = .Height
                    If iNewCtrlWidth > iOldCtrlWidth Then Me.Width = Me.Width + iNewCtrlWidth - iOldCtrlWidth
                    If iNewCtrlHeight > iOldCtrlHeight Then Me.Height = Me.Height + iNewCtrlHeight - iOldCtrlHeight
                End With
            End If
    End Select
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If Mid(Me.Tag, 1, 2) = "加载" Then
        Me.Tag = Mid(Me.Tag, 3)
        initForm
    End If
    Me.txt编码.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    Call cmdCancel_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If InStr("TXTBOX,GRDTABLE,VISFORM", UCase(Me.ActiveControl.Name)) = 0 And KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
'    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    Dim tmpCtrl As Control
    On Error Resume Next
    
    If Me.Width < MINWIDTH Then Me.Width = MINWIDTH
    If Me.Height < MINHEIGHT Then Me.Height = MINHEIGHT
    
    With lblNote
        .Width = Me.ScaleWidth - 255 - .Left
    End With
    
    With fraMain
        .Top = lblNote.Top + lblNote.Height + 50 '+ Me.cmdCancel.Top
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - FraButtom.Height - .Top
    End With
    With fraLine(0)
        .Width = fraMain.Width - .Left
    End With
    With Me.txt名称
        .Width = fraMain.Width - .Left - 300
    End With
    With Me.txt说明
        .Width = fraMain.Width - .Left - 300
    End With
    With Me.CmdEdit
        .Left = fraMain.Width - 300 - .Width
    End With
    With picSample
        .Width = fraMain.Width - .Left - 300
        .Height = fraMain.Height - .Top - 50 '- Me.cmdCancel.Top
    End With
    
    With FraButtom
        .Top = Me.ScaleHeight - .Height
        .Width = Me.ScaleWidth - .Left
    End With
    With fraLine(1)
        .Width = FraButtom.Width - .Left
    End With
    With Me.cmdCancel
        .Left = FraButtom.Width - .Width - Me.cmdHelp.Left
    End With
    With Me.cmdOK
        .Left = Me.cmdCancel.Left - 50 - .Width
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub grdTable_DblClick(ByVal nRow As Long, ByVal nCol As Long)
    grdTable.StartEdit False, True, False
End Sub

Private Sub grdTable_EndEdit(EditString As String, Cancel As Integer)
    Dim iDecPos As Integer
    On Error Resume Next
    With grdTable
        If IsNumeric(EditString) Then
            iDecPos = InStr(EditString, ".")
            If iDecPos > 0 And iDecPos < Len(EditString) Then
                .NumberFormat = "#." + String(Len(EditString) - iDecPos, "0")
            Else
                .NumberFormat = "General"
            End If
        Else
            .NumberFormat = "General"
        End If
        .TextRC(.Row, .Col) = EditString
        
        .SetRowHeightAuto .Row, 1, .Row, .MaxCol, True
    End With
End Sub

Private Sub grdTable_GotFocus()
    With grdTable
        .Row = IIf(.Row <= .FixedRows, .FixedRows + 1, .Row)
        .Col = IIf(.Col <= .FixedCols, .FixedCols + 1, .Col)
        
        .ShowActiveCell
        bNotRunSelChange = False
    End With
End Sub

Private Sub grdTable_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab Then Me.cmdOK.SetFocus
End Sub

Private Sub grdTable_KeyPress(KeyAscii As Integer)
    Dim objCellFormat As TTF160Ctl.F1CellFormat
    On Error Resume Next
    With grdTable
        Set objCellFormat = .GetCellFormat
        If Len(objCellFormat.ValidationText) > 0 Then
            KeyAscii = 0
        End If
    End With
End Sub

Private Sub grdTable_LostFocus()
    bNotRunSelChange = True
End Sub

Private Sub grdTable_SelChange()
    Dim objCellFormat As TTF160Ctl.F1CellFormat
    Dim aVisItemInfo() As String
    
    On Error Resume Next
    '非用户操作触发的，不处理
    If bNotRunSelChange Then Exit Sub
    If Not Me.Visible Or Me.ActiveControl.Name <> "grdTable" Then Exit Sub
    With grdTable
        Set objCellFormat = .GetCellFormat
        If Len(objCellFormat.ValidationText) > 0 Then
            aVisItemInfo = Split(objCellFormat.ValidationText, ",")
            Me.VisItem(aVisItemInfo(1)).SetFocus
        End If
    End With
End Sub

Private Sub grdTable_StartEdit(EditString As String, Cancel As Integer)
    Dim objCellFormat As TTF160Ctl.F1CellFormat
    On Error Resume Next
    With grdTable
        Set objCellFormat = .GetCellFormat
        If Len(objCellFormat.ValidationText) > 0 Then
            Cancel = True
        End If
    End With
End Sub

Private Sub grdTable_TopLeftChanged()
    '非用户操作触发的，不处理
    If bNotRunSelChange Then Exit Sub
    
    bNotRunSelChange = True
    Proc_Table_TopLeftChanged grdTable
    bNotRunSelChange = False
End Sub

Private Sub PicFlag_DblClick()
    CmdEdit_Click
End Sub

Private Sub picSample_Resize()
    Dim tmpCtrl As Control
    On Error Resume Next
    Select Case iElementType
        Case 0
            Set tmpCtrl = txtBox
        Case 1
            Set tmpCtrl = grdTable
        Case 2
            Set tmpCtrl = VisForm
        Case 3
            Set tmpCtrl = PicFlag
        Case 4
            Set tmpCtrl = SpecPaper
    End Select
    With tmpCtrl
        .Left = 0: .Top = 0
        .Width = picSample.Width - 30: .Height = picSample.Height - 30
    End With
    If iElementType = 3 Then ShowFlagInOjbect PicFlag, lngElementID, aPicFlag
End Sub

Private Sub txtBox_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = vbCtrlMask Then
        txtBox.Tag = "1"
        Me.cmdOK.SetFocus
    End If
End Sub

Private Sub txtBox_KeyPress(KeyAscii As Integer)
    If txtBox.Tag = "1" Then
        KeyAscii = 0
        txtBox.Tag = ""
        Exit Sub
    End If
End Sub

Private Sub txt编码_GotFocus()
    Me.txt编码.SelStart = 0: Me.txt编码.SelLength = 100
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

Private Sub txt名称_GotFocus()
    Me.txt名称.SelStart = 0: Me.txt名称.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt名称_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt说明_GotFocus()
    Me.txt说明.SelStart = 0: Me.txt说明.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt说明_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt说明_LostFocus()
    Me.txt说明.Text = Replace(Me.txt说明, Chr(vbKeyReturn), "")
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub VisForm_NextControl()
    Me.cmdOK.SetFocus
End Sub

Private Sub VisItem_GotFocus(Index As Integer)
    Dim aCellInfo() As String

    On Error Resume Next
    aCellInfo = Split(VisItem(Index).Tag, ",")
    
    grdTable.SetActiveCell aCellInfo(0), aCellInfo(1): DoEvents
End Sub

Private Sub VisItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim aCellInfo() As String
    
    On Error Resume Next
    If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then
        grdTable.SetFocus
        zlCommFun.PressKey CByte(KeyCode)
    End If
End Sub

Private Sub BeginShowProgress()
    With picProgBar
        .Width = Me.FraButtom.Width - 2 * .Left
        .Visible = True
    End With
    With prbRefresh
        .Width = Me.picProgBar.Width - 50
    End With
End Sub
