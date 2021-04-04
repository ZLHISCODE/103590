VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBlackList 
   AutoRedraw      =   -1  'True
   Caption         =   "特殊病人管理"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   Icon            =   "frmBlackList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9885
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   810
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   9765
         _ExtentX        =   17224
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Description     =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Description     =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "登记"
               Key             =   "Add"
               Description     =   "登记"
               Object.ToolTipText     =   "登记特殊病人"
               Object.Tag             =   "登记"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modi"
               Description     =   "修改"
               Object.ToolTipText     =   "修改特殊病人"
               Object.Tag             =   "修改"
               ImageKey        =   "Modi"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "撤消"
               Key             =   "Del"
               Description     =   "撤消"
               Object.ToolTipText     =   "撤消病人的登记"
               Object.Tag             =   "撤消"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Edit_"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "卡片"
               Key             =   "View"
               Description     =   "卡片"
               Object.ToolTipText     =   "以卡片方式查阅当前病人信息"
               Object.Tag             =   "卡片"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Filter"
               Description     =   "过滤"
               Object.ToolTipText     =   "在当前病人清单中过滤满足条件的病人"
               Object.Tag             =   "过滤"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "定位"
               Key             =   "Go"
               Description     =   "定位"
               Object.ToolTipText     =   "定位到满点条件的病人上"
               Object.Tag             =   "定位"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
         Begin VB.CheckBox chkShowDel 
            Caption         =   "显示已撤消病人"
            Height          =   195
            Left            =   8175
            TabIndex        =   4
            Top             =   240
            Width           =   1560
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   6675
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBlackList.frx":06EA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12356
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshPati 
      Height          =   5970
      Left            =   0
      TabIndex        =   0
      Top             =   705
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   10530
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      BackColorSel    =   12632256
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmBlackList.frx":0F7C
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   45
      Top             =   450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":1296
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":14B0
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":16CA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":18E4
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":1AFE
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":1D18
            Key             =   "Merge"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":2412
            Key             =   "View"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":2B0C
            Key             =   "Patis"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":3206
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":3420
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":363A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":3854
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   645
      Top             =   450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":3A6E
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":3C88
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":3EA2
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":40BC
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":42D6
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":44F0
            Key             =   "Merge"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":4BEA
            Key             =   "View"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":52E4
            Key             =   "Patis"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":59DE
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":5BF8
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":5E12
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":602C
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBlackList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mstrPrivs As String
Private mrsPati As ADODB.Recordset
Private mstrFilter As String
Private mfrmFilter As frmPatiFilter
Private mfrmFind As frmPatiFind
Private mlngGo As Long, mblnGo As Boolean
Private mblnDown As Boolean
Private mlngCurRow As Long, mlngTopRow As Long

Private Type Type_SQLCondition
    Default As Boolean          '是否是缺省进入，此时没有条件值,缺省值在mstrFilter中
    登记时间B As Date
    登记时间E As Date
    出生时间B As Date
    出生时间E As Date
    入院时间B As Date
    入院时间E As Date
    出院时间B As Date
    出院时间E As Date
    住院号 As String
    性别 As String
    编号 As String
    区域 As String
    Patient As String
End Type
Private SQLCondition As Type_SQLCondition

Private Sub chkShowDel_Click()
    If Visible Then Call ShowPatis(mstrFilter)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF5
            Call ShowPatis(mstrFilter)
        Case vbKeyF3
            '始终从当前行开始查找
            If tbr.Buttons("Go").Enabled Then Call SeekPati(False)
        Case vbKeyReturn
            Call tbr_ButtonClick(tbr.Buttons("View"))
        Case vbKeyEscape
            mblnGo = False
    End Select
End Sub

Private Sub Form_Load()
    Call SetHeader
    RestoreWinState Me, App.ProductName
    
    Set mfrmFilter = New frmPatiFilter
    Set mfrmFind = New frmPatiFind
    
    If InStr(mstrPrivs, "特殊病人管理") = 0 Then
        tbr.Buttons("Add").Visible = False
        tbr.Buttons("Modi").Visible = False
        tbr.Buttons("Del").Visible = False
        tbr.Buttons("Edit_").Visible = False
    End If
    
    '刷新名单
    mstrFilter = ""
    Call ShowPatis(mstrFilter)
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long '工具条占用高度
    Dim staH As Long '状态栏占用高度
    
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub

    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    
    With mshPati
        .Left = 0
        .Top = cbrH
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - cbrH - staH
    End With
    
    chkShowDel.Top = tbr.Top + (tbr.Height - chkShowDel.Height) / 2
    If Me.ScaleWidth - chkShowDel.Width - 100 < 6000 Then
        chkShowDel.Left = 6000
    Else
        chkShowDel.Left = Me.ScaleWidth - chkShowDel.Width - 100
    End If
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmFind
    Unload mfrmFilter
    
    Set mrsPati = Nothing
    SaveWinState Me, App.ProductName
End Sub

Private Sub mshPati_DblClick()
    Call tbr_ButtonClick(tbr.Buttons("View"))
End Sub

Private Sub mshPati_EnterCell()
    mshPati.ForeColorSel = mshPati.CellForeColor
    Call SetMenuEnabled
    mlngGo = mshPati.Row
    mlngCurRow = mshPati.Row: mlngTopRow = mshPati.TopRow
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Integer
    For i = 0 To mshPati.Cols - 1
        If mshPati.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
    GetColNum = -1
End Function

Private Sub SetMenuEnabled()
'功能：根据当前记录情况设置菜单可用状态
    Dim lng病人ID As Long
    
    If glngSys Like "8??" Then
        lng病人ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("客户ID")))
    Else
        lng病人ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID")))
    End If
    
    tbr.Buttons("Print").Enabled = lng病人ID <> 0
    tbr.Buttons("Preview").Enabled = lng病人ID <> 0
    tbr.Buttons("Modi").Enabled = lng病人ID <> 0
    tbr.Buttons("Del").Enabled = lng病人ID <> 0 And mshPati.TextMatrix(mshPati.Row, GetColNum("撤消时间")) = ""
    tbr.Buttons("View").Enabled = lng病人ID <> 0
    tbr.Buttons("Go").Enabled = lng病人ID <> 0
End Sub

Private Sub mshPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And tbr.Buttons("Del").Enabled And tbr.Buttons("Del").Visible Then
        Call tbr_ButtonClick(tbr.Buttons("Del"))
    End If
End Sub

Private Sub mshPati_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then mblnDown = True
End Sub

Private Sub mshPati_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshPati.MouseRow = 0 Then
        mshPati.MousePointer = 99
    Else
        mshPati.MousePointer = 0
    End If
End Sub

Private Sub mshPati_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshPati.MouseCol
    
    If Button = 1 And mshPati.MousePointer = 99 And mblnDown Then '双击最大化时会执行
        mblnDown = False
        If mshPati.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshPati.TextMatrix(1, GetColNum("病人ID")) = "" Then Exit Sub
        Set mshPati.DataSource = Nothing
        mrsPati.Sort = mshPati.TextMatrix(0, lngCol) & IIf(mshPati.ColData(lngCol) = 0, "", " DESC")
        mshPati.ColData(lngCol) = (mshPati.ColData(lngCol) + 1) Mod 2
        Call ShowPatis("", True)
    End If
End Sub

Private Sub DeletePati(ByVal lngRow As Long)
    Dim lng编号 As Long
    
    lng编号 = Val(mshPati.TextMatrix(lngRow, GetColNum("编号")))
    If lng编号 = 0 Then Exit Sub
    
    If frmBlackListEdit.ShowMe(Me, mstrPrivs, lng编号, True) Then
        Call ShowPatis(mstrFilter)
    End If
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim lng编号 As Long, lng病人ID As Long, blnOK As Boolean
    
    Select Case Button.Key
        Case "Preview"
            Call OutputList(2)
        Case "Print"
            Call OutputList(1)
        Case "Add"
            If frmBlackListEdit.ShowMe(Me, mstrPrivs) Then
                Call ShowPatis(mstrFilter)
            End If
        Case "Modi"
            lng编号 = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("编号")))
            If lng编号 <> 0 Then
                If frmBlackListEdit.ShowMe(Me, mstrPrivs, lng编号) Then
                    Call ShowPatis(mstrFilter)
                End If
            End If
        Case "Del"
            Call DeletePati(mshPati.Row)
        Case "View"
            lng病人ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID")))
            If lng病人ID <> 0 Then
                If CreatePublicPatient() Then
                    Call gobjPublicPatient.ReadPatiDegreeCard(Me, lng病人ID, Val(mshPati.TextMatrix(mshPati.Row, GetColNum("主页ID"))))
                End If
                mshPati.Refresh
            End If
        Case "Go"
            blnOK = gblnOK
            mfrmFind.mbytType = 0 '按所有病人条件查找
            mfrmFind.Show 1, Me
            If gblnOK Then Call SeekPati(mfrmFind.optHead)
            gblnOK = blnOK
        Case "Filter"
            blnOK = gblnOK
            
            mfrmFilter.mbytType = 0 '按所有病人条件过滤
            mfrmFilter.mbytInFun = 1
            mfrmFilter.Show 1, Me
            If gblnOK Then
                 With mfrmFilter
                    mstrFilter = .mstrFilter
                    SQLCondition.登记时间B = .dtp登记B
                    SQLCondition.登记时间E = .dtp登记E
                    SQLCondition.出生时间B = .dtp出生B
                    SQLCondition.出生时间E = .dtp出生E
                    
                    SQLCondition.入院时间B = .dtp入院B
                    SQLCondition.入院时间E = .dtp入院E
                    SQLCondition.出院时间B = .dtp出院B
                    SQLCondition.出院时间E = .dtp出院E
                    
                    SQLCondition.住院号 = Trim(.txt住院号.Text)
                    SQLCondition.性别 = zlCommFun.GetNeedName(.cbo性别.Text)
                    SQLCondition.编号 = Trim(.txt编号.Text)
                    SQLCondition.区域 = zlCommFun.GetNeedName(.txt区域.Text)
                    
                    If .PatiIdentify.GetCurCard.名称 = "姓名" And .mlngPatiId = 0 And (.chk登记.Value = 1 Or .chk入院.Value = 1 Or .chk出院.Value = 1) Then    '姓名
                        SQLCondition.Patient = Trim(.PatiIdentify.Text) & "%"
                    Else
                        SQLCondition.Patient = IIf(.mlngPatiId <> 0, .mlngPatiId, .PatiIdentify.Text)
                    End If
                End With
                
                Call ShowPatis(mstrFilter)
            End If
            gblnOK = blnOK
        Case "Help"
            ShowHelp App.ProductName, Me.hwnd, Me.Name
        Case "Quit"
            Unload Me
    End Select
End Sub

Private Function GetPatiType(ByVal lngRow As Long) As Byte
'功能：0-所有,1-在院,2-出院,3-门诊
    With mshPati
        If .TextMatrix(.Row, GetColNum("入院时间")) <> "" And .TextMatrix(.Row, GetColNum("出院时间")) = "" Then
            GetPatiType = 1
        ElseIf .TextMatrix(.Row, GetColNum("出院时间")) <> "" Then
            GetPatiType = 2
        Else
            GetPatiType = 3
        End If
    End With
End Function

Private Sub OutputList(bytStyle As Byte)
'功能：输入出列表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    intRow = mshPati.Row
    
    '表头
    objOut.Title.Text = "特殊病人名单"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    mshPati.Redraw = False
    Set objOut.Body = mshPati
    
    '输出
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    mshPati.Row = intRow
    mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
    mshPati.Redraw = True
End Sub

Private Sub SetHeader(Optional blnWidth As Boolean = True)
    Dim strHead As String
    Dim i As Integer
    
    strHead = "编号,1,600|病人ID,1,750|标识号,1,750|姓名,1,700|性别,1,500|年龄,1,500|加入原因,1,2500|加入时间,1,1100|登记人,1,700|撤消原因,1,2500|撤消时间,1,1100|撤消人,1,700|费别,1,850|科室,1,850|床号,1,500|入院时间,1,1000|出院时间,1,1000|住院次数,4,850|主页ID,1,0|身份证号,1,0|门诊号,1,0|住院号,1,0|就诊卡号,1,0"

    With mshPati
        .Redraw = False
        
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible And blnWidth Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        .RowHeight(0) = 320
        
        '恢复上次行
        If mlngCurRow = 0 Then mlngCurRow = 1
        If mlngTopRow = 0 Then mlngTopRow = 1
        If mlngCurRow <= .Rows - 1 Then
            .Row = mlngCurRow
        Else
            .Row = .Rows - 1
        End If
        If mlngTopRow <= .Rows - 1 Then
            .TopRow = mlngTopRow
        Else
            .TopRow = .Row
        End If
        
        .Col = 0: .ColSel = .Cols - 1
        Call mshPati_EnterCell

        .Redraw = True
    End With
End Sub

Private Sub ShowPatis(ByVal strIF As String, Optional blnSort As Boolean)
    Dim Curdate As Date, strSQL As String
    
    On Error GoTo errH
    
    If Not blnSort Then
        '缺省显示所有特殊病人
        If strIF = "" Then
            strIF = " And C.加入时间 Between trunc(sysdate-30) And trunc(sysdate)+1"
        Else
            strIF = Replace(strIF, "A.登记时间", "C.加入时间")
        End If
        
        If chkShowDel.Value = 0 Then
            strIF = strIF & " And C.撤消时间 is NULL"
        End If
        
        strSQL = _
            "Select C.编号,A.病人ID,Decode(Nvl(A.住院次数,0),0,A.门诊号,A.住院号) as 标识号," & _
            " A.姓名,A.性别,A.年龄,C.加入原因,To_Char(C.加入时间,'MM-DD HH24:MI') as 加入时间,C.登记人," & _
            " C.撤消原因,To_Char(C.撤消时间,'MM-DD HH24:MI') as 撤消时间,C.撤消人," & _
            " Decode(Nvl(A.主页ID,0),0,A.费别,P.费别) as 费别,D.名称 as 科室,P.出院病床 as 床号," & _
            " To_Char(P.入院日期,'YYYY-MM-DD') as 入院时间,To_Char(P.出院日期,'YYYY-MM-DD') as 出院时间," & _
            " A.住院次数,P.主页ID,A.身份证号,A.门诊号,A.住院号,A.就诊卡号" & _
            " From 病案主页 P,病人信息 A,特殊病人 C,部门表 D" & _
            " Where A.病人ID=P.病人ID(+) And Nvl(A.主页ID,0)=P.主页ID(+)" & _
            " And A.病人ID=C.病人ID And P.出院科室ID=D.ID(+)" & strIF & _
            " Order by C.加入时间 Desc"

        Call zlCommFun.ShowFlash("正在读取特殊病人名单,请稍候 ...", Me)
        Me.Refresh
        With SQLCondition
            Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, 0, "", .登记时间B, .登记时间E, .出生时间B, .出生时间E, _
                .入院时间B, .入院时间E, .出院时间B, .出院时间E, .住院号, .性别, .区域, .编号, .Patient)
        End With
    End If
    
    mshPati.Clear
    mshPati.Rows = 2
    
    If mrsPati.EOF Then
        Call SetHeader(False)
        stbThis.Panels(2).Text = "当前设置没有过滤出任何病人"
    Else
        Set mshPati.DataSource = mrsPati
        Call SetHeader(False)
        mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
        stbThis.Panels(2) = "共 " & mrsPati.RecordCount & " 个病人"
    End If
    Call mshPati_EnterCell
    
    If Not blnSort Then Call zlCommFun.StopFlash
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SeekPati(blnHead As Boolean)
    Dim i As Long
    Dim blnFill As Boolean
    
    Screen.MousePointer = 11
    mblnGo = True
    stbThis.Panels(2).Text = "正在定位满足条件的病人,按ESC终止 ..."
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshPati.Rows - 1
        DoEvents
        
        '比较条件
        blnFill = True
        With mfrmFind
            If .txt病人ID.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("病人ID")) = .txt病人ID.Text
            End If
            If .txt就诊卡.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("就诊卡号")) = .txt就诊卡.Text
            End If
            If .txt门诊号.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("门诊号")) = .txt门诊号.Text
            End If
            If .txt住院号.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("住院号")) = .txt住院号.Text
            End If
            If .txt床号.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("床号")) = .txt床号.Text
            End If
            If .txt姓名.Text <> "" Then
                blnFill = blnFill And UCase(mshPati.TextMatrix(i, GetColNum("姓名"))) Like "*" & UCase(.txt姓名.Text) & "*"
            End If
            If .txt身份证.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("身份证号")) = .txt身份证.Text
            End If
        End With
        
        '满足则退出
        If blnFill Then
            mlngGo = i + 1
            mshPati.Row = i: mshPati.TopRow = i
            mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
            stbThis.Panels(2).Text = "找到一条记录"
            Screen.MousePointer = 0: Exit Sub
        End If
        
        '按ESC取消
        If mblnGo = False Then
            stbThis.Panels(2).Text = "用户取消定位操作"
            Screen.MousePointer = 0: Exit Sub
        End If
    Next
    mlngGo = 1
    stbThis.Panels(2).Text = "已定位到名单尾部"
    Screen.MousePointer = 0
End Sub
