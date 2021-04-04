VERSION 5.00
Begin VB.Form frmReportPathology 
   BorderStyle     =   0  'None
   ClientHeight    =   2355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7725
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2355
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame frmCell 
      Caption         =   "细胞项目："
      Height          =   1095
      Left            =   0
      TabIndex        =   13
      Top             =   1200
      Width           =   7695
      Begin VB.TextBox txt腺上皮细胞病变 
         Height          =   300
         Left            =   3480
         TabIndex        =   8
         Top             =   320
         Width           =   1300
      End
      Begin VB.TextBox txt鳞状上皮细胞病变 
         Height          =   300
         Left            =   3480
         TabIndex        =   9
         Top             =   697
         Width           =   1300
      End
      Begin VB.TextBox txt炎症细胞 
         Height          =   300
         Left            =   6000
         TabIndex        =   11
         Top             =   690
         Width           =   1300
      End
      Begin VB.TextBox txt鳞状细胞 
         Height          =   300
         Left            =   6000
         TabIndex        =   10
         Top             =   315
         Width           =   1300
      End
      Begin VB.ComboBox cbo化生细胞 
         Height          =   300
         ItemData        =   "frmReportPathology.frx":0000
         Left            =   960
         List            =   "frmReportPathology.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   697
         Width           =   800
      End
      Begin VB.ComboBox cbo颈管细胞 
         Height          =   300
         ItemData        =   "frmReportPathology.frx":0016
         Left            =   960
         List            =   "frmReportPathology.frx":0020
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   320
         Width           =   800
      End
      Begin VB.Label Label12 
         Caption         =   "腺上皮细胞病变："
         Height          =   255
         Left            =   1920
         TabIndex        =   25
         Top             =   345
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "鳞状上皮细胞病变："
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "化生细胞："
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "颈管细胞："
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   345
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "炎症细胞："
         Height          =   255
         Left            =   5040
         TabIndex        =   21
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label7 
         Caption         =   "鳞状细胞："
         Height          =   255
         Left            =   5040
         TabIndex        =   20
         Top             =   345
         Width           =   900
      End
   End
   Begin VB.Frame frmMicrobe 
      Caption         =   "微生物信息："
      Height          =   1095
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   7695
      Begin VB.ComboBox cboHPV感染 
         Height          =   300
         ItemData        =   "frmReportPathology.frx":002C
         Left            =   6720
         List            =   "frmReportPathology.frx":0036
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   800
      End
      Begin VB.ComboBox cbo疱疹病毒感染 
         Height          =   300
         ItemData        =   "frmReportPathology.frx":0042
         Left            =   3960
         List            =   "frmReportPathology.frx":004C
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   800
      End
      Begin VB.ComboBox cbo球杆菌感染 
         Height          =   300
         ItemData        =   "frmReportPathology.frx":0058
         Left            =   1320
         List            =   "frmReportPathology.frx":0062
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   800
      End
      Begin VB.ComboBox cbo放线菌感染 
         Height          =   300
         ItemData        =   "frmReportPathology.frx":006E
         Left            =   6720
         List            =   "frmReportPathology.frx":0078
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   800
      End
      Begin VB.ComboBox cbo念珠菌感染 
         Height          =   300
         ItemData        =   "frmReportPathology.frx":0084
         Left            =   3960
         List            =   "frmReportPathology.frx":008E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   800
      End
      Begin VB.ComboBox cbo滴虫感染 
         Height          =   300
         ItemData        =   "frmReportPathology.frx":009A
         Left            =   1320
         List            =   "frmReportPathology.frx":00A4
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   800
      End
      Begin VB.Label Label6 
         Caption         =   "HPV感染："
         Height          =   255
         Left            =   5520
         TabIndex        =   19
         Top             =   750
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "疱疹病毒感染："
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "球杆菌感染："
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   743
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "放线菌感染："
         Height          =   255
         Left            =   5520
         TabIndex        =   16
         Top             =   390
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "念珠菌感染："
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Top             =   390
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "滴虫感染："
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   383
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmReportPathology"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnSingleWindow As Boolean     '是否使用独立窗口显示报告编辑器，True-独立窗口显示；False-嵌入式显示
Private mlngAdviceID As Long    '医嘱ID
Private mlngReportID As Long    '病历文件号
Private mintEditType As Integer '病历状态 0 创建，1书写，2 修订
Private mblnCheckModity As Boolean      '是否启动内容修改记录
Private mblnEditable As Boolean         '是否可以编辑内容
Private mblnMoved As Boolean            '是否已经转储

'病理专科报告要素
Private Const Report_Element_滴虫感染 = "滴虫感染"
Private Const Report_Element_球杆菌感染 = "球杆菌感染"
Private Const Report_Element_念珠菌感染 = "念珠菌感染"
Private Const Report_Element_疱疹病毒感染 = "疱疹病毒感染"
Private Const Report_Element_放线菌感染 = "放线菌感染"
Private Const Report_Element_HPV感染 = "HPV感染"
Private Const Report_Element_颈管细胞 = "颈管细胞"
Private Const Report_Element_化生细胞 = "化生细胞"
Private Const Report_Element_腺上皮细胞病变 = "腺上皮细胞病变"
Private Const Report_Element_鳞状上皮细胞病变 = "鳞状上皮细胞病变"
Private Const Report_Element_鳞状细胞 = "鳞状细胞"
Private Const Report_Element_炎症细胞 = "炎症细胞"

Public pModified As Boolean     '记录是否有修改

Public Sub zlRefresh(frmParentReport As frmReport, ByVal lngAdviceID As Long, lngReportID As Long, _
    blnSingleWindow As Boolean, blnEditable As Boolean, ByVal blnMoved As Boolean)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    mlngAdviceID = lngAdviceID
    mlngReportID = lngReportID
    mblnSingleWindow = blnSingleWindow
    mblnEditable = blnEditable
    mblnMoved = blnMoved
    
    '初始化控件
    Call InitControls
    
    mblnCheckModity = False         '关闭内容修改记录
    pModified = False
    
    strSql = "Select 内容文本,要素名称 From 电子病历内容 Where 文件ID=[1] And 对象类型=4 And 终止版=0 And 替换域=0"
    If mblnMoved = True Then
        strSql = Replace(strSql, "电子病历内容", "H电子病历内容")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngReportID)
    
    While rsTemp.EOF = False
        Select Case Nvl(rsTemp!要素名称)
            Case Report_Element_滴虫感染
                cbo滴虫感染.Text = Nvl(rsTemp!内容文本, "否")
            Case Report_Element_球杆菌感染
                cbo球杆菌感染.Text = Nvl(rsTemp!内容文本, "否")
            Case Report_Element_念珠菌感染
                cbo念珠菌感染.Text = Nvl(rsTemp!内容文本, "否")
            Case Report_Element_疱疹病毒感染
                cbo疱疹病毒感染.Text = Nvl(rsTemp!内容文本, "否")
            Case Report_Element_放线菌感染
                cbo放线菌感染.Text = Nvl(rsTemp!内容文本, "否")
            Case Report_Element_HPV感染
                cboHPV感染.Text = Nvl(rsTemp!内容文本, "否")
            Case Report_Element_颈管细胞
                cbo颈管细胞.Text = Nvl(rsTemp!内容文本, "否")
            Case Report_Element_化生细胞
                cbo化生细胞.Text = Nvl(rsTemp!内容文本, "否")
            Case Report_Element_腺上皮细胞病变
                txt腺上皮细胞病变.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_鳞状上皮细胞病变
                txt鳞状上皮细胞病变.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_鳞状细胞
                txt鳞状细胞.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_炎症细胞
                txt炎症细胞.Text = Nvl(rsTemp!内容文本)
        End Select
        rsTemp.MoveNext
    Wend
    
    '设置界面控件是否可以编辑
    frmMicrobe.Enabled = mblnEditable
    frmCell.Enabled = mblnEditable
    
    mblnCheckModity = True         '启动内容修改记录
End Sub

Public Function getElementString() As String
    Dim strElements As String
    
    strElements = SPLITER_REPORT & Report_Element_滴虫感染 & SPLITER_ELEMENT & cbo滴虫感染.Text & SPLITER_REPORT & _
                    Report_Element_球杆菌感染 & SPLITER_ELEMENT & cbo球杆菌感染.Text & SPLITER_REPORT & _
                    Report_Element_念珠菌感染 & SPLITER_ELEMENT & cbo念珠菌感染.Text & SPLITER_REPORT & _
                    Report_Element_疱疹病毒感染 & SPLITER_ELEMENT & cbo疱疹病毒感染.Text & SPLITER_REPORT & _
                    Report_Element_放线菌感染 & SPLITER_ELEMENT & cbo放线菌感染.Text & SPLITER_REPORT & _
                    Report_Element_HPV感染 & SPLITER_ELEMENT & cboHPV感染.Text & SPLITER_REPORT & _
                    Report_Element_颈管细胞 & SPLITER_ELEMENT & cbo颈管细胞.Text & SPLITER_REPORT & _
                    Report_Element_化生细胞 & SPLITER_ELEMENT & cbo化生细胞.Text & SPLITER_REPORT & _
                    Report_Element_腺上皮细胞病变 & SPLITER_ELEMENT & txt腺上皮细胞病变.Text & SPLITER_REPORT & _
                    Report_Element_鳞状上皮细胞病变 & SPLITER_ELEMENT & txt鳞状上皮细胞病变.Text & SPLITER_REPORT & _
                    Report_Element_鳞状细胞 & SPLITER_ELEMENT & txt鳞状细胞.Text & SPLITER_REPORT & _
                    Report_Element_炎症细胞 & SPLITER_ELEMENT & txt炎症细胞.Text
    getElementString = strElements
End Function

Private Sub InitControls()
    cbo滴虫感染.ListIndex = 1
    cbo球杆菌感染.ListIndex = 1
    cbo念珠菌感染.ListIndex = 1
    cbo疱疹病毒感染.ListIndex = 1
    cbo放线菌感染.ListIndex = 1
    cboHPV感染.ListIndex = 1
    cbo颈管细胞.ListIndex = 1
    cbo化生细胞.ListIndex = 1
    
    txt腺上皮细胞病变.Text = ""
    txt鳞状上皮细胞病变.Text = ""
    txt鳞状细胞.Text = ""
    txt炎症细胞.Text = ""
End Sub

Private Sub cboHPV感染_DropDown()
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cboHPV感染_KeyDown(KeyCode As Integer, Shift As Integer)
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo滴虫感染_DropDown()
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo滴虫感染_KeyDown(KeyCode As Integer, Shift As Integer)
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo放线菌感染_DropDown()
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo放线菌感染_KeyDown(KeyCode As Integer, Shift As Integer)
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo化生细胞_DropDown()
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo化生细胞_KeyDown(KeyCode As Integer, Shift As Integer)
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo颈管细胞_DropDown()
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo颈管细胞_KeyDown(KeyCode As Integer, Shift As Integer)
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo念珠菌感染_DropDown()
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo念珠菌感染_KeyDown(KeyCode As Integer, Shift As Integer)
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo疱疹病毒感染_DropDown()
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo疱疹病毒感染_KeyDown(KeyCode As Integer, Shift As Integer)
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo球杆菌感染_DropDown()
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo球杆菌感染_KeyDown(KeyCode As Integer, Shift As Integer)
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            zlCommFun.PressKey vbKeyTab
    End Select
End Sub

Private Sub Form_Resize()
    Dim lngTemp As Long
    
    '摆放控件位置
    If Me.Width > 7500 Then    '摆成2排
        '微生物信息
        Label3.Left = cbo念珠菌感染.Left + cbo念珠菌感染.Width + 500
        Label3.Top = Label2.Top
        cbo放线菌感染.Left = Label3.Left + Label3.Width + 50
        cbo放线菌感染.Top = cbo念珠菌感染.Top
        
        Label6.Left = Label3.Left
        Label6.Top = Label5.Top
        cboHPV感染.Left = cbo放线菌感染.Left
        cboHPV感染.Top = cbo疱疹病毒感染.Top
        
        '细胞项目
        Label7.Left = txt腺上皮细胞病变.Left + txt腺上皮细胞病变.Width + 300
        Label7.Top = Label12.Top
        txt鳞状细胞.Left = Label7.Left + Label7.Width + 50
        txt鳞状细胞.Top = txt腺上皮细胞病变.Top
        
        Label8.Left = txt鳞状上皮细胞病变.Left + txt鳞状上皮细胞病变.Width + 300
        Label8.Top = Label11.Top
        txt炎症细胞.Left = txt鳞状细胞.Left
        txt炎症细胞.Top = txt鳞状上皮细胞病变.Top
    Else        '摆成3排
        '微生物信息
        Label3.Left = Label4.Left
        Label3.Top = Label4.Top + Label4.Height + 200
        cbo放线菌感染.Left = cbo球杆菌感染.Left
        cbo放线菌感染.Top = Label3.Top - 50
        
        Label6.Left = Label5.Left
        Label6.Top = Label5.Top + Label5.Height + 200
        cboHPV感染.Left = cbo疱疹病毒感染.Left
        cboHPV感染.Top = Label6.Top - 50
        
        '细胞项目
        Label7.Left = Label10.Left
        Label7.Top = Label10.Top + Label10.Height + 200
        txt鳞状细胞.Left = Label7.Left + Label7.Width '+ 10
        txt鳞状细胞.Top = Label7.Top - 50
        
        Label8.Left = txt鳞状细胞.Left + txt鳞状细胞.Width + 200
        Label8.Top = Label11.Top + Label11.Height + 200
        txt炎症细胞.Left = Label8.Left + Label8.Width ' + 10
        txt炎症细胞.Top = Label8.Top - 50
    End If
    
    '摆放外框
    frmMicrobe.Left = 0
    frmMicrobe.Top = 0
    lngTemp = Me.Width - 100
    frmMicrobe.Width = IIf(lngTemp > 0, lngTemp, 0)
    frmMicrobe.Height = Me.cboHPV感染.Top + Me.cboHPV感染.Height + 100
    
    frmCell.Left = 0
    frmCell.Top = frmMicrobe.Height + 50
    frmCell.Width = frmMicrobe.Width
    frmCell.Height = Me.txt炎症细胞.Top + Me.txt炎症细胞.Height + 100
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strRegPath As String
    
    If mblnSingleWindow = True Then
        strRegPath = "公共模块\" & App.ProductName & "\frmReport\SingleWindow"
    Else
        strRegPath = "公共模块\" & App.ProductName & "\frmReport"
    End If
    
    SaveSetting "ZLSOFT", strRegPath, "CY22", Me.Height
End Sub

Private Sub txt鳞状上皮细胞病变_Change()
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub txt鳞状上皮细胞病变_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt鳞状上皮细胞病变_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt鳞状细胞_Change()
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub txt鳞状细胞_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt鳞状细胞_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt腺上皮细胞病变_Change()
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub txt腺上皮细胞病变_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt腺上皮细胞病变_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt炎症细胞_Change()
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub txt炎症细胞_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt炎症细胞_LostFocus()
    Call zlCommFun.OpenIme
End Sub
