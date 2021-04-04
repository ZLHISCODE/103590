VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServiceRoom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "库房选择"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   Icon            =   "frmServiceRoom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4560
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox chk库房 
      Appearance      =   0  'Flat
      Caption         =   "全选"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3785
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   675
   End
   Begin VB.CommandButton cmd取消 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3360
      TabIndex        =   2
      Top             =   2640
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2160
      TabIndex        =   1
      Top             =   2640
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvw存储库房 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3836
      View            =   2
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "名称"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceRoom.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceRoom.frx":D0B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceRoom.frx":13916
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmServiceRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstr药品分类 As String
Private mstr存储库房 As String
Private mstr存储库房ID As String
Private mstrArr存储库房() As String
Private mstrStationNo As String
Private mstrPrivs As String
Private mbln无药库药房性质部门 As Boolean

Private Sub Init存储库房()
    Dim rsTemp As New ADODB.Recordset
    Dim rsOther As New ADODB.Recordset
    Const str西药 As String = "'西药%'"
    Const str中药 As String = "'中药%'"
    Const str成药 As String = "'成药%'"
    Dim mstr全部库房ID As String
    Dim dbl所有库房 As Boolean
    
    On Error GoTo ErrHandle
    
    If InStr(1, ";" & mstrPrivs & ";", ";所有库房;") > 0 Then dbl所有库房 = True
    
    '根据药品的用途分类提取所允许存储的库房
    gstrSql = " Select ID,编码,名称 From 部门表 " & _
              " Where ID in (select distinct 部门id from 部门性质说明 where 工作性质 like "
    If mstr药品分类 = "西成药" Then
        gstrSql = gstrSql & str西药
    ElseIf mstr药品分类 = "中成药" Then
        gstrSql = gstrSql & str成药
    Else
        gstrSql = gstrSql & str中药
    End If
    gstrSql = gstrSql & " or 工作性质='制剂室')"
    
    gstrSql = gstrSql & " and (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "根据药品的用途分类提取所允许存储的库房(其他库房)")
    mstr全部库房ID = ""
    Do While Not rsTemp.EOF
        mstr全部库房ID = mstr全部库房ID & "," & rsTemp!ID
        rsTemp.MoveNext
    Loop
    If mstr全部库房ID <> "" Then
        mbln无药库药房性质部门 = False
    Else
        mbln无药库药房性质部门 = True
        Exit Sub
    End If
    
    If Not dbl所有库房 Then
        '取当前用户所属库房
        gstrSql = gstrSql & " And Id In(Select 部门ID From 部门人员 Where 人员id=[1]) "
    End If
    
    gstrSql = gstrSql & "order by id"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "根据药品的用途分类提取所允许存储的库房", UserInfo.ID)
    
    lvw存储库房.ListItems.Clear

    With rsTemp
        Do While Not .EOF
            lvw存储库房.ListItems.Add , "K" & !ID, !名称, , 2
            .MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim intItem As Integer
    Dim intItems As Integer
    '取得存储库房
    mstr存储库房ID = ""
    mstr存储库房 = ""
    intItems = Me.lvw存储库房.ListItems.Count
    For intItem = 1 To intItems
        If lvw存储库房.ListItems(intItem).Checked Then
            mstr存储库房ID = mstr存储库房ID & "!!" & Mid(lvw存储库房.ListItems(intItem).Key, 2) & "|"
            mstr存储库房 = mstr存储库房 & "|" & Mid(lvw存储库房.ListItems(intItem).Text, 1)
        End If
    Next
    mstr存储库房 = Mid(mstr存储库房, 2)
    mstr存储库房ID = Mid(mstr存储库房ID, 3)
    Call frmBatchUpdate.ShowRoom(mstr存储库房, mstr存储库房ID)
    Unload Me
End Sub

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call Init存储库房
    
    If mbln无药库药房性质部门 = True Then
        MsgBox "请先设置具有药库药房性质的部门。", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    Call Change存储库房
End Sub

Public Sub ShowMe(ByVal frmParent As Object, ByVal str药品分类 As String, ByVal str存储库房 As String, ByVal strPrivs As String)
    mstr药品分类 = str药品分类
    mstr存储库房 = str存储库房
    mstrPrivs = strPrivs

    Me.Show 1, frmParent
End Sub
Private Sub chk库房_Click()
'库房全选按钮
    If chk库房.Value = 2 Then Exit Sub
    Call SetSelect(lvw存储库房, chk库房.Value)
End Sub
Private Sub SetSelect(ByVal lvwObj As Object, Optional ByVal BlnSelect As Boolean = True)
'全选功能
    Dim intSelect As Integer
    With lvwObj
        For intSelect = 1 To .ListItems.Count
            .ListItems(intSelect).Checked = BlnSelect
        Next
    End With
End Sub
Private Sub lvw存储库房_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'具体选择的存储库房
    Call ItemCheck(lvw存储库房, Item, chk库房)
End Sub
Private Sub ItemCheck(ByVal lvwObj As Object, ByVal Item As MSComctlLib.ListItem, ByVal chkObj As CheckBox)
'纪录选择的库房
    Dim lngCheck As Long, blnCheck As Boolean, intCount As Integer
    
    intCount = 0
    With lvwObj
        For lngCheck = 1 To .ListItems.Count
            If .ListItems(lngCheck).Checked = True Then
                intCount = intCount + 1
            End If
        Next
        
        If intCount = lvwObj.ListItems.Count Then
            chkObj.Value = 1
        ElseIf intCount > 0 Then
            chkObj.Value = 2
        Else
            chkObj.Value = 0
        End If
    End With
End Sub

Private Sub Change存储库房()
    Dim i As Integer, j As Integer
    Dim intSelect As Integer
    mstrArr存储库房 = Split(mstr存储库房, "|")
    
    For i = LBound(mstrArr存储库房) To UBound(mstrArr存储库房)
        For intSelect = 1 To lvw存储库房.ListItems.Count
            If mstrArr存储库房(i) = lvw存储库房.ListItems(intSelect).Text Then
                lvw存储库房.ListItems(intSelect).Checked = True
                j = j + 1
            End If
        Next
    Next
    
    If j = lvw存储库房.ListItems.Count Then
        chk库房.Value = 1
    ElseIf j > 0 And j < lvw存储库房.ListItems.Count Then
        chk库房.Value = 2
    End If
End Sub


