VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "当前位置"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4515
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4515
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ImageList imgMain 
      Left            =   1275
      Top             =   2385
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelClient.frx":0000
            Key             =   "dep"
            Object.Tag             =   "dep"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   1980
      Left            =   150
      TabIndex        =   2
      Top             =   570
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   3493
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgMain"
      SmallIcons      =   "imgMain"
      ColHdrIcons     =   "imgMain"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "部门名称"
         Object.Width           =   6704
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "站点编号"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3090
      TabIndex        =   1
      Top             =   2610
      Width           =   1230
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "请选择你当前计算机位置所在的部门："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   195
      TabIndex        =   0
      Top             =   270
      Width           =   3570
   End
End
Attribute VB_Name = "frmSelClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstr部门 As String
Dim mstr名称 As String
Dim mstrCurIndex As String
Public gstr站点 As String
Public gstrCur站点 As String

Private Sub cmdOK_Click()
    If lvwMain.ListItems.Count <> 0 Then
    If ObjPtr(lvwMain.SelectedItem) = 0 Then
        If lvwMain.Enabled Then lvwMain.SetFocus
    End If
    
    gstr站点 = ""
    gstrCur站点 = ""
    With lvwMain.SelectedItem
        gstr站点 = .SubItems(1)
        gstrCur站点 = .Text

        If gstr站点 = "" Then
            MsgBox "请选择一个计算机所在的部门!", vbInformation, "提示"
            If lvwMain.Enabled Then lvwMain.SetFocus
            Exit Sub
        End If
    End With
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    '加载头信息
    Call LoadListView(mstr部门, mstr名称, mstrCurIndex)
    
End Sub

Public Sub ShowEdit(ByVal str部门 As String, ByVal str名称 As String, ByVal strCurIndex As String)
    '--功能：显示选择计算机位置所在部门
    mstr部门 = str部门
    mstr名称 = str名称
    mstrCurIndex = strCurIndex
    Me.Show 1
End Sub

Private Sub LoadListView(ByVal str部门, str名称, strCurIndex As String)
    Dim i As Integer
    Dim strSplit部门() As String, strSplit名称() As String
    Dim mList As MSComctlLib.ListItem
    On Error Resume Next
    With lvwMain
        .ListItems.Clear
        strSplit部门 = Split(mstr部门, ",")
        strSplit名称 = Split(mstr名称, ",")
        For i = 0 To UBound(strSplit名称) - 1
            Set mList = .ListItems.Add(, , strSplit名称(i), "dep", "dep")
            mList.SubItems(1) = strSplit部门(i)
        Next
        
        If .Enabled Then .SetFocus
        
        If lvwMain.ListItems.Count > 0 Then
            If strCurIndex = "" Then
                lvwMain.ListItems(1).Selected = True
            Else
                lvwMain.ListItems(1).Selected = True
                For i = 1 To lvwMain.ListItems.Count
                    If strCurIndex = lvwMain.ListItems(i).SubItems(1) Then
                        lvwMain.ListItems(i).Selected = True
                        Exit For
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdOK_Click
    End If
End Sub

Private Sub lvwMain_DblClick()
    cmdOK_Click
End Sub
