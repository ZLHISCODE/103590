VERSION 5.00
Begin VB.Form frmTechnicStudy 
   Caption         =   "检查项目关联"
   ClientHeight    =   6330
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5670
   Icon            =   "frmTechnicStudy.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   5670
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdSure 
      Caption         =   "确 定(&S)"
      Height          =   375
      Left            =   3390
      Picture         =   "frmTechnicStudy.frx":000C
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5835
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消(&C)"
      Height          =   375
      Left            =   4485
      Picture         =   "frmTechnicStudy.frx":0156
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5835
      Width           =   1100
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   5670
      TabIndex        =   1
      Top             =   0
      Width           =   5670
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmTechnicStudy.frx":02A0
         Height          =   660
         Left            =   225
         TabIndex        =   2
         Top             =   165
         Width           =   5265
      End
   End
   Begin zl9PACSWork.ucFlexGrid ufgStudy 
      Height          =   4755
      Left            =   90
      TabIndex        =   0
      Top             =   975
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   8387
      DefaultCols     =   ""
      ColNames        =   "|ID,hide,key|项目名称>名称,w3000,rowcheck|项目编码>编码,read,w1200|影像类别,w1000|分组ID,hide|"
      KeyName         =   "ID"
      DisCellColor    =   16777215
      IsCopyAdoMode   =   0   'False
      IsEjectConfig   =   -1  'True
      HeadFontSize    =   10.5
      HeadFontCharset =   134
      HeadFontWeight  =   400
      HeadColor       =   0
      DataFontSize    =   10.5
      DataFontCharset =   134
      DataFontWeight  =   400
      DataColor       =   -2147483640
      ExtendLastCol   =   -1  'True
   End
End
Attribute VB_Name = "frmTechnicStudy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngGroupId As Long
Private mblnOK As Boolean

Public Function ShowStudyAssociation(ByVal lngGroupId As Long, objOwner As Object) As Boolean
'显示检查项目关联
    mlngGroupId = lngGroupId
    mblnOK = False
    
    ShowStudyAssociation = False
    
    Call LoadStudyPro
    
    Me.Show 1, objOwner
    
    ShowStudyAssociation = mblnOK
    
End Function


Private Sub LoadStudyPro()
'载入检查项目
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select a.ID,a.名称,a.编码,b.影像类别,b.分组id " & _
            " from 诊疗项目目录 a, 影像检查项目 b, 影像设备目录 c, 医技执行房间 d, 影像执行分组 e " & _
            " where a.id =b.诊疗项目id and b.影像类别=c.影像类别 and c.设备号=d.检查设备 and d.分组id=e.id and (b.分组Id=[1] or b.分组ID is null) and e.id=[1] "
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询影像检查项目", mlngGroupId)
    
    Call ufgStudy.ClearListData
    If rsData.RecordCount <= 0 Then Exit Sub
    
    rsData.Sort = "分组ID Desc,名称 Asc"
    Set ufgStudy.AdoData = rsData
    
    ufgStudy.GridRows = ufgStudy.AdoData.RecordCount + 1
    Call ufgStudy.RefreshData
End Sub


Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdSure_Click()
On Error GoTo errHandle
    Dim strSql As String
    Dim strIds As String
    Dim i As Long
    
    strIds = ""
    
    For i = 1 To ufgStudy.GridRows - 1
        If ufgStudy.GetRowCheck(i) Then
            If strIds <> "" Then strIds = strIds & ","
            strIds = strIds & ufgStudy.KeyValue(i)
        End If
    Next i
    
    strSql = "zl_影像分组关联_Association(" & mlngGroupId & ",'" & strIds & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "影像分组关联")
    
    mblnOK = True
    
    Unload Me
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
'    'Debug Code
'    InitDebugObject 1290, Me, "zlhis", "HIS"
'    mlngGroupId = 29
'
'    LoadStudyPro
'    'Debug End
End Sub

Private Sub ufgStudy_OnNewRow(ByVal Row As Long)
    If Val(ufgStudy.Text(Row, "分组ID")) = mlngGroupId Then Call ufgStudy.SetRowCheck(Row, True)
End Sub

