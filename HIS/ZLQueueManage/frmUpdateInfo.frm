VERSION 5.00
Begin VB.Form frmUpdateInfo 
   Caption         =   "修改队列信息"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4365
   Icon            =   "frmUpdateInfo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   4365
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4095
      Begin VB.ComboBox cbo医生 
         Height          =   300
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1800
         Width           =   2295
      End
      Begin VB.ComboBox cbo诊室 
         Height          =   300
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txt患者姓名 
         Height          =   350
         Left            =   1560
         TabIndex        =   6
         Top             =   840
         Width           =   2295
      End
      Begin VB.ComboBox cboQueueName 
         Height          =   300
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "医生姓名"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "诊室 "
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "患者姓名"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "队列名称"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   350
      Left            =   2280
      TabIndex        =   1
      Top             =   2640
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   350
      Left            =   480
      TabIndex        =   0
      Top             =   2640
      Width           =   1100
   End
End
Attribute VB_Name = "frmUpdateInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mrsRoomData As ADODB.Recordset
Public mrsDoctorData As ADODB.Recordset
Public mlngCurrentQueueId As Long
Public mblnIsAllowChange As Boolean
Public mblnIsAlreadyProcess As Boolean


Private Const C_STR_MSGINF As String = "修改排队信息"

Private mstr队列名称 As String
Private mstr患者姓名 As String
Private mstr诊室 As String
Private mstr医生姓名 As String

Public Event OnQueueChange(ByVal lngQueueId As Long, ByVal strQueue As String, ByVal strPatient As String, ByVal strRoom As String, ByVal strDoctor As String, ByRef blnIsAllowChange As Boolean, ByRef blnIsAlreadyProcess As Boolean)


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub LoadQueueName(ByRef astr队列名称() As String)
    
End Sub

Public Function zlShowMe(frmParent As Form, ByRef astr队列名称() As String, ByRef str队列名称 As String, str患者姓名 As String, _
            ByRef str诊室 As String, ByRef str医生姓名 As String) As Boolean
    Dim i As Integer
    
    mstr队列名称 = str队列名称
    mstr患者姓名 = str患者姓名
    mstr诊室 = str诊室
    mstr医生姓名 = str医生姓名

    On Error GoTo err
    
    cboQueueName.Clear
    
    If SafeArrayGetDim(astr队列名称) <> 0 Then
        For i = 1 To UBound(astr队列名称)
            cboQueueName.AddItem astr队列名称(i)
            If astr队列名称(i) = str队列名称 Then cboQueueName.ListIndex = i - 1
        Next i
        
        If cboQueueName.ListIndex = -1 Then Exit Function
        
        txt患者姓名 = mstr患者姓名
        
        '加载诊室cbo内容
        cbo诊室.Clear
        If Not mrsRoomData Is Nothing Then
            If mrsRoomData.RecordCount < 1 Then
                cbo诊室.Enabled = False
                MsgBox "备选诊室无数据", vbInformation, C_STR_MSGINF
            End If
            For i = 1 To mrsRoomData.RecordCount
                cbo诊室.AddItem Nvl(mrsRoomData!RoomCode) & "-" & Nvl(mrsRoomData!RoomName)
                cbo诊室.ItemData(i - 1) = Nvl(mrsRoomData!RoomID)
                
                If Nvl(mrsRoomData!RoomName) = mstr诊室 Then
                    cbo诊室.ListIndex = i - 1
                End If
                mrsRoomData.MoveNext
            Next
        Else
            cbo诊室.Enabled = False
            MsgBox "备选诊室无数据", vbInformation, C_STR_MSGINF
        End If
        
        '加载医生cbo内容
        cbo医生.Clear
        If Not mrsDoctorData Is Nothing Then
        
            If mrsDoctorData.RecordCount < 1 Then
                cbo医生.Enabled = False
                MsgBox "备选医生无数据", vbInformation, C_STR_MSGINF
            End If
            
            For i = 1 To mrsDoctorData.RecordCount
                cbo医生.AddItem Nvl(mrsDoctorData!DoctorIdCode) & "-" & Nvl(mrsDoctorData!DoctorIdName)
                cbo医生.ItemData(i - 1) = Nvl(mrsDoctorData!DoctorId)
                
                If Nvl(mrsDoctorData!DoctorIdName) = mstr医生姓名 Then
                    cbo医生.ListIndex = i - 1
                End If
                mrsDoctorData.MoveNext
            Next
        Else
            cbo医生.Enabled = False
            MsgBox "备选医生无数据", vbInformation, C_STR_MSGINF
        End If

        
        Me.Show 1, frmParent

        If mstr队列名称 <> str队列名称 Or mstr患者姓名 <> str患者姓名 Or mstr诊室 <> str诊室 Or mstr医生姓名 <> str医生姓名 Then
            str队列名称 = mstr队列名称
            str患者姓名 = mstr患者姓名
            str医生姓名 = mstr医生姓名
            str诊室 = mstr诊室

            zlShowMe = True
            
        End If
    End If
       
    
      
    Exit Function
    
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function getNameByCbo(ByVal strText As String) As String
'获得cbo选中内容的名字
    On Error GoTo errh
    
    getNameByCbo = ""
    If InStr(strText, "-") < 1 Then Exit Function
   
    getNameByCbo = Mid(strText, InStr(strText, "-") + 1, Len(strText))
    Exit Function
    
errh:
    Resume
    getNameByCbo = ""
End Function

Private Function getCodeByCbo(ByVal strText As String) As Long
'获得cbo选中内容的简码
    On Error GoTo errh
    
    getCodeByCbo = 0
    If InStr(strText, "-") < 1 Then Exit Function
    
    getCodeByCbo = Val(Mid(strText, 1, InStr(strText, "-") - 1))
    Exit Function
      
errh:
    Resume
    getCodeByCbo = 0
End Function

Private Sub cmdOK_Click()
    
    mstr队列名称 = cboQueueName.Text
    mstr患者姓名 = txt患者姓名.Text
    
    If mstr医生姓名 <> getNameByCbo(cbo医生.Text) And cbo医生.Enabled = True Then mstr医生姓名 = getNameByCbo(cbo医生.Text)

    If mstr诊室 <> getNameByCbo(cbo诊室.Text) And cbo诊室.Enabled = True Then mstr诊室 = getNameByCbo(cbo诊室.Text)

    RaiseEvent OnQueueChange(mlngCurrentQueueId, mstr队列名称, mstr患者姓名, mstr诊室, mstr医生姓名, mblnIsAllowChange, mblnIsAlreadyProcess)
    
    If mblnIsAllowChange = False Then Exit Sub
    
    Unload Me
End Sub

