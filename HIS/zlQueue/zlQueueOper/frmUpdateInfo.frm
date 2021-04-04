VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUpdateInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "修改信息"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4125
   Icon            =   "frmUpdateInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtData 
      Height          =   300
      Index           =   0
      Left            =   1320
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ComboBox cboData 
      Height          =   300
      Index           =   0
      Left            =   1320
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   600
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ComboBox cbx诊室 
      Height          =   300
      Left            =   7080
      TabIndex        =   13
      Top             =   2760
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker dtp排队时间 
      Height          =   375
      Left            =   7080
      TabIndex        =   12
      Top             =   3705
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd hh:mm:ss"
      Format          =   50987011
      CurrentDate     =   41836
   End
   Begin VB.TextBox txt备注 
      Height          =   350
      Left            =   7080
      TabIndex        =   9
      Top             =   3210
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox txt排队号码 
      Height          =   350
      Left            =   7080
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox txt患者姓名 
      Height          =   350
      Left            =   7080
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox txt排队标记 
      Height          =   350
      Left            =   7080
      TabIndex        =   2
      Top             =   4215
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消(&C)"
      Height          =   385
      Left            =   2760
      TabIndex        =   1
      Top             =   3240
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确 定(&S)"
      Height          =   385
      Left            =   1320
      TabIndex        =   0
      Top             =   3240
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtpData 
      Height          =   300
      Index           =   0
      Left            =   1320
      TabIndex        =   17
      Top             =   1080
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd hh:mm:ss"
      Format          =   50987011
      CurrentDate     =   41836
   End
   Begin VB.Label labData 
      Caption         =   "----------"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "排队时间："
      Height          =   255
      Left            =   6120
      TabIndex        =   11
      Top             =   3810
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "备    注："
      Height          =   255
      Left            =   6120
      TabIndex        =   10
      Top             =   3285
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "排队号码："
      Height          =   255
      Left            =   6120
      TabIndex        =   8
      Top             =   2370
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "患者姓名："
      Height          =   255
      Left            =   6120
      TabIndex        =   6
      Top             =   1890
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "诊    室： "
      Height          =   255
      Left            =   6120
      TabIndex        =   5
      Top             =   2820
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "排队标记："
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   4290
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmUpdateInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public str患者姓名 As String
Public str排队号码 As String
Public str诊室 As String
Public str备注 As String
Public str排队时间 As String
Public str排队标记 As String

Private mobjInputCfg As Dictionary
Private mobjReturn As Dictionary

Public blnOk As Boolean
Private mlngQueueId As Long

Private mobjQueueOper As clsQueueOperation


Private Sub cmdCancel_Click()
'取消更新
On Error GoTo errHandle
    blnOk = False
    Unload Me
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Function LoadQueueInf(ByVal lngQueueId As Long) As Boolean
'载入选择的对立信息
    Dim rsData As ADODB.Recordset
    
    LoadQueueInf = False
    
    Set rsData = mobjQueueOper.GetQueueInf(lngQueueId)
    If rsData.RecordCount <= 0 Then Exit Function
    
    txt患者姓名.Text = Nvl(rsData!患者姓名)
    txt排队号码.Text = Nvl(rsData!排队号码)
    cbx诊室.Text = Nvl(rsData!诊室)
    txt备注.Text = Nvl(rsData!备注)
    dtp排队时间.value = IIf(Nvl(rsData!排队时间) = "", Now, Nvl(rsData!排队时间))
    txt排队标记.Text = Nvl(rsData!排队标记)
    
    LoadQueueInf = True
    
End Function

Private Sub InitInputFace(objInputCfg As Dictionary, ByRef strQueryField As String)
'初始化录入界面
    Dim i As Long
    Dim j As Long
    Dim strKey As Variant
    Dim strContext As String
    Dim strType As String
    Dim aryData() As String
    Dim blnCombobox As Boolean
    Dim objLastInput As Object
    Dim objCurInput As Object
    
    If objInputCfg.Count <= 0 Then Exit Sub
    
    i = 1
    
    Set objLastInput = Nothing
    
    For Each strKey In objInputCfg.Keys
        strContext = objInputCfg.Item(strKey)
        blnCombobox = False
        
        If strQueryField <> "" Then strQueryField = strQueryField & ","
        strQueryField = strQueryField & strKey
        
        If InStr(strContext, ":") > 0 Then
            strType = Mid(strContext, 1, InStr(strContext, ":") - 1)
            aryData = Split(Replace(strContext, strType & ":", ""), ",")
            
            blnCombobox = True
        Else
            strType = strContext
        End If
        
        If blnCombobox Then
            Load cboData(i)
  
            For j = 0 To UBound(aryData)
                If aryData(j) <> "" Then cboData(i).AddItem aryData(j)
            Next j
    
            Set objCurInput = cboData(i)
        Else
            Select Case UCase(strType)
                Case "STRING", "NUMBER"
                    Load txtData(i)
                    
                    Set objCurInput = txtData(i)
                Case "DATE", "DATETIME"
                    Load dtpData(i)
                    
                    dtpData(i).CustomFormat = IIf(UCase(strType) = "DATE", "yyyy-MM-dd", "yyyy-MM-dd hh:mm:ss")
                    
                    Set objCurInput = dtpData(i)
            End Select
        End If
        
        objCurInput.Left = 1320
        objCurInput.Tag = strKey
        
        If objLastInput Is Nothing Then
            objCurInput.Top = 240
        Else
            objCurInput.Top = objLastInput.Top + objLastInput.Height + 120
        End If
        objCurInput.Visible = True
        
        Set objLastInput = objCurInput
        
        Load labData(i)
        
        labData(i).Top = objCurInput.Top + 60
        labData(i).Left = 360
        labData(i).Caption = strKey
        labData(i).Visible = True
        
        i = i + 1
    Next
    
    If Not (objLastInput Is Nothing) Then
        cmdOK.Left = 1320
        cmdOK.Top = objLastInput.Top + objLastInput.Height + 120
        
        cmdCancel.Left = 2760
        cmdCancel.Top = cmdOK.Top
        
        Me.Height = cmdOK.Top + cmdOK.Height + 120 + 600
    End If
End Sub


Private Sub LoadInputValue(ByVal lngQueueId As Long, ByVal strQueryField As String)
'载入对应的数据值
    Dim rsData As ADODB.Recordset
    Dim objInputCon As Object
    
    Set rsData = mobjQueueOper.GetQueueInf(lngQueueId, strQueryField)
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    '配置text值
    For Each objInputCon In txtData
        If Not objInputCon Is Nothing Then
            If objInputCon.Tag <> "" Then objInputCon.Text = Nvl(rsData(objInputCon.Tag))
        End If
    Next
    
    '配置cbo值
    For Each objInputCon In cboData
        If Not objInputCon Is Nothing Then
            If objInputCon.Tag <> "" Then objInputCon.Text = Nvl(rsData(objInputCon.Tag))
        End If
    Next
    
    '配置日期值
    For Each objInputCon In dtpData
        If Not objInputCon Is Nothing Then
            If objInputCon.Tag <> "" Then objInputCon.value = IIf(Nvl(rsData(objInputCon.Tag)) = "", Now, Nvl(rsData(objInputCon.Tag)))
        End If
    Next
End Sub

Public Function zlShowMe(ByVal lngQueueId As Long, objInputCfg As Dictionary, objReturn As Dictionary, objQueueOper As clsQueueOperation, objOwner As Object) As Boolean
    Dim i As Long
    Dim aryRoom() As String
    Dim strQueryField As String
    
    zlShowMe = False
    
    blnOk = False
    mlngQueueId = lngQueueId
    
    Set mobjInputCfg = objInputCfg
    Set mobjReturn = objReturn
    
    Set mobjQueueOper = objQueueOper
    

    Call InitInputFace(objInputCfg, strQueryField)
    
    Call LoadInputValue(lngQueueId, strQueryField)
    
    Call LoadQueueInf(lngQueueId)
        
    Me.Show 1, objOwner
    
    zlShowMe = blnOk
End Function


Private Sub GetUpdateValue(objReturn As Dictionary, ByRef strUpdate As String)
'获取更新的值
    Dim objInputCon As Object
    
    '配置text值
    For Each objInputCon In txtData
        If Not objInputCon Is Nothing Then
            If objInputCon.Tag <> "" Then
                objReturn.Add CStr(objInputCon.Tag), objInputCon.Text
                If strUpdate <> "" Then strUpdate = strUpdate & ","
                strUpdate = strUpdate & objInputCon.Tag & "='" & objInputCon.Text & "'"
            End If
        End If
    Next
    
    '配置cbo值
    For Each objInputCon In cboData
        If Not objInputCon Is Nothing Then
            If objInputCon.Tag <> "" Then
                objReturn.Add objInputCon.Tag, objInputCon.Text
                If strUpdate <> "" Then strUpdate = strUpdate & ","
                strUpdate = strUpdate & objInputCon.Tag & "='" & objInputCon.Text & "'"
            End If
        End If
    Next
    
    '配置日期值
    For Each objInputCon In dtpData
        If Not objInputCon Is Nothing Then
            If objInputCon.Tag <> "" Then
                objReturn.Add objInputCon.Tag, objInputCon.value
                If strUpdate <> "" Then strUpdate = strUpdate & ","
                strUpdate = strUpdate & objInputCon.Tag & "=" & To_Date(objInputCon.value)
            End If
        End If
    Next
    
    strUpdate = Replace(strUpdate, "'", "''")
End Sub


Private Sub cmdOK_Click()
'确认更新
On Error GoTo errHandle
    Dim strUpdate As String
    
    Call GetUpdateValue(mobjReturn, strUpdate)
                
    Call mobjQueueOper.UpdateQueue(mlngQueueId, strUpdate)

    blnOk = True
    
    Unload Me
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

