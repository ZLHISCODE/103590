VERSION 5.00
Begin VB.Form frmFileRequest 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6555
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picTitle 
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   6165
      TabIndex        =   1
      Top             =   0
      Width           =   6165
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "总检医生:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   105
         TabIndex        =   5
         Top             =   60
         Width           =   885
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   1
         Left            =   1005
         TabIndex        =   4
         Top             =   60
         Width           =   90
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "总检时间:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   3705
         TabIndex        =   3
         Top             =   60
         Width           =   885
      End
      Begin VB.Label lbl 
         Height          =   180
         Index           =   3
         Left            =   4620
         TabIndex        =   2
         Top             =   60
         Width           =   1440
      End
   End
   Begin zl9CISCore.ctrlPatientFile ProFile1 
      Height          =   4455
      Left            =   420
      TabIndex        =   0
      Top             =   795
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   7858
      Border_Width    =   0
   End
End
Attribute VB_Name = "frmFileRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnStartUp As Boolean
Private mlngKey As Long
Private mfrmMain As Object

Public Function zlMenuClick(ByVal frmMain As Object, ByVal lngKey As Long, ByVal strMenuItem As String) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能：
    '参数：lngKey 档案ID
    '--------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHand
    
    mlngKey = lngKey
    Set mfrmMain = frmMain
    
    Select Case strMenuItem
    Case "刷新"
    
        Call zlClearData
        Call RefreshData(strMenuItem)
        
    End Select
    
    zlMenuClick = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Sub zlClearData(Optional ByVal strPart As String = "所有")
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    
    ProFile1.ShowFile "", , , , -1 '清除病历内容
    lbl(1).Caption = ""
    lbl(3).Caption = ""
    
End Sub

Public Property Get Body(ByVal lngIndex As Long) As Object
'    Set Body = vsf
End Property


Private Function RefreshData(ByVal strMenu As String) As Boolean
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim blnDataMoved As Boolean
    
    On Error GoTo errHand
    
    Select Case strMenu
    Case "刷新"
        
        mfrmMain.MousePointer = vbHourglass
        DoEvents
        
        blnDataMoved = False
        If mlngKey = 0 Then
            Call zlClearData
        Else
            
            blnDataMoved = False
            strSQL = "Select 书写日期,书写人 From 病人病历记录 WHERE ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngKey)
            If rs.BOF Then
                blnDataMoved = True
            Else
                lbl(3).Caption = Format(zlCommFun.NVL(rs("书写日期").Value), "yyyy-MM-dd HH:mm")
                lbl(1).Caption = zlCommFun.NVL(rs("书写人").Value)
            End If

            strSQL = "SELECT 医嘱id from 病人医嘱发送 WHERE 报告id=[1]"
            If blnDataMoved Then
                strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
            End If
            
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngKey)
            If rs.BOF = False Then
                ProFile1.ShowFile mlngKey, , , , , , , , , rs("医嘱id").Value, , blnDataMoved
            Else
                ProFile1.ShowFile mlngKey, , , , , , , , , , , blnDataMoved
            End If

        End If
        mfrmMain.MousePointer = vbDefault
    End Select
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function InitLoad() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据，发生在窗体的Load事件
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHand
    
        
    InitLoad = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

'（３）窗体及其控件的事件处理******************************************************************************************
Private Sub Form_Activate()
    
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True
    
    Call InitLoad
        
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    picTitle.Move 0, 0, Me.ScaleWidth
    
    With ProFile1
        .Left = 0
        .Top = picTitle.Top + picTitle.Height
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - picTitle.Height
    End With
    
End Sub


Private Sub picTitle_Paint()
    zlControl.PicShowFlat picTitle, 1
End Sub

Private Sub picTitle_Resize()
    On Error Resume Next
    
    lbl(3).Move picTitle.Width - lbl(3).Width - 60, lbl(3).Top
    lbl(2).Move lbl(3).Left - lbl(2).Width, lbl(3).Top
    
End Sub
