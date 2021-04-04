VERSION 5.00
Begin VB.Form frm等待返回北京尚洋 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "等待返回......"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdReadCenter 
      Cancel          =   -1  'True
      Caption         =   "读取(&R)"
      Height          =   400
      Left            =   1800
      TabIndex        =   2
      Top             =   1320
      Width           =   1100
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   2097
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   630
      Width           =   1965
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   4080
      TabIndex        =   4
      Top             =   1320
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   2940
      TabIndex        =   3
      Top             =   1320
      Width           =   1100
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   23
      Picture         =   "frm等待返回北京尚洋.frx":0000
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   355
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1065
      Width           =   5325
   End
   Begin VB.Timer TimeAvi 
      Interval        =   50
      Left            =   240
      Top             =   120
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "收费号码"
      Height          =   180
      Left            =   1309
      TabIndex        =   0
      Top             =   690
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "等待医保返回结算数据......"
      Height          =   180
      Left            =   1309
      TabIndex        =   6
      Top             =   255
      Width           =   2340
   End
End
Attribute VB_Name = "frm等待返回北京尚洋"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFile As String
Private mbytType As Integer

Public Function waitReturn(strFile As String, bytType As Integer) As String
    mstrFile = strFile
    mbytType = bytType
    
    Me.Show vbModal
    waitReturn = mstrFile
End Function

Private Sub cmdCancel_Click()
    If MsgBox("你确定要取消吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    mstrFile = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
'    Dim rsTemp As New ADODB.Recordset
'    mstrFile = Text1.Text
'    Set rsTemp = gcn尚洋.Execute("Select * From MED_RECEIPT_RECORD_MASTER Where RESIDENCE_NO='" & mstrFile & "'")
'    If rsTemp.EOF Then
'        MsgBox "中间数据库中没有找到指定收费号码的数据", vbInformation, gstrSysName
'        Exit Sub
'    End If
'    Me.Hide
 Dim rsTemp As New ADODB.Recordset
    mstrFile = Text1.Text
 
    ' Set rsTemp = gcn尚洋.Execute("Select * From MED_RECEIPT_RECORD_MASTER Where RESIDENCE_NO='" & mstrFile & "'")
'    范志英修改06-22-15:43(结算完毕之后读取)
    Set rsTemp = gcn尚洋.Execute("Select * From MED_RECEIPT_RECORD_MASTER Where CHARGE_NUMBER='" & mstrFile & "'")
    If rsTemp.EOF Then
        MsgBox "中间数据库中没有找到指定收费号码的数据", vbInformation, gstrSysName
        Exit Sub
    End If
    Me.Hide
End Sub

Private Sub cmdReadCenter_Click()
    Dim rsTemp As New ADODB.Recordset
    If cmdOK.Enabled = False Then
        mstrFile = Text1.Text
        
        If gint是否职工 = 0 Then
           Set rsTemp = gcn尚洋.Execute("Select * From MED_RECEIPT_RECORD_MASTER Where VISIT_NUMBER='" & mstrFile & "'")
        Else
            If mbytType = 1 And gint是否职工 = 1 Then
                '住院职工医保新的读取住院结算结果方式：用住院号+医保机构代码查询
                Set rsTemp = gcn尚洋.Execute("Select * From MED_RECEIPT_RECORD_MASTER Where RESIDENCE_NO='" & mstrFile & "'" & _
                   " AND HANDLE_DATE >= SYSDATE - 3 AND HANDLE_DATE < SYSDATE + 1 AND HOSPITAL_NUMBER ='" & gstr医保机构编码 & "'")
            Else
                
                Set rsTemp = gcn尚洋.Execute("Select * From MED_RECEIPT_RECORD_MASTER Where VISIT_NUMBER='" & mstrFile & "'" & _
                   " AND HANDLE_DATE >= SYSDATE - 3 AND HANDLE_DATE < SYSDATE + 1 AND HOSPITAL_NUMBER ='" & gstr医保机构编码 & "'")
            End If
        End If
        
        If rsTemp.EOF Then
            Exit Sub
        End If
        cmdOK.Enabled = True
        Text1.Text = rsTemp!CHARGE_NUMBER
        mstrFile = Text1.Text
    End If
    cmdReadCenter.Enabled = Not cmdOK.Enabled
End Sub

Private Sub Form_Load()
    cmdOK.Enabled = False
    Text1.Text = mstrFile
End Sub

Private Sub TimeAvi_Timer()
    Static i As Long
    i = i + 20
    If i >= Picture1.ScaleWidth Then i = 1
    
    Picture1.PaintPicture Picture1.Picture, i, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight, 0, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight
    Picture1.PaintPicture Picture1.Picture, 0, 0, i, Picture1.ScaleHeight, Picture1.ScaleWidth - i, 0, i, Picture1.ScaleHeight
End Sub
