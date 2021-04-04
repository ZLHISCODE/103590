VERSION 5.00
Begin VB.Form frmAddFile 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "数据文件添加"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   6525
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2460
      Width           =   6525
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Default         =   -1  'True
         Height          =   345
         Left            =   4080
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   345
         Left            =   5280
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblPgs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   195
         Width           =   45
      End
      Begin VB.Line lineBottom 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   10000
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line lineBottom 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   10000
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.TextBox txtDataFile 
      Height          =   300
      Left            =   1710
      TabIndex        =   2
      Top             =   1560
      Width           =   3945
   End
   Begin VB.CheckBox chkSpaceExtd 
      Caption         =   "自动扩展空间"
      Height          =   270
      Left            =   2880
      TabIndex        =   5
      ToolTipText     =   "AUTOEXTEND ON Next (表空间大小/10)M"
      Top             =   1965
      Value           =   1  'Checked
      Width           =   1425
   End
   Begin VB.TextBox txtSpaceSize 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1710
      MaxLength       =   6
      TabIndex        =   3
      Text            =   "500"
      Top             =   1950
      Width           =   735
   End
   Begin VB.TextBox txtTableSpace 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1710
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   2160
   End
   Begin VB.TextBox txtFileAmount 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   1710
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "1"
      Top             =   1230
      Width           =   300
   End
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "（文件名按文件末尾数字递增）"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2640
      TabIndex        =   13
      Top             =   1290
      Width           =   2520
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "为当前表空间添加数据文件"
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   240
      Width           =   3975
   End
   Begin VB.Image Img 
      Height          =   480
      Left            =   240
      Picture         =   "frmAddFile.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblDataFile 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "第一个文件"
      Height          =   180
      Left            =   720
      TabIndex        =   10
      Top             =   1620
      Width           =   900
   End
   Begin VB.Label lblBakSpace 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "数据表空间名"
      Height          =   225
      Left            =   480
      TabIndex        =   8
      Top             =   900
      Width           =   1125
   End
   Begin VB.Label lblFileAmount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "共添加          个文件"
      Height          =   195
      Index           =   0
      Left            =   1065
      TabIndex        =   4
      Top             =   1290
      Width           =   1530
   End
   Begin VB.Label lblFileSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "初始大小                     M"
      Height          =   195
      Left            =   855
      TabIndex        =   9
      Top             =   2010
      Width           =   1785
   End
End
Attribute VB_Name = "frmAddFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnCreate As Boolean

Public Function ShowAddFile(ByVal strTableSpace As String) As Boolean
    
    txtTableSpace.Text = strTableSpace
    txtDataFile.Text = GetFileName(, strTableSpace)
    
    Me.Show 1
    ShowAddFile = mblnCreate
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function GetFileName(Optional ByVal strFile As String, Optional ByVal strTableSpace As String) As String
    '根据当前的数据文件名称,获取下一个数据文件
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strTmp As String, i As Integer
    
    If strFile = "" Then
        strSQL = "Select Max(File_Name) Max_File From Dba_Data_Files Where Tablespace_Name =[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "获取数据文件名", strTableSpace)
        strFile = rsTmp!Max_file
    End If
    
    If InStr(1, strFile, ".DBF") > 0 Then
        strFile = Left(strFile, InStr(1, strFile, ".DBF") - 1)
    End If
    
    If IsNumeric(Right(strFile, 4)) Then
        '后四位为数字,可能是形如 ZLHD2017\2018 这种按年份为规则的备份数据文件
        strFile = strFile & "_01.DBF"
    Else
        '否则,取末端数字+1
        i = 1
        Do While IsNumeric(Right(strFile, i))
            i = i + 1
        Loop
        
        If i = 1 Then
            '没有数字
            strFile = strFile & "01.DBF"
        Else
            strTmp = Format(Val(Right(strFile, i - 1)) + 1, Lpad("", i - 1, "0"))
            strFile = Left(strFile, Len(strFile) - i + 1) & strTmp & ".DBF"
        End If
    End If
    
    GetFileName = strFile
End Function


Private Sub cmdOK_Click()
    
    '数据检查
    If Len(Trim(txtDataFile.Text)) = 0 Then
        MsgBox "请定义" & txtTableSpace.Text & "表空间的数据文件。", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If Val(txtSpaceSize.Text) > 32000 Then
        MsgBox "表空间" & txtTableSpace.Text & "超过32G了。", vbExclamation, gstrSysName
        Exit Sub
    End If

    
    If AddDatafile(txtTableSpace.Text, txtDataFile.Text, txtFileAmount.Text, txtSpaceSize.Text, chkSpaceExtd.Value) Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub txtDataFile_GotFocus()
    txtDataFile.SelStart = Len(txtDataFile.Text)
End Sub

Private Sub txtDataFile_KeyPress(KeyAscii As Integer)
    OnlyStrCK KeyAscii, "\", "_", "/"
End Sub

Private Sub txtFileAmount_GotFocus()
    txtFileAmount.SelStart = Len(txtFileAmount.Text)
End Sub

Private Sub txtFileAmount_KeyPress(KeyAscii As Integer)
    OnlyIntCK KeyAscii
End Sub

Private Sub txtSpaceSize_GotFocus()
    txtSpaceSize.SelStart = Len(txtSpaceSize.Text)
End Sub

Private Function AddDatafile(ByVal strTableSpace As String, ByVal strFile As String, ByVal intNum As Integer, ByVal lngSize As Long, ByVal blnAutoExtend As Boolean) As Boolean
    '为表空间添加数据文件
    '参数:strTableSpace - 表空间名称,strFile - 首个文件名 , intNum - 添加文件个数 ,lngSize  - 初始文件大小, blnAutoExtend - 是否自动拓展
    Dim strErrMsg As String, strSQL As String
    Dim strNextFile As String, i As Integer, strTmp As String
    
    On Error Resume Next
    
    lblPgs.Caption = "正在创建数据文件．．．"
    
    For i = 1 To intNum
        If strNextFile = "" Then
            strNextFile = strFile
        Else
            strNextFile = GetFileName(strNextFile)
        End If
        
        strTmp = IIf(InStr(1, strNextFile, "\") > 0, "\", "/")
        strTmp = Mid(strNextFile, InStrRev(strNextFile, strTmp) + 1, InStr(1, strNextFile, ".") - 1)
        lblPgs.Caption = "正在创建数据文件" & strTmp & "．．．"
        lblPgs.Refresh
        
        strSQL = "Alter TableSpace " & strTableSpace & " Add DataFile '" & strNextFile & "' Size " & lngSize & "M  AutoExtend  " & IIf(blnAutoExtend, "On", "Flase")
        gcnOracle.Execute strSQL
        
        If Err.Number <> 0 Then
            strErrMsg = "添加数据文件 " & strTmp & "发生错误， 错误原因 ：" & vbNewLine & Err.Description
            
             If MsgBox(strErrMsg & vbNewLine & "是否继续创建其他数据文件？点击是将继续，点击取消将退出当前操作。", vbYesNo, "错误") = vbYes Then
                strErrMsg = ""
                Err.Clear
            Else
                lblPgs.Caption = "操作被取消"
                Exit Function
            End If
        End If
    Next
    
    mblnCreate = True
    AddDatafile = mblnCreate
End Function

Private Sub txtSpaceSize_KeyPress(KeyAscii As Integer)
    OnlyIntCK KeyAscii
End Sub
