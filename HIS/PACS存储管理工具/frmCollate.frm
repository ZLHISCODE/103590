VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCollate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "图像校对"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9495
   Icon            =   "frmCollate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   60
      TabIndex        =   2
      Top             =   30
      Width           =   9315
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   8040
         TabIndex        =   9
         Top             =   300
         Width           =   1100
      End
      Begin VB.CommandButton CmdModify 
         Caption         =   "修正(&M)"
         Height          =   350
         Left            =   6750
         TabIndex        =   8
         Top             =   300
         Width           =   1100
      End
      Begin VB.CommandButton CmdCollate 
         Caption         =   "校对(&L)"
         Height          =   350
         Left            =   5430
         TabIndex        =   7
         Top             =   300
         Width           =   1100
      End
      Begin MSComCtl2.DTPicker DTPBegin 
         Height          =   285
         Left            =   660
         TabIndex        =   3
         Top             =   330
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   503
         _Version        =   393216
         Format          =   96010241
         CurrentDate     =   38670
      End
      Begin MSComCtl2.DTPicker DTPEnd 
         Height          =   285
         Left            =   3150
         TabIndex        =   4
         Top             =   330
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   503
         _Version        =   393216
         Format          =   96010241
         CurrentDate     =   38670
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "到"
         Height          =   180
         Left            =   2880
         TabIndex        =   6
         Top             =   390
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "从"
         Height          =   180
         Left            =   390
         TabIndex        =   5
         Top             =   390
         Width           =   180
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Top             =   5760
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lvwCollateList 
      Height          =   4785
      Left            =   60
      TabIndex        =   1
      Top             =   930
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   8440
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmCollate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCollate_Click()
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim clsCheckFtpFile As New clsFTP
    Dim objItem As ListItem
    
    On Error GoTo FindFileError
    
    strSQL = "Select d.ip地址 As IP地址1,d.ftp用户名 As 用户名1 , d.ftp密码  As 密码1,decode(d.FTP目录,'','/','/'||d.FTP目录||'/') as FTP目录1," & _
             " decode(位置一,'','',to_char(a.接收日期,'YYYYMMDD')||'/'||a.检查UID ) As 虚拟目录1, " & _
             " e.ip地址 As IP地址2,e.ftp用户名 As 用户名2 , e.ftp密码  As 密码2,decode(e.FTP目录,'','/','/'||e.FTP目录||'/') as FTP目录2," & _
             " decode(位置二,'','',to_char(a.接收日期,'YYYYMMDD')||'/'||a.检查UID ) As 虚拟目录2, " & _
             " c.图像uid " & _
             " From 影像检查记录 a , 影像检查序列 b , 影像检查图象 c , 影像设备目录 d , 影像设备目录 e " & _
             " Where a.检查uid = b.检查uid And b.序列uid = c.序列uid And a.位置一 = d.设备号(+) And a.位置二 = e.设备号(+) " & _
             "       And 接收日期 Between [1] and [2]"
                         
    Set rsTmp = OpenSQLRecord(strSQL, "图像存储管理", CDate(Format(Me.DTPBegin, "yyyy-MM-dd")), CDate(Format(Me.DTPEnd, "yyyy-MM-dd")))
    
    ProgressBar.Value = 0
    ProgressBar.Min = 0
    
    If rsTmp.RecordCount > 0 Then
        ProgressBar.Max = rsTmp.RecordCount
    End If
    
    lvwCollateList.ListItems.Clear
    Me.MousePointer = 11
    zl9comlib.ZlCommFun.ShowFlash "请稍后正在校对文件", Me
    Do While Not rsTmp.EOF
        
        ProgressBar.Value = ProgressBar.Value + 1
        
        If ZlCommFun.NVL(rsTmp("虚拟目录1")) <> "" Then
            clsCheckFtpFile.strIPAddress = rsTmp("IP地址1")
            clsCheckFtpFile.strUser = ZlCommFun.NVL(rsTmp("用户名1"))
            clsCheckFtpFile.strPsw = ZlCommFun.NVL(rsTmp("密码1"))
            If clsCheckFtpFile.FuncFileExist(rsTmp("FTP目录1") & rsTmp("虚拟目录1"), rsTmp("图像UID")) = False Then
                With lvwCollateList
                    Set objItem = .ListItems.Add(, "A" & rsTmp("FTP目录1") & rsTmp("虚拟目录1") & rsTmp("图像UID"), rsTmp("IP地址1"))
                    objItem.SubItems(1) = IIf(rsTmp("FTP目录1") = "/", "", rsTmp("FTP目录1"))
                    objItem.SubItems(2) = rsTmp("虚拟目录1")
                    objItem.SubItems(3) = rsTmp("图像UID")
                    objItem.Tag = ZlCommFun.NVL(rsTmp("用户名1")) & ";" & ZlCommFun.NVL(rsTmp("密码1"))
                End With
            End If
        End If
        
        If ZlCommFun.NVL(rsTmp("虚拟目录2")) <> "" Then
            clsCheckFtpFile.strIPAddress = rsTmp("IP地址2")
            clsCheckFtpFile.strUser = ZlCommFun.NVL(rsTmp("用户名2"))
            clsCheckFtpFile.strPsw = ZlCommFun.NVL(rsTmp("密码2"))
            If clsCheckFtpFile.FuncFileExist(rsTmp("FTP目录2") & rsTmp("虚拟目录2"), rsTmp("图像UID")) = False Then
                With lvwCollateList
                    Set objItem = .ListItems.Add(, "A" & rsTmp("FTP目录2") & rsTmp("虚拟目录2") & rsTmp("图像UID"), rsTmp("IP地址2"))
                    objItem.SubItems(1) = IIf(rsTmp("FTP目录2") = "/", "", rsTmp("FTP目录2"))
                    objItem.SubItems(2) = rsTmp("虚拟目录2")
                    objItem.SubItems(3) = rsTmp("图像UID")
                    objItem.Tag = ZlCommFun.NVL(rsTmp("用户名2")) & ";" & ZlCommFun.NVL(rsTmp("密码2"))
                End With
            End If
        End If
        DoEvents
        rsTmp.MoveNext
    Loop
    Me.MousePointer = 0
    zl9comlib.ZlCommFun.StopFlash
    Exit Sub
FindFileError:
    Me.MousePointer = 0
    zl9comlib.ZlCommFun.StopFlash
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdModify_Click()
    Dim i As Long
    Dim objFtp As New clsFTP
    Dim objItem As ListItem
    Dim strSrcPath As String
    Dim strVirtualPath As String
    Dim strFileName As String
    Dim strUser As String
    Dim strPassword As String
    Dim strIP As String
    Dim intTmp As Integer
    Dim j As Integer
    
    
    On Error GoTo ModifyFileError
    
    If Me.lvwCollateList.ListItems.Count = 0 Then Exit Sub
    
    ProgressBar.Value = 0
    ProgressBar.Min = 0
    If lvwCollateList.ListItems.Count > 0 Then
        ProgressBar.Max = lvwCollateList.ListItems.Count
    Else
        Exit Sub
    End If
    Me.MousePointer = 11
    zl9comlib.ZlCommFun.ShowFlash "正在修正文件，请耐心等待！", Me
    With lvwCollateList
        For i = 1 To .ListItems.Count
            ProgressBar.Visible = ProgressBar.Visible + 1
            strVirtualPath = IIf(.ListItems(i - j).SubItems(1) = "", "/", .ListItems(i - j).SubItems(1))
            strFileName = .ListItems(i - j).SubItems(3)
            strIP = .ListItems(i - j).Text
            
            If InStr(.ListItems(i - j).Tag, ";") <> 0 Then
                strUser = Mid(.ListItems(i - j).Tag, 1, InStr(.ListItems(i - j).Tag, ";") - 1)
            End If
            If Len(Trim(.ListItems(i - j).Tag)) > InStr(.ListItems(i - j).Tag, ";") Then
                strPassword = Mid(.ListItems(i - j).Tag, InStr(.ListItems(i - j).Tag, ";") + 1)
            End If
            
            objFtp.strIPAddress = strIP
            objFtp.strUser = strUser
            objFtp.strPsw = strPassword
            
            strSrcPath = objFtp.FuncSearchFiles(strVirtualPath, strFileName)
            If Len(Trim(strSrcPath)) > 0 Then
                strSrcPath = Split(Split(strSrcPath, "|")(0), ";")(1)
                
                If objFtp.FuncDupDir(strSrcPath, strVirtualPath & .ListItems(i - j).SubItems(2), False) Then
                    .ListItems.Remove (i - j)
                    j = j + 1
                End If
            End If
            DoEvents
        Next
        Me.MousePointer = 0
    End With
    zl9comlib.ZlCommFun.StopFlash
    Exit Sub
ModifyFileError:
    Me.MousePointer = 0
    zl9comlib.ZlCommFun.StopFlash
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Sub

Private Sub Form_Load()
    Dim intAvgWidth As Integer
    With lvwCollateList
        intAvgWidth = .Width / 8
        .ColumnHeaders.Add , "A", "IP地址", intAvgWidth
        .ColumnHeaders.Add , "B", "FTP目录", intAvgWidth
        .ColumnHeaders.Add , "C", "虚拟目", intAvgWidth * 4
        .ColumnHeaders.Add , "D", "文件名", intAvgWidth * 2
    End With
    
    Me.DTPBegin = Now - 30
    Me.DTPEnd = Now
    
    Call RestoreWinState(Me, App.ProductName)
End Sub


Private Sub Form_Resize()
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub
