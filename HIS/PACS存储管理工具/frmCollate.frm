VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCollate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ͼ��У��"
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
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   60
      TabIndex        =   2
      Top             =   30
      Width           =   9315
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   8040
         TabIndex        =   9
         Top             =   300
         Width           =   1100
      End
      Begin VB.CommandButton CmdModify 
         Caption         =   "����(&M)"
         Height          =   350
         Left            =   6750
         TabIndex        =   8
         Top             =   300
         Width           =   1100
      End
      Begin VB.CommandButton CmdCollate 
         Caption         =   "У��(&L)"
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
         Caption         =   "��"
         Height          =   180
         Left            =   2880
         TabIndex        =   6
         Top             =   390
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��"
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
    
    strSQL = "Select d.ip��ַ As IP��ַ1,d.ftp�û��� As �û���1 , d.ftp����  As ����1,decode(d.FTPĿ¼,'','/','/'||d.FTPĿ¼||'/') as FTPĿ¼1," & _
             " decode(λ��һ,'','',to_char(a.��������,'YYYYMMDD')||'/'||a.���UID ) As ����Ŀ¼1, " & _
             " e.ip��ַ As IP��ַ2,e.ftp�û��� As �û���2 , e.ftp����  As ����2,decode(e.FTPĿ¼,'','/','/'||e.FTPĿ¼||'/') as FTPĿ¼2," & _
             " decode(λ�ö�,'','',to_char(a.��������,'YYYYMMDD')||'/'||a.���UID ) As ����Ŀ¼2, " & _
             " c.ͼ��uid " & _
             " From Ӱ�����¼ a , Ӱ�������� b , Ӱ����ͼ�� c , Ӱ���豸Ŀ¼ d , Ӱ���豸Ŀ¼ e " & _
             " Where a.���uid = b.���uid And b.����uid = c.����uid And a.λ��һ = d.�豸��(+) And a.λ�ö� = e.�豸��(+) " & _
             "       And �������� Between [1] and [2]"
                         
    Set rsTmp = OpenSQLRecord(strSQL, "ͼ��洢����", CDate(Format(Me.DTPBegin, "yyyy-MM-dd")), CDate(Format(Me.DTPEnd, "yyyy-MM-dd")))
    
    ProgressBar.Value = 0
    ProgressBar.Min = 0
    
    If rsTmp.RecordCount > 0 Then
        ProgressBar.Max = rsTmp.RecordCount
    End If
    
    lvwCollateList.ListItems.Clear
    Me.MousePointer = 11
    zl9comlib.ZlCommFun.ShowFlash "���Ժ�����У���ļ�", Me
    Do While Not rsTmp.EOF
        
        ProgressBar.Value = ProgressBar.Value + 1
        
        If ZlCommFun.NVL(rsTmp("����Ŀ¼1")) <> "" Then
            clsCheckFtpFile.strIPAddress = rsTmp("IP��ַ1")
            clsCheckFtpFile.strUser = ZlCommFun.NVL(rsTmp("�û���1"))
            clsCheckFtpFile.strPsw = ZlCommFun.NVL(rsTmp("����1"))
            If clsCheckFtpFile.FuncFileExist(rsTmp("FTPĿ¼1") & rsTmp("����Ŀ¼1"), rsTmp("ͼ��UID")) = False Then
                With lvwCollateList
                    Set objItem = .ListItems.Add(, "A" & rsTmp("FTPĿ¼1") & rsTmp("����Ŀ¼1") & rsTmp("ͼ��UID"), rsTmp("IP��ַ1"))
                    objItem.SubItems(1) = IIf(rsTmp("FTPĿ¼1") = "/", "", rsTmp("FTPĿ¼1"))
                    objItem.SubItems(2) = rsTmp("����Ŀ¼1")
                    objItem.SubItems(3) = rsTmp("ͼ��UID")
                    objItem.Tag = ZlCommFun.NVL(rsTmp("�û���1")) & ";" & ZlCommFun.NVL(rsTmp("����1"))
                End With
            End If
        End If
        
        If ZlCommFun.NVL(rsTmp("����Ŀ¼2")) <> "" Then
            clsCheckFtpFile.strIPAddress = rsTmp("IP��ַ2")
            clsCheckFtpFile.strUser = ZlCommFun.NVL(rsTmp("�û���2"))
            clsCheckFtpFile.strPsw = ZlCommFun.NVL(rsTmp("����2"))
            If clsCheckFtpFile.FuncFileExist(rsTmp("FTPĿ¼2") & rsTmp("����Ŀ¼2"), rsTmp("ͼ��UID")) = False Then
                With lvwCollateList
                    Set objItem = .ListItems.Add(, "A" & rsTmp("FTPĿ¼2") & rsTmp("����Ŀ¼2") & rsTmp("ͼ��UID"), rsTmp("IP��ַ2"))
                    objItem.SubItems(1) = IIf(rsTmp("FTPĿ¼2") = "/", "", rsTmp("FTPĿ¼2"))
                    objItem.SubItems(2) = rsTmp("����Ŀ¼2")
                    objItem.SubItems(3) = rsTmp("ͼ��UID")
                    objItem.Tag = ZlCommFun.NVL(rsTmp("�û���2")) & ";" & ZlCommFun.NVL(rsTmp("����2"))
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
    zl9comlib.ZlCommFun.ShowFlash "���������ļ��������ĵȴ���", Me
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
        .ColumnHeaders.Add , "A", "IP��ַ", intAvgWidth
        .ColumnHeaders.Add , "B", "FTPĿ¼", intAvgWidth
        .ColumnHeaders.Add , "C", "����Ŀ", intAvgWidth * 4
        .ColumnHeaders.Add , "D", "�ļ���", intAvgWidth * 2
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
