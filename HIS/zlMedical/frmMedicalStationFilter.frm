VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMedicalStationFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤查找"
   ClientHeight    =   5310
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8595
   Icon            =   "frmMedicalStationFilter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame3 
      Caption         =   "过滤条件"
      Height          =   4575
      Left            =   90
      TabIndex        =   33
      Top             =   90
      Width           =   5100
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   285
         Width           =   3855
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1140
         TabIndex        =   7
         Top             =   1005
         Width           =   3855
      End
      Begin VB.ComboBox cbo 
         BackColor       =   &H80000018&
         Height          =   300
         Index           =   4
         Left            =   1140
         TabIndex        =   25
         Top             =   3195
         Width           =   3855
      End
      Begin VB.ComboBox cbo 
         BackColor       =   &H80000018&
         Height          =   300
         Index           =   3
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2850
         Width           =   3855
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "X"
         Height          =   255
         Index           =   0
         Left            =   4725
         TabIndex        =   21
         Top             =   2535
         Width           =   255
      End
      Begin VB.ComboBox cbo 
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   9
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   645
         Width           =   1080
      End
      Begin VB.ComboBox cbo 
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   10
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1380
         Width           =   1080
      End
      Begin VB.TextBox txtCover 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   2265
         TabIndex        =   37
         Top             =   675
         Width           =   1215
      End
      Begin VB.TextBox txtCover 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Index           =   1
         Left            =   3750
         TabIndex        =   36
         Top             =   675
         Width           =   1215
      End
      Begin VB.TextBox txtCover 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Index           =   2
         Left            =   2265
         TabIndex        =   35
         Top             =   1410
         Width           =   1215
      End
      Begin VB.TextBox txtCover 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Index           =   3
         Left            =   3750
         TabIndex        =   34
         Top             =   1410
         Width           =   1215
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&P"
         Height          =   255
         Index           =   1
         Left            =   4455
         TabIndex        =   20
         Top             =   2535
         Width           =   255
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   7
         Left            =   1140
         TabIndex        =   13
         Top             =   1755
         Width           =   3855
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&P"
         Height          =   255
         Index           =   0
         Left            =   4455
         TabIndex        =   16
         Top             =   2160
         Width           =   255
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "X"
         Height          =   255
         Index           =   1
         Left            =   4725
         TabIndex        =   17
         Top             =   2160
         Width           =   255
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   5
         Left            =   3720
         TabIndex        =   5
         Top             =   645
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   75563011
         CurrentDate     =   38210
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   2
         Left            =   2235
         TabIndex        =   10
         Top             =   1380
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   75563011
         CurrentDate     =   38210
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   3
         Left            =   3720
         TabIndex        =   11
         Top             =   1380
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   75563011
         CurrentDate     =   38210
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   4
         Left            =   2235
         TabIndex        =   4
         Top             =   645
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   75563011
         CurrentDate     =   38210
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1140
         TabIndex        =   15
         Top             =   2130
         Width           =   3855
      End
      Begin VB.TextBox txt 
         BackColor       =   &H80000018&
         Height          =   300
         Index           =   15
         Left            =   1140
         TabIndex        =   19
         Top             =   2505
         Width           =   3855
      End
      Begin VB.ComboBox cbo 
         BackColor       =   &H80000018&
         Height          =   300
         Index           =   5
         Left            =   1140
         TabIndex        =   26
         Top             =   3540
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "体检科室(&1)"
         Height          =   180
         Index           =   4
         Left            =   105
         TabIndex        =   0
         Top             =   345
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "体 检 单(&3)"
         Height          =   180
         Index           =   3
         Left            =   105
         TabIndex        =   6
         Top             =   1050
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Index           =   8
         Left            =   3525
         TabIndex        =   39
         Top             =   705
         Width           =   180
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "体检日期(&2)"
         Height          =   180
         Index           =   10
         Left            =   105
         TabIndex        =   2
         Top             =   705
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Index           =   11
         Left            =   3525
         TabIndex        =   38
         Top             =   1455
         Width           =   180
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "比较方式(&8)"
         Height          =   180
         Index           =   23
         Left            =   90
         TabIndex        =   22
         Top             =   2910
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "体检结果(&9)"
         Height          =   180
         Index           =   24
         Left            =   90
         TabIndex        =   24
         Top             =   3255
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "体检项目(&7)"
         Height          =   180
         Index           =   12
         Left            =   90
         TabIndex        =   18
         Top             =   2565
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "完成时间(&4)"
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   8
         Top             =   1425
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "病人姓名(&5)"
         Height          =   180
         Index           =   20
         Left            =   90
         TabIndex        =   12
         Top             =   1815
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "团体名称(&6)"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   14
         Top             =   2190
         Width           =   990
      End
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   3855
      Top             =   4860
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2685
      Top             =   4860
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationFilter.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationFilter.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationFilter.frx":1178
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationFilter.frx":696A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationFilter.frx":737C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationFilter.frx":CE66
            Key             =   "Query"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "常用过滤查找(&Y)"
      Height          =   4575
      Left            =   5205
      TabIndex        =   30
      Top             =   90
      Width           =   3330
      Begin MSComctlLib.ListView lvw 
         Height          =   3855
         Left            =   90
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   210
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   6800
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         Icons           =   "ils16"
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   4974
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbr 
         Height          =   330
         Left            =   90
         TabIndex        =   32
         Top             =   4155
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ils16"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "New"
               Object.ToolTipText     =   "新增查询模板"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Modify"
               Object.ToolTipText     =   "更新查询模板"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Delete"
               Object.ToolTipText     =   "删除查询模板"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Export"
               Object.ToolTipText     =   "导出到文件"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Import"
               Object.ToolTipText     =   "从文件导入"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6105
      TabIndex        =   27
      Top             =   4830
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7335
      TabIndex        =   28
      Top             =   4830
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   135
      TabIndex        =   29
      Top             =   4830
      Width           =   1100
   End
End
Attribute VB_Name = "frmMedicalStationFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private mblnOK As Boolean
Private mblnStartUp As Boolean
Private mlngUpKey As Long
Private mlngKey As Long
Private mfrmMain As Form
Private mlngLoop As Long
Private mRs As New ADODB.Recordset
Private mstrSQL As String
Private mstrCondition As String

Private Type Items
    团体名称 As String
    体检项目 As String
End Type

Private usrSaveGroup As Items

Public Function FileExists(sSource As String) As Boolean
'功能：一个判断文件是否存在的函数
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   
   On Error Resume Next
   hFile = FindFirstFile(sSource, WFD)
   If Err <> 0 Then
        Err.Clear
        Exit Function
   End If
   FileExists = hFile <> INVALID_HANDLE_VALUE
   
   Call FindClose(hFile)
   
End Function

Private Sub ExportCondition()
    Dim strFile As String, strFileTemp As String
    Dim strSectoin As String, strCommand As String, lngTemp As Long
    Dim objSys As New Scripting.FileSystemObject, objSource As Scripting.TextStream, objDest As Scripting.TextStream
    
    
    If lvw.ListItems.Count = 0 Then Exit Sub
    
    '首先得到要导出的文件名
    On Error Resume Next
    dlgFile.Filter = "常用条件(*.cdt)|*.cdt"
    dlgFile.Flags = cdlOFNOverwritePrompt
    dlgFile.CancelError = True
    dlgFile.ShowSave
    If Err <> 0 Then Exit Sub
    strFile = dlgFile.FileName
    
    On Error GoTo errHandle
    
    '接着导出成临时文件
    strFileTemp = Space(256)
    lngTemp = GetTempPath(256, strFileTemp)
    strFileTemp = Mid(strFileTemp, 1, InStr(strFileTemp, Chr(0)) - 1)
    strFileTemp = strFileTemp & dlgFile.FileTitle
    
    strSectoin = "HKEY_CURRENT_USER\SOFTWARE\VB AND VBA PROGRAM SETTINGS\ZLSOFT\私有模块\" & App.ProductName & "\过滤查找"
    strCommand = "REGEDIT /E " & strFileTemp & " """ & strSectoin & """"
    Call RunExternal(strCommand)
    If FileExists(strFileTemp) = False Then
        MsgBox "条件导出失败！", vbInformation, gstrSysName
        Exit Sub
    End If
    '再经过处理得到正式文件
    Set objSource = objSys.OpenTextFile(strFileTemp, ForReading, , TristateMixed)
    Set objDest = objSys.OpenTextFile(strFile, ForWriting, True)
    
    Do Until objSource.AtEndOfStream
        strCommand = objSource.ReadLine
        If InStr(strCommand, strSectoin) > 0 Then
            objDest.WriteLine "[注册表位置]"
        Else
            objDest.WriteLine strCommand
        End If
    Loop
    objSource.Close
    objDest.Close

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ImportCondition()
    Dim strFile As String, strFileTemp As String
    Dim strSectoin As String, strCommand As String, lngTemp As Long
    Dim objSys As New Scripting.FileSystemObject, objSource As Scripting.TextStream, objDest As Scripting.TextStream
    
    
    If MsgBox("导入将覆盖现有条件，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    '首先得到要导出的文件名
    On Error Resume Next
    dlgFile.Filter = "常用条件(*.cdt)|*.cdt"
    dlgFile.Flags = cdlOFNFileMustExist
    dlgFile.CancelError = True
    dlgFile.ShowOpen
    If Err <> 0 Then Exit Sub
    strFile = dlgFile.FileName
    
    On Error GoTo errHandle
    strSectoin = "HKEY_CURRENT_USER\SOFTWARE\VB AND VBA PROGRAM SETTINGS\ZLSOFT\私有模块\" & App.ProductName & "\过滤查找"
    
    strFileTemp = Space(256)
    lngTemp = GetTempPath(256, strFileTemp)
    strFileTemp = Mid(strFileTemp, 1, InStr(strFileTemp, Chr(0)) - 1)
    strFileTemp = strFileTemp & dlgFile.FileTitle
    
    '接着导出成临时文件
    Set objSource = objSys.OpenTextFile(strFile, ForReading, , TristateMixed)
    Set objDest = objSys.OpenTextFile(strFileTemp, ForWriting, True)
    
    Do Until objSource.AtEndOfStream
        strCommand = objSource.ReadLine
        If strCommand = "[注册表位置]" Then
            objDest.WriteLine "[" & strSectoin & "]"
        Else
            objDest.WriteLine strCommand
        End If
    Loop
    objSource.Close
    objDest.Close

    '再经过处理得到正式文件
    
    strCommand = "REGEDIT /S " & strFileTemp
    Call RunExternal(strCommand)
    
    Call Form_Load
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function RunExternal(ByVal strCommand As String) As Long
'功能：运行外部程序，并且要等待其结束后再返回
'参数：strCommand    外部程序
'返回：程序的返回值
    Dim lngID As Long, lngProcess As Long, lngReturn  As Long
    
    On Error GoTo errHandle
    lngID = Shell(strCommand, vbHide)
    lngProcess = OpenProcess(Process_Query_Information, False, lngID)
    Do
        Sleep 100
        GetExitCodeProcess lngProcess, lngReturn
        DoEvents
    Loop While lngReturn = Still_Active
    CloseHandle lngProcess
    
    RunExternal = lngReturn
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub RestoreCondition(ByVal strTag As String)
    '-----------------------------------------------------------------------------------------
    '功能:
    '-----------------------------------------------------------------------------------------
    Dim varTmp As Variant
    
    If strTag = "" Then Exit Sub
    
    varTmp = Split(strTag, "^")
    
    On Error Resume Next
    
    zlControl.CboLocate cbo(1), Val(varTmp(0)), True
    
    cbo(9).Text = varTmp(1)
    If cbo(9).Text = "指  定" Then
        dtp(4).Value = Format(varTmp(2), dtp(4).CustomFormat)
        dtp(5).Value = Format(varTmp(3), dtp(5).CustomFormat)
    End If
    
    txt(0).Text = varTmp(4)
    

    cbo(10).Text = varTmp(5)
    If cbo(10).Text = "指  定" Then
        dtp(2).Value = Format(varTmp(6), dtp(2).CustomFormat)
        dtp(3).Value = Format(varTmp(7), dtp(3).CustomFormat)
    End If
    
    txt(7).Text = varTmp(8)
    cmd(0).Tag = Val(varTmp(9))
    txt(1).Text = varTmp(10)
    
    mlngKey = Val(varTmp(11))
    cmd(1).Tag = Val(varTmp(12))
    txt(15).Text = varTmp(13)
    cbo(3).Text = varTmp(14)
    cbo(4).Text = varTmp(15)
    cbo(5).Text = varTmp(16)
            
    
    usrSaveGroup.体检项目 = txt(15).Text
    usrSaveGroup.团体名称 = txt(1).Text
    
    
End Sub

Private Function SaveCondition() As String
    '-----------------------------------------------------------------------------------------
    '功能:
    '-----------------------------------------------------------------------------------------
    Dim strTmp As String
    
    '格式:
    
    strTmp = IIf(cbo(1).ListIndex >= 0, cbo(1).ItemData(cbo(1).ListIndex), "")

    strTmp = strTmp & "^" & cbo(9).Text
    If cbo(9).Text = "指  定" Then
        strTmp = strTmp & "^" & Format(dtp(4).Value, "yyyy-mm-dd") & "^" & Format(dtp(5).Value, "yyyy-mm-dd")
    Else
        strTmp = strTmp & "^" & txtCover(0).Text & "^" & txtCover(1).Text
    End If
    strTmp = strTmp & "^" & txt(0).Text

    strTmp = strTmp & "^" & cbo(10).Text
    If cbo(10).Text = "指  定" Then
        strTmp = strTmp & "^" & Format(dtp(2).Value, "yyyy-mm-dd") & "^" & Format(dtp(3).Value, "yyyy-mm-dd")
    Else
        strTmp = strTmp & "^" & txtCover(2).Text & "^" & txtCover(3).Text
    End If
    
    strTmp = strTmp & "^" & txt(7).Text
    strTmp = strTmp & "^" & Val(cmd(0).Tag)
    strTmp = strTmp & "^" & txt(1).Text
    
    strTmp = strTmp & "^" & mlngKey
    strTmp = strTmp & "^" & Val(cmd(1).Tag)
    strTmp = strTmp & "^" & txt(15).Text
    
    strTmp = strTmp & "^" & cbo(3).Text
    strTmp = strTmp & "^" & cbo(4).Text
    strTmp = strTmp & "^" & cbo(5).Text
        
    SaveCondition = strTmp
End Function

Private Sub AdjustDateShow(ByVal Index As Integer)
    '-----------------------------------------------------------------------------------------
    '功能:
    '-----------------------------------------------------------------------------------------
    Dim lngDtp1 As Long
    Dim lngDtp2 As Long
    Dim lngTxt1 As Long
    Dim lngTxt2 As Long
    
    Select Case Index
    Case 9
        lngDtp1 = 4
        lngDtp2 = 5
        
        lngTxt1 = 0
        lngTxt2 = 1
    Case 10
        lngDtp1 = 2
        lngDtp2 = 3
        
        lngTxt1 = 2
        lngTxt2 = 3
    Case 11
        lngDtp1 = 6
        lngDtp2 = 7
        
        lngTxt1 = 4
        lngTxt2 = 5
    Case 12
        lngDtp1 = 0
        lngDtp2 = 1
        
        lngTxt1 = 6
        lngTxt2 = 7
    End Select
    
    
    If cbo(Index).Text = "指  定" Then
        txtCover(lngTxt1).Visible = False
        txtCover(lngTxt2).Visible = False
        dtp(lngDtp1).Enabled = True
        dtp(lngDtp2).Enabled = True
    Else
        txtCover(lngTxt1).Visible = True
        txtCover(lngTxt2).Visible = True
        dtp(lngDtp1).Enabled = False
        dtp(lngDtp2).Enabled = False
    End If
    
    txtCover(lngTxt1).Text = ""
    txtCover(lngTxt2).Text = ""
    
    If cbo(Index).Text <> "所  有" And cbo(Index).Text <> "指  定" Then
        txtCover(lngTxt1).Text = Format(GetDateTime(cbo(Index).Text, 1), "yyyy-mm-dd")
        txtCover(lngTxt2).Text = Format(GetDateTime(cbo(Index).Text, 2), "yyyy-mm-dd")
    End If
    
End Sub

Private Sub FillOperate(ByVal bytMode As Byte)
    '-----------------------------------------------------------------------------------------
    '功能:
    '-----------------------------------------------------------------------------------------
    Dim strText As String
    
    strText = cbo(3).Text
    
    cbo(3).Clear
    cbo(4).Clear
    Select Case bytMode
    Case 0  '数字型
        cbo(3).AddItem "等于"
        cbo(3).AddItem "大于"
        cbo(3).AddItem "小于"
        cbo(3).AddItem "大于等于"
        cbo(3).AddItem "小于等于"
        cbo(3).AddItem "不等于"
        cbo(3).AddItem "在范围内"
    Case 1, 2 '文字型
        cbo(3).AddItem "等于"
        cbo(3).AddItem "大于"
        cbo(3).AddItem "小于"
        cbo(3).AddItem "大于等于"
        cbo(3).AddItem "小于等于"
        cbo(3).AddItem "不等于"
        cbo(3).AddItem "包含"
    Case 3  '阴阳型(罗辑型)
        cbo(3).AddItem "等于"
        cbo(3).AddItem "不等于"
        
        cbo(4).AddItem "阴性"
        cbo(4).AddItem "阳性"
        cbo(4).ListIndex = 0
    End Select
    
    On Error Resume Next
    
    cbo(3).Text = strText
    If cbo(3).ListCount > 0 And cbo(3).ListIndex = -1 Then cbo(3).ListIndex = 0
    
End Sub

Private Function ShowOpenTree(ByVal intIndex As Integer) As Byte
    '-----------------------------------------------------------------------------------------
    '功能:打开树型+列表结构的诊疗项目数据
    '返回:出错返回2;成功返回1;取消返回0
    '-----------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim sglX As Single
    Dim sglY As Single
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim objPoint As POINTAPI
    
    On Error GoTo errHand
            
    ShowOpenTree = 2
    Select Case intIndex
        Case 1
            strLvw = "编码,1200,0,1;名称,1800,0,0;临床意义,1800,0,0"
            strSQL = "SELECT * FROM (" & _
                        "(select -1 AS ID,0 AS 上级id,'' AS 编码,'所有项目' AS 名称,'' AS 临床意义,'' AS 数值域,0 AS 末级,0 AS 排序,0 AS 类型 from dual UNION ALL " & _
                        "Select DISTINCT ID," & _
                                        "DECODE(上级ID,NULL,-1,上级ID) AS 上级ID," & _
                                        "编码," & _
                                        "名称," & _
                                        "'' as 临床意义," & _
                                        "'' as 数值域," & _
                                        "0 as 末级," & _
                                        "DECODE(上级ID,Null,ID * POWER(10, 20),上级ID * POWER(10, 20) + ID) As 排序,0 AS 类型 " & _
                                  "From 诊治所见分类 " & _
                                 "Start With ID IN " & _
                                               "( " & _
                                               "SELECT 分类id from 诊治所见项目 A " & _
                                               "where A.ID IN (SELECT A.所见项id " & _
                                                              "FROM 病历所见单 A, 病历元素目录 B " & _
                                                              "WHERE A.元素id = B.ID AND B.类型 = 2 AND B.适用 LIKE '%1') " & _
                                               "Union " & _
                                               "SELECT 分类id from 诊治所见项目 A " & _
                                               "where A.ID IN (SELECT DISTINCT 报告项目id from 检验报告项目 A) " & _
                                               ") " & _
                                "Connect by Prior 上级ID = ID) "
                                
            strSQL = strSQL & _
                        "Union All " & _
                        "(SELECT ID, 分类id AS 上级id, 编码, 中文名 AS 名称, 临床意义, 数值域,1 AS 末级,1 AS 排序,类型 " & _
                          "from 诊治所见项目 A " & _
                         "where A.ID IN " & _
                               "(SELECT A.所见项id FROM 病历所见单 A, 病历元素目录 B WHERE A.元素id = B.ID AND B.类型 = 2 AND B.适用 LIKE '%1') " & _
                        "Union " & _
                        "SELECT ID, DECODE(分类id,NULL,-1,分类id) AS 上级id, 编码, 中文名 AS 名称, 临床意义, 数值域,1 AS 末级,1 AS 排序,类型 " & _
                          "from 诊治所见项目 A " & _
                         "where A.ID IN " & _
                               "(SELECT DISTINCT 报告项目id from 检验报告项目 A)) " & _
                        ") A ORDER BY A.末级,A.编码"
                        
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            
            If rs.BOF Then Exit Function
            Call ClientToScreen(txt(15).hWnd, objPoint)
            
            If frmSelectDialog.ShowSelect(Me, 3, rs, strLvw, "请选择一个诊治项目", objPoint.x * 15 - 30, objPoint.y * 15 + txt(15).Height - 30, 9000, 3600, txt(15).Height, , Me.Name & "\诊治项目树型选择", , False) Then
                
                txt(15).Text = zlCommFun.NVL(rs("名称").Value)
                mlngKey = zlCommFun.NVL(rs("ID").Value)
                cmd(1).Tag = zlCommFun.NVL(rs("类型").Value, 0)
                txt(15).Tag = ""
                
                usrSaveGroup.体检项目 = txt(15).Text
                
                ShowOpenTree = 1
                            
            End If
            
            txt(15).SetFocus
        Case 0
            '查询合约单位
            
            Call ClientToScreen(txt(1).hWnd, objPoint)
            
            gstrSQL = "select -1 AS ID,NULL+0 AS 上级id,'' AS 编码,'所有' AS 名称,'' as 简码,'' as 地址,0 AS 末级,'' AS 联系人,'' AS 电话,'' AS 电子邮件,'' AS 开户银行,'' AS 帐号,'' AS 地址,'' AS 说明 from dual " & _
                        "Union All " & _
                        "select ID,DECODE(上级id,NULL,-1,0,-1,上级id) AS 上级id,编码,名称,简码,地址,0 AS 末级,联系人,电话,电子邮件,开户银行,帐号,地址,说明 from 合约单位  where 末级<>1 " & _
                        "Start With 上级id is null connect by prior ID=上级id " & _
                        "Union All " & _
                        "select ID,DECODE(上级id,NULL,-1,0,-1,上级id) AS 上级id,编码,名称,简码,地址,1 AS 末级,联系人,电话,电子邮件,开户银行,帐号,地址,说明 from 合约单位  where 末级=1"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            
            strLvw = "编码,900,0,1;名称,1500,0,1;简码,900,0,1;地址,3000,0,1"
            If frmSelectDialog.ShowSelect(Me, 3, rs, strLvw, "请在下表中选择一个团体/单位", objPoint.x * 15 - 30, objPoint.y * 15 + txt(1).Height - 30, 8790, 5100, txt(1).Height, , Me.Name & "\体检团体选择", , False) Then
            
                txt(1).Text = zlCommFun.NVL(rs("名称").Value)
                cmd(0).Tag = zlCommFun.NVL(rs("ID").Value, 0)
                
                usrSaveGroup.团体名称 = txt(1).Text
                                                
                ShowOpenTree = 1
                
            End If
        
            txt(1).SetFocus
        
    End Select
    
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim lngLoop As Long
    
    On Error GoTo errHand
                        
    '1.体检部门
    cbo(1).Clear
    For lngLoop = 0 To mfrmMain.cboDept.ListCount - 1
        cbo(1).AddItem mfrmMain.cboDept.List(lngLoop)
        cbo(1).ItemData(cbo(1).NewIndex) = mfrmMain.cboDept.ItemData(lngLoop)
    Next
    cbo(1).ListIndex = mfrmMain.cboDept.ListIndex

    dtp(2).Value = Format(zlDatabase.Currentdate, dtp(2).CustomFormat)
    dtp(3).Value = Format(zlDatabase.Currentdate, dtp(3).CustomFormat)
    dtp(4).Value = Format(zlDatabase.Currentdate, dtp(4).CustomFormat)
    dtp(5).Value = Format(zlDatabase.Currentdate, dtp(5).CustomFormat)
        
    cbo(3).AddItem "等于"
    cbo(3).AddItem "大于"
    cbo(3).AddItem "小于"
    cbo(3).AddItem "大于等于"
    cbo(3).AddItem "小于等于"
    cbo(3).AddItem "不等于"
    cbo(3).AddItem "包含"
    cbo(3).AddItem "在范围内"
    cbo(3).ListIndex = -1
        
    For lngLoop = 9 To 10
        If lngLoop <> 9 Then cbo(lngLoop).AddItem "所  有"
        cbo(lngLoop).AddItem "今  天"
        cbo(lngLoop).AddItem "昨  天"
        cbo(lngLoop).AddItem "本  周"
        cbo(lngLoop).AddItem "本  月"
        cbo(lngLoop).AddItem "本  季"
        cbo(lngLoop).AddItem "本半年"
        cbo(lngLoop).AddItem "本  年"
        cbo(lngLoop).AddItem "前三天"
        cbo(lngLoop).AddItem "前一周"
        cbo(lngLoop).AddItem "前半月"
        cbo(lngLoop).AddItem "前一月"
        cbo(lngLoop).AddItem "前二月"
        cbo(lngLoop).AddItem "前三月"
        cbo(lngLoop).AddItem "前半年"
        cbo(lngLoop).AddItem "指  定"
        cbo(lngLoop).ListIndex = 0
    Next
        
    InitData = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then Resume
    
End Function

Public Function ShowEdit(ByVal frmMain As Form, ByRef strCondition As String) As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    
    mblnStartUp = True
    mblnOK = False
        
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
    
    If strCondition <> "" Then Call RestoreCondition(strCondition)
        
    Me.Show 1, frmMain
    
    strCondition = mstrCondition
    ShowEdit = mblnOK
    
End Function

Private Sub cbo_Click(Index As Integer)
    Select Case Index
    Case 3
        cbo(5).Visible = (cbo(Index).List(cbo(Index).ListIndex) = "在范围内")
    Case 9, 10, 11, 12
        Call AdjustDateShow(Index)
    End Select

End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
        If Index = 4 Then
            Select Case Val(cmd(0).Tag)
            Case 1
                KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789.")
            Case 2
            End Select
        End If
    End If
    
End Sub

Private Sub cbo_Validate(Index As Integer, Cancel As Boolean)
    If Index = 4 Then
        Select Case Val(cmd(0).Tag)
        Case 3
            For mlngLoop = 0 To cbo(Index).ListCount - 1
                If cbo(Index).Text = cbo(Index).List(mlngLoop) Then
                    Exit For
                End If
            Next
            If mlngLoop >= cbo(Index).ListCount Then
                MsgBox "输入的结果必须在下拉列表内容之内！", vbInformation, gstrSysName
                Cancel = True
            End If
        End Select
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    Select Case Index
        Case 0
        
            gstrSQL = GetPublicSQL(SQL.体检团体选择)
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            If ShowTxtSelect(Me, txt(1), "编码,900,0,1;名称,1500,0,1;简码,900,0,1;地址,3000,0,1", Me.Name & "\体检团体选择", "请在下表中选择一个团体/单位。", rsData, rs, 8790, 5100) Then
                
                txt(1).Text = zlCommFun.NVL(rs("名称").Value)
                cmd(0).Tag = zlCommFun.NVL(rs("ID").Value, 0)
                txt(1).Tag = ""
                
                usrSaveGroup.团体名称 = txt(1).Text
                        
            End If
        
            txt(1).SetFocus
        Case 1
        
            gstrSQL = GetPublicSQL(SQL.诊治项目选择)
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            
            If ShowTxtSelect(Me, txt(1), "编码,1200,0,1;名称,1800,0,0;临床意义,1800,0,0", Me.Name & "\诊治项目选择", "请选择一个诊治项目。", rsData, rs, 8790, 5100) Then
                
                txt(15).Text = zlCommFun.NVL(rs("名称").Value)
                mlngKey = zlCommFun.NVL(rs("ID").Value)
                cmd(1).Tag = zlCommFun.NVL(rs("类型").Value, 0)
                txt(15).Tag = ""
                
                usrSaveGroup.体检项目 = txt(15).Text
                                        
                Call FillOperate(Val(cmd(1).Tag))
                
            End If
        
            txt(15).SetFocus
            
    End Select
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdClear_Click(Index As Integer)
    
    Select Case Index
        Case 1
            cmd(0).Tag = ""
            txt(1).Text = ""
            txt(1).Tag = ""
        Case 0
            cmd(1).Tag = ""
            txt(15).Text = ""
            txt(15).Tag = ""
            mlngKey = 0
    End Select
    
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()

    mstrCondition = SaveCondition
    mblnOK = True
    Unload Me
    
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub Form_Load()
    Dim lngLoop As Long
    Dim strSectoin  As String
    Dim objItem As ListItem
    Dim strTmp As String
    
    lvw.ListItems.Clear
    
    strSectoin = "私有模块\" & App.ProductName & "\过滤查找"
    
    For lngLoop = 1 To CLng(Val(GetSetting("ZLSOFT", strSectoin, "查找项数", "0")))
        
        strTmp = GetSetting("ZLSOFT", strSectoin, "过滤查找" & lngLoop, "")
        
        If Trim(strTmp) <> "" And InStr(strTmp, "|") > 0 Then
            Set objItem = lvw.ListItems.Add(, , Mid(strTmp, 1, InStr(strTmp, "|") - 1), "Query", "Query")
            objItem.Tag = Mid(strTmp, InStr(strTmp, "|") + 1)
        End If
    Next
            
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim lngLoop As Long
    Dim strSectoin As String
    
    '进行了常用条件的更改
    strSectoin = "私有模块\" & App.ProductName & "\过滤查找"
    
    On Error Resume Next '如果没有该键值，就会出错
    DeleteSetting "ZLSOFT", strSectoin '删除以前的条件
    On Error GoTo 0
    
    Call SaveSetting("ZLSOFT", strSectoin, "查找项数", lvw.ListItems.Count)
    
    For lngLoop = 1 To lvw.ListItems.Count
        
        Call SaveSetting("ZLSOFT", strSectoin, "过滤查找" & lngLoop, lvw.ListItems(lngLoop).Text & "|" & lvw.ListItems(lngLoop).Tag)

    Next
    
End Sub


Private Sub lvw_DblClick()
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    Call RestoreCondition(lvw.SelectedItem.Tag)
    
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim lngIndex As Long
    Dim objItem As ListItem
    Dim strTmp As String
    
    Select Case Button.Key
    Case "New"
        
        Set objItem = lvw.ListItems.Add(, , "新查询" & (lvw.ListItems.Count + 1), "Query", "Query")
        objItem.Tag = SaveCondition

    Case "Modify"
        If Not (lvw.SelectedItem Is Nothing) Then lvw.SelectedItem.Tag = SaveCondition
    Case "Delete"
        If lvw.SelectedItem Is Nothing Then Exit Sub
        
        lngIndex = lvw.SelectedItem.Index
        lvw.ListItems.Remove lngIndex
        Call NextLvwPos(lvw, lngIndex)
        
    Case "Export"
        Call Form_Unload(False)
        Call ExportCondition
    Case "Import"
        Call ImportCondition
    End Select
End Sub

Private Sub txt_Change(Index As Integer)
    If Index <> 15 And Index <> 1 Then Exit Sub
    
    txt(Index).Tag = "Changed"
    
End Sub

Private Sub txt_GotFocus(Index As Integer)
    Select Case Index
    Case 2, 3, 4, 5, 7
        zlCommFun.OpenIme True
    End Select
    
    zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strText As String
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        
        
        If Index = 15 Or Index = 1 Then
            If txt(Index).Tag <> "" Then
                txt(Index).Tag = ""
                
                If Index = 1 Then
                    gstrSQL = GetPublicSQL(SQL.团体过滤选择)
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "%" & UCase(txt(Index).Text) & "%")
                    
                    If ShowTxtFilter(Me, txt(Index), "名称,1800,0,0;编码,900,0,0;简码,900,0,0;联系人,900,0,0;电话,1200,0,0", Me.Name & "\团体过滤选择", "请从下面选择一个团体单位", rsData, rs) Then
                        
                        txt(1).Text = zlCommFun.NVL(rs("名称"))
                        cmd(0).Tag = zlCommFun.NVL(rs("ID"))
                        txt(1).Tag = ""
                        usrSaveGroup.团体名称 = txt(1).Text

                    Else
                        txt(1).Text = usrSaveGroup.团体名称
                        Exit Sub
                    End If
                End If
                
                If Index = 15 Then
                    
                    strText = UCase(txt(Index).Text) & "%"
                    If ParamInfo.项目输入匹配方式 = 0 Then strTmp = "%" & strText
                    
                    gstrSQL = GetPublicSQL(SQL.诊治项目过滤选择)
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strText, strTmp)
                    
                    If ShowTxtFilter(Me, txt(Index), "编码,900,0,1;名称,2400,0,0;英文名,1200,0,0;临床意义,900,0,0", Me.Name & "\诊治项目过滤选择", "请从下表中选择一个项目", rsData, rs) Then
                        
                        txt(15).Text = zlCommFun.NVL(rs("名称").Value)
                        mlngKey = zlCommFun.NVL(rs("ID").Value)
                        cmd(1).Tag = zlCommFun.NVL(rs("类型").Value, 0)
                        txt(15).Tag = ""
                        usrSaveGroup.体检项目 = txt(15).Text
                        
                        Call FillOperate(Val(cmd(1).Tag))
                        
                    Else
                        txt(15).Text = usrSaveGroup.体检项目
                        Exit Sub
                    End If
                End If
            Else
                zlCommFun.PressKey vbKeyTab
                zlCommFun.PressKey vbKeyTab
                zlCommFun.PressKey vbKeyTab
            End If
            txt(Index).Tag = ""
        Else
            zlCommFun.PressKey vbKeyTab
        End If
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        Select Case Index
        Case 9, 10, 11
            KeyAscii = FilterKeyAscii(KeyAscii, 1)
        Case 0
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789AZXCVBNMASDFGHJKLPOIUYTREWQ-,")
        Case 15
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End Select
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 2, 3, 4, 5, 7
        zlCommFun.OpenIme False
    Case 0
                        
        Dim intYear As Integer
        Dim strYear As String
        '自动补齐单据号
        If (UCase(Left(txt(Index).Text, 1)) < "A" Or UCase(Left(txt(Index).Text, 1)) > "Z") And Trim(txt(Index).Text) <> "" Then
            intYear = Format(zlDatabase.Currentdate, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            txt(Index).Text = strYear & Right("0000000" & txt(Index).Text, 7)
        End If
        
    End Select
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
        
    If (txt(Index).Tag = "Changed") And Index = 15 Then
        txt(Index).Text = usrSaveGroup.体检项目
    End If
    
    If (txt(Index).Tag = "Changed") And Index = 1 Then
        txt(Index).Text = usrSaveGroup.团体名称
    End If
    
End Sub
