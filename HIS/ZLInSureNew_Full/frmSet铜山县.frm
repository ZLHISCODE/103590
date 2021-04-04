VERSION 5.00
Begin VB.Form frmSet铜山县 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保参数设置"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5235
      TabIndex        =   5
      Top             =   2730
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4110
      TabIndex        =   4
      Top             =   2730
      Width           =   1100
   End
   Begin VB.Frame fraIC 
      Caption         =   "IC卡操作"
      Height          =   2250
      Left            =   75
      TabIndex        =   0
      Top             =   150
      Width           =   6240
      Begin VB.CommandButton cmd自编码 
         Caption         =   "同步自编码"
         Height          =   390
         Left            =   4650
         TabIndex        =   17
         Top             =   1740
         Width           =   1320
      End
      Begin VB.TextBox txtPass 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1335
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   1095
         Width           =   1425
      End
      Begin VB.TextBox txtUser 
         Height          =   270
         Left            =   1335
         TabIndex        =   13
         Top             =   720
         Width           =   1425
      End
      Begin VB.CommandButton cmd未对码 
         Caption         =   "查未对码项目"
         Height          =   390
         Left            =   4650
         TabIndex        =   12
         Top             =   1200
         Width           =   1320
      End
      Begin VB.CommandButton cmd诊疗库 
         Caption         =   "诊疗库更新"
         Height          =   390
         Left            =   4650
         TabIndex        =   11
         Top             =   772
         Width           =   1320
      End
      Begin VB.CommandButton cmd药品库 
         Caption         =   "药品库更新"
         Height          =   390
         Left            =   4650
         TabIndex        =   10
         Top             =   315
         Width           =   1320
      End
      Begin VB.CommandButton cmd医院编码 
         Caption         =   "医院编码更新"
         Height          =   390
         Left            =   3120
         TabIndex        =   9
         Top             =   315
         Width           =   1320
      End
      Begin VB.CommandButton cmd住院病种 
         Caption         =   "住院病种更新"
         Height          =   360
         Left            =   3120
         TabIndex        =   8
         Top             =   772
         Width           =   1350
      End
      Begin VB.CommandButton cmd门诊病种更新 
         Caption         =   "门诊病种更新"
         Height          =   390
         Left            =   3120
         TabIndex        =   6
         Top             =   1200
         Width           =   1320
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   1335
         MaxLength       =   40
         TabIndex        =   1
         Text            =   "1"
         Top             =   315
         Width           =   360
      End
      Begin VB.Label Label4 
         Caption         =   "请注意：同步自编码功能只适用于自编码和收费细目ID相同的单位。"
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   570
         TabIndex        =   18
         Top             =   1725
         Width           =   3765
      End
      Begin VB.Label Label3 
         Caption         =   "管理员密码"
         Height          =   165
         Left            =   270
         TabIndex        =   16
         Top             =   1170
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "医保管理员"
         Height          =   165
         Left            =   270
         TabIndex        =   15
         Top             =   780
         Width           =   1005
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "号串口"
         Height          =   180
         Index           =   4
         Left            =   1740
         TabIndex        =   3
         Top             =   375
         Width           =   540
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "当前串口(&D)"
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   2
         Top             =   375
         Width           =   990
      End
   End
   Begin VB.Label Label1 
      Caption         =   "病种更新时间较长，请耐心等候。"
      Height          =   315
      Left            =   90
      TabIndex        =   7
      Top             =   2730
      Width           =   3630
   End
End
Attribute VB_Name = "frmSet铜山县"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mblnChange As Boolean
Private mlngIcdev As Long
Private st%


Dim strUser As String, strServer As String, strPass As String
Dim strFileName As String
Dim objStream As TextStream, lngReturn As Long
Dim objFileSystem As New FileSystemObject, lngID As Long
Dim strLine As String, rsBzgx As New ADODB.Recordset
Dim lngRount As Long
Dim lng可更新 As Long

Private Const P_FILENAME = 167782162

 
Private Sub cmd门诊病种更新_Click()

On Error GoTo errHand
    strUser = "tsxyb"
    strServer = GetSetting("ZLSOFT", "注册信息\登陆信息", "SERVER", "")
    strPass = "tsxyb"
    lngRount = 0
    If OraDataOpen(gcn铜山县, strServer, strUser, strPass, False) = False Then
        MsgBox "无法连接到中间库，请检查保险参数是否设置正确！", vbInformation, gstrSysName
        Exit Sub
    End If
    Label1.Caption = "开始下载门诊病种!"
    '如果有文件,则提示,是否覆盖,否则就下载
    DoEvents
    strFileName = App.Path & "\MZBZ.TXT"
    If Not objFileSystem.FileExists(strFileName) Then
        '下载门诊病种
        Call 医保初始化_铜山县
        If tsx_createparams(1024, 1024) = -1 Then
            MsgBox "分配内存空间失败!" & tsx_getlasterr(), vbInformation, gstrSysName
            Exit Sub
        End If
        
        lngReturn = tsx_setstringparam(P_FILENAME, 0, strFileName)
        If tsx_jkcall("D_MZBZB") = -1 Then
            MsgBox tsx_getlasterr()
            Exit Sub
        End If
        Label1.Caption = "门诊病种下载完毕!"
        lngReturn = tsx_destroyparams()
    Else
        '提示是
        If MsgBox("已有门诊病种文件,是否重新下载?", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            '下载
            Call 医保初始化_铜山县
            If tsx_createparams(1024, 1024) = -1 Then
                MsgBox "分配内存空间失败!" & tsx_getlasterr(), vbInformation, gstrSysName
                Exit Sub
            End If
            
            lngReturn = tsx_setstringparam(P_FILENAME, 0, strFileName)
            If tsx_jkcall("D_MZBZB") = -1 Then
                MsgBox tsx_getlasterr()
                Exit Sub
            End If
            Label1.Caption = "门诊病种下载完毕!"
            lngReturn = tsx_destroyparams()
        End If
    End If
    DoEvents

    strFileName = App.Path & "\MZBZ.TXT"
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForReading)
    
    Do Until objStream.AtEndOfLine
        strLine = objStream.ReadLine
        gstrSQL = "Select * from Mzbz where 病种编码='" & _
                    Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 1, 2), vbUnicode)) & "'"
        Call OpenRecordset_OtherBase(rsBzgx, "保险病种", , gcn铜山县)
        
        If rsBzgx.EOF Then
            gstrSQL = "Select 保险病种_ID.nextval as ID from dual"
            Set rsBzgx = zlDatabase.OpenSQLRecord(gstrSQL, "保险病种")
            lngID = rsBzgx!ID
            gstrSQL = "Insert into MZBZ(ID,病种编码,病种名称,拼音码) values(" & lngID & ",'" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 1, 2), vbUnicode)) & "','" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 3, 20), vbUnicode)) & "','" & _
                                        Mid(Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 24, 10), vbUnicode)), 2) & "')"
        End If
        gcn铜山县.Execute gstrSQL
        DoEvents
        lngRount = lngRount + 1
        Label1.Caption = "已增加" & lngRount & "条记录"
    Loop
    Label1.Caption = "完成门诊病种更新,本次共增加" & lngRount & "条记录"
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Sub

Private Sub cmd未对码_Click()

    Dim rsYpzlk As New ADODB.Recordset, rsSfxm As New ADODB.Recordset
    Dim lngRount As Long, strSqlTemp As String
On Error GoTo errHand

    cmd未对码.Enabled = False
    lngRount = 0
    
    strUser = "tsxyb"
    strServer = GetSetting("ZLSOFT", "注册信息\登陆信息", "SERVER", "")
    strPass = "tsxyb"
    lngRount = 0
    
    If OraDataOpen(gcn铜山县, strServer, strUser, strPass, False) = False Then
        MsgBox "无法连接到中间库，请检查保险参数是否设置正确！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strSqlTemp = "Delete ypzlk where substr(支付类别,1,1)<>'1' and 收费细目ID is not null"
    gcn铜山县.Execute strSqlTemp
    
    strSqlTemp = "update ypzlk set 支付类别=substr(支付类别,2) where substr(支付类别,1,1)='1'"
    gcn铜山县.Execute strSqlTemp
    
    gstrSQL = "Select a.*," & _
              "Decode(Nvl(撤档时间, To_Date('3000-01-01', 'YYYY-MM-DD')), To_Date('3000-01-01', 'YYYY-MM-DD'), 0, 1) As 停用 " & _
              " From 收费细目 a"
    Set rsSfxm = zlDatabase.OpenSQLRecord(gstrSQL, "收费细目")
    strFileName = App.Path & "\未对码项目.TXT"
    Call objFileSystem.CreateTextFile(strFileName)
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
    cmd未对码.Tag = 0
    
    Do Until rsSfxm.EOF
        gstrSQL = "Select * from ypzlk where 收费细目ID=" & rsSfxm!ID
        Call OpenRecordset_OtherBase(rsYpzlk, "ypzlk", , gcn铜山县)
        If rsYpzlk.EOF Then
            '写文件
            If rsSfxm!停用 = 0 Then
                objStream.WriteLine rsSfxm!类别 & "  " & rsSfxm!ID & "  " & rsSfxm!编码 & "  " & rsSfxm!名称
                lngRount = lngRount + 1
            End If
            gstrSQL = "ZL_保险支付项目_Delete(" & rsSfxm!ID & "," & TYPE_铜山县 & ")"
        Else
            gstrSQL = "ZL_保险支付项目_Modify(" & rsSfxm!ID & "," & TYPE_铜山县 & ",NUll,'" & _
                           rsYpzlk!自编码 & "','" & rsSfxm!名称 & "','" & rsYpzlk!支付类别 & "',1)"
                           
        End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保险支付项目")
        rsSfxm.MoveNext
        DoEvents
        cmd未对码.Tag = cmd未对码.Tag + 1
        Label1.Caption = "已检查" & cmd未对码.Tag & "条记录"
    Loop
    
    objStream.WriteLine "有" & lngRount & "条记录未对码"
    objStream.Close
    Set objStream = Nothing
    
    If lng可更新 = 3 Then
        cmd未对码.Enabled = False
        cmd药品库.Enabled = True
        cmd诊疗库.Enabled = True
        cmd自编码.Enabled = True
        lng可更新 = 0
    End If
    
    Shell "notepad.exe " & strFileName, vbMaximizedFocus

    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmd药品库_Click()
    Dim str类别 As String
On Error GoTo errHand


    strUser = "tsxyb"
    strServer = GetSetting("ZLSOFT", "注册信息\登陆信息", "SERVER", "")
    strPass = "tsxyb"
    lngRount = 0
    
    
    If OraDataOpen(gcn铜山县, strServer, strUser, strPass, False) = False Then
        MsgBox "无法连接到中间库，请检查保险参数是否设置正确！", vbInformation, gstrSysName
        Exit Sub
    End If
    Label1.Caption = "开始下载药品库!"
    '如果有文件,则提示,是否覆盖,否则就下载
    DoEvents
    strFileName = App.Path & "\YPK.TXT"
    If Not objFileSystem.FileExists(strFileName) Then
        MsgBox "请从医保前台程序中导出药品库,将其命名为YPK.TXT放到" & App.Path & "目录下。再执行此功能！", vbInformation, gstrSysName
        Exit Sub
    Else
        If MsgBox("请确认从医保前台导出的文件YPK.TXT放到了" & App.Path & "目录下", vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    lng可更新 = lng可更新 + 1
    cmd药品库.Enabled = False
    
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForReading)
    
    Do Until objStream.AtEndOfLine
        strLine = objStream.ReadLine
        gstrSQL = "Select * from YPZLK where 自编码='" & _
                    Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 1, 10), vbUnicode)) & "'"
        Call OpenRecordset_OtherBase(rsBzgx, "医院编码", , gcn铜山县)
        
        If rsBzgx.EOF Then
            Select Case Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 21, 1), vbUnicode))
                Case 1
                    str类别 = "1甲类"
                Case 2
                    str类别 = "1乙类"
                Case Else
                    str类别 = "1自费"
            End Select
            gstrSQL = "Insert into YPZLK(自编码,医保编码,支付类别,特批标志) values('" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 1, 10), vbUnicode)) & "','" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 11, 10), vbUnicode)) & "','" & _
                                        str类别 & "','" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 22, 1), vbUnicode)) & "')"
        Else
            Select Case Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 21, 1), vbUnicode))
                Case 1
                    str类别 = "1甲类"
                Case 2
                    str类别 = "1乙类"
                Case Else
                    str类别 = "1自费"
            End Select
            gstrSQL = "Update YPZLK Set 医保编码='" & Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 11, 10), vbUnicode)) & "'," & _
                                       "支付类别='" & str类别 & "'," & _
                                       "特批标志='" & Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 22, 1), vbUnicode)) & "' " & _
                                       " Where 自编码='" & Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 1, 10), vbUnicode)) & "'"
            
        End If
        gcn铜山县.Execute gstrSQL
        DoEvents
        lngRount = lngRount + 1
        Label1.Caption = "已增加" & lngRount & "条记录"
        
    Loop
    Label1.Caption = "完成药品库更新,本次共增加" & lngRount & "条记录"
    objStream.Close
    Set objStream = Nothing
    

    If lng可更新 = 3 Then
        cmd未对码.Enabled = True
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmd医院编码_Click()
On Error GoTo errHand
    strUser = "tsxyb"
    strServer = GetSetting("ZLSOFT", "注册信息\登陆信息", "SERVER", "")
    strPass = "tsxyb"
    lngRount = 0
    If OraDataOpen(gcn铜山县, strServer, strUser, strPass, False) = False Then
        MsgBox "无法连接到中间库，请检查保险参数是否设置正确！", vbInformation, gstrSysName
        Exit Sub
    End If
    Label1.Caption = "开始下载医院编码!"
    '如果有文件,则提示,是否覆盖,否则就下载
    DoEvents

    strFileName = App.Path & "\YYDA.TXT"
    If Not objFileSystem.FileExists(strFileName) Then
        '下载住院病种
        Call 医保初始化_铜山县
        If tsx_createparams(1024, 1024) = -1 Then
            Label1.Caption = "分配内存空间失败!" & tsx_getlasterr()
            Exit Sub
        End If
        
        lngReturn = tsx_setstringparam(P_FILENAME, 0, strFileName)
        If tsx_jkcall("D_YYDA") = -1 Then
            MsgBox tsx_getlasterr()
            Exit Sub
        End If
        Label1.Caption = "医院编码下载完毕!开始写入本地数据库中！"
        lngReturn = tsx_destroyparams()
    Else
        '提示是
        If MsgBox("已有医院编码文件,是否重新下载?", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            '下载
            Call 医保初始化_铜山县
            If tsx_createparams(1024, 1024) = -1 Then
                MsgBox "分配内存空间失败!" & tsx_getlasterr(), vbInformation, gstrSysName
                Exit Sub
            End If
            
            lngReturn = tsx_setstringparam(P_FILENAME, 0, strFileName)
            If tsx_jkcall("D_YYDA") = -1 Then
                MsgBox tsx_getlasterr()
                Exit Sub
            End If
            Label1.Caption = "医院编码下载完毕!开始写入本地数据库中！"
            lngReturn = tsx_destroyparams()
        End If
    End If
    DoEvents
    strFileName = App.Path & "\YYDA.TXT"
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForReading)
    
    Do Until objStream.AtEndOfLine
        strLine = objStream.ReadLine
        gstrSQL = "Select * from YYDA where 医院编码='" & _
                    Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 1, 2), vbUnicode)) & "'"
        Call OpenRecordset_OtherBase(rsBzgx, "医院编码", , gcn铜山县)
        
        If rsBzgx.EOF Then
            gstrSQL = "Insert into YYDA(医院编码,医院名称) values('" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 1, 2), vbUnicode)) & "','" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 3, 25), vbUnicode)) & "')"
        End If
        gcn铜山县.Execute gstrSQL
        DoEvents
        lngRount = lngRount + 1
        Label1.Caption = "已增加" & lngRount & "条记录"
        
    Loop
    Label1.Caption = "完成医院编码更新,本次共增加" & lngRount & "条记录"
    objStream.Close
    Set objStream = Nothing
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Sub

Private Sub cmd诊疗库_Click()
    Dim str类别 As String
    
On Error GoTo errHand

    strUser = "tsxyb"
    strServer = GetSetting("ZLSOFT", "注册信息\登陆信息", "SERVER", "")
    strPass = "tsxyb"
    lngRount = 0
    
    If OraDataOpen(gcn铜山县, strServer, strUser, strPass, False) = False Then
        MsgBox "无法连接到中间库，请检查保险参数是否设置正确！", vbInformation, gstrSysName
        Exit Sub
    End If
    Label1.Caption = "开始下载药品库!"
    '如果有文件,则提示,是否覆盖,否则就下载
    DoEvents
    strFileName = App.Path & "\ZLK.TXT"
    If Not objFileSystem.FileExists(strFileName) Then
        MsgBox "请从医保前台程序中导出药品库,将其命名为ZLK.TXT放到" & App.Path & "目录下。再执行此功能！", vbInformation, gstrSysName
        Exit Sub
    Else
        If MsgBox("请确认从医保前台导出的文件ZLK.TXT放到了" & App.Path & "目录下", vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Sub
        End If
        
    End If
    
    lng可更新 = lng可更新 + 1
    cmd诊疗库.Enabled = False
    
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForReading)
    
    Do Until objStream.AtEndOfLine
        strLine = objStream.ReadLine
        gstrSQL = "Select * from YPZLK where 自编码='" & _
                    Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 1, 10), vbUnicode)) & "'"
        Call OpenRecordset_OtherBase(rsBzgx, "医院编码", , gcn铜山县)
        
        If rsBzgx.EOF Then
            Select Case Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 21, 1), vbUnicode))
                Case 1
                    str类别 = "1甲类"
                Case 2
                    str类别 = "1乙类"
                Case Else
                    str类别 = "1自费"
            End Select
            gstrSQL = "Insert into YPZLK(自编码,医保编码,支付类别,特批标志) values('" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 1, 10), vbUnicode)) & "','" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 11, 10), vbUnicode)) & "','" & _
                                        str类别 & "','" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 22, 1), vbUnicode)) & "')"
        Else
            Select Case Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 21, 1), vbUnicode))
                Case 1
                    str类别 = "1甲类"
                Case 2
                    str类别 = "1乙类"
                Case Else
                    str类别 = "1自费"
            End Select
            gstrSQL = "Update YPZLK Set 医保编码='" & Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 11, 10), vbUnicode)) & "'," & _
                                       "支付类别='" & str类别 & "'," & _
                                       "特批标志='" & Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 22, 1), vbUnicode)) & "' " & _
                                       " Where 自编码='" & Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 1, 10), vbUnicode)) & "'"
        End If
        gcn铜山县.Execute gstrSQL
        DoEvents
        lngRount = lngRount + 1
        Label1.Caption = "已增加" & lngRount & "条记录"
        
    Loop
    Label1.Caption = "完成诊疗库更新,本次共增加" & lngRount & "条记录"
    objStream.Close
    Set objStream = Nothing
    If lng可更新 = 3 Then
        cmd未对码.Enabled = True
    End If

    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmd住院病种_Click()
On Error GoTo errHand

    strUser = "tsxyb"
    strServer = GetSetting("ZLSOFT", "注册信息\登陆信息", "SERVER", "")
    strPass = "tsxyb"
    lngRount = 0
    If OraDataOpen(gcn铜山县, strServer, strUser, strPass, False) = False Then
        MsgBox "无法连接到中间库，请检查保险参数是否设置正确！", vbInformation, gstrSysName
        Exit Sub
    End If
    Label1.Caption = "开始下载住院病种!"
    '如果有文件,则提示,是否覆盖,否则就下载
    DoEvents

    strFileName = App.Path & "\ICD10.TXT"
    If Not objFileSystem.FileExists(strFileName) Then
        '下载住院病种
        Call 医保初始化_铜山县
        If tsx_createparams(1024, 1024) = -1 Then
            Label1.Caption = "分配内存空间失败!" & tsx_getlasterr()
            Exit Sub
        End If
        
        lngReturn = tsx_setstringparam(P_FILENAME, 0, strFileName)
        If tsx_jkcall("D_ICD10") = -1 Then
            MsgBox tsx_getlasterr()
            Exit Sub
        End If
        Label1.Caption = "住院病种下载完毕!开始写入HIS数据库中！"
        lngReturn = tsx_destroyparams()
    Else
        '提示是
        If MsgBox("已有住院病种文件,是否重新下载?", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            '下载
            Call 医保初始化_铜山县
            If tsx_createparams(1024, 1024) = -1 Then
                MsgBox "分配内存空间失败!" & tsx_getlasterr(), vbInformation, gstrSysName
                Exit Sub
            End If
            
            lngReturn = tsx_setstringparam(P_FILENAME, 0, strFileName)
            If tsx_jkcall("D_ICD10") = -1 Then
                MsgBox tsx_getlasterr()
                Exit Sub
            End If
            Label1.Caption = "住院病种下载完毕!开始写入HIS数据库中！"
            lngReturn = tsx_destroyparams()
        End If
    End If
    DoEvents
    strFileName = App.Path & "\ICD10.TXT"
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForReading)
    
    Do Until objStream.AtEndOfLine
        strLine = objStream.ReadLine
        gstrSQL = "Select * from ICD10 where 病种编码='" & _
                    Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 1, 10), vbUnicode)) & "'"
        Call OpenRecordset_OtherBase(rsBzgx, "保险病种", , gcn铜山县)
        
        If rsBzgx.EOF Then
            gstrSQL = "Select 保险病种_ID.nextval as ID from dual"
            Set rsBzgx = zlDatabase.OpenSQLRecord(gstrSQL, "保险病种")
            lngID = rsBzgx!ID
            gstrSQL = "Insert into ICD10(ID,病种编码,病种名称,拼音码) values(" & lngID & ",'" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 1, 10), vbUnicode)) & "','" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 11, 20), vbUnicode)) & "','" & _
                                        Trim(StrConv(Mid(StrConv(strLine, vbFromUnicode), 31, 10), vbUnicode)) & "')"
        End If
        gcn铜山县.Execute gstrSQL
        DoEvents
        lngRount = lngRount + 1
        Label1.Caption = "已增加" & lngRount & "条记录"
        
    Loop
    Label1.Caption = "完成住院病种更新,本次共增加" & lngRount & "条记录"
    objStream.Close
    Set objStream = Nothing
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Sub

Private Sub cmd自编码_Click()
    
On Error GoTo errHand
    strUser = "tsxyb"
    strServer = GetSetting("ZLSOFT", "注册信息\登陆信息", "SERVER", "")
    strPass = "tsxyb"
    lngRount = 0
    lng可更新 = lng可更新 + 1
    cmd自编码.Enabled = False
    
    If OraDataOpen(gcn铜山县, strServer, strUser, strPass, False) = False Then
        MsgBox "无法连接到中间库，请先执行初始化脚本！", vbInformation, gstrSysName
        'Exit Function
    Else
        gstrSQL = "Update ypzlk set 收费细目ID=自编码 Where 收费细目ID is null"
        gcn铜山县.Execute gstrSQL
    End If
    If lng可更新 = 3 Then
        cmd未对码.Enabled = True
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    lng可更新 = 0
    cmd未对码.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If IsValid = False Then Exit Sub
    If SaveData = False Then Exit Sub
    
    mblnOK = True
    mblnChange = False
    Unload Me
End Sub

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    
    If Not IsNumeric(txtEdit(4).Text) Then
        MsgBox "请将串口号输入数字信息", vbInformation, gstrSysName
        Exit Function
    End If
    
    On Error Resume Next
'    mlngIcdev = init_com(txtEdit(4).Text - 1) 'Init COM2
'    If mlngIcdev <> 0 Then
'        If MsgBox("串口初始化失败，请检查串口。是否继续保存？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
'            txtEdit(4).SetFocus
'            Exit Function
'        End If
'    End If
'    st = close_com()
    IsValid = True
End Function

Public Function 参数设置() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim str参数值 As String
    
    mblnOK = False
    On Error Resume Next
    txtEdit(4).Text = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "当前使用的串口") + 1
    
    strUser = "tsxyb"
    strServer = GetSetting("ZLSOFT", "注册信息\登陆信息", "SERVER", "")
    strPass = "tsxyb"
    lngRount = 0
    If OraDataOpen(gcn铜山县, strServer, strUser, strPass, False) = False Then
        MsgBox "无法连接到中间库，请先执行初始化脚本！", vbInformation, gstrSysName
        'Exit Function
    End If
        
    gstrSQL = "Select * from czry where P_GLY=1"
    Call OpenRecordset_OtherBase(rsTemp, "czry", , gcn铜山县)
    If rsTemp.EOF = False Then
        txtUser.Text = rsTemp!P_RYH
        txtPass.Text = rsTemp!P_MM
    End If
    
    mblnChange = False
    frmSet铜山县.Show vbModal, frm医保类别
    参数设置 = mblnOK
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SaveData() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lst As ListItem
    
    '将当前使用的串口写入注册表之中
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "当前使用的串口", CStr(txtEdit(4).Text - 1)
    If Len(Trim(txtUser.Text)) > 0 Then
        
        gstrSQL = "Delete czry where P_GLY=1"
        gcn铜山县.Execute gstrSQL
        gstrSQL = "Insert into czry(P_RYH,P_XM,P_MM,P_GLY) values('" & Trim(txtUser.Text) & _
                "','管理员','" & Trim(txtPass.Text) & "',1)"
        gcn铜山县.Execute gstrSQL
    
    End If
    
    
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
    If Index = 4 Then
        If CheckIsInclude(UCase(Chr(KeyAscii)), "正整数") = True Then KeyAscii = 0
    End If
End Sub

Private Function CheckIsInclude(strSource As String, strTarge As String) As Boolean
    '检查strSource中的每一个字符是否在strTarge中
    Dim i As Long
    CheckIsInclude = False
    
    Select Case strTarge
    Case "日期"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "时间"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+-_)(*&^%$#@!`~"
    Case "日期时间"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+_)(*&^%$#@!`~"
    Case "整数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "小数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "正整数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "正小数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "可打印字符"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/."":;|\=+-_)(*&^%$#@!`~0123456789"
    End Select
    For i = 1 To Len(strSource)
        If InStr(strTarge, Mid(strSource, i, 1)) <= 0 Then Exit Function
    Next
    CheckIsInclude = True
End Function

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
    If Index = 4 Then
        If Not IsNumeric(txtEdit(4).Text) Then
            MsgBox "请将串口号输入数字信息", vbInformation, gstrSysName
        End If
    End If
End Sub




