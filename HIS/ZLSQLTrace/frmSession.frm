VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSession 
   AutoRedraw      =   -1  'True
   Caption         =   "�Ự"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   8655
   Icon            =   "frmSession.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5265
   ScaleWidth      =   8655
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picSession 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      ScaleHeight     =   3015
      ScaleWidth      =   8655
      TabIndex        =   17
      Top             =   0
      Width           =   8655
      Begin VB.CheckBox chkLike 
         Caption         =   "ģ������"
         Height          =   255
         Left            =   2280
         TabIndex        =   21
         Top             =   83
         Width           =   1095
      End
      Begin VB.CheckBox chkTraceOnly 
         Caption         =   "ֻ��ʾ�Ѹ��ٵ��û�"
         Height          =   255
         Left            =   3600
         TabIndex        =   22
         Top             =   83
         Width           =   2055
      End
      Begin VB.TextBox txtLocate 
         Height          =   300
         Left            =   1380
         TabIndex        =   20
         Top             =   60
         Width           =   855
      End
      Begin VSFlex8Ctl.VSFlexGrid vsSession 
         Height          =   2400
         Left            =   0
         TabIndex        =   18
         Top             =   480
         Width           =   8655
         _cx             =   15266
         _cy             =   4233
         Appearance      =   2
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmSession.frx":058A
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   1
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin MSComDlg.CommonDialog cdgFile 
            Left            =   3090
            Top             =   1290
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComctlLib.ImageList imgSession 
            Left            =   2310
            Top             =   1260
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   4
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSession.frx":06BD
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSession.frx":0C57
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSession.frx":11F1
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSession.frx":178B
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin VB.Label lblPrompt 
         Caption         =   "��ʾ"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5760
         TabIndex        =   23
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label lblLocate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�û�������(&S)"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   1170
      End
   End
   Begin VB.Frame fraUD 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   30
      MousePointer    =   7  'Size N S
      TabIndex        =   11
      Top             =   3120
      Width           =   7110
   End
   Begin VB.Frame fraTrace 
      Caption         =   " Trace �ļ����� "
      Height          =   2000
      Left            =   30
      TabIndex        =   10
      Top             =   3240
      Width           =   8595
      Begin ZLSQLTrace.ccXPButton cmdOut 
         Height          =   240
         Left            =   8175
         TabIndex        =   3
         Top             =   645
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   423
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ZLSQLTrace.ccXPButton cmdIn 
         Height          =   240
         Left            =   6285
         TabIndex        =   1
         Top             =   285
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   423
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ZLSQLTrace.ccXPButton cmdTrace 
         Height          =   345
         Left            =   7305
         TabIndex        =   8
         Top             =   1005
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "����(&A)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cboCount 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3450
         TabIndex        =   5
         Text            =   "cboCount"
         Top             =   1035
         Width           =   975
      End
      Begin VB.CheckBox chkNoSys 
         Caption         =   "�ſ�ϵͳ���"
         Height          =   195
         Left            =   4515
         TabIndex        =   6
         Top             =   1088
         Value           =   1  'Checked
         Width           =   1380
      End
      Begin VB.CheckBox chkOpen 
         Caption         =   "�Զ���"
         Height          =   195
         Left            =   6240
         TabIndex        =   7
         Top             =   1088
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.ComboBox cboSort 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1035
         Width           =   1620
      End
      Begin VB.TextBox txtCmd 
         Height          =   465
         IMEMode         =   2  'OFF
         Left            =   990
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1425
         Width           =   7470
      End
      Begin VB.TextBox txtOut 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   990
         TabIndex        =   2
         Top             =   615
         Width           =   7470
      End
      Begin VB.TextBox txtIn 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   990
         TabIndex        =   0
         Top             =   255
         Width           =   5550
      End
      Begin ZLSQLTrace.ccXPButton cmdFile 
         Height          =   300
         Left            =   6600
         TabIndex        =   24
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         Caption         =   "��ȡZLTRACE�ļ�"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   2685
         TabIndex        =   16
         Top             =   1095
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʽ"
         Height          =   180
         Left            =   210
         TabIndex        =   15
         Top             =   1095
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ı�"
         Height          =   180
         Index           =   1
         Left            =   210
         TabIndex        =   14
         Top             =   1470
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ŀ���ļ�"
         Height          =   180
         Left            =   210
         TabIndex        =   13
         Top             =   675
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Դ�ļ�"
         Height          =   180
         Left            =   210
         TabIndex        =   12
         Top             =   315
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event PopSessionMenu() '�Ӵ��������¼�
Public Event OpenNewFile(ByVal File As String)
Public Event UpdateStatus(ByVal strStatus As String)
Private mcolTrace As New Collection
Private mstrDest As String
Private mstrDBName As String
Private mblnEv As Boolean
Public mlngCount As Long

Private mcolCon As New Collection

Public Sub ShowMe(frmMain As Object)
    mlngCount = 0
    Me.Show
End Sub

Public Sub DoCommand(ByVal DoID As CommandBarIDCond, Optional ByVal blnIsZlTraceFile)
'���ܣ��Ӵ�������ִ�нӿ�
    Dim intEv As Integer, strTmp As String
    Dim cnTmp As New adodb.Connection, strConnect As String
    Dim rstmp As adodb.Recordset, strSql As String, strNewInstance

    On Error GoTo errh
    
    With vsSession
        Select Case DoID
        Case conMenu_Edit_Trace_1, conMenu_Edit_Trace_4, conMenu_Edit_Trace_8, conMenu_Edit_Trace_12
            
            If gblnIsRac Then
                If .TextMatrix(.Row, .ColIndex("Inst_ID")) <> gintInstId Then
                    Set cnTmp = GetConnection(.TextMatrix(.Row, .ColIndex("Inst_ID")))
                    If cnTmp.ConnectionString <> "" Then
                         Set gcnOracle = cnTmp
                         gintInstId = Val(.TextMatrix(.Row, .ColIndex("Inst_ID")))
                    Else
                        strSql = "Select Inst_Name From Gv$active_Instances Where Inst_Id = [1]"
                        Set rstmp = OpenSQLRecord(strSql, "GETINSTANCESNAME", .TextMatrix(.Row, .ColIndex("Inst_ID")))
                        strNewInstance = rstmp!Inst_Name
                        Set cnTmp = gcnOracle
                        strConnect = gcnOracle.ConnectionString
reLogin:
                        MsgBox "��ǰ��¼ʵ����������ʵ��(" & strNewInstance & ")��һ��,�����µ�¼��ʵ��(" & strNewInstance & ")����ִ�и��١�"
                        frmUserLogin.Show 1
                        
                        '������ȡ��,�Ͳ�ִ�и��ٷ���
                        If gcnOracle Is Nothing Then
                            Set gcnOracle = cnTmp
                            Exit Sub
                        End If
                        '����ٴε�¼ͬһ��ʵ��,�͵�������Ҫ�����µ�¼
                        If gcnOracle.ConnectionString = strConnect Then GoTo reLogin
                        mcolCon.Add gcnOracle, "_" & gintInstId   '���漯��
                    End If
                End If
            End If
            
            'If MsgBox("ȷ��Ҫ�� �û��� = " & .TextMatrix(.Row, vssession.colindex("�û���")) & ",SID = " & .TextMatrix(.Row, vssession.colindex("SID")) & ",Serial# = " & .TextMatrix(.Row, vssession.colindex("Serial#")) & " �ĻỰ���и�����", vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbNo Then Exit Sub
            Me.Tag = "���ڸ���"
                                    
            intEv = Decode(DoID, conMenu_Edit_Trace_1, 1, conMenu_Edit_Trace_4, 4, conMenu_Edit_Trace_8, 8, conMenu_Edit_Trace_12, 12)
            If DoID = conMenu_Edit_Trace_1 Then
                '���ַ�ʽ���ܿ�Щ
                gcnOracle.Execute "SYS.DBMS_System.Set_SQL_Trace_In_Session(" & .TextMatrix(.Row, vsSession.ColIndex("SID")) & "," & .TextMatrix(.Row, vsSession.ColIndex("Serial#")) & ",True)", , adCmdStoredProc
            Else
                gcnOracle.Execute "SYS.DBMS_System.Set_Bool_Param_In_Session(" & .TextMatrix(.Row, vsSession.ColIndex("SID")) & "," & .TextMatrix(.Row, vsSession.ColIndex("Serial#")) & ",'Timed_Statistics',True)", , adCmdStoredProc
                gcnOracle.Execute "SYS.DBMS_System.Set_Ev(" & .TextMatrix(.Row, vsSession.ColIndex("SID")) & "," & .TextMatrix(.Row, vsSession.ColIndex("Serial#")) & ",10046," & intEv & ",'')", , adCmdStoredProc
            End If
            
            mlngCount = mlngCount + 1
            lblPrompt.Caption = "�����û�" & .TextMatrix(.Row, vsSession.ColIndex("�û���")) & "(" & .TextMatrix(.Row, vsSession.ColIndex("SID")) & "," & .TextMatrix(.Row, vsSession.ColIndex("Serial#")) & ")�������ɹ���"
            .RowData(.Row) = intEv
            Set .Cell(flexcpPicture, .Row, 0) = imgSession.ListImages(Decode(intEv, 1, 1, 4, 2, 8, 3, 12, 4)).Picture
            .TextMatrix(.Row, .ColIndex("����״̬")) = "�Ѹ���"
            .TextMatrix(.Row, .ColIndex("���ٵȼ�")) = intEv
            Err.Clear: On Error Resume Next
            mcolTrace.Add intEv, "_" & .TextMatrix(.Row, vsSession.ColIndex("SID")) & "_" & .TextMatrix(.Row, vsSession.ColIndex("Serial#"))
            Err.Clear: On Error GoTo errh
        Case conMenu_Edit_TraceOff
            'If MsgBox("ȷ��Ҫֹͣ�� �û��� = " & .TextMatrix(.Row, vssession.colindex("�û���")) & ",SID = " & .TextMatrix(.Row, vssession.colindex("SID")) & ",Serial# = " & .TextMatrix(.Row, vssession.colindex("Serial#")) & " �ĻỰ���и�����", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
            Me.Tag = ""
            
            gcnOracle.Execute "SYS.DBMS_System.Set_Bool_Param_In_Session(" & .TextMatrix(.Row, vsSession.ColIndex("SID")) & "," & .TextMatrix(.Row, vsSession.ColIndex("Serial#")) & ",'Timed_Statistics',False)", , adCmdStoredProc
            gcnOracle.Execute "SYS.DBMS_System.Set_Ev(" & .TextMatrix(.Row, vsSession.ColIndex("SID")) & "," & .TextMatrix(.Row, vsSession.ColIndex("Serial#")) & ",10046,0,'')", , adCmdStoredProc
            'ǿ�а����ֶ�ͣ��,��Set_SQL_Trace_In_SessionҪ���ں���
            gcnOracle.Execute "SYS.DBMS_System.Set_SQL_Trace_In_Session(" & .TextMatrix(.Row, vsSession.ColIndex("SID")) & "," & .TextMatrix(.Row, vsSession.ColIndex("Serial#")) & ",False)", , adCmdStoredProc
                        
            mlngCount = mlngCount - 1
            lblPrompt.Caption = "ֹͣ�����û�" & .TextMatrix(.Row, vsSession.ColIndex("�û���")) & "(" & .TextMatrix(.Row, vsSession.ColIndex("SID")) & "," & .TextMatrix(.Row, vsSession.ColIndex("Serial#")) & ")�������ɹ���"
            .RowData(.Row) = 0
            Set .Cell(flexcpPicture, .Row, 0) = Nothing
            .TextMatrix(.Row, .ColIndex("����״̬")) = ""
            .TextMatrix(.Row, .ColIndex("���ٵȼ�")) = ""
            Err.Clear: On Error Resume Next
            mcolTrace.Remove "_" & .TextMatrix(.Row, vsSession.ColIndex("SID")) & "_" & .TextMatrix(.Row, vsSession.ColIndex("Serial#"))
            Err.Clear: On Error GoTo errh
            
            If gstrFilePath = "" Then
                gstrFilePath = GetDirName
                If gstrFilePath = "" Then
                    lblPrompt.Caption = "δѡ�񱣴�·�����޷����档"
                    Exit Sub
                End If
                Call SaveSetting("ZLSOFT\����ģ��\ZLDBATools", "Setting", "TraceFilePath", gstrFilePath)
            End If
            
            strTmp = GetTraceFile
            If strTmp <> "" Then
                txtIn.Text = strTmp
            End If
            
            
        
        Case conMenu_Edit_ChangeReg
            strTmp = GetDirName
            If strTmp <> "" Then
                gstrFilePath = strTmp
                Call SaveSetting("ZLSOFT\����ģ��\ZLDBATools", "Setting", "TraceFilePath", gstrFilePath)
            End If
            
        Case conMenu_View_Refresh
            Call LoadSession
        End Select
    End With
    
    Exit Sub
errh:
    MsgBox Err.Number & ":" & vbCrLf & vbCrLf & Err.Description, vbCritical, App.Title
    If 0 = 1 Then
        Resume
    End If
End Sub

Public Function GetCommand(ByVal DoID As CommandBarIDCond) As Boolean
'���ܣ��Ӵ�������״̬�ӿ�
    Select Case DoID
    Case conMenu_Edit_Trace
        GetCommand = mblnEv
        
    Case conMenu_Edit_Trace_1, conMenu_Edit_Trace_4, conMenu_Edit_Trace_8, conMenu_Edit_Trace_12
        If vsSession.Rows > vsSession.FixedRows Then
            GetCommand = vsSession.RowData(vsSession.Row) = 0 And mblnEv
        End If
    Case conMenu_Edit_TraceOff
        GetCommand = Me.Tag = "���ڸ���" Or Me.Tag = "�Ѹ���"
    Case conMenu_View_Refresh
        GetCommand = gcnOracle.State = 1
    End Select
End Function

Private Sub cboCount_Change()
    Call MakeCmdLine
End Sub

Private Sub cboCount_Click()
    Call MakeCmdLine
End Sub

Private Sub cboCount_GotFocus()
    cboCount.SelStart = 0: cboCount.SelLength = Len(cboCount.Text)
End Sub

Private Sub cboCount_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub cboSort_Click()
    Call MakeCmdLine
End Sub

Private Sub chkNoSys_Click()
    Call MakeCmdLine
End Sub

Private Sub chkTraceOnly_Click()
    Dim i As Long
    Dim lngCount As Long
    
    With vsSession
        For i = .FixedRows To .Rows - .FixedRows
            If chkTraceOnly.Value = 1 Then
                If .Cell(flexcpPicture, i, 0) Is Nothing Then
                    .RowHidden(i) = True
                Else
                    .RowHidden(i) = False
                    lngCount = lngCount + 1
                End If
            Else
                .RowHidden(i) = False
                lngCount = lngCount + 1
            End If
        Next
        lblPrompt.Caption = "��" & lngCount & "���û��Ự��"
    End With
End Sub

Private Sub cmdFile_Click()
    Dim strTmp As String

    If gstrFilePath = "" Then
        gstrFilePath = GetDirName
        If gstrFilePath = "" Then
            lblPrompt.Caption = "δѡ�񱣴�·�����޷����档"
            Exit Sub
        End If
        Call SaveSetting("ZLSOFT\����ģ��\ZLDBATools", "Setting", "TraceFilePath", gstrFilePath)
    End If
    
    strTmp = GetTraceFile(True)
    If strTmp <> "" Then
        txtIn.Text = strTmp
    End If
    
End Sub

Private Sub cmdIn_Click()
    With Me.cdgFile
        .DialogTitle = "ѡ��Ҫ������ SQL Trace �ļ�"
        .Filter = "SQL Trace(*.trc)|*.trc"
        .flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
        
        If txtIn.Text <> "" Then
            .InitDir = gobjFile.GetParentFolderName(txtIn.Text)
            .FileName = gobjFile.GetFileName(txtIn.Text)
        Else
            .InitDir = GetSetting("ZLSOFT\����ģ��\ZLDBATools", "Setting", "Input", mstrDest)
            .FileName = ""
        End If
        
        .CancelError = True
        On Error GoTo errh
        .ShowOpen
        
        SaveSetting "ZLSOFT\����ģ��\ZLDBATools", "Setting", "Input", gobjFile.GetParentFolderName(.FileName)
        txtIn.Text = .FileName: txtIn.SetFocus
    End With
errh:
End Sub

Private Sub cmdOut_Click()
    With Me.cdgFile
        .DialogTitle = "ȷ��Ҫ�������ɵı����ļ�"
        .Filter = "SQL Trace(*.log)|*.log"
        .flags = &H200000 Or &H4 Or &H2 Or &H800 Or &H4000
                
        If txtOut.Text <> "" Then
            .InitDir = gobjFile.GetParentFolderName(txtOut.Text)
            .FileName = gobjFile.GetFileName(txtOut.Text)
        Else
            .InitDir = GetSetting("ZLSOFT\����ģ��\ZLDBATools", "Setting", "Output", mstrDest)
            If txtIn.Text <> "" Then
                .FileName = ChangeType(gobjFile.GetFileName(txtIn.Text), ".trc", ".log")
            Else
                .FileName = ""
            End If
        End If
        
        .CancelError = True
        On Error GoTo errh
        .ShowSave

        SaveSetting "ZLSOFT\����ģ��\ZLDBATools", "Setting", "Output", gobjFile.GetParentFolderName(.FileName)
        txtOut.Text = .FileName: txtOut.SetFocus
    End With
errh:
End Sub

Private Sub cmdTrace_Click()
    Dim lngTemp As Long, lngProcess As Long
    
    If txtIn.Text = "" Then
        MsgBox "��ȷ��Ҫ������ SQL Trace �ļ���", vbInformation, App.Title
        txtIn.SetFocus: Exit Sub
    End If
    If Not gobjFile.FileExists(txtIn.Text) Then
        MsgBox "ָ��Ҫ������ SQL Trace �ļ������ڡ�", vbInformation, App.Title
        txtIn.SetFocus: Exit Sub
    End If
    If txtOut.Text = "" Then
        MsgBox "��ȷ�����������ɵ� SQL Trace �ļ���", vbInformation, App.Title
        txtOut.SetFocus: Exit Sub
    End If
    If txtCmd.Text = "" Then
        MsgBox "�޷����н�����", vbInformation, App.Title
        txtIn.SetFocus: Exit Sub
    End If
        
    Screen.MousePointer = 11
    On Error GoTo errh
    lngTemp = Shell(txtCmd.Tag, vbHide)
    
    'Ϊʲô�еĻ�������û��Ӧ
    lngProcess = OpenProcess(Process_Query_Information, False, lngTemp)
    Do
        GetExitCodeProcess lngProcess, lngTemp
    Loop While lngTemp = Still_Active
    CloseHandle lngProcess
    
    '����Ƿ�����ɹ�
    Screen.MousePointer = 0
    If Dir(txtOut.Text) = "" Then
        MsgBox "��Դ�ļ��������ļ������ܰ������ļ�������š�" & vbNewLine & "���޸ĺ����½�����"
    Else
        If chkOpen.Value = 0 Then
            Screen.MousePointer = 0
            MsgBox "�ļ�������ɡ�", vbInformation, App.Title
        Else
            RaiseEvent OpenNewFile(txtOut.Text)
        End If
    End If
  
    Exit Sub
errh:
    Screen.MousePointer = 0
    MsgBox Err.Number & ":" & vbCrLf & vbCrLf & Err.Description, vbCritical, App.Title
End Sub

Private Sub Form_Activate()
    RaiseEvent UpdateStatus("���� " & vsSession.Rows - vsSession.FixedRows & " ���Ự" & "|��ǰ�û�:" & gstrDBUser)
End Sub

Private Sub Form_Load()
    Dim i As Long, strCol As String
    
    With cboSort
        .AddItem ""
        .AddItem "(����)ִ�д���:prscnt  number of times parse was called"
        .AddItem "(����)CPUʱ�� :prscpu  cpu time parsing"
        .AddItem "(����)��ʱ��  :prsela  elapsed time parsing"
        .AddItem "(����)�����  :prsdsk  number of disk reads during parse"
        .AddItem "(����)һ�¶�  :prsqry  number of buffers for consistent read during parse"
        .AddItem "(����)��ǰ��  :prscu   number of buffers for current read during parse"
        .AddItem "(����)Ӳ����  :prsmis  number of misses in library cache during parse"
        .AddItem "(ִ��)ִ�д���:execnt  number of execute was called"
        .AddItem "(ִ��)CPUʱ�� :execpu  cpu time spent executing"
        .AddItem "(ִ��)�ܵ�ʱ��:exeela  elapsed time executing"
        .AddItem "(ִ��)�����  :exedsk  number of disk reads during execute"
        .AddItem "(ִ��)һ�¶�  :exeqry  number of buffers for consistent read during execute"
        .AddItem "(ִ��)��ǰ��  :execu   number of buffers for current read during execute"
        .AddItem "(ִ��)��¼��  :exerow  number of rows processed during execute"
        .AddItem "(ִ��)Ӳ����  :exemis  number of library cache misses during execute"
        .AddItem "(��ȡ)ִ�д���:fchcnt  number of times fetch was called"
        .AddItem "(��ȡ)CPUʱ�� :fchcpu  cpu time spent fetching"
        
        .AddItem "(��ȡ)�ܵ�ʱ��:fchela  elapsed time fetching"
        
        .AddItem "(��ȡ)�����  :fchdsk  number of disk reads during fetch"
        .AddItem "(��ȡ)һ�¶�  :fchqry  number of buffers for consistent read during fetch"
        .AddItem "(��ȡ)��ǰ��  :fchcu   number of buffers for current read during fetch"
        .AddItem "(��ȡ)��¼��  :fchrow  number of rows fetched"
        .ListIndex = 18
        SendMessage .hWnd, CB_SETDROPPEDWIDTH, 7000 / Screen.TwipsPerPixelX, 0
        SetWindowPos .hWnd, 0, 0, 0, .Width / Screen.TwipsPerPixelX, 5000 / Screen.TwipsPerPixelY, &H2
        
        Set gcolSort = New Collection
        For i = 1 To .ListCount - 1
            gcolSort.Add Trim(Split(.List(i), ":")(0)), "_" & UCase(Split(Split(.List(i), ":")(1), " ")(0))
        Next
    End With
    
    With cboCount
        .AddItem ""
        .AddItem "ǰ50��": .ItemData(.NewIndex) = 50
        .AddItem "ǰ100��": .ItemData(.NewIndex) = 100
        .AddItem "ǰ200��": .ItemData(.NewIndex) = 200
        .ListIndex = 0
    End With
    lblPrompt.Caption = ""
    
    Call InitData
    
    strCol = "  ,2000,1;" & IIf(gblnIsRac, "Inst_ID,2000,1;", "") & "�û���,2000,1;SID,1500,4;Serial#,1500,1;״̬,1500,1;����,1500,1;����,1500,1;����վ,500,1;ϵͳ�û�,500,1;��¼����,500,1;��¼ʱ��,1500,1;����״̬,500,1;���ٵȼ�,500,1"
    Call InitTable(vsSession, strCol)
    Call LoadSession
    
    '��ע����ȡ·��
    gstrFilePath = GetSetting("ZLSOFT\����ģ��\ZLDBATools", "Setting", "TraceFilePath")
    
    If gstrDBUser <> "" Then
        gblnZlhis = CheckZlhis
    End If
    
    cmdFile.Visible = gstrDBUser <> "" And gblnZlhis
    
    mcolCon.Add gcnOracle, "_" & gintInstId   '���漯��
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    
    picSession.Left = 0
    picSession.Top = 0
    picSession.Width = Me.ScaleWidth
    picSession.Height = Me.ScaleHeight - fraTrace.Height - fraUD.Height - 45
    
    
    fraUD.Left = 0
    fraUD.Top = picSession.Top + picSession.Height
    fraUD.Width = Me.ScaleWidth
    
    fraTrace.Left = 60
    fraTrace.Top = fraUD.Top + fraUD.Height
    fraTrace.Width = Me.ScaleWidth - fraTrace.Left * 2
    
    If fraTrace.Width - txtIn.Left - 200 >= 7300 Then
        txtIn.Width = fraTrace.Width - txtIn.Left - IIf(cmdFile.Visible, cmdFile.Width + 245, 200)
        cmdFile.Left = txtIn.Left + txtIn.Width + 45
        cmdIn.Left = txtIn.Left + txtIn.Width - cmdIn.Width - 30
        txtOut.Width = fraTrace.Width - txtOut.Left - 200
        cmdOut.Left = txtOut.Left + txtOut.Width - cmdOut.Width - 30
        
        cmdTrace.Left = txtOut.Left + txtOut.Width - cmdTrace.Width - 30
        chkOpen.Left = cmdTrace.Left - chkOpen.Width - 100
        txtCmd.Width = txtOut.Width
    End If
End Sub

Private Sub fraUD_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Exit Sub
    If Button = 1 Then
        If picSession.Height + y < 1000 Or fraTrace.Height - y < 2000 Then Exit Sub
        fraUD.Top = fraUD.Top + y
        picSession.Height = picSession.Height + y
        fraTrace.Top = fraTrace.Top + y
        fraTrace.Height = fraTrace.Height - y
    End If
End Sub

Private Sub InitData()
    Dim rstmp As adodb.Recordset
    Dim cmdTmp As adodb.Command
    Dim strSql As String
    
    If gcnOracle.State = 0 Then
        mblnEv = False: Exit Sub
    End If
    
    Err.Clear: On Error Resume Next
    'ORA-00942: �����ͼ������
    
    'Trace�ļ�����Ŀ¼
    strSql = "Select Value as Ŀ¼ From v$Parameter Where Name='user_dump_dest'"
    Set rstmp = New adodb.Recordset
    rstmp.CursorLocation = adUseClient
    rstmp.Open strSql, gcnOracle, adOpenKeyset
    If Not rstmp.EOF Then mstrDest = Nvl(rstmp!Ŀ¼)
    
    '��ǰ���ݿ�ʵ����
    strSql = "Select SYS_CONTEXT('USERENV','DB_NAME') as ���� From Dual"
    Set rstmp = New adodb.Recordset
    rstmp.CursorLocation = adUseClient
    rstmp.Open strSql, gcnOracle, adOpenKeyset
    If Not rstmp.EOF Then mstrDBName = Nvl(rstmp!����)
    
    '����Ƿ���ʹ��SYS.DBMS_SYSTEM��Ȩ��
    mblnEv = True
    Err.Clear
    Set cmdTmp = New adodb.Command
    Set cmdTmp.ActiveConnection = gcnOracle
    cmdTmp.CommandType = adCmdStoredProc
    cmdTmp.CommandText = "SYS.DBMS_System.Read_Ev"
    cmdTmp.Parameters.Append cmdTmp.CreateParameter("IEv", adNumeric, adParamInput, , 10046)
    cmdTmp.Parameters.Append cmdTmp.CreateParameter("OEv", adNumeric, adParamOutput)
    cmdTmp.Execute
    If InStr(Err.Description, "PLS-00201") > 0 Then mblnEv = False
End Sub

Private Sub LoadSession()
    Dim rstmp As New adodb.Recordset
    Dim strSql As String, i As Long
    Dim strPre As String

    If gcnOracle.State = 0 Then
        vsSession.Rows = vsSession.FixedRows
        Exit Sub
    End If
    
    On Error GoTo errh
    
    Screen.MousePointer = 11
    If gblnZlhis Then
        strSql = "Select B.�û���,A.����,C.���� as ���� From ��Ա�� A,�ϻ���Ա�� B,���ű� C,������Ա D " & vbCrLf & _
                 "Where A.ID=B.��ԱID And A.ID = D.��ԱID And D.ȱʡ = 1 And D.����ID = C.ID"
        strSql = "Select A.*,B.����,b.���� From gv$Session A,(" & strSql & ") B Where A.UserName=B.�û���(+) And A.UserName Is Not Null  And A.AUDSID <> userenv('sessionid') Order By A.UserName"
    Else
        strSql = "Select A.*,Null as ����,Null as ���� From gv$Session A Where A.UserName Is Not Null And A.AUDSID <> userenv('sessionid') Order By A.Logon_Time,A.SID"
    End If
    rstmp.Open strSql, gcnOracle, adOpenKeyset, adLockOptimistic
    
    With vsSession
        strPre = .TextMatrix(.Row, vsSession.ColIndex("SID")) & "," & .TextMatrix(.Row, vsSession.ColIndex("Serial#"))
        
        .Redraw = flexRDNone
        .Rows = .FixedRows '�������
        .Rows = .FixedRows + rstmp.RecordCount
        For i = 1 To rstmp.RecordCount
            If gblnIsRac Then
                .TextMatrix(i, .ColIndex("Inst_ID")) = rstmp!INST_ID
            End If
            .TextMatrix(i, .ColIndex("�û���")) = rstmp!UserName
            .TextMatrix(i, .ColIndex("SID")) = rstmp!SID
            .TextMatrix(i, .ColIndex("Serial#")) = rstmp.Fields("Serial#").Value
            .TextMatrix(i, .ColIndex("״̬")) = Decode(rstmp!Status, "ACTIVE", "��ǰ", "INACTIVE", "����", rstmp!Status)
            .TextMatrix(i, .ColIndex("����")) = Nvl(rstmp!����)
            .TextMatrix(i, .ColIndex("����")) = Nvl(rstmp!����)
            .TextMatrix(i, .ColIndex("����վ")) = Nvl(rstmp!Machine)
            .TextMatrix(i, .ColIndex("ϵͳ�û�")) = Nvl(rstmp!OSUser)
            .TextMatrix(i, .ColIndex("��¼����")) = Nvl(rstmp!Program) & IIf(Not IsNull(rstmp!Action), ":" & rstmp!Action, "")
            .TextMatrix(i, .ColIndex("��¼ʱ��")) = rstmp!Logon_Time
            
            '����Ƿ��д��ڸ���״̬�ĻỰ
            .TextMatrix(i, .ColIndex("����״̬")) = IIf(Nvl(rstmp!SQL_TRACE) = "ENABLED", "�ѿ�������", "")
'            ���ٵȼ���Ӧ
'            �ȼ�    SQL_TRACE   SQL_TRACE_WAITS    SQL_TRACE_BINDS
'            1         ENABLED           FALSE                           FALSE
'            4         ENABLED           FALSE                           TRUE
'            8         ENABLED           TRUE                             FALSE
'            12       ENABLED           TRUE                            TRUE
            
            If Nvl(rstmp!SQL_TRACE) = "ENABLED" Then
                Select Case Nvl(rstmp!SQL_TRACE_WAITS)
                    Case "FALSE"
                        .TextMatrix(i, .ColIndex("���ٵȼ�")) = IIf(rstmp!SQL_TRACE_BINDS = "FALSE", 1, 4)
                    Case "TRUE"
                        .TextMatrix(i, .ColIndex("���ٵȼ�")) = IIf(rstmp!SQL_TRACE_BINDS = "FALSE", 8, 12)
                End Select
                Set .Cell(flexcpPicture, i, 0) = imgSession.ListImages(Decode(Val(.TextMatrix(i, .ColIndex("���ٵȼ�"))), 1, 1, 4, 2, 8, 3, 12, 4)).Picture
            End If
            
            Err.Clear: On Error Resume Next
            .RowData(i) = mcolTrace("_" & rstmp!SID & "_" & rstmp.Fields("Serial#").Value)
            Err.Clear: On Error GoTo errh
            If .RowData(i) <> 0 Then
                Set .Cell(flexcpPicture, i, 0) = imgSession.ListImages(Decode(.RowData(i), 1, 1, 4, 2, 8, 3, 12, 4)).Picture
            End If

            rstmp.MoveNext
        Next
        If .Rows > 2 Then
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, 0) = 4
            .AutoSize 1, 9
            .Row = 1: .Col = 1
            .ShowCell .Row, .Col
        End If
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDDirect
        
        lblPrompt.Caption = "��" & .Rows - .FixedRows & "���û��Ự��"
    End With
    Screen.MousePointer = 0
    Exit Sub
errh:
    Screen.MousePointer = 0
    If InStr(Err.Description, "ORA-00942") > 0 Then
        MsgBox "��ǰ��¼�û�û�ж�ȡ���ݿ�Ự��Ϣ v$Session ��Ȩ�ޡ�", vbCritical, App.Title
    Else
        MsgBox Err.Number & ":" & vbCrLf & vbCrLf & Err.Description, vbCritical, App.Title
    End If
    If 0 = 1 Then
        Resume
    End If
End Sub

Private Sub picSession_Resize()
    On Error Resume Next
    
    vsSession.Top = txtLocate.Top + txtLocate.Height + 60
    vsSession.Left = picSession.Left
    vsSession.Width = picSession.Width
    vsSession.Height = picSession.Height - txtLocate.Height - 60
    
    lblPrompt.Width = vsSession.Width - lblPrompt.Left
End Sub

Private Sub txtCmd_GotFocus()
    txtCmd.SelStart = 0: txtCmd.SelLength = Len(txtCmd.Text)
End Sub

Private Sub txtIn_Change()
    If txtIn.Text <> "" Then
        If gobjFile.FileExists(txtIn.Text) Then
            If txtOut.Text = "" Then
                txtOut.Text = ChangeType(txtIn.Text, ".trc", ".log")
            Else
                txtOut.Text = gobjFile.GetParentFolderName(txtOut.Text) & "\" & ChangeType(gobjFile.GetFileName(txtIn.Text), ".trc", ".log")
            End If
        End If
    End If

    Call MakeCmdLine
End Sub

Private Sub txtIn_GotFocus()
    txtIn.SelStart = 0: txtIn.SelLength = Len(txtIn.Text)
End Sub

Private Sub txtLocate_Change()
    txtLocate.Tag = ""
End Sub

Private Sub txtLocate_GotFocus()
    txtLocate.SelStart = 0
    txtLocate.SelLength = Len(txtLocate.Text)
End Sub

Private Sub txtLocate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Len(Trim(txtLocate.Text)) > 0 Then
        Dim lngRow As Long
        Dim strText As String
        Dim lngStart As Long
        
        If txtLocate.Tag = "" Then
            lngStart = vsSession.FixedRows
        Else
            lngStart = Val(txtLocate.Tag)
        End If
        
        strText = UCase(Trim(txtLocate.Text))
        If chkLike.Value = 1 Then
            lngRow = vsSession.FindRow(strText, lngStart, vsSession.ColIndex("�û���"), , False)
        Else
            lngRow = vsSession.FindRow(strText, lngStart, vsSession.ColIndex("�û���"))
        End If
        If lngRow > 0 Then
            txtLocate.Tag = lngRow + 1
            vsSession.Row = lngRow
            vsSession.TopRow = lngRow
        Else
            If lngStart = vsSession.FixedRows Then
                lblPrompt.Caption = "û���ҵ�ƥ����û�����"
            Else
                lblPrompt.Caption = "��ǰ�������һ��������û���ˡ�"
                txtLocate.Tag = ""
            End If
        End If
        Call txtLocate_GotFocus
    End If
End Sub

Private Sub txtOut_Change()
    Call MakeCmdLine
End Sub

Private Sub txtOut_GotFocus()
    txtOut.SelStart = 0: txtOut.SelLength = Len(txtOut.Text)
End Sub

Private Sub vsSession_AfterSort(ByVal Col As Long, Order As Integer)
    Dim lngRow As Long
    Dim i As Long
    
    If chkTraceOnly.Tag <> "" Then
        With vsSession
            For i = .FixedRows To .Rows - .FixedRows
                If .TextMatrix(i, vsSession.ColIndex("SID")) & "," & .TextMatrix(i, vsSession.ColIndex("Serial#")) = chkTraceOnly.Tag Then
                    .Row = i
                    .TopRow = i
                End If
            Next
        End With
    End If
End Sub

Private Sub vsSession_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If Me.Tag = "���ڸ���" Then
        Cancel = True
        Exit Sub
    End If
    
    '�������ڸ��ٵĻỰ,��tag���Ϊ"�Ѹ���",�������ð�ť������
    With vsSession
        If .Redraw = flexRDNone Then Exit Sub
        If .TextMatrix(NewRow, .ColIndex("����״̬")) = "�ѿ�������" Then
            Me.Tag = "�Ѹ���"
        Else
            Me.Tag = ""
        End If
    End With
End Sub

Private Sub vsSession_BeforeSort(ByVal Col As Long, Order As Integer)
    With vsSession
        If .Row > .FixedRows Then
            chkTraceOnly.Tag = .TextMatrix(.Row, vsSession.ColIndex("SID")) & "," & .TextMatrix(.Row, vsSession.ColIndex("Serial#"))
        End If
    End With
End Sub

Private Sub vsSession_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsSession_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngRow As Long
    
    With vsSession
        lngRow = .MouseRow
        If Button = 2 And .MouseRow >= .FixedRows Then
            .Row = lngRow
            RaiseEvent PopSessionMenu
        End If
    End With
End Sub

Private Sub MakeCmdLine()
    Dim strIn As String, strOut As String
    Dim strCmd As String
    
    If txtIn.Text = "" Or txtOut.Text = "" Then
        txtCmd.Text = "": txtCmd.Tag = ""
    Else
        strCmd = "tkprof " & txtIn.Text & " " & txtOut.Text
        If cboSort.Text <> "" Then
            strCmd = strCmd & " sort=" & Split(Split(cboSort.Text, ":")(1), " ")(0)
        End If
        If cboCount.Text <> "" Then
            If cboCount.ListIndex = -1 Then
                If Int(Val(cboCount.Text)) > 0 Then
                    strCmd = strCmd & " print=" & Int(Val(cboCount.Text))
                End If
            ElseIf cboCount.ItemData(cboCount.ListIndex) > 0 Then
                strCmd = strCmd & " print=" & cboCount.ItemData(cboCount.ListIndex)
            End If
        End If
        
        If chkNoSys.Value = 1 Then strCmd = strCmd & " sys=no"
        txtCmd.Text = strCmd
        
        '������ļ�·��
        strIn = GetShortName(gobjFile.GetParentFolderName(txtIn.Text)) & "\" & gobjFile.GetFileName(txtIn.Text)
        strOut = GetShortName(gobjFile.GetParentFolderName(txtOut.Text)) & "\" & gobjFile.GetFileName(txtOut.Text)
        txtCmd.Tag = Replace(Replace(txtCmd.Text, txtIn.Text, strIn), txtOut.Text, strOut)
    End If
End Sub

Private Function ChangeType(ByVal strFile As String, ByVal strS As String, ByVal strO As String) As String
    If UCase(strFile) Like "*" & UCase(strS) Then
        strFile = Left(strFile, Len(strFile) - Len(strS)) & strO
    End If
    ChangeType = strFile
End Function

Private Function GetTraceFile(Optional ByVal blnIsZlTraceFile As Boolean) As String
    '����:��ȡ������Trace�ļ���ת��������
    '����˵����  blnIsZlTraceFile -����10046�¼����ٵ���־�ļ�
    
    Dim rstmp As adodb.Recordset, strSql As String
    Dim strFileName As String, strFilePath As String
    Dim objText As TextStream
    Dim intInstID As Integer
    
    On Error GoTo errh
    
    If gblnIsRac Then
        intInstID = Val(vsSession.TextMatrix(vsSession.Row, vsSession.ColIndex("Inst_ID")))
    Else
        intInstID = 1
    End If
    If blnIsZlTraceFile Then
    '���blnIsZlTraceFileΪTrue��˵����Ҫ��zlreginfo��ȡ�ļ���
        strSql = "Select ���� Filename, Value Filepath From zlRegInfo," & IIf(gblnIsRac, "G", "") & "V$parameter Where ��Ŀ = 'TRACE�ļ�' And Name = 'user_dump_dest'" & _
                IIf(gblnIsRac, " And Inst_ID = " & intInstID, "")
        Set rstmp = OpenSQLRecord(strSql, "DoCommand")
        
        If rstmp.RecordCount = 0 Then
            MsgBox "��ȡTrace�ļ�ʧ�ܣ����ȵ�¼����̨��ȡTrace�ļ���": Exit Function
        End If
        
        If rstmp!FileName = "" Then
            MsgBox "��ȡTrace�ļ�ʧ�ܣ����ȵ�¼����̨��ȡTrace�ļ���": Exit Function
        End If
        
    Else
    '����SID��Serial#ȷ��������Trace�ļ�·��
        strSql = "Select a.Value FilePath, c.Instance_Name || '_ora_' || d.Spid || '.trc' FileName" & vbNewLine & _
                        "From (Select Value From " & IIf(gblnIsRac, "G", "") & "V$parameter Where Name = 'user_dump_dest'" & IIf(gblnIsRac, " And Inst_ID = " & intInstID, "") & ") A," & vbNewLine & _
                        "     (Select Instance_Name From " & IIf(gblnIsRac, "G", "") & "V$instance " & IIf(gblnIsRac, " Where Inst_ID = " & intInstID, "") & ") C," & vbNewLine & _
                        "     (Select Spid From " & IIf(gblnIsRac, "G", "") & "V$session S, " & IIf(gblnIsRac, "G", "") & "V$process P " & vbNewLine & _
                        "Where s.Paddr = p.Addr And s.Sid = [1] And s.Serial# = [2] " & IIf(gblnIsRac, " And S.Inst_ID = " & intInstID & " And P.Inst_ID = " & intInstID, "") & " ) D"
        Set rstmp = OpenSQLRecord(strSql, "DoCommand", vsSession.TextMatrix(vsSession.Row, vsSession.ColIndex("Sid")), vsSession.TextMatrix(vsSession.Row, vsSession.ColIndex("Serial#")))
    End If
                    
    strFileName = rstmp!FileName
    strFilePath = rstmp!FilePath
    

    
    '����Ƿ���Ҫɾ����
    strSql = "Select 1 From User_Tables Where table_name ='TRACEFILE'"
    Set rstmp = OpenSQLRecord(strSql, "DoCommand")
    If rstmp.RecordCount > 0 Then
        strSql = "Drop table tracefile Purge"
        gcnOracle.Execute strSql
        strSql = "Drop directory zltracefile"
        gcnOracle.Execute strSql
    End If
    
    '����·��
    strSql = "create or replace directory zltracefile as '" & strFilePath & "'"
    gcnOracle.Execute strSql
    '�����ⲿ��
    strSql = "create  table TRACEFILE" & vbNewLine & _
                    " (TEXT varchar2(4000))" & vbNewLine & _
                    " organization external (" & vbNewLine & _
                    " type oracle_loader default directory zltracefile" & vbNewLine & _
                    " access parameters (" & vbNewLine & _
                    " records delimited by newline nobadfile nodiscardfile  nologfile  FIELDS (text (1:4000) CHAR) )" & vbNewLine & _
                    " location('" & strFileName & "')" & vbNewLine & _
                    " ) reject limit Unlimited"
                    
    gcnOracle.Execute strSql

    '��ѯ�ⲿ��
    strSql = "Select Text from TRACEFILE"
    Set rstmp = OpenSQLRecord(strSql, "DoCommand")
    
    '�������������ݴ������ش���
    Set objText = gobjFile.CreateTextFile(gstrFilePath & "\" & strFileName, True)
    
    Do While Not rstmp.EOF
        objText.WriteLine "" & rstmp!Text
        rstmp.MoveNext
    Loop
    objText.Close

    GetTraceFile = gstrFilePath & "\" & strFileName
    Exit Function
    
errh:
    GetTraceFile = ""
    If InStr(Err.Description, "ORA-29400") > 0 Then
        MsgBox "������·��" & strFilePath & "�²�����Trace�ļ�" & strFileName & "��" & vbNewLine & "ע���类���ٻỰû��ִ��SQL��䣬���������Trace�ļ���"
        Exit Function
    End If
    MsgBox Err.Description
    
    If 0 = 1 Then
        Resume
    End If
End Function

Public Sub InitTable(vsf As VSFlexGrid, strCol As String)
'����: ��ʼ����ͷ
    Dim arrHead As Variant
    Dim i As Long
    
    arrHead = Split(strCol, ";")
   
    With vsf
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .Cols = UBound(arrHead) + 1
        .Rows = .FixedRows
        .Editable = flexEDNone
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)
            .ColKey(i) = Split(arrHead(i), ",")(0)
            
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(i) = False
                .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(i) = True
                .ColWidth(i) = 0
            End If
        Next
        .Redraw = True
    End With
End Sub

Private Function GetConnection(intInstID As Integer) As adodb.Connection
    '����: ����ʵ���Ż�ȡ��Ӧ���Ӷ���
    Dim cnResult As New adodb.Connection
    
    On Error Resume Next
    Set cnResult = mcolCon("_" & intInstID)
    
    Set GetConnection = cnResult
End Function
