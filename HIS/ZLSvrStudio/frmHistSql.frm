VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmHistSql 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "��ʷSql���"
   ClientHeight    =   1305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12735
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
   ScaleHeight     =   1305
   ScaleWidth      =   12735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraExecute 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ִ����"
      Height          =   615
      Left            =   6960
      TabIndex        =   10
      Top             =   0
      Width           =   4095
      Begin VB.OptionButton optUser 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ǰ�û�(ZLHIS)"
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optSess 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ǰ�Ự(333,151)"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.Frame fraTime 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ʱ�䷶Χ"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   4335
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy/MM/dd HH:mm"
         Format          =   221184003
         CurrentDate     =   42961
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   315
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy/MM/dd HH:mm"
         Format          =   221184003
         CurrentDate     =   42961
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   195
         Left            =   2040
         TabIndex        =   9
         Top             =   300
         Width           =   180
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   45
      End
   End
   Begin VB.Frame fraData 
      BackColor       =   &H00FFFFFF&
      Caption         =   "������Դ"
      Height          =   615
      Left            =   4560
      TabIndex        =   2
      Top             =   0
      Width           =   2295
      Begin VB.OptionButton optHist 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ʷ��"
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         ToolTipText     =   "ȱʡ������,SQL�����ִ�к�һСʱ�Żᱣ������ʷ����"
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optCur 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ǰ����"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "����(&C)"
      Height          =   345
      Left            =   11280
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfSQL 
      Height          =   540
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4995
      _cx             =   8811
      _cy             =   952
      Appearance      =   2
      BorderStyle     =   1
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
      BackColorFixed  =   15921906
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   8421504
      GridColorFixed  =   8421504
      TreeColor       =   0
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   280
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
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
      Editable        =   2
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
   End
End
Attribute VB_Name = "frmHistSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintType As Integer     'ִ�мƻ���Դ:  1-v$��ͼ 2-��ʷ����
Private mdtStart As Date
Private mdtEnd As Date
Private mlngSid As Long
Private mlngSerial As Long
Private mstrUser As String  'ͨ���Ƿ����û����ж��ǵ������廹��Ƕ�봰��,�����������봫���û���

Public Sub SetUser(ByVal strUser As String)
    '�˷������ڲ�ѯ�û�����ʷSQLʱ�����û���
    mstrUser = strUser
    optUser.Caption = "��ǰ�û�(" & strUser & ")"
End Sub

Public Sub SetSid(ByVal lngSid As Long, ByVal lngSerial As Long)
    '�˷����û��������Ự����,Ҫ���� sid+Serial#
    mlngSid = lngSid
    mlngSerial = lngSerial
    optSess.Caption = "��ǰ�Ự(" & lngSid & "," & lngSerial & ")"
End Sub

Public Sub ShowMe(ByVal lngSid As Long, ByVal lngSerial As Long, ByVal dtStart As Date, ByVal dtEnd As Date, ByVal intSource As String)
    mintType = intSource
    mlngSid = lngSid
    mlngSerial = lngSerial
    mdtStart = dtStart
    mdtEnd = dtEnd
    Me.Show
End Sub

Private Sub cmdFind_Click()
    LoadSQL
End Sub

Private Sub Form_Load()
    Dim strCol As String
    
    If mstrUser <> "" Then
        dtpStart.value = date
        dtpEnd.value = date + 1
    Else
        dtpStart.value = mdtStart
        dtpEnd.value = mdtEnd
        optCur.value = mintType = 1
        optHist.value = mintType = 2
        optSess.Caption = "��ǰ�Ự(" & mlngSid & "," & mlngSerial & ")"
    End If
    
    If Val(gstrOracleBigVer) > 10 Then
        strCol = "  ,500,1;Sid,1200,1;�û�,1005,1;������,1980,1;��������,1365,1;ִ��ʱ��,1550,1;�ȴ�ʱ��(��s),1300,1;�ȴ��¼�,2005,1;" & _
                    "�����Ự,1200,1;SQL_ID,1300,1;SQL�ı�,1500,1"
    Else
        strCol = "  ,500,1;Sid,1200,1;�û�,1005,1;������,1980,1;��������,1365,1;����ʱ��,1550,1;�ȴ�ʱ��(��s),1300,1;�ȴ��¼�,2005,1;" & _
                    "�����Ự,1200,1;SQL_ID,1300,1;SQL�ı�,1500,1"
    End If
    
    InitTable vsfSQL, strCol
    vsfSQL.Rows = 1
    vsfSQL.TextMatrix(0, 1) = "Sid,Serial#"
    
    If mstrUser = "" Then   '������û���Ϊ��,˵���ǵ�������,ֻ������Sid
        fraExecute.Visible = False
        cmdFind.Left = fraData.Left + fraData.Width + 60
        LoadSQL
    End If
    
    If Val(gstrOracleBigVer) = 10 Then
        fraTime.Caption = "����ʱ��"
    Else
        fraTime.Caption = "���ִ��ʱ��"
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    vsfSQL.Width = Me.ScaleWidth - vsfSQL.Left - 120
    vsfSQL.Height = Me.ScaleHeight - vsfSQL.Top - 60
End Sub


Private Sub LoadSQL()
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim i As Long, strSID As String
    Dim dtStart As Date, dtEnd As Date
    
    On Error GoTo errH
    
    ShowFlash "���ڼ�����ʷSQL���..."
    dtStart = CDate(Format(dtpStart.value, "yyyy-MM-dd hh:mm:ss"))
    dtEnd = CDate(Format(dtpEnd.value, "yyyy-MM-dd hh:mm:ss"))
    
    If optCur.value Then    '��ȡ��ǰ����
        strSQL = "Select A.Session_Id || ',' || A.Session_Serial# As Sid,B.Machine,B.Username," & vbNewLine & _
                        IIf(mstrUser = "", "A.Time_Waited WaitTime ,", " A.Wait_TIME WaitTime,") & vbNewLine & _
                        "                       decode(A.Blocking_Session || ',' ||A.Blocking_Session_Serial#,',','',A.Blocking_Session || ',' ||A.Blocking_Session_Serial#) As Blocking_Session," & vbNewLine & _
                        "                A.Program, A.Sql_Id, C.Sql_Text," & vbNewLine & _
                        "                       To_Char( " & IIf(gstrOracleBigVer = 10, "A.Sample_Time", "A.sql_exec_start") & ", 'yyyy/mm/dd hh24:mi') sql_exec_start,A.Event" & vbNewLine & _
                        IIf(gblnRac, "From GV$active_Session_History A ,GV$Session B,GV$sql C", "From V$active_Session_History A ,V$Session B,v$sql C") & vbNewLine & _
                        "Where a.SESSION_ID = B.SID And A.SESSION_SERIAL# =B.SERIAL# And A.SQL_ID= C.SQL_ID" & vbNewLine & _
                        "And " & IIf(gstrOracleBigVer = 10, "A.Sample_Time", "A.sql_exec_start") & " Between [1] And [2] " & vbNewLine & _
                        IIf(mstrUser <> "", "And B.UserName = [3]", "") & vbNewLine & _
                        IIf(optSess.value, "And A.Session_Id =[4] And  A.Session_Serial# =[5]", "") & vbNewLine & _
                        "Order By " & IIf(gstrOracleBigVer = 10, "A.Sample_Time", "A.sql_exec_start") & " Desc ,A.Session_Id, A.Session_Serial#"
    Else    '��ȡ��ʷ����
        strSQL = "Select A.Session_Id || ',' || A.Session_Serial# As Sid,  " & IIf(gstrOracleBigVer = 10, " '' as Machine, ", "A.Machine,") & " c.Username, " & vbNewLine & _
                        IIf(mstrUser = "", "A.Time_Waited WaitTime ,", " A.Wait_TIME WaitTime,") & vbNewLine & _
                        "            decode(A.Blocking_Session || ',' ||A.Blocking_Session_Serial#,',','',A.Blocking_Session || ',' ||A.Blocking_Session_Serial#) As Blocking_Session," & vbNewLine & _
                        "             A.Program, d.Sql_Id, d.Sql_Text," & vbNewLine & _
                        "             To_Char( " & IIf(gstrOracleBigVer = 10, "A.Sample_Time", "A.sql_exec_start") & ", 'yyyy/mm/dd hh24:mi') sql_exec_start,Event" & vbNewLine & _
                        "From Dba_Hist_Active_Sess_History A, Dba_Users C, Dba_Hist_Sqltext D" & vbNewLine & _
                        "Where A.User_Id = c.User_Id  And A.Sql_Id(+) = d.Sql_Id And A.Sql_Id Is Not Null" & vbNewLine & _
                        "      And " & IIf(gstrOracleBigVer = 10, "A.Sample_Time", "A.sql_exec_start") & "  between [1] and [2]" & vbNewLine & _
                         IIf(mstrUser <> "", "And c.Username = [3]", "") & vbNewLine & _
                        IIf(optSess.value, "And A.Session_Id =[4] And  A.Session_Serial# =[5]", "") & vbNewLine & _
                        "Order By " & IIf(gstrOracleBigVer = 10, "A.Sample_Time", "A.sql_exec_start") & " Desc ,A.Session_Id, A.Session_Serial#"
    End If
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡSQL��Ϣ", dtpStart, dtpEnd, mstrUser, mlngSid, mlngSerial)
    
    With vsfSQL
        .Redraw = flexRDNone
        'Ƕ��ʽ����,ͨ���Ự��ѯ,ʡ���ظ���Ϣ
        .ColHidden(0) = optSess.value
        .ColHidden(.ColIndex("Sid")) = mstrUser <> "" And optSess.value
        .ColHidden(.ColIndex("�û�")) = mstrUser <> "" Or optUser.value
        If Val(gstrOracleBigVer) = 10 Then
            .ColHidden(.ColIndex("������")) = True
        Else
            .ColHidden(.ColIndex("������")) = optCur.value
        End If
        .ColHidden(.ColIndex("��������")) = mstrUser <> "" And optSess.value
        
        If rsTmp.RecordCount = 0 Then
            .Rows = 1
            ShowFlash ""
            If mstrUser = "" Then '��������û����ϢҪ������ʾ
                .Rows = 2
                .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = "�ûỰû�к�ʱ������ASH��¼��"
                .MergeCells = flexMergeRestrictRows
                .MergeRow(-1) = True
                .Select 1, 0
            End If
            .Redraw = flexRDDirect
            Exit Sub
        End If
        
        .MergeCol(.ColIndex("Sid-Serial#")) = True
        .MergeCells = flexMergeRestrictAll
        
        .OutlineBar = flexOutlineBarSimpleLeaf
        .SubtotalPosition = flexSTAbove
        .Rows = 1: i = 1
        .Rows = rsTmp.RecordCount + 1
        .ComboList = "..."
        'Sid",Serial#,1200,1;�û�,1000,1;������,1000,1;��������,1000,1;��¼ʱ��,1500,1;�ȴ�ʱ��,1000,1; �����Ự,1200,1;SQL_ID,800,1;SQL�ı�,2000,1
        Do While Not rsTmp.EOF
            If optUser.value Then
            .IsSubtotal(i) = True   '��ѯ��ǰ�Ựʱ,����Sid���з���
                If strSID = rsTmp!Sid & "" Then
                    .RowOutlineLevel(i) = 2
                Else
                    strSID = rsTmp!Sid & ""
                    .RowOutlineLevel(i) = 1
                End If
            End If
            .TextMatrix(i, .ColIndex("Sid")) = rsTmp!Sid & ""
            .TextMatrix(i, .ColIndex("�û�")) = rsTmp!USERNAME & ""
            .TextMatrix(i, .ColIndex("��������")) = rsTmp!Program & ""
            If Val(gstrOracleBigVer) < 11 Then
                .TextMatrix(i, .ColIndex("����ʱ��")) = rsTmp!sql_exec_start & ""
                If optCur.value Then .TextMatrix(i, .ColIndex("������")) = rsTmp!Machine & ""
            Else
                .TextMatrix(i, .ColIndex("ִ��ʱ��")) = rsTmp!sql_exec_start & ""
                .TextMatrix(i, .ColIndex("������")) = rsTmp!Machine & ""
            End If
            .TextMatrix(i, .ColIndex("�ȴ�ʱ��(��s)")) = rsTmp!WaitTime & ""
            .TextMatrix(i, .ColIndex("�ȴ��¼�")) = rsTmp!Event & ""
            .TextMatrix(i, .ColIndex("�����Ự")) = rsTmp!Blocking_Session & ""
            .TextMatrix(i, .ColIndex("SQL_ID")) = rsTmp!Sql_Id & ""
            .TextMatrix(i, .ColIndex("SQL�ı�")) = rsTmp!sql_text & ""
            i = i + 1
            rsTmp.MoveNext
        Loop
        
        .ColAlignment(.ColIndex("�ȴ�ʱ��(��s)")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("ִ��ʱ��")) = flexAlignCenterCenter
        
        .Redraw = flexRDDirect
        If .Rows > 1 Then .Select 1, 0
    End With
    
    ShowFlash ""
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    ShowFlash ""
    MsgBox "��ȡ��ʷSQL��䷢������" & vbNewLine & err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mstrUser <> "" Then
        Unload frmHistSqlParent
    End If
End Sub

Private Sub vsfSQL_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfSQL
        If .TextMatrix(Row, Col) <> "" And (Col = .ColIndex("SQL�ı�") Or (Col = .ColIndex("�����Ự") And mstrUser <> "")) Then
            Cancel = False
        Else
            Cancel = True
        End If
    End With
End Sub

Private Sub vsfSQL_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    With vsfSQL
        If .TextMatrix(NewRow, NewCol) <> "" And (NewCol = .ColIndex("SQL�ı�") Or (NewCol = .ColIndex("�����Ự") And mstrUser <> "")) Then
            '��ʾ��ť�����ж�:
            '1.sql�ı���Ϊ��
            '2.�����Ự��Ϊ�յ��ҷǵ�������(ͨ���Ƿ���SID�ж�,û�д���User,˵���ǵ�������)
            .ComboList() = "..."
            .FocusRect = flexFocusSolid
        Else
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsfSQL_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim lngSid As Long, lngSerial As Long
    
    With vsfSQL
        Select Case Col
        
        Case .ColIndex("SQL�ı�")
            If optCur.value Then
                frmHistSqlPlan.ShowMe .TextMatrix(Row, .ColIndex("SQL_ID")), 1  '��ǰ����,intType=1
            Else
                frmHistSqlPlan.ShowMe .TextMatrix(Row, .ColIndex("SQL_ID")), 2
            End If
        Case .ColIndex("�����Ự")
            lngSid = Split(.TextMatrix(Row, Col), ",")(0)
            lngSerial = Split(.TextMatrix(Row, Col), ",")(1)
            If optCur.value Then
                frmHistSqlParent.ShowMe lngSid, lngSerial, dtpStart.value, dtpEnd.value, 1
            Else
                frmHistSqlParent.ShowMe lngSid, lngSerial, dtpStart.value, dtpEnd.value, 2
            End If
        End Select
    End With
End Sub

Private Sub vsfSQL_DblClick()
    With vsfSQL
        If .Rows = 1 Then Exit Sub
        If .Rows = 0 Then Exit Sub
        
        If .Col <> .ColIndex("�����Ự") Or .TextMatrix(.Row, .ColIndex("�����Ự")) = "" Then
            If optCur.value Then
                frmHistSqlPlan.ShowMe .TextMatrix(.Row, .ColIndex("SQL_ID")), 1  '��ǰ����,intType=1
            Else
                frmHistSqlPlan.ShowMe .TextMatrix(.Row, .ColIndex("SQL_ID")), 2
            End If
        End If
    End With
End Sub
