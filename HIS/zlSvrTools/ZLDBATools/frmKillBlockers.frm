VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmKillBlockers 
   BorderStyle     =   0  'None
   Caption         =   "�����û��ȴ�������"
   ClientHeight    =   7995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16425
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   16425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pctBtm 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   0
      ScaleHeight     =   5055
      ScaleWidth      =   16425
      TabIndex        =   8
      Top             =   2940
      Width           =   16425
      Begin VSFlex8Ctl.VSFlexGrid vsfWaiters 
         Height          =   4815
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   12375
         _cx             =   21828
         _cy             =   8493
         Appearance      =   1
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   32768
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   380
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
         ExplorerBar     =   0
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
      End
      Begin VB.Label lblWaiters 
         Caption         =   "��������"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.PictureBox pctTop 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2895
      ScaleWidth      =   16425
      TabIndex        =   0
      Top             =   0
      Width           =   16425
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "ˢ��(&R)"
         Height          =   350
         Left            =   11400
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdKill 
         Caption         =   "������ǰ�����߻Ự"
         Height          =   350
         Index           =   0
         Left            =   3600
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtPreSQL 
         Height          =   2055
         Left            =   7800
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   4695
      End
      Begin VB.CommandButton cmdKill 
         Caption         =   "�������������߻Ự"
         Height          =   350
         Index           =   1
         Left            =   5640
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfBlockers 
         Height          =   2055
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   7575
         _cx             =   13361
         _cy             =   3625
         Appearance      =   1
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   32768
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   380
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
         ExplorerBar     =   0
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
      End
      Begin VB.Label lblSQL 
         Caption         =   "���һ�����е�SQL"
         Height          =   255
         Left            =   7800
         TabIndex        =   7
         Top             =   330
         Width           =   4695
      End
      Begin VB.Label lblBlocker 
         Caption         =   "������"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   330
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmKillBlockers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytOracleVer As Byte     'Oracle��汾,0-10.2���°汾��1-10.2�汾��2-11g�汾
Private mcolConn As New Collection  '���ʵ�������Ӷ��󼯺�
Private mstrOwner As String 'ZLHIS�������û�

Public Sub ShowMe()
    Me.Show
End Sub

Private Function GetConnection(ByVal lngKey As Long) As Connection
    Dim connTmp As ADODB.Connection
    
    On Error Resume Next
    Set connTmp = mcolConn("C_" & lngKey)
    
    Set GetConnection = connTmp
End Function

Private Sub cmdKill_Click(Index As Integer)
    Dim strTmp As String, strSql As String, strMethod As String
    Dim i As Long, lngRow As Long
    Dim lngB As Long, lngE As Long
    
    Dim lngCurr_INST_ID As Long, lngThis_INST_ID As Long
    Dim connInstance As ADODB.Connection
    
    If mbytOracleVer <> 2 Then
        lngCurr_INST_ID = CLng(gintInstId)
    End If
    If mbytOracleVer = 0 Then
        strMethod = "Kill"
    Else
        strMethod = "DISCONNECT"
    End If
   
    On Error Resume Next
    With vsfBlockers
        If Index = 0 Then
            lngB = .Row
            lngE = .Row
        Else
            lngB = .FixedRows
            lngE = .Rows - 1
        End If
    
        For i = lngB To lngE
            lngThis_INST_ID = Val(.TextMatrix(i, .ColIndex("INST_ID")))
            strTmp = .TextMatrix(i, .ColIndex("SID"))
            
            If strTmp <> "" Then
                '11g���²�֧��ɱ����ǰ���ӻỰ֮�������ʵ���ĻỰ
                If mbytOracleVer <> 2 And gblnRAC And lngCurr_INST_ID <> lngThis_INST_ID Then
                    Set connInstance = GetConnection(lngThis_INST_ID)
                    If connInstance Is Nothing Then
                        If frmUserCheckLogin.ShowLogin(connInstance, lngThis_INST_ID, gstrUserName, gstrPassword) Then
                            mcolConn.Add connInstance, "C_" & lngThis_INST_ID
                            '�����¼ʧ�ܣ���ô����Ự�����ᱻ���
                        End If
                    End If
                    
                    If Not connInstance Is Nothing Then
                        strSql = "alter system " & strMethod & " session '" & strTmp & "' IMMEDIATE"
                        connInstance.Execute strSql
                    End If
                Else
                    strSql = "alter system " & strMethod & " session '" & strTmp & IIf(gblnRAC And mbytOracleVer = 2, ",@" & lngThis_INST_ID, "") & "' IMMEDIATE"
                    gcnOracle.Execute strSql
                End If
            End If
        Next
        Call cmdRefresh_Click
        
        If Index = 0 Then
            lblSql.Caption = "�������߻Ự��" & strTmp & "�����������ɹ���"
        Else
            lblSql.Caption = "��ȫ�������߻Ự���������ɹ���"
        End If
    End With
       
    Exit Sub
errH:
    MsgBox Err.Description, vbExclamation
End Sub


Private Sub cmdRefresh_Click()
    Call LoadBlockers
    
    If Me.Visible And Me.Enabled Then vsfBlockers.SetFocus
End Sub


Private Function GetSysOwner() As String
    Dim rstmp As ADODB.Recordset, strSql As String
 
    strSql = "Select ������ From zltools.Zlsystems Where ��� = 100"
    
    On Error Resume Next    '����û�а�װZLHIS,Ҳ����ʹ��
    Set rstmp = OpenSQLRecord(strSql, Me.Caption)
    If rstmp.RecordCount > 0 Then
        GetSysOwner = rstmp!������
    End If
End Function

Private Sub Form_Activate()
    If vsfBlockers.Enabled Then Call vsfBlockers.SetFocus
End Sub

Private Sub Form_load()
    Dim strHead As String
    
    cmdKill(0).Enabled = False
    cmdKill(1).Enabled = False
    
   If Not gblnIsZlhis Then
        lblSql.Caption = "��ǰ���ݿ�û�а�װZLHIS��׼�棬���ֹ��ܲ��ܳ��ʹ�á�"
    End If
    
    mstrOwner = GetSysOwner

   If gblnIsZlhis Then
        strHead = "����;��Ա;SID,800,1;UserName,850,1;������,1800,1;������,1000,1;SQL;INST_ID" & IIf(gblnRAC, ",800,1", "")
    Else
        strHead = "����,1200,1;��Ա,1000,1;SID,800,1;UserName,850,1;������,1800,1;������,1000,1;SQL;INST_ID" & IIf(gblnRAC, ",800,1", "")
    End If
    Call InitTable(vsfBlockers, strHead)
    
   If gblnIsZlhis Then
        strHead = "����;��Ա;SID,800,1;UserName,850,1;������,2800,1;������,2800,1;SQL;������,1250,1;��ģʽ,1000,1;������ģʽ,1200,1;������,1000,1;INST_ID" & IIf(gblnRAC, ",800,1", "")
    Else
        strHead = "����,1200,1;��Ա,1000,1;SID,800,1;UserName,850,1;������,2800,1;������,2800,1;SQL;������,1250,1;��ģʽ,1000,1;������ģʽ,1200,1;������,1000,1;INST_ID" & IIf(gblnRAC, ",800,1", "")
    End If
    Call InitTable(vsfWaiters, strHead)
    
    If gblnRAC Then
        Call CheckAndCreateGLOBALView
    End If
    
    '10.2��ʼ֧��distinct session
    If gstrVerNum < "10.2.0.0.0" Then
        mbytOracleVer = 0
    '11��ʼ֧��ɱ������ʵ���ĻỰ
    ElseIf gstrVerNum > "11.0.0.0.0" Then
        mbytOracleVer = 2
    Else
        mbytOracleVer = 1
    End If
        
    
    Call cmdRefresh_Click
End Sub

Private Sub InitTable(vsfTmp As VSFlexGrid, strHead As String)
'���ܣ���ʼ�����
    Dim arrHead As Variant, i As Long
    
    arrHead = Split(strHead, ";")
    With vsfTmp
        .Clear
        .FixedRows = 1: .FixedCols = 0
        
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1  'ȱʡ��һ�հ���
        
        For i = 0 To UBound(arrHead)
            .ColKey(i) = Split(arrHead(i), ",")(0)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
    End With
End Sub

Private Sub Form_Resize()
    pctTop.Height = Me.ScaleHeight / 2
    pctBtm.Height = Abs(Me.ScaleHeight - pctTop.Height)
End Sub

Private Sub pctTop_Resize()
    vsfBlockers.Height = Abs(pctTop.Height - vsfBlockers.Top)
    vsfBlockers.Width = pctTop.Width / 2
    cmdKill(1).Left = Abs(vsfBlockers.Width + vsfBlockers.Left - cmdKill(1).Width)
    cmdKill(0).Left = Abs(cmdKill(1).Left - cmdKill(0).Width - 65)
    txtPreSQL.Height = vsfBlockers.Height
    txtPreSQL.Left = vsfBlockers.Width + vsfBlockers.Left + 120
    txtPreSQL.Width = Abs(pctTop.Width - txtPreSQL.Left - 120)
    lblSql.Left = txtPreSQL.Left
    
    cmdRefresh.Left = Abs(txtPreSQL.Left + txtPreSQL.Width - cmdRefresh.Width)
    cmdRefresh.Top = cmdKill(0).Top
End Sub

Private Sub pctbtm_Resize()
    vsfWaiters.Height = Abs(pctBtm.Height - vsfWaiters.Top)
    vsfWaiters.Width = Abs(pctBtm.Width - vsfWaiters.Left - 120)
End Sub



Private Sub vsfBlockers_RowColChange()
    Dim lngBlockerSID As Long, strSqlID As String
    Dim lngThis_INST_ID As Long
    
    With vsfBlockers
        If .Row < .FixedRows Or .Visible = False Then Exit Sub
  
        lngBlockerSID = Val(.TextMatrix(.Row, .ColIndex("SID")))
        
        If lngBlockerSID <> 0 Then
            Call LoadWaiters(lngBlockerSID)
            
            strSqlID = .TextMatrix(.Row, .ColIndex("SQL"))
            
            If strSqlID <> "" Then
                lngThis_INST_ID = Val(.TextMatrix(.Row, .ColIndex("INST_ID")))
                lblSql.Caption = "�����Ự��" & .TextMatrix(.Row, .ColIndex("SID")) & "�����һ�����е�SQL"
                Call LoadSQL(strSqlID, lngThis_INST_ID)
            Else
                Call ClearSQLText
            End If
        End If
    End With
End Sub

Private Sub vsfWaiters_RowColChange()
    Dim lngBlockerSID As Long, strSqlID As String
    Dim lngThis_INST_ID As Long
    
    With vsfWaiters
        If .Row < .FixedRows Then Exit Sub

        strSqlID = .TextMatrix(.Row, .ColIndex("SQL"))
        
        If strSqlID <> "" Then
            lngThis_INST_ID = Val(.TextMatrix(.Row, .ColIndex("INST_ID")))
            
            lblSql.Caption = "�������Ự��" & .TextMatrix(.Row, .ColIndex("SID")) & "�����һ�����е�SQL"
            Call LoadSQL(strSqlID, lngThis_INST_ID)
        Else
            Call ClearSQLText
        End If
    End With
End Sub

Private Sub LoadSQL(strSqlID As String, lngThis_INST_ID As Long)
    Dim rstmp As ADODB.Recordset, strSql As String
    Dim strTmp As String
    
    strSql = "Select Sql_Text From " & IIf(gblnRAC, "G", "") & "V$sqltext Where Sql_Id = [1] " & IIf(gblnRAC, "And INST_ID = " & lngThis_INST_ID, "") & "Order By Piece"
    On Error GoTo errH
    Set rstmp = OpenSQLRecord(strSql, Me.Caption, strSqlID)
        
    Do While Not rstmp.EOF
        strTmp = strTmp & IIf(strTmp = "", "", vbCrLf) & rstmp!Sql_Text
        rstmp.MoveNext
    Loop
    txtPreSQL.Text = strTmp
    
    Exit Sub
errH:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub LoadBlockers()
    Dim rstmp As ADODB.Recordset, strSql As String
    Dim lngRow As Long
    
   If Not gblnIsZlhis Then
        strSql = "Select '' As ����, '' as ����, a.Sid||','||a.Serial# as SID," & IIf(gblnRAC, "a.INST_ID,", "0 as INST_ID,") & "a.Username, a.Machine, a.Program, Nvl(a.SQL_ID,a.PREV_SQL_ID) as SQLID" & vbNewLine & _
                "From " & IIf(gblnRAC, "G", "") & "V$session A, Dba_" & IIf(gblnRAC, "GLOBAL_", "") & "Blockers B" & vbNewLine & _
                "Where a.Sid = b.Holding_Session And a.Username Is Not Null And a.Taddr Is Not Null Order by 1,2"
    Else
        strSql = "Select f.���� As ����, d.����, a.Sid||','||a.Serial# as SID," & IIf(gblnRAC, "a.INST_ID,", "0 as INST_ID,") & "a.Username, a.Machine, a.Program, Nvl(a.SQL_ID,a.PREV_SQL_ID) as SQLID" & vbNewLine & _
                "From " & IIf(gblnRAC, "G", "") & "V$session A, Dba_" & IIf(gblnRAC, "GLOBAL_", "") & "Blockers B, " & _
                mstrOwner & ".�ϻ���Ա�� C, " & mstrOwner & ".��Ա�� D, " & mstrOwner & ".������Ա E, " & mstrOwner & ".���ű� F" & vbNewLine & _
                "Where a.Sid = b.Holding_Session And a.Username Is Not Null And a.Taddr Is Not Null And a.Username = c.�û���(+) And c.��Աid = d.Id(+) And d.Id = e.��Աid(+) And e.ȱʡ(+) = 1 And" & vbNewLine & _
                "      e.����id = f.Id(+) Order by 1,2"
    End If
    
    On Error GoTo errH
    Set rstmp = OpenSQLRecord(strSql, Me.Caption)
    
    If rstmp.RecordCount = 0 Then
        Call ClearVsf(vsfBlockers, "")
        Call ClearVsf(vsfWaiters, "")
        txtPreSQL.Text = ""
        cmdKill(0).Enabled = rstmp.RecordCount > 0
        cmdKill(1).Enabled = rstmp.RecordCount > 1
        Exit Sub
    End If
    
    With vsfBlockers
        .Rows = .FixedRows
        If rstmp.RecordCount = 0 Then
            .Rows = .FixedRows
            vsfWaiters.Rows = vsfWaiters.FixedRows
            Call ClearSQLText
        Else
            .Rows = .FixedRows + rstmp.RecordCount
        End If
        lngRow = .FixedRows
        
        Do While Not rstmp.EOF
            
            .TextMatrix(lngRow, .ColIndex("����")) = "" & rstmp!����
            .TextMatrix(lngRow, .ColIndex("��Ա")) = "" & rstmp!����
            .TextMatrix(lngRow, .ColIndex("SID")) = "" & rstmp!Sid
            .TextMatrix(lngRow, .ColIndex("INST_ID")) = "" & rstmp!Inst_ID
            .TextMatrix(lngRow, .ColIndex("UserName")) = "" & rstmp!Username
            .TextMatrix(lngRow, .ColIndex("������")) = "" & rstmp!Machine
            .TextMatrix(lngRow, .ColIndex("������")) = "" & rstmp!Program
            .TextMatrix(lngRow, .ColIndex("SQL")) = "" & rstmp!SQLID
            lngRow = lngRow + 1
            rstmp.MoveNext
        Loop
        .Row = .FixedRows
    End With
    
    
    cmdKill(0).Enabled = rstmp.RecordCount > 0
    cmdKill(1).Enabled = rstmp.RecordCount > 1
    
    Exit Sub
errH:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub ClearSQLText()
    
    lblSql.Caption = "�Ự���һ�����е�SQL"
    txtPreSQL.Text = ""
End Sub

Private Sub LoadWaiters(lngBlockerSID As Long)
    Dim rstmp As ADODB.Recordset, strSql As String
    Dim lngRow As Long
    
   If Not gblnIsZlhis Then
        strSql = "Select '' As ����, '' As ����, a.Sid||','||a.Serial# as SID," & IIf(gblnRAC, "a.INST_ID,", "0 as INST_ID,") & "a.Username, a.Machine, a.Program, Nvl(a.SQL_ID,a.PREV_SQL_ID) as SQLID" & vbNewLine & _
                ",b.LOCK_TYPE,b.MODE_HELD,b.MODE_REQUESTED,g.object_name as ������" & vbNewLine & _
                "From " & IIf(gblnRAC, "G", "") & "V$session A, Dba_" & IIf(gblnRAC, "GLOBAL_", "") & "Waiters B, dba_objects g,gv$lock h" & vbNewLine & _
                "Where a.Sid = b.WAITING_SESSION And a.Username Is Not Null And a.Taddr Is Not Null And b.HOLDING_SESSION = [1] And h.id1  = g.Object_Id And h.sid =a.sid Order by 1,2"

    Else
        strSql = "Select f.���� As ����, d.����, a.Sid||','||a.Serial# as SID," & IIf(gblnRAC, "a.INST_ID,", "0 as INST_ID,") & "a.Username, a.Machine, a.Program, Nvl(a.SQL_ID,a.PREV_SQL_ID) as SQLID" & vbNewLine & _
                ",b.LOCK_TYPE,b.MODE_HELD,b.MODE_REQUESTED,g.object_name as ������" & vbNewLine & _
                "From " & IIf(gblnRAC, "G", "") & "V$session A, Dba_" & IIf(gblnRAC, "GLOBAL_", "") & "Waiters B, " & _
                mstrOwner & ".�ϻ���Ա�� C, " & mstrOwner & ".��Ա�� D, " & mstrOwner & ".������Ա E, " & mstrOwner & ".���ű� F,dba_objects g,gv$lock h" & vbNewLine & _
                "Where a.Sid = b.WAITING_SESSION And a.Username Is Not Null And a.Taddr Is Not Null And a.Username = c.�û���(+) And c.��Աid = d.Id(+) And d.Id = e.��Աid(+) And e.ȱʡ(+) = 1 And" & vbNewLine & _
                "      e.����id = f.Id(+) And b.HOLDING_SESSION = [1] And h.id1  = g.Object_Id And h.sid =a.sid Order by 1,2"
    End If
    
    On Error GoTo errH
    Set rstmp = OpenSQLRecord(strSql, Me.Caption, lngBlockerSID)

    With vsfWaiters
        .Rows = .FixedRows
        If rstmp.RecordCount = 0 Then
            .Rows = .FixedRows
        Else
            .Rows = .FixedRows + rstmp.RecordCount
        End If
        lngRow = .FixedRows
        
        Do While Not rstmp.EOF
            
            .TextMatrix(lngRow, .ColIndex("����")) = "" & rstmp!����
            .TextMatrix(lngRow, .ColIndex("��Ա")) = "" & rstmp!����
            .TextMatrix(lngRow, .ColIndex("SID")) = "" & rstmp!Sid
            .TextMatrix(lngRow, .ColIndex("INST_ID")) = "" & rstmp!Inst_ID
            .TextMatrix(lngRow, .ColIndex("UserName")) = "" & rstmp!Username
            .TextMatrix(lngRow, .ColIndex("������")) = "" & rstmp!Machine
            .TextMatrix(lngRow, .ColIndex("������")) = "" & rstmp!Program
            .TextMatrix(lngRow, .ColIndex("SQL")) = "" & rstmp!SQLID
            
            .TextMatrix(lngRow, .ColIndex("������")) = "" & rstmp!LOCK_TYPE
            .TextMatrix(lngRow, .ColIndex("��ģʽ")) = "" & rstmp!MODE_HELD
            .TextMatrix(lngRow, .ColIndex("������ģʽ")) = "" & rstmp!MODE_REQUESTED
            .TextMatrix(lngRow, .ColIndex("������")) = "" & rstmp!������
            
            lngRow = lngRow + 1
            rstmp.MoveNext
        Loop
    End With
    
    Exit Sub
errH:
    MsgBox Err.Description, vbExclamation
End Sub



Private Sub vsfWaiters_GotFocus()
    vsfWaiters.BackColorSel = &H8000000D
    vsfWaiters.ForeColorSel = vbWhite
    Call vsfWaiters_RowColChange
End Sub

Private Sub vsfWaiters_LostFocus()
    vsfWaiters.BackColorSel = &HFAEADA
    vsfWaiters.ForeColorSel = vbBlack
End Sub

Private Sub vsfBlockers_GotFocus()
    vsfBlockers.BackColorSel = &H8000000D
    vsfBlockers.ForeColorSel = vbWhite
    Call vsfBlockers_RowColChange
End Sub

Private Sub vsfBlockers_LostFocus()
    vsfBlockers.BackColorSel = &HFAEADA
    vsfBlockers.ForeColorSel = vbBlack
End Sub

Private Sub CheckAndCreateGLOBALView()
'���ܣ����RAC�����µ�������ͼ������Ȩ�޵Ĵ���
    Dim rstmp As ADODB.Recordset, strSql As String
 
    strSql = "select 1 from DBA_GLOBAL_BLOCKERS"
    Err.Clear
    On Error Resume Next
    Set rstmp = OpenSQLRecord(strSql, Me.Caption)
    If Err.Number <> 0 Then
        Err.Clear
        strSql = "create or replace view DBA_GLOBAL_BLOCKERS" & vbNewLine & _
                "as Select DISTINCT h.Sid Holding_Session" & vbNewLine & _
                "From Gv$lock W, Gv$lock H" & vbNewLine & _
                "Where h.Block != 0 And h.Lmode != 0 And h.Lmode != 1 And w.Request != 0 And w.Type = h.Type And w.Id1 = h.Id1 And" & vbNewLine & _
                "      w.Id2 = h.Id2 And w.Addr != h.Addr"
        gcnOracle.Execute strSql
    
        strSql = "create or replace public synonym DBA_GLOBAL_BLOCKERS for DBA_GLOBAL_BLOCKERS"
        gcnOracle.Execute strSql
        strSql = "grant select on DBA_GLOBAL_BLOCKERS to select_catalog_role"
        gcnOracle.Execute strSql
        strSql = "grant select on DBA_GLOBAL_BLOCKERS to dba"
        gcnOracle.Execute strSql
    End If
    
    strSql = "select 1 from DBA_GLOBAL_WAITERS"
    On Error Resume Next
    Set rstmp = OpenSQLRecord(strSql, Me.Caption)
    If Err.Number <> 0 Then
        Err.Clear
        strSql = "create or replace view DBA_GLOBAL_WAITERS" & vbNewLine & _
                "as select w.sid waiting_session,h.sid holding_session," & vbNewLine & _
                "        decode(w.type," & vbNewLine & _
                "                'MR', 'Media Recovery','RT', 'Redo Thread', 'UN', 'User Name','TX', 'Transaction', 'TM', 'DML'," & vbNewLine & _
                "                'UL', 'PL/SQL User Lock','DX', 'Distributed Xaction','CF', 'Control File','IS', 'Instance State'," & vbNewLine & _
                "                'FS', 'File Set','IR', 'Instance Recovery','ST', 'Disk Space Transaction','TS', 'Temp Segment'," & vbNewLine & _
                "                'IV', 'Library Cache Invalidation','LS', 'Log Start or Switch','RW', 'Row Wait','SQ', 'Sequence Number'," & vbNewLine & _
                "                'TE', 'Extend Table','TT', 'Temp Table',w.type) lock_type," & vbNewLine & _
                "        decode(h.lmode," & vbNewLine & _
                "                0, 'None', 1, 'Null',2, 'Row-S (SS)',3, 'Row-X (SX)'," & vbNewLine & _
                "                4, 'Share',5, 'S/Row-X (SSX)',6, 'Exclusive',to_char(h.lmode)) mode_held," & vbNewLine & _
                "        decode(w.request," & vbNewLine & _
                "                0, 'None',1, 'Null',2, 'Row-S (SS)',3, 'Row-X (SX)',4, 'Share',5, 'S/Row-X (SSX)'," & vbNewLine & _
                "                6, 'Exclusive',to_char(w.request)) mode_requested, to_char(w.id1) lock_id1, to_char(w.id2) lock_id2" & vbNewLine & _
                "  from gv$lock w, gv$lock h" & vbNewLine & _
                " where h.block      !=  0 and  h.lmode      !=  0 and  h.lmode      !=  1" & vbNewLine & _
                "  and  w.request    !=  0 and  w.type       =  h.type and  w.id1        =  h.id1" & vbNewLine & _
                "  and  w.id2        =  h.id2  and  w.addr       != h.addr"

        gcnOracle.Execute strSql
        
        strSql = "create or replace public synonym DBA_GLOBAL_WAITERS for DBA_GLOBAL_WAITERS"
        gcnOracle.Execute strSql
        strSql = "grant select on DBA_GLOBAL_WAITERS to select_catalog_role"
        gcnOracle.Execute strSql
        strSql = "grant select on DBA_GLOBAL_WAITERS to dba"
        gcnOracle.Execute strSql
    End If
    
End Sub

