VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form Form1 
   Caption         =   "ʹ��clsQueueOperation����������"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12900
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   12900
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   2955
      Left            =   75
      TabIndex        =   11
      Top             =   135
      Width           =   12780
      _cx             =   22542
      _cy             =   5212
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin VB.CommandButton Command10 
      Caption         =   "2.��Ӳ��Ŷ�    "
      Height          =   465
      Left            =   1901
      TabIndex        =   10
      Top             =   3165
      Width           =   1620
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9.��������      "
      Height          =   510
      Left            =   5475
      TabIndex        =   9
      Top             =   3810
      Width           =   1620
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8.�㲥          "
      Height          =   510
      Left            =   3690
      TabIndex        =   8
      Top             =   3810
      Width           =   1620
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7.��ɺ���      "
      Height          =   510
      Left            =   1890
      TabIndex        =   7
      Top             =   3810
      Width           =   1620
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6.�����Ŷ�      "
      Height          =   510
      Left            =   105
      TabIndex        =   6
      Top             =   3810
      Width           =   1620
   End
   Begin VB.CommandButton Command5 
      Caption         =   "10.����������� "
      Height          =   510
      Left            =   7290
      TabIndex        =   5
      Top             =   3810
      Width           =   1620
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   480
      Top             =   4620
   End
   Begin VB.CommandButton Command4 
      Caption         =   "5.�����ǰҵ���µ��Ŷ�����"
      Height          =   465
      Left            =   7290
      TabIndex        =   3
      Top             =   3165
      Width           =   1620
   End
   Begin VB.CommandButton Command3 
      Caption         =   "4.�����������  "
      Height          =   465
      Left            =   5493
      TabIndex        =   2
      Top             =   3165
      Width           =   1620
   End
   Begin VB.CommandButton Command2 
      Caption         =   "1.���          "
      Height          =   465
      Left            =   105
      TabIndex        =   1
      Top             =   3165
      Width           =   1620
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3.��Ӻ�ֱ�Ӻ���"
      Height          =   465
      Left            =   3697
      TabIndex        =   0
      Top             =   3165
      Width           =   1620
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   345
      Left            =   2775
      TabIndex        =   4
      Top             =   5070
      Width           =   4485
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mlngQueueId As Long = 962

Private objQueue As clsQueueOperation
Attribute objQueue.VB_VarHelpID = -1


Private Sub RefreshQueueData()
    Dim rsData As ADODB.Recordset
    Dim strSql As String
    Dim i As Long
    
    strSql = "select * from �ŶӽкŶ��� order by " & objQueue.GetCustomOrderWhere
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯ�����Ŷӽк�����")
    
    Call VSFlexGrid1.Clear
    VSFlexGrid1.Rows = 0
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    VSFlexGrid1.Cols = rsData.Fields.Count
    VSFlexGrid1.Rows = rsData.RecordCount + 1
    
    VSFlexGrid1.FixedRows = 1
    
    VSFlexGrid1.AutoResize = True
    
    '��ʼ����ͷ
    For i = 0 To rsData.Fields.Count - 1
        VSFlexGrid1.Cell(flexcpText, 0, i) = rsData.Fields(i).Name
        VSFlexGrid1.ColWidth(i) = 1400
    Next i
    
    '��������
    While Not rsData.EOF
        For i = 0 To rsData.Fields.Count - 1
            VSFlexGrid1.Cell(flexcpText, rsData.AbsolutePosition, i) = rsData.Fields(i).Value
        Next i
        
        VSFlexGrid1.Cell(flexcpData, rsData.AbsolutePosition, 0) = Val(rsData!ID)
        
        Call rsData.MoveNext
    Wend
End Sub

Private Sub Command1_Click()
'����QUEUE1���в�����
    Dim lngQueueId As Long
    Dim lngVoiceId As Long
    Dim strNewNo As String
     
    lngQueueId = objQueue.InsertQueue("QUEUE1", , , "QUEUE1" & Format(Now, "hh:mm:ss"))
    
    '��ʼ�Ŷ�
    Call objQueue.LineQueue(lngQueueId, strNewNo)
    
    '˳������
    lngVoiceId = objQueue.SpecifiedCall(lngQueueId)
    
    '��������
    Call objQueue.PlayQueueVoice(lngVoiceId, lngQueueId, False)
    
    Call RefreshQueueData
End Sub

Private Sub Command10_Click()
'����QUEUE2����
    Dim lngQueueId As Long
    Dim strNewNo As String
     
    lngQueueId = objQueue.InsertQueue("QUEUE2", , , "QUEUE2" & Format(Now, "hh:mm:ss"))
    
    '��ʼ�Ŷ�
    Call objQueue.LineQueue(lngQueueId, strNewNo)
    
    Call RefreshQueueData
End Sub

Private Sub Command2_Click()
'����QUEUE2����
    Dim lngQueueId As Long
     
    lngQueueId = objQueue.InsertQueue("QUEUE2", , , "QUEUE3" & Format(Now, "hh:mm:ss"))
    
    Call RefreshQueueData
End Sub

Private Sub Command3_Click()
    Call objQueue.ClearQueueData("QUEUE2", True)
    
    Call RefreshQueueData
End Sub

Private Sub Command4_Click()
    Call objQueue.ClearBusinessData(True)
    
    Call RefreshQueueData
End Sub

Private Sub Command5_Click()
    Call objQueue.ClearVoiceData(, True)
End Sub

Private Sub Command6_Click()
'�����ŶӲ�����
    Dim lngQueueId As Long
    
    If VSFlexGrid1.RowSel <= 0 Then
        MsgBox "��ѡ����Ҫ���ŵ����ݡ�"
        Exit Sub
    End If
    
    
    lngQueueId = Val(VSFlexGrid1.Cell(flexcpData, VSFlexGrid1.RowSel, 0))
    If lngQueueId <= 0 Then Exit Sub
    
    Call objQueue.RestoreQueue(lngQueueId)
    
    Call RefreshQueueData
End Sub

Private Sub Command7_Click()
    '���
    Dim lngQueueId As Long
    
    If VSFlexGrid1.RowSel <= 0 Then
        MsgBox "��ѡ����Ҫ��ɵ����ݡ�"
        Exit Sub
    End If
    
    
    lngQueueId = Val(VSFlexGrid1.Cell(flexcpData, VSFlexGrid1.RowSel, 0))
    If lngQueueId <= 0 Then Exit Sub
    
    Call objQueue.CompleteQueue(lngQueueId)
    
    Call RefreshQueueData
End Sub

Private Sub Command8_Click()
    '�㲥
    Dim lngQueueId As Long
    
    If VSFlexGrid1.RowSel <= 0 Then
        MsgBox "��ѡ����Ҫ�㲥�����ݡ�"
        Exit Sub
    End If
    
    
    lngQueueId = Val(VSFlexGrid1.Cell(flexcpData, VSFlexGrid1.RowSel, 0))
    If lngQueueId <= 0 Then Exit Sub
    
    Call objQueue.BroadcastCall(lngQueueId)
    
    Call RefreshQueueData
End Sub

Private Sub Command9_Click()
    Dim lngQueueId As Long
    
    If VSFlexGrid1.RowSel <= 0 Then
        MsgBox "��ѡ����Ҫ���µ����ݡ�"
        Exit Sub
    End If
    
    
    lngQueueId = Val(VSFlexGrid1.Cell(flexcpData, VSFlexGrid1.RowSel, 0))
    If lngQueueId <= 0 Then Exit Sub
    
    Call objQueue.UpdateQueue(lngQueueId, "��������='TEST_" & Format(Now, "hh:mm:ss") & "'")
    
    Call RefreshQueueData
End Sub

Private Sub Form_Load()
    '��ʼ��clsQueueOperation����
    Call InitDebugObject(1160, Me, "zlhis", "HIS")
    
    Set objQueue = New clsQueueOperation
    Call objQueue.InitQueue(gcnOracle, 1)
    
    objQueue.VoiceType = "Girl XiaoKun"
    objQueue.PlaySpeed = 0
    
    Call RefreshQueueData
End Sub




Public Sub InitDebugObject(ByVal lngModuleNum As Long, ByVal frmMain As Object, ByVal strUser As String, ByVal strPwd As String)
'��ʼ������״̬�µ��������
    Set gcnOracle = New ADODB.Connection
    
    Call OraDataOpen("", strUser, strPwd)
    
    glngSys = 100
'    gstrPrivs = ";PACS�����ӡ;PACS����ɾ��;PACS������д;PACS�������Ʊ���;PACS�����޶�;PACS���˱���;�ɼ���������;��������;�洢����;��������;����;��鱨��;���Ǽ�;������;��ɫͨ��;�Ŷӽк�;���ͼ��;ȡ������;ȡ��������;ɾ����ʱӰ��;��Ƶ�ɼ�;���;���п���;ͼ�����;δ�ɷѱ���;�ļ�����;�ޱ������;Ӱ���ʿ�;������������;Excel���;"
'    glngModul = lngModuleNum
    
    UserInfo.ID = 281
    UserInfo.���� = "������"
    UserInfo.�û��� = "ZLHIS"
    UserInfo.��� = "1123"
    UserInfo.���� = "WGY"
    UserInfo.����ID = "65"
    
    
    Call InitCommon(gcnOracle)
    
'    Call RegCheck
        
'    Call gobjKernel.InitCISKernel(gcnOracle, frmMain, glngSys, gstrPrivs) '��ʼ��ҽ�����������Ĳ���
'    Call gobjRichEPR.InitRichEPR(gcnOracle, frmMain, glngSys, False)
End Sub


Private Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strSql As String
    Dim strError As String
    
    On Error Resume Next
    Err = 0
    DoEvents
    With gcnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If Err <> 0 Then
            '���������Ϣ
            strError = Err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE�����ã������������ݿ�ʵ���Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "�û�" & UCase(strUserName) & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ�����������¼��)��", vbExclamation, gstrSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "�����û�������������ָ�������޷���¼��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "�����û��Ѿ������ã��޷���¼��", vbInformation, gstrSysName
            Else
                MsgBox strError, vbInformation, gstrSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo Errhand
    
    gstrDBUser = UCase(strUserName)
    SetDbUser gstrDBUser
    
    OraDataOpen = True
    Exit Function
    
Errhand:
    If ErrCenter() = 1 Then Resume
    OraDataOpen = False
    Err = 0
End Function

Private Sub Form_Resize()
On Error Resume Next
    VSFlexGrid1.Width = Me.ScaleWidth - 80
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call objQueue.StopVoice
    Set objQueue = Nothing
End Sub


Private Sub Timer2_Timer()
    If objQueue Is Nothing Then Exit Sub
    
    Label2.Caption = "�Ŷ�����:" & objQueue.GetStateCount("QUEUE1", qsQueueing) & _
                    "    �Ѻ�������:" & objQueue.GetStateCount("QUEUE1", qsCalled) & _
                    "    ���������:" & objQueue.GetStateCount("QUEUE1", qsComplete)
End Sub
