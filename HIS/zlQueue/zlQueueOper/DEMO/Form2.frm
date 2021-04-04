VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{1B83D023-3CA6-4181-A286-20352E645AE2}#2.2#0"; "zlQueueOper.ocx"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13155
   LinkTopic       =   "Form2"
   ScaleHeight     =   6930
   ScaleWidth      =   13155
   StartUpPosition =   3  '����ȱʡ
   Begin zlQueueOper.UcQueue UcQueueStation1 
      Height          =   4770
      Left            =   90
      TabIndex        =   8
      Top             =   210
      Width           =   12690
      _ExtentX        =   22384
      _ExtentY        =   8414
      Interval        =   30000
      ValidDays       =   0
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   105
      ScaleHeight     =   1590
      ScaleWidth      =   12675
      TabIndex        =   0
      Top             =   5160
      Width           =   12735
      Begin VB.CommandButton Command2 
         Caption         =   "���ò�������"
         Height          =   390
         Left            =   8190
         TabIndex        =   9
         Top             =   405
         Width           =   1500
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ˢ������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6765
         TabIndex        =   7
         Top             =   420
         Width           =   1200
      End
      Begin VB.CommandButton Command1 
         Caption         =   "վ������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5430
         TabIndex        =   6
         Top             =   390
         Width           =   1005
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3945
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   405
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1650
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   405
         Width           =   1245
      End
      Begin VB.Label Label3 
         Caption         =   "����վ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2955
         TabIndex        =   4
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label Label2 
         Caption         =   "����վ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   750
         TabIndex        =   3
         Top             =   480
         Width           =   900
      End
      Begin XtremeCommandBars.CommandBars cbrMain 
         Left            =   90
         Top             =   15
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   60
         TabIndex        =   1
         Top             =   1110
         Width           =   12570
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call UcQueueStation1.zlExecuteCommandBars(Control)
End Sub

Private Sub Command1_Click()
    UcQueueStation1.QueueOper.LocalStation = Text1.Text
    UcQueueStation1.QueueOper.PlayStation = Text2.Text
End Sub

Private Sub Command2_Click()
    Call ConfigDefaultQueueData
End Sub

Private Sub Command3_Click()
    Call UcQueueStation1.RefreshQueueData
End Sub

Private Sub Form_Load()
    Dim cbrMenuBar As CommandBarPopup
    
    Call OraDataOpen("002133-1033ORCLMSG", "ZLHIS", TranPasswd("aqa"))
    
    
    Call InitCommon(gcnOracle)
    Call SetDbUser("ZLHIS")
    
    Call ConnectMip(Me.hWnd)
    
    '������Ϣ����
    Call UcQueueStation1.UseMsgCenter(100, 1290)
    
    UcQueueStation1.QueryQueueNames = "���Զ���,QUEUE1"           '���δ���ô����ԣ�����ʾ��ҵ�������µ����ж�������
    UcQueueStation1.ReportNum = "Test"
    
    UcQueueStation1.DataFields = "ID,��������,�Ŷ����,�ŶӺ���,�Ŷ�״̬,��������,����,ҽ������,�Ŷ�ʱ��,��ע,����1,����ҽ��,����ʱ��,����,����1"
    UcQueueStation1.DisplayQueueFields = "�ŶӺ���,��������,����,ҽ������,�Ŷ�ʱ��,��ע,����1,����ʱ��,�Ŷ�״̬"
    UcQueueStation1.DisplayCallFields = "�ŶӺ���,��������,����ʱ��"    '���û�����ô����ԣ�����ʾ�����ֶ�
        
'    UcQueueStation1.CustomOrderField = "�Ŷ����" ' "��������"  '���ö��е������ֶΣ����δ���ã���ʹ�����ݿ��Ĭ������ʽ
    
    UcQueueStation1.GroupField = "��������"                      '�����Ŷӽкŵķ��鷽ʽ
    
'    UcQueueStation1.IsShowBars = True                           '�����Ƿ���ʾ�ŶӽкŵĲ�������������������ã���Ĭ����ʾ������
    
'    UcQueueStation1.IsShowCalledQueue = True                    '�����Ƿ���ʾ�Ѻ��ж���,��������ã���Ĭ����ʾ������
    
    UcQueueStation1.FindWayEx = "�����,סԺ��,�����,�籣��,����"
    
    Call UcQueueStation1.InitQueue(gcnOracle, 1, Me, App.ProductName, "ZLHIS", ",���,˳��,ֱ��,�㲥,����,���,����,����,��ͣ,����,�ָ�,���,ˢ��,����,��λ,����,�޸�,����,")
    
    UcQueueStation1.QueueOper.VoiceType = 1
    
    
    Call UcQueueStation1.ApplyVoiceConfig
    
    Call UcQueueStation1.StartVoice
    
    Set UcQueueStation1.Font = Me.Font
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "�Ŷ�")
    
    Call UcQueueStation1.zlCreateMenuBars(cbrMenuBar, True)
End Sub

Private Sub ConfigDefaultQueueData()
    Dim objQueue As clsQueueOperation
    Dim i As Long
    Dim lngQueueId As Long
    Dim strNewQueueNo As String
    
    Set objQueue = UcQueueStation1.QueueOper

    For i = 1 To 10
        lngQueueId = objQueue.InsertQueue("���Զ���", , , "��" & i & Format(Now, "hh:mm:ss") & Rnd)
        Call objQueue.LineQueue(lngQueueId, strNewQueueNo)
    Next i

    Call UcQueueStation1.RefreshQueueData
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    UcQueueStation1.Left = 0
    UcQueueStation1.Top = 0
    UcQueueStation1.Width = Me.ScaleWidth
    UcQueueStation1.Height = Me.ScaleHeight - Picture1.Height
    
    Picture1.Top = UcQueueStation1.Height
    Picture1.Left = 0
    Picture1.Width = Me.ScaleWidth
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call UcQueueStation1.StopVoice
    
    Call DisConnectMip
End Sub

Private Sub UcQueueStation1_OnCallPreBefore(ByVal lngQueueId As Long, ByVal lngCallWay As TCallWay, strCallContext As String, blnCancel As Boolean)
'    MsgBox "���в�����" & lngQueueId
End Sub

Private Sub UcQueueStation1_OnCmdBarExecute(objControl As Object, ByRef blnUseCustom As Boolean)
    Dim strName As String
    Dim lngRowIndex As Long
    Dim ID As Long
    
    Select Case objControl.ID
        Case 7890
            blnUseCustom = True
            
            lngRowIndex = UcQueueStation1.GetCalledQueueIndex()
            
            If lngRowIndex >= 0 Then
                strName = UcQueueStation1.GetListValue(qftCalledQueue, lngRowIndex, "��������")
            End If
            
            '��ȡ����ID�󣬺�������ﻼ��
            ID = UcQueueStation1.GetListValue(qftCalledQueue, lngRowIndex, "ID")
            Call UcQueueStation1.QueueOper.WaitRoomCall(ID)
            
            MsgBox "���԰�ť������  ������" & strName
    End Select
End Sub

Private Sub UcQueueStation1_OnCmdBarInit(CmdBar As Object)
    CmdBar.Controls.Add(1, 7890, "����", 7).IconId = 721
End Sub


Private Sub UcQueueStation1_OnCustomFindButton(ByVal lngQueueId As Long)
    MsgBox "�Զ������ִ�У�" & lngQueueId
End Sub

Private Sub UcQueueStation1_OnCmdBarUpdate(objControl As Object)
    If objControl.ID = 7890 Then
        objControl.Enabled = UcQueueStation1.CurQueueType = qftCalledQueue
    End If
End Sub

Private Sub UcQueueStation1_OnColumnInit(objQueueList As Object, objReportColumn As Object)
    If objReportColumn.Caption = "�Ŷ�ʱ��" Then
'        objReportColumn.Width = 200
        objReportColumn.Icon = 721
    End If
End Sub

Private Sub UcQueueStation1_OnFindData(ByVal strFindWay As String, ByVal strFindValue As String, txtFind As Object, rsData As ADODB.Recordset, blnUseCustom As Boolean)
    Dim strSql As String
    
    Select Case strFindWay
        Case "�����"
            strSql = "select * from �ŶӽкŶ���"
        Case "סԺ��"
            strSql = "select * from �ŶӽкŶ���"
        Case Else
            Exit Sub
    End Select
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "�����Ŷ�����", strFindValue)
    
    blnUseCustom = True
End Sub

Private Sub UcQueueStation1_OnItemDblClick(ByVal lngListType As zlQueueOper.TQueueFromType, ByVal lngQueueId As Long, objReoprtRow As Object, objReportItem As Object)
    MsgBox UcQueueStation1.GetListValue(lngListType, objReoprtRow.Index, "��������")
End Sub



Private Sub UcQueueStation1_OnPlayVoiceAfter(ByVal lngCallId As Long, ByVal lngQueueId As Long, ByVal strCallContext As String)
    Label1.Caption = "��ɺ��У���������Ϊ[" & strCallContext & "]"
End Sub



Private Sub UcQueueStation1_OnQueryQueueData(rsData As ADODB.Recordset, blnUseCustom As Boolean)
    Dim strSql As String
    '�ŶӺ���,��������,����,����,ҽ������,�Ŷ�״̬,�Ŷ�ʱ��,��ע
    
    strSql = "select " & UcQueueStation1.GetValidCols("ID,�ŶӺ���,��������,����,����,ҽ������,�Ŷ�״̬,�Ŷ�ʱ��,��ע, �Ŷ����, '����1' as ����1,'����2' as ����2, ����ʱ��", "") & _
            " from �ŶӽкŶ��� where ҵ������=" & UcQueueStation1.WorkType & _
          " and �Ŷ�ʱ�� between sysdate - 1 and sysdate and �������� in ('" & Replace(UcQueueStation1.QueryQueueNames, ",", "','") & "') order by �Ŷ���� "
            
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ȡ�Ŷ�����")
    
    blnUseCustom = True
    
        
End Sub




Private Sub UcQueueStation1_OnReadAfter(rsData As ADODB.Recordset, ByVal lngListType As TQueueFromType, objReportRecord As Object)
    If InStr(rsData!�Ŷ����, ".") > 0 Then
        objReportRecord(0).Icon = 3560
        
    End If

    Label1.Caption = "OnReadAfter�¼�ִ�У���������Ϊ" & lngListType
End Sub

Private Sub UcQueueStation1_OnSelectionChanged(ByVal lngListType As zlQueueOper.TQueueFromType, ByVal lngQueueId As Long, objQueueList As Object, objReportRow As Object)
    If objReportRow.GroupRow = True Then Exit Sub
    
    Label1.Caption = "OnSelectionChanged�¼�ִ�У�" & objReportRow.Record(1).Value
End Sub

Private Sub UcQueueStation1_OnWorkBefore(ByVal lngQueueId As Long, ByVal lngOperationType As TOperationType, blnCancel As Boolean)
    Select Case lngOperationType
        Case TOperationType.otDiagnose
            Label1.Caption = "OnWorkBefore�¼�ִ�У���������Ϊ����"
            
            Call UcQueueStation1.QueueOper.AbstainQueue(lngQueueId)
        Case TOperationType.otAbstain
            Label1.Caption = "OnWorkBefore�¼�ִ�У���������Ϊ����"
        Case TOperationType.otPause
            Label1.Caption = "OnWorkBefore�¼�ִ�У���������Ϊ��ͣ"
        Case TOperationType.otComplete
            Label1.Caption = "OnWorkBefore�¼�ִ�У���������Ϊ���"
        Case TOperationType.otRestore
            Label1.Caption = "OnWorkBefore�¼�ִ�У���������Ϊ����"
        Case TOperationType.otStart
            Label1.Caption = "OnWorkBefore�¼�ִ�У���������Ϊ�ָ�"
        Case TOperationType.otPrintNo
            MsgBox "ִ�д�Ŵ���ID" & lngQueueId
        Case Else
            Label1.Caption = "OnWorkBefore�¼�ִ�У���������Ϊ" & lngOperationType
    End Select
End Sub
