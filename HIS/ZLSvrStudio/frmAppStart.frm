VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppStart 
   BackColor       =   &H80000005&
   Caption         =   "ϵͳװж����"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9090
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmAppStart.frx":0000
   ScaleHeight     =   5385
   ScaleWidth      =   9090
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdReCalc 
      Caption         =   "���½���ֵ(&R)"
      Height          =   350
      Left            =   3885
      TabIndex        =   11
      Top             =   3660
      Width           =   1500
   End
   Begin VB.TextBox txtMem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   2100
      TabIndex        =   9
      Top             =   3465
      Width           =   630
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "����(&C)��"
      Height          =   350
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Top             =   2325
      Width           =   1275
   End
   Begin MSComctlLib.ImageList imgSys 
      Left            =   4170
      Top             =   1140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppStart.frx":04F9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwSys 
      Height          =   1380
      Left            =   960
      TabIndex        =   2
      Top             =   870
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   2434
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "Img��ͼ��"
      SmallIcons      =   "imgSys"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ϵͳ����"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "�汾��"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "���"
         Text            =   "���"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "������"
         Text            =   "������"
         Object.Width           =   1411
      EndProperty
   End
   Begin MSComctlLib.ListView lvwPara 
      Height          =   1230
      Left            =   960
      TabIndex        =   4
      Top             =   4020
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   2170
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "Img��ͼ��"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "��ǰֵ"
         Object.Width           =   1942
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "˵��"
         Object.Width           =   16581
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "��������"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ת��ΪM"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "��ж(&M)��"
      Height          =   350
      Index           =   1
      Left            =   960
      TabIndex        =   7
      Top             =   2670
      Width           =   1275
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "��ֲ(&R)��"
      Height          =   350
      Index           =   2
      Left            =   960
      TabIndex        =   8
      Top             =   3015
      Width           =   1275
   End
   Begin VB.Label lblMem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������ڴ�(M)       (���������Ƿ�����ʱ,���޸�Ϊ�������������ڴ��С���Ը���׼ȷ�Ľ���ֵ��)"
      Height          =   180
      Left            =   930
      TabIndex        =   10
      Top             =   3510
      Width           =   8190
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��ȡ�úϷ���Ӧ��ϵͳ�����ļ������Ч��Ȩ֮�󣬿��Դ����µ�ϵͳ��"
      ForeColor       =   &H80000008&
      Height          =   1020
      Left            =   2310
      TabIndex        =   6
      Top             =   2340
      Width           =   2700
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblPara 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����Ҫ���������ݿ����"
      Height          =   180
      Left            =   945
      TabIndex        =   3
      Top             =   3795
      Width           =   2340
   End
   Begin VB.Label lblSys 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Ѱ�װӦ��ϵͳ"
      Height          =   180
      Left            =   960
      TabIndex        =   1
      Top             =   675
      Width           =   1260
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ϵͳװж����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   0
      Top             =   105
      Width           =   1440
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   240
      Picture         =   "frmAppStart.frx":114B
      Top             =   645
      Width           =   480
   End
End
Attribute VB_Name = "frmAppStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTemp As New ADODB.Recordset
Dim strSql As String
Dim objItem As ListItem
Dim intCount As Integer

Private mintVersion As Integer

Private Sub cmdFunction_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    For intCount = 0 To cmdFunction.UBound
        If intCount = Index Then
            cmdFunction(intCount).FontBold = True
            cmdFunction(intCount).SetFocus
            Select Case intCount
            Case 0
                lblNote.Caption = "    ��ȡ�úϷ���Ӧ��ϵͳ�����ļ������Ч��Ȩ֮�󣬿��Դ����µ�ϵͳ��"
            Case 1
                lblNote.Caption = "    �Բ����õ�ϵͳ���ɸ��ݰ�װ�ļ����в�ж���Խ���ϵͳ�ĸ��ɡ�"
            Case 2
                lblNote.Caption = "    ���ȷ���Ѿ���������ʽ(���ֹ�ִ��Import)��װ��Ӧ��ϵͳ��Ӧ�ýṹ�����ݣ�����ͨ��������ֲ��ϵͳ�������ݡ�"
            End Select
        Else
            cmdFunction(intCount).FontBold = False
        End If
    Next
End Sub

Private Sub cmdFunction_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call cmdFunction_MouseMove(Index, 0, 0, 0, 0)
    Select Case Index
    Case 0
        frmAppCreate.Show 1, frmMDIMain
        Call SysCreated
    Case 1
        Dim strLinkSys As String
        If lvwSys.SelectedItem Is Nothing Then Exit Sub
        If lvwSys.SelectedItem.Selected = False Then Exit Sub
        strLinkSys = ""
        
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Loadandunload.Get_Share_name", Mid(lvwSys.SelectedItem.Key, 2))
        
        With rsTemp
            Do While Not .EOF
                strLinkSys = strLinkSys & vbCrLf & .Fields(0).value
                .MoveNext
            Loop
        End With
        If strLinkSys <> "" Then
            MsgBox "���ڵ�ǰϵͳ������ϵͳ��������ֱ�Ӳ�ж��" & strLinkSys, vbExclamation, gstrSysName
            Exit Sub
        End If
        frmAppRemove.Show 1, frmMDIMain
        Call SysCreated
    Case 2
        frmAppReplant.Show 1, frmMDIMain
        Call SysCreated
    End Select

End Sub

Private Sub cmdReCalc_Click()
    If IsNumeric(txtMem) Then
        If Val(txtMem) < 256 Or Val(txtMem) > 10000 Then
            MsgBox "�������ڴ�Ӧ��256��10000֮��!", vbInformation, gstrSysName
        Else
            Call SysPara
        End If
    Else
        MsgBox "�������ڴ�ӦΪ��������!", vbInformation, gstrSysName
    End If
End Sub

Private Sub Form_Activate()
    If Not gblnDBA Then
        frmMDIMain.stbThis.Panels(2).Text = "���谲װ����жӦ��ϵͳ����ʹ�û������Ȩ�޵�DBA�û�����ע�����"
    End If
End Sub

Private Sub Form_Deactivate()
    frmMDIMain.stbThis.Panels(2).Text = ""
End Sub

Private Sub Form_Load()

    If Not gblnDBA Then
        For intCount = 0 To cmdFunction.UBound
            cmdFunction(intCount).Enabled = False
        Next
    End If
    
    mintVersion = GetOracleVersion
    '�����ڴ��С
    Dim mem As MEMORYSTATUS
    GlobalMemoryStatus mem
    'MsgBox "physical   Memory   is:" & mem.dwTotalPhys
    txtMem = Format(Val(mem.dwTotalPhys) / 1024 / 1024, "0")
    
    '��дϵͳ����
    Call SysPara

    '��д�Ѱ�װϵͳ�嵥
    Call SysCreated

End Sub

Private Sub Form_Resize()
    Dim sngHeight As Single
    
    On Error Resume Next
    With imgMain
        .Top = 700
        .Left = ScaleLeft + 200
    End With
    lblSys.Left = imgMain.Left + imgMain.Width + 200
    
    lvwSys.Left = lblSys.Left
    lvwSys.Width = ScaleWidth - lvwSys.Left - 200
    
    For intCount = 0 To cmdFunction.UBound
        cmdFunction(intCount).Left = lblSys.Left
    Next
    lblNote.Left = cmdFunction(0).Left + cmdFunction(0).Width + 100
    lblNote.Width = ScaleWidth - lblNote.Left - 200
    
    lblPara.Left = imgMain.Left + imgMain.Width + 200
    lvwPara.Left = lblPara.Left
    lvwPara.Width = ScaleWidth - lvwPara.Left - 200
    
    
    '���ø߶�
    sngHeight = IIf(ScaleHeight < 5400, 5400, ScaleHeight) '��С�߶�
    lvwPara.Height = 4050
    lvwPara.Top = sngHeight - lvwPara.Height - 200
    lblPara.Top = lvwPara.Top - lblPara.Height - 30
    
    cmdReCalc.Left = lvwPara.Left + lvwPara.Width - cmdReCalc.Width - 60
    cmdReCalc.Top = lvwPara.Top - cmdReCalc.Height - 15
    
    txtMem.Top = cmdReCalc.Top - lblMem.Height - 105
    txtMem.Left = lblPara.Left + 1170
    
    lblMem.Top = cmdReCalc.Top - lblMem.Height - 85
    lblMem.Left = lblPara.Left
    
    lblNote.Top = lblMem.Top - lblNote.Height - 200
    cmdFunction(0).Top = lblNote.Top - 15
    cmdFunction(1).Top = cmdFunction(0).Top + 345
    cmdFunction(2).Top = cmdFunction(1).Top + 345
    
    lvwSys.Height = lblNote.Top - lvwSys.Top - 90
    
End Sub


Private Sub SysCreated()
    lvwSys.ListItems.Clear
        
    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
    With rsTemp
        Do Until .EOF
            Set objItem = lvwSys.ListItems.Add(, "S" & !���, !����, , 1)
            objItem.SubItems(1) = IIf(IsNull(.Fields("�汾��").value), "", .Fields("�汾��").value)
            objItem.SubItems(2) = !���
            objItem.SubItems(3) = IIf(IsNull(.Fields("������").value), "", .Fields("������").value)
            .MoveNext
        Loop
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMDIMain.stbThis.Panels(2).Text = ""
    mintVersion = 0
    txtMem = ""
End Sub

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = True
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�

'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As zlPrintLvw
    
    Set objPrint = New zlPrintLvw
    If ActiveControl Is lvwPara Then
        objPrint.Title.Text = "Ҫ���������ݿ����"
        Set objPrint.Body.objData = lvwPara
    Else
        objPrint.Title.Text = "�Ѱ�װӦ��ϵͳ"
        Set objPrint.Body.objData = lvwSys
    End If
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(CurrentDate, "yyyy��MM��dd��")
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

Private Sub SysPara()
    Dim strParas As String
    Dim blnCBO As Boolean 'Ϊ�治��ʾ ,optimizer_index_cost_adj �� optimizer_index_caching
    Dim blnSGA As Boolean 'Ϊ�治��ʾ Db_cache_size,Shared_pool_size,java_pool_size
    blnCBO = True
    blnSGA = False
    
    lvwPara.ListItems.Clear
    
    With lvwPara
        
        Set objItem = lvwPara.ListItems.Add(, "open_cursors", "open_cursors")
        objItem.SubItems(2) = ">=60"
        objItem.SubItems(3) = "ÿ�������пɴ�SQL�α�������������ִ��һ�����ӵĴ�����ʱ��������Ҫ�򿪴�����SQL�α�"
        objItem.SubItems(4) = "����"
        objItem.SubItems(5) = "��ת��"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
                
        Set objItem = lvwPara.ListItems.Add(, "session_cached_cursors", "session_cached_cursors")
        objItem.SubItems(2) = ">=10"
        objItem.SubItems(3) = "ÿ���Ự����Ŀͻ����α�����,�˲���Ӱ��SQL�������,�Ӵ����ֵ�����SQL���ܣ�������ø���ķ������ڴ�"
        objItem.SubItems(4) = "����"
        objItem.SubItems(5) = "��ת��"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        
        Set objItem = lvwPara.ListItems.Add(, "max_enabled_roles", "max_enabled_roles")
        objItem.SubItems(2) = ">=40"
        objItem.SubItems(3) = "����Ҫ�����϶�Ľ�ɫʱ�����޸�����ɫ������"
        objItem.SubItems(4) = "����"
        objItem.SubItems(5) = "��ת��"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        
        Set objItem = lvwPara.ListItems.Add(, "processes", "processes")
        objItem.SubItems(2) = ">=150"
        objItem.SubItems(3) = "���ݿ�ʵ���Ĳ�����������������������ٽ����ƿ����ӵĲ���������"
        objItem.SubItems(4) = "����"
        objItem.SubItems(5) = "��ת��"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        
        Set objItem = lvwPara.ListItems.Add(, "sessions", "sessions")
        objItem.SubItems(2) = ">=150"
        objItem.SubItems(3) = "���ݿ�ʵ���Ĳ����Ự����������������ٽ����ƿ����ӵĲ����Ự��"
        objItem.SubItems(4) = "����"
        objItem.SubItems(5) = "��ת��"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        
        Set objItem = lvwPara.ListItems.Add(, "job_queue_processes", "job_queue_processes")
        objItem.SubItems(2) = ">=10"
        objItem.SubItems(3) = "����ϵͳ�����е��Զ���ҵ�������ݿ������õ��Զ���ҵ��Ŀ����(������36)"
        objItem.SubItems(4) = "����"
        objItem.SubItems(5) = "��ת��"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        
        
        Set objItem = lvwPara.ListItems.Add(, "compatible", "compatible")
        objItem.SubItems(2) = ">=10.0.3"
        objItem.SubItems(3) = "���ݲ�����ZLHIS��׼���ƷҪ�����Ͱ汾Ϊ10.0.3"
        objItem.SubItems(4) = "���ݰ汾��"
        objItem.SubItems(5) = "��ת��"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        
        Set objItem = lvwPara.ListItems.Add(, "optimizer_mode", "optimizer_mode")
        objItem.SubItems(2) = "ALL_ROWS"
        objItem.SubItems(3) = "�Ż���ģʽ����������Ϊall_rows"
        objItem.SubItems(4) = "�ı�"
        objItem.SubItems(5) = "ǿ��"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        
        '����ALL_ROWSʱ����ʾ
        Set objItem = lvwPara.ListItems.Add(, "optimizer_index_cost_adj", "optimizer_index_cost_adj")
        objItem.SubItems(2) = "20"
        objItem.SubItems(3) = "CBO�Ż���ģʽ��,����SQLִ�мƻ��ĳɱ�ʱ,��������ڱ�ɨ��ĳɱ���������,ֵԽС,����ɨ��Ĺ���ɱ���Խ��"
        objItem.SubItems(4) = "����"
        objItem.SubItems(5) = "ǿ��"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"

        Set objItem = lvwPara.ListItems.Add(, "optimizer_index_caching", "optimizer_index_caching")
        objItem.SubItems(2) = "80"
        objItem.SubItems(3) = "CBO�Ż���ģʽ��,����SQLִ�мƻ��ĳɱ�ʱ,�������ڴ��еĹ������,��Ӱ��Ƕ��ѭ����in-list����,����ɨ��Ĺ���ɱ���Խ��"
        objItem.SubItems(4) = "����"
        objItem.SubItems(5) = "ǿ��"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        
        Set objItem = lvwPara.ListItems.Add(, "cursor_sharing", "cursor_sharing")
        objItem.SubItems(2) = "EXACT"
        objItem.SubItems(3) = "�˲���Ӱ��SQL�Ľ���,����ΪEXACT(��ȷƥ��)"
        objItem.SubItems(4) = "�ı�"
        objItem.SubItems(5) = "ǿ��"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        
        
        
        '�ڴ����ò���
        '----------------------------------------------------------------------------------------------
        Set objItem = lvwPara.ListItems.Add(, "log_buffer", "log_buffer")
        objItem.SubItems(2) = ">=" & Val(209715200 / 1024 / 1024) & "M"
        objItem.SubItems(3) = "��־��������С���ڴ��������ݴ������磺�����ű��е�����������Ӱ��ϴ󣬽��鲻����200M���޸ĺ�������ʵ��"
        objItem.SubItems(4) = "����"
        objItem.SubItems(5) = "ת��ǿ��"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        
        Set objItem = lvwPara.ListItems.Add(, "parallel_execution_message_size", "parallel_execution_message_size")
        objItem.SubItems(2) = ">=8192"
        objItem.SubItems(3) = "����ִ����Ϣ�Ĵ�С������8192ʱ�����ò���DDL�ؽ�����ʱ���ܻᱨ��"
        objItem.SubItems(4) = "����"
        objItem.SubItems(5) = "��ת��"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        
        If mintVersion >= 90 Then
            Set objItem = lvwPara.ListItems.Add(, "db_cache_size", "db_cache_size")
            'objItem.SubItems(2) = ">=26214400"
            objItem.SubItems(2) = ">=" & Format((Val(txtMem) * 1024 * 1024 * 0.3) / 1024 / 1024, "0")
            objItem.SubItems(3) = "���ݻ���ش�С(M),���ݻ����Ӧ�����ܵش�,��������ΪSGA��80%"
            objItem.SubItems(4) = "����"
            objItem.SubItems(5) = "ת��ǿ��"
        Else
            Set objItem = lvwPara.ListItems.Add(, "db_block_buffers", "db_block_buffers")
            objItem.SubItems(2) = ">=" & Format((Val(txtMem) * 1024 * 1024 * 0.25) / 8192, "0")
            objItem.SubItems(3) = "�Կ��С��ʾ�����ݻ�����,���ݻ����Ӧ�����ܵش�,��������ΪSGA��80%"
            objItem.SubItems(4) = "����"
            objItem.SubItems(5) = "��ת��"
             
        End If
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        
        Set objItem = lvwPara.ListItems.Add(, "shared_pool_size", "shared_pool_size")
        objItem.SubItems(2) = ">=" & Format((Val(txtMem) * 1024 * 1024 * 0.1) / 1024 / 1024, "0")
        objItem.SubItems(3) = "�����(����SQL��仺�桢ϵͳ�����ֵ仺���)���ڴ���(M)������Ϊ�����ڴ��10-30%,�����̫��,����Ӱ������"
        objItem.SubItems(4) = "����"
        objItem.SubItems(5) = "ת��ǿ��"
        strParas = strParas & ",'" & Trim(objItem.Key) & "'"
                
        '-- 9I ����
        If mintVersion >= 90 Then
            Set objItem = lvwPara.ListItems.Add(, "workarea_size_policy", "workarea_size_policy")
            objItem.SubItems(2) = "AUTO"
            objItem.SubItems(3) = "ָPGA�Ĺ���ģʽ������Ϊ�Զ�(Auto),������������Щ�����ĺ���ֵsort_area_size,hash_area_size,bitmap_merge_area_size"
            objItem.SubItems(4) = "�ı�"
            objItem.SubItems(5) = "ǿ��"
            strParas = strParas & ",'" & Trim(objItem.Key) & "'"
            
            '
            Set objItem = lvwPara.ListItems.Add(, "pga_aggregate_target", "pga_aggregate_target")
            objItem.SubItems(2) = ">0"
            objItem.SubItems(3) = "���лỰ���õ�˽���ڴ�������ÿ��������100M���ڵ�5%�����������Ϊ�ܵ������ڴ�*80%*20%"
            objItem.SubItems(4) = "����"
            objItem.SubItems(5) = "ת��ǿ��"
            strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        End If
     
        If mintVersion >= 100 Then
            Set objItem = lvwPara.ListItems.Add(, "sga_target", "sga_target")
            objItem.SubItems(2) = ">0"
            objItem.SubItems(3) = "���ݻ���͹���صȹ����ڴ��������0��ʾ�Զ������������Ϊ�ܵ������ڴ�*80%*80%������޸�ֵ����SGA_MAX_SIZE���������޸ĺ��߲�����ʵ��"
            objItem.SubItems(4) = "����"
            objItem.SubItems(5) = "ת��ǿ��"
            strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        End If
        
        '11G������
        If mintVersion > 100 Then
            Set objItem = lvwPara.ListItems.Add(, "memory_target", "memory_target")
            objItem.SubItems(2) = ">0"
            objItem.SubItems(3) = "���ݿ�ʵ�����ڴ�������0��ʾ�Զ�������������Ϊ�����ڴ��80%������޸�ֵ����memory_max_target���������޸ĺ��߲�����ʵ����"
            objItem.SubItems(4) = "����"
            objItem.SubItems(5) = "ת��ǿ��"
            strParas = strParas & ",'" & Trim(objItem.Key) & "'"
        End If
    End With
    
    Dim strParList As String
    strParList = Replace(strParas, "'", "")
    Dim lngϵ�� As Long, bln����Ҫ�� As Boolean
    With rsTemp
        On Error Resume Next
        strSql = "select lower(name) as name,value" & _
                " from v$parameter" & _
                " where name in (" & Mid(strParas, 2) & ")"
        If .State = adStateOpen Then .Close
        .Open strSql, gcnOracle, adOpenKeyset
        Do While Not .EOF
            bln����Ҫ�� = False
            Set objItem = lvwPara.ListItems(.Fields("name").value)
            objItem.SubItems(1) = .Fields("value").value
            Select Case objItem.SubItems(4)
            Case "����"
                lngϵ�� = 1
                If objItem.SubItems(5) = "ת��ǿ��" Then lngϵ�� = 1048576  '(1024 * 1024)
                objItem.SubItems(1) = Format(Val(.Fields("value").value) / lngϵ��, "0") & IIf(lngϵ�� > 1, "M", "")
                            
                If .Fields("value").value < Val(Mid(objItem.SubItems(2), 3)) * lngϵ�� Then
                    If objItem.Key = "sga_target" Then
                        blnSGA = True
                    End If
                    If objItem.SubItems(5) = "ǿ��" Or objItem.SubItems(5) = "ת��ǿ��" Then
                        objItem.ForeColor = vbBlue
                        objItem.ListSubItems(1).ForeColor = vbBlue
                        objItem.ListSubItems(2).ForeColor = vbBlue
                        objItem.ListSubItems(3).ForeColor = vbBlue
                    End If
                End If
            Case "�ı�"
                If objItem.Key = "optimizer_mode" Then
                    blnCBO = .Fields("value").value = "ALL_ROWS"
                End If
                
                If UCase(.Fields("value").value) <> objItem.SubItems(2) Then
                    If objItem.SubItems(5) = "ǿ��" Then
                        objItem.ForeColor = vbBlue
                        objItem.ListSubItems(1).ForeColor = vbBlue
                        objItem.ListSubItems(2).ForeColor = vbBlue
                        objItem.ListSubItems(3).ForeColor = vbBlue
                    Else
                        bln����Ҫ�� = True
                    End If
                End If
            Case "���ݰ汾��"
                If Val(Replace(.Fields("value").value, ".", "")) < Val(Replace(Mid(objItem.SubItems(2), 3), ".", "")) Then
                    objItem.ForeColor = vbBlue
                    objItem.ListSubItems(1).ForeColor = vbBlue
                    objItem.ListSubItems(2).ForeColor = vbBlue
                    objItem.ListSubItems(3).ForeColor = vbBlue
                End If
            End Select
            If bln����Ҫ�� Then
                '����Ҫ��(��)
                objItem.ForeColor = RGB(255, 0, 0)
                objItem.ListSubItems(1).ForeColor = RGB(255, 0, 0)
                objItem.ListSubItems(2).ForeColor = RGB(255, 0, 0)
                objItem.ListSubItems(3).ForeColor = RGB(255, 0, 0)
            End If
            .MoveNext
        Loop
        
        
        '-- �������ʾ����
        Dim i As Integer
 
        For i = 1 To lvwPara.ListItems.Count
            If blnCBO = False Then
                If i < lvwPara.ListItems.Count Then
                    If InStr("optimizer_index_cost_adj,optimizer_index_caching", lvwPara.ListItems(i).Key) > 0 Then
                         lvwPara.ListItems.Remove i
                         i = i - 1
                    End If
                End If
            End If
            
            If blnSGA = True Then
                If i < lvwPara.ListItems.Count Then
                    If InStr("db_cache_size,shared_pool_size,java_pool_size", lvwPara.ListItems(i).Key) > 0 Then
                        lvwPara.ListItems.Remove i
                        i = i - 1
                    End If
                End If
            End If
        Next

        
    End With
    
End Sub

Private Sub txtMem_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789" & Chr(8), Chr(KeyAscii)) <= 0 Then KeyAscii = 0
End Sub


