VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmStatus 
   BackColor       =   &H80000005&
   Caption         =   "����״̬���"
   ClientHeight    =   6000
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   9105
   ControlBox      =   0   'False
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmStatus.frx":000C
   ScaleHeight     =   6000
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.TabControl tabMore 
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   3480
      Width           =   2415
      _Version        =   589884
      _ExtentX        =   4260
      _ExtentY        =   450
      _StockProps     =   64
   End
   Begin VB.CommandButton cmdKillSession 
      Caption         =   "�����Ự(&D)"
      Height          =   300
      Left            =   2145
      TabIndex        =   4
      Top             =   3120
      Width           =   1200
   End
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1560
      TabIndex        =   2
      Top             =   1170
      Width           =   3975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "ˢ��(&R)"
      Height          =   300
      Left            =   885
      TabIndex        =   3
      Top             =   3120
      Width           =   1200
   End
   Begin VB.Timer TimerRefresh 
      Interval        =   30000
      Left            =   165
      Top             =   2355
   End
   Begin MSComctlLib.ListView lvw�ϻ��û� 
      Height          =   1290
      Left            =   885
      TabIndex        =   0
      Top             =   1455
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   2275
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImgСͼ��"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Audsid"
         Text            =   "Audsid"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Sid,Serial#"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "����"
         Text            =   "����"
         Object.Width           =   2671
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "������"
         Text            =   "������"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "����"
         Text            =   "����"
         Object.Width           =   1667
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "�û�"
         Text            =   "�û�"
         Object.Width           =   1401
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "��������"
         Text            =   "��������"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "����ϵͳ�û���"
         Text            =   "����ϵͳ�û���"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Key             =   "����ʱ��"
         Text            =   "����ʱ��"
         Object.Width           =   3678
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Key             =   "״̬"
         Text            =   "״̬"
         Object.Width           =   1295
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Key             =   "ʵ��"
         Text            =   "ʵ��"
         Object.Width           =   1085
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Key             =   "SID"
         Text            =   "SID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Key             =   "SERIAL#"
         Text            =   "SERIAL#"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgСͼ�� 
      Left            =   5640
      Top             =   2715
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatus.frx":0505
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatus.frx":065F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblConError 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   180
      Left            =   4095
      TabIndex        =   13
      Top             =   2820
      Width           =   90
   End
   Begin VB.Label lblConNormal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   180
      Left            =   2835
      TabIndex        =   12
      Top             =   2820
      Width           =   90
   End
   Begin VB.Label lblConCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   180
      Left            =   1575
      TabIndex        =   11
      Top             =   2820
      Width           =   90
   End
   Begin VB.Label lblCon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Ự����     ������     �쳣��    "
      Height          =   180
      Left            =   885
      TabIndex        =   10
      Top             =   2820
      Width           =   3060
   End
   Begin VB.Label lblMain 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "״̬˵������ǰ-��ʾ������Լ���"
      Height          =   180
      Index           =   1
      Left            =   885
      TabIndex        =   9
      Top             =   870
      Width           =   2790
   End
   Begin VB.Label lblMain 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "          ����-���������������û���"
      Height          =   180
      Index           =   2
      Left            =   2700
      TabIndex        =   8
      Top             =   870
      Width           =   3150
   End
   Begin VB.Label lblMain 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "          ����-�����ѷ�������ĻỰ��"
      Height          =   180
      Index           =   3
      Left            =   4890
      TabIndex        =   7
      Top             =   870
      Width           =   3330
   End
   Begin VB.Label lblFind 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����(&Z)"
      Height          =   180
      Left            =   885
      TabIndex        =   1
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label lblMain 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ǰ�����ݿ����ӵ��û��Ự�嵥"
      Height          =   180
      Index           =   0
      Left            =   885
      TabIndex        =   6
      Top             =   630
      Width           =   2700
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����״̬���"
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
      Left            =   200
      TabIndex        =   5
      Top             =   100
      Width           =   1440
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   210
      Picture         =   "frmStatus.frx":07B9
      Top             =   645
      Width           =   480
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintCount As Integer            '��¼���Ҷ�λ��һ�ε�λ��
Private mstrNowAudSID As String      '�洢��ǰѡ����

Private mintLastTime  As Integer    '��¼���ӵĳ���ʱ��,���ڳ�ʱ��Ͽ�����
Private mstrConnStat As String  '��¼����״̬,1.��ʼ 2.ֹͣ

Private mfrmChild As New frmHistSql


Private Sub LoadSession()
    Dim rsTemp As New ADODB.Recordset
    Dim objItem As ListItem, lngInterval As Long, strSID As String
    Dim lngCount As Long
    Dim blnFindTag As Boolean
    
    On Error GoTo errHandle
    lngInterval = TimerRefresh.Interval
    TimerRefresh.Interval = 0
    
    With rsTemp
        gstrSQL = "Select Audsid, Program, Username, Osuser, ����, ����, Status, Terminal, ����ʱ��, Sid, Serial#, t.�Ա�" & IIf(gblnRac, ", Inst_Id", "") & vbNewLine & _
                            "From (Select u.Audsid, u.Program, u.Username, u.Osuser, u.Status, u.Terminal," & vbNewLine & _
                            "              To_Char(u.Logon_Time, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��, u.Sid, u.Serial#" & IIf(gblnRac, ", Inst_Id", "") & vbNewLine & _
                            "       From " & IIf(gblnRac, "G", "") & "v$session U" & vbNewLine & _
                            "       Where u.Username Is Not Null) A," & vbNewLine & _
                            "     (Select w.�û���, p.����, b.���� As ����, p.�Ա�" & vbNewLine & _
                            "       From ��Ա�� P, �ϻ���Ա�� W, ���ű� B, ������Ա C" & vbNewLine & _
                            "       Where p.Id = c.��Աid And b.Id = c.����id And c.ȱʡ = 1 And" & vbNewLine & _
                            "             ((p.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or p.����ʱ�� Is Null) And p.Id = w.��Աid)) T" & vbNewLine & _
                            "Where a.Username = t.�û���(+)" & vbNewLine & _
                            "Order By ����, Username, Terminal, Audsid, ����ʱ��"

        If .State = adStateOpen Then .Close
        
        .Open gstrSQL, gcnOracle, adOpenKeyset
        lvw�ϻ��û�.ListItems.Clear
        Do While Not .EOF
            If strSID <> .Fields("AUDSID").value Then
                strSID = .Fields("AUDSID").value
                Set objItem = lvw�ϻ��û�.ListItems.Add(, , .Fields("AUDSID").value, _
                    , IIf(.Fields("�Ա�").value = "Ů", 2, 1))
                objItem.SubItems(1) = "" & .Fields("SID").value & "," & .Fields("SERIAL#").value
                objItem.SubItems(2) = "" & .Fields("����").value
                objItem.SubItems(3) = "" & .Fields("Terminal").value
                objItem.SubItems(4) = "" & .Fields("����").value
                objItem.SubItems(5) = "" & .Fields("USERNAME").value
                objItem.SubItems(6) = "" & .Fields("PROGRAM").value
                objItem.SubItems(7) = "" & .Fields("OsUser").value
                objItem.SubItems(8) = "" & .Fields("����ʱ��").value
                Select Case "" & .Fields("STATUS").value
                Case "ACTIVE"
                    objItem.SubItems(9) = "��ǰ"
                Case "INACTIVE"
                    objItem.SubItems(10) = "����"
                Case Else
                    objItem.SubItems(11) = "����"
                End Select
                If gblnRac Then
                    objItem.SubItems(10) = "" & .Fields("INST_ID").value
                Else
                    objItem.SubItems(10) = 1
                End If
                objItem.SubItems(11) = "" & .Fields("SID").value
                objItem.SubItems(12) = "" & .Fields("SERIAL#").value
            End If
            .MoveNext
        Loop
    End With
    
    With lvw�ϻ��û�
        .ColumnHeaders(10).Width = IIf(gblnRac, 615, 0)
        lblConCount.Caption = .ListItems.Count
        lblConNormal.Caption = 0
        lblConError.Caption = 0
        For lngCount = 1 To .ListItems.Count
            If .ListItems(lngCount).SubItems(8) = "����" Then
                lblConError.Caption = lblConError.Caption + 1
            Else
                lblConNormal.Caption = lblConNormal.Caption + 1
            End If
            If .ListItems(lngCount).Text = mstrNowAudSID Then
                .ListItems(lngCount).Selected = True
                .ListItems(lngCount).EnsureVisible
                blnFindTag = True
            End If
        Next
        If .ListItems.Count > 0 Then
            If blnFindTag = False Then
                .ListItems(1).Selected = True
                .ListItems(1).EnsureVisible
                mstrNowAudSID = .ListItems(1).Text
            End If
        Else
            mstrNowAudSID = 0
        End If
    End With
    TimerRefresh.Interval = lngInterval
    Call lvw�ϻ��û�_ItemClick(lvw�ϻ��û�.ListItems(1))
    Exit Sub
errHandle:
    MsgBox "[" & err.Number & "]" & err.Description, vbExclamation, Me.Caption
    If 0 = 1 Then
        Resume
    End If
End Sub

Private Sub cmdKillSession_Click()
    Dim lngCount As Long
    Dim blnSelect As Boolean
    Dim strNote As String
    Dim i As Long

    TimerRefresh.Enabled = False
    With lvw�ϻ��û�
        lngCount = 0
        blnSelect = False
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked = True Then lngCount = lngCount + 1
            If .ListItems(i).Selected = True Then blnSelect = True
        Next
        
        If lngCount > 0 Then
            If MsgBox("�����Ѿ���ѡ�ĻỰ���ܵ��¶���û�δ��������ݶ�ʧ����ȷ��Ҫ�������� " & lngCount & " ���Ự��", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                For i = 1 To .ListItems.Count
                    If .ListItems(i).Checked = True Then
                        If .ListItems(i).SubItems(9) <> "��ǰ" Then
                            Call KillSessions(Trim(.ListItems(i).SubItems(11)) & "," & Trim(.ListItems(i).SubItems(12)))
                            strNote = strNote & ";�û�:" & .ListItems(i).SubItems(5) & ",��������:" & .ListItems(i).SubItems(6)
                        Else
                            Call MsgBox("���ܽ���״̬Ϊ""��ǰ""�ĻỰ", vbInformation, gstrSysName)
                        End If
                    End If
                Next
            End If
        Else
            .SetFocus
            For i = 1 To .ListItems.Count
                If .ListItems(i).Selected = True Then Exit For
            Next
            If .ListItems(i).SubItems(8) <> "��ǰ" Then
                If MsgBox("�����Ự���ܵ��¸��û�δ��������ݶ�ʧ����ȷ��Ҫ���������û� " & .ListItems(i).SubItems(5) & "(AUDSID:" & .ListItems(i).Text & ") �ĻỰ��", vbQuestion + vbYesNo, gstrSysName) = vbYes And blnSelect = True Then
                    Call KillSessions(Trim(.ListItems(i).SubItems(11)) & "," & Trim(.ListItems(i).SubItems(12)))
                    strNote = strNote & ";" & .ListItems(i).SubItems(5) & "," & .ListItems(i).SubItems(6)
                End If
            Else
                Call MsgBox("���ܽ���״̬Ϊ""��ǰ""�ĻỰ", vbInformation, gstrSysName)
            End If
        End If
        
        If strNote <> "" Then
            '������Ҫ������־
            Call SaveAuditLog(2, "�����Ự", Mid(strNote, 2))
        End If
        Call cmdRefresh_Click
        TimerRefresh.Enabled = True
    End With
End Sub

Private Sub cmdRefresh_Click()
    Call LoadSession
    lvw�ϻ��û�.SetFocus
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    
    Call LoadSession
    '�������ʼ��
    txtFind.Text = "�������û���������": txtFind.ForeColor = vbGrayText: mintCount = 0
    
    Call InitTab
    Exit Sub
errHandle:
    MsgBox err.Description, vbQuestion, gstrSysName
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    cmdRefresh.Top = txtFind.Top - 30
    cmdRefresh.Left = txtFind.Left + txtFind.Width + 50
    
    cmdKillSession.Top = txtFind.Top - 30
    cmdKillSession.Left = cmdRefresh.Left + cmdRefresh.Width + 50
    
    With lvw�ϻ��û�
        If gblnDBA Then
            .Height = (Me.ScaleHeight) / 2
        Else
            .Height = Me.ScaleHeight - .Top - 350
        End If
        .Width = frmMDIMain.Width - frmMDIMain.sbFunc.Width - 1500
        lblCon.Left = .Left
        lblCon.Top = .Top + .Height + 100
        lblConCount.Left = lblCon.Left + 700
        lblConCount.Top = lblCon.Top
        lblConNormal.Left = lblCon.Left + 1700
        lblConNormal.Top = lblCon.Top
        lblConError.Left = lblCon.Left + 2680
        lblConError.Top = lblCon.Top
    End With
    
    If Not gblnDBA Then Exit Sub
    tabMore.Left = lblCon.Left
    tabMore.Width = lvw�ϻ��û�.Width
    tabMore.Top = lblCon.Top + lblCon.Height + 120
    tabMore.Height = Me.ScaleHeight - tabMore.Top - 200
End Sub

Private Sub lvw�ϻ��û�_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvw�ϻ��û�.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub lvw�ϻ��û�_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Item.Selected = True
End Sub

Private Sub lvw�ϻ��û�_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim lngPort As Long, strPgm As String
    
    If lvw�ϻ��û�.SelectedItem Is Nothing Then Exit Sub
    If lvw�ϻ��û�.SelectedItem.Selected = False Then Exit Sub
    mstrNowAudSID = Item.Text
'    Item.Checked = IIf(Item.Checked, False, True)
'    frmMDIMain.stbThis.Panels(2).Text = "�û���" & lvw�ϻ��û�.SelectedItem.Text & _
'            "  ����ʱ�䣺" & lvw�ϻ��û�.SelectedItem.SubItems(1)
    mfrmChild.SetUser lvw�ϻ��û�.SelectedItem.SubItems(5)
    mfrmChild.SetSid lvw�ϻ��û�.SelectedItem.SubItems(11), lvw�ϻ��û�.SelectedItem.SubItems(12)
End Sub

Private Sub TimerRefresh_Timer()
    Call LoadSession
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
    objPrint.Title.Text = "�û��Ự�嵥"
    Set objPrint.Body.objData = lvw�ϻ��û�
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

Private Sub txtFind_Change()
    mintCount = 0
End Sub

Private Sub txtFind_GotFocus()
    If txtFind.ForeColor = vbGrayText Then
        txtFind.Text = ""
        txtFind.ForeColor = vbBlack
    Else
        txtFind.SelStart = 0
        txtFind.SelLength = Len(txtFind.Text)
    End If
End Sub

Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim intRow As Integer

    If KeyCode = vbKeyReturn And txtFind.Text <> "" Then
        txtFind.Text = Replace(txtFind.Text, " ", "")
        With lvw�ϻ��û�
            For intRow = mintCount + 1 To .ListItems.Count
                If InStr(UCase(.ListItems(intRow).SubItems(4)), UCase(txtFind.Text)) > 0 Or InStr(.ListItems(intRow).SubItems(5), UCase(txtFind.Text)) > 0 Then
                    mintCount = intRow
                    lvw�ϻ��û�_ItemClick .ListItems(intRow)
                    .ListItems(intRow).Selected = True
                    .ListItems(intRow).EnsureVisible
                    Exit For
                End If
            Next

            If intRow = (.ListItems.Count + 1) Then
                If mintCount = 0 Then
                    Call MsgBox("δ�ҵ��롰" & txtFind.Text & "��ƥ�����Ŀ�������������û�����������", vbInformation, gstrSysName)
                    txtFind.Text = "": txtFind.SetFocus
                Else
                    mintCount = 0
                End If
            End If
        End With
    End If
End Sub

Private Sub txtFind_LostFocus()
    If txtFind.Text = "" Then
        txtFind.Text = "�������û���������"
        txtFind.ForeColor = vbGrayText
    End If
End Sub

Private Sub InitTab()
    tabMore.Visible = gblnDBA
    With tabMore
        .InsertItem 0, "��ʷSQL���", mfrmChild.hwnd, 0
        .Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
    End With
End Sub


