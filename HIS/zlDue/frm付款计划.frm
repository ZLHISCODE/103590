VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frm����ƻ� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ƻ�����"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Tag             =   "0"
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5760
      Left            =   6510
      TabIndex        =   18
      Top             =   495
      Width           =   60
   End
   Begin ZL9BillEdit.BillEdit mshList 
      Height          =   2760
      Left            =   75
      TabIndex        =   0
      Top             =   2595
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   4868
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   5
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   6720
      TabIndex        =   4
      Top             =   4860
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6720
      TabIndex        =   3
      Top             =   960
      Width           =   1100
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "����(&R)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6720
      TabIndex        =   2
      Top             =   1530
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6720
      TabIndex        =   1
      Top             =   555
      Width           =   1100
   End
   Begin VB.Frame fraTemp 
      Caption         =   "Ӧ������Ϣ"
      Height          =   1515
      Left            =   75
      TabIndex        =   5
      Top             =   660
      Width           =   6300
      Begin VB.TextBox txtInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   3570
         TabIndex        =   16
         Top             =   1020
         Width           =   1470
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   1050
         TabIndex        =   10
         Top             =   1020
         Width           =   1470
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   1050
         TabIndex        =   9
         Top             =   630
         Width           =   5100
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   4695
         TabIndex        =   7
         Top             =   255
         Width           =   1455
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1050
         TabIndex        =   6
         Top             =   255
         Width           =   2955
      End
      Begin VB.Label lblTemp 
         AutoSize        =   -1  'True
         Caption         =   "δ�ƻ����"
         Height          =   180
         Index           =   4
         Left            =   2610
         TabIndex        =   17
         Top             =   1110
         Width           =   900
      End
      Begin VB.Label lblTemp 
         AutoSize        =   -1  'True
         Caption         =   "Ӧ�����"
         Height          =   180
         Index           =   3
         Left            =   285
         TabIndex        =   13
         Top             =   1110
         Width           =   720
      End
      Begin VB.Label lblTemp 
         AutoSize        =   -1  'True
         Caption         =   "��Ʊ��"
         Height          =   180
         Index           =   1
         Left            =   4110
         TabIndex        =   12
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lblTemp 
         AutoSize        =   -1  'True
         Caption         =   "Ӧ�������"
         Height          =   180
         Index           =   0
         Left            =   105
         TabIndex        =   11
         Top             =   330
         Width           =   900
      End
      Begin VB.Label lblTemp 
         AutoSize        =   -1  'True
         Caption         =   "ժ    Ҫ"
         Height          =   180
         Index           =   2
         Left            =   285
         TabIndex        =   8
         Top             =   720
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   -45
      TabIndex        =   19
      Top             =   495
      Width           =   6615
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "frm����ƻ�.frx":0000
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblInfor 
      Caption         =   "����Ӧ����¼�����ƶ����󸶿�ƻ���"
      Height          =   165
      Left            =   645
      TabIndex        =   20
      Top             =   330
      Width           =   5655
   End
   Begin VB.Label lblTemp 
      AutoSize        =   -1  'True
      Caption         =   "Ӧ�������֧���ƻ�"
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   5
      Left            =   2295
      TabIndex        =   15
      Top             =   2370
      Width           =   1620
   End
   Begin VB.Label lblTemp 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   75
      TabIndex        =   14
      Top             =   2295
      Width           =   6300
   End
End
Attribute VB_Name = "frm����ƻ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mlngID As Long, mstrDate As String
Dim mblnOK As Boolean
Dim mblnChange As Boolean
Private Enum HeadCol
        �ƻ�֧������ = 0
        ֧�����
        ִ��
        �ƻ���
        �ƶ�����
End Enum

Public Sub �ƻ�(ByVal FrmMain As Object, lngID As Long, Optional ByRef blnSussces As Boolean)
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�ƶ��ƻ�
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------

    mlngID = lngID
    Call initCard
    
    Me.Show vbModal, FrmMain
    blnSussces = mblnOK
End Sub
Private Sub initCard()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ʼ�ƻ�
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rslist As New ADODB.Recordset, lngRow As Long, strSQL As String
    Dim lngID As Long
    'by lesfeng 2009-12-2 �����Ż�
    strSQL = "" & _
            "   Select ��λID,max(decode(��¼״̬,1,ID,3,iD,0)) as ID, " & _
            "       max(��Ʊ��) ��Ʊ��,max(ժҪ) ժҪ,sum(nvl(��Ʊ���,0)) ��Ʊ��� " & _
            "   From Ӧ����¼  a  " & _
            "   Where  exists  (Select ��¼����,NO,nvl(��Ŀid,0),nvl(���,0) From Ӧ����¼  " & _
            "                   where ��¼����=a.��¼���� and no=a.no and nvl(��Ŀid,0)=nvl(a.��Ŀid,0) and " & _
            "                           nvl(���,0)=nvl(a.���,0) and nvl(ϵͳ��ʶ,0)=nvl(a.ϵͳ��ʶ,0) and ID=[1]" & _
            "                           And ��¼����<>-1 and ����� is not null )" & _
            "   Group by ��λID,��¼����,NO,��Ŀid,���"
    
    
    Err = 0
    On Error GoTo ErrHand:
    Set rslist = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
    If rslist.EOF Then
        MsgBox "�üƻ�δ�ҵ�,�����Ѿ�������ɾ��!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    lngID = Val(Nvl(rslist!ID))
    txtInfo(0) = IIf(IsNull(rslist!ID), "", rslist!ID)
    txtInfo(1) = IIf(IsNull(rslist!��Ʊ��), "", rslist!��Ʊ��)
    txtInfo(2) = IIf(IsNull(rslist!ժҪ), "", rslist!ժҪ)
    txtInfo(3) = IIf(IsNull(rslist!��Ʊ���), "0.00", Format(rslist!��Ʊ���, "###0.00;-###0.00;0.00;0.00"))
    txtInfo(4) = IIf(IsNull(rslist!��Ʊ���), "0.00", Format(rslist!��Ʊ���, "###0.00;-###0.00;0.00;0.00"))
    txtInfo(0).Tag = rslist!��λID
    
    Call initGrid
    
    strSQL = "" & _
        "   Select �ƻ�����,�ƻ����,Decode(�������,Null,' ','��') As ִ��,�ƻ���,�ƶ����� " & _
        "   From Ӧ����¼ " & _
        "   Where ID=[1] And ��¼����=-1 " & _
        "   Order By �ƻ����"
    Set rslist = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    Dim dblTmp As Double
    
    With mshList
        .ClearBill
        .Rows = rslist.RecordCount + 2
        lngRow = 1
        While Not rslist.EOF
            .TextMatrix(lngRow, 0) = Format(rslist(0), "yyyy-MM-dd")
            .TextMatrix(lngRow, 1) = Format(rslist(1), "0.00")
            .TextMatrix(lngRow, 2) = Nvl(rslist(2))
            .TextMatrix(lngRow, 3) = Nvl(rslist(3))
            .TextMatrix(lngRow, 4) = Format(rslist(4), "yyyy-MM-dd")
            dblTmp = dblTmp + Nvl(rslist(1), 0)
            rslist.MoveNext
            lngRow = lngRow + 1
        Wend
        
        .Value = Format(zldatabase.Currentdate, "yyyy-mm-dd")
        
    End With
    txtInfo(4) = Format(Val(txtInfo(4)) - dblTmp, "###0.00;-###0.00;0;0")
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Sub
Private Sub initGrid()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ʼ������
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    With mshList
        .Cols = 5
        .TextMatrix(0, HeadCol.�ƻ�֧������) = "�ƻ�֧������"
        .TextMatrix(0, HeadCol.֧�����) = "֧�����"
        .TextMatrix(0, HeadCol.ִ��) = "ִ��"
        .TextMatrix(0, HeadCol.�ƻ���) = "�ƻ���"
        .TextMatrix(0, HeadCol.�ƶ�����) = "�ƶ�����"
        
        .ColAlignment(HeadCol.�ƻ�֧������) = flexAlignCenterCenter
        .ColAlignment(HeadCol.֧�����) = flexAlignRightCenter
        .ColAlignment(HeadCol.ִ��) = flexAlignCenterCenter
        .ColAlignment(HeadCol.�ƻ���) = flexAlignLeftCenter
        .ColAlignment(HeadCol.�ƶ�����) = flexAlignCenterCenter
                
        
        .ColWidth(HeadCol.�ƻ�֧������) = 1600
        .ColWidth(HeadCol.֧�����) = 1200
        .ColWidth(HeadCol.ִ��) = 700
        .ColWidth(HeadCol.�ƻ���) = 1000
        .ColWidth(HeadCol.�ƶ�����) = 1600
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��
        .ColData(HeadCol.�ƻ�֧������) = 2
        .ColData(HeadCol.֧�����) = 4
        .ColData(HeadCol.�ƻ���) = 5
        .ColData(HeadCol.�ƶ�����) = 5
        .ColData(HeadCol.ִ��) = 5
        .CmdVisible = True
        .PrimaryCol = HeadCol.�ƻ�֧������
        .LocateCol = HeadCol.�ƻ�֧������
        .Active = True
                                
    End With
    
End Sub
Private Sub cmdExit_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdHelp_Click()
   ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdReset_Click()
    Call initCard
    cmdSave.Enabled = False
    cmdReset.Enabled = False
    mblnChange = False
    mshList.SetFocus
End Sub

Private Sub cmdSave_Click()
    '��֤����
    If IsValid = False Then Exit Sub
        
    '��������
    If Save() Then
        cmdSave.Enabled = False
        cmdReset.Enabled = False
        mblnChange = False
    End If
    mblnOK = True
    mshList.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And mshList.Col = 4 Then
        Me.Tag = 1
    Else
        Me.Tag = 0
    End If
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim blnSaveFlag As Integer
    Dim blnYes As Boolean
    If mblnChange = False Then
        Exit Sub
    End If
    
    ShowMsgbox "���Ѿ���������Ϣ,�������˳��Ļ�," & vbCrLf & "�����ĵ����ݽ����ܱ���,���Ҫ�˳���?", True, blnYes
    If blnYes = True Then
        Exit Sub
    End If
    Cancel = 1
    mshList.SetFocus
End Sub
Private Function IsValid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��֤���ݵĺϷ���
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long, strSQL As String
    Dim dblTmp As Double
    Dim lngRow As Long
    Dim strTmp As String
    If Val(txtInfo(4)) < 0 Then
        ShowMsgbox "�ƻ����������Ӧ��������޸�"
        IsValid = False
        Exit Function
    End If
    With mshList
        For lngLoop = 1 To mshList.Rows - 1
            If Trim(mshList.TextMatrix(lngLoop, HeadCol.�ƻ�֧������)) = "" Then Exit For
            
            Err = 0
            On Error Resume Next
            strSQL = Format(CDate(mshList.TextMatrix(lngLoop, 0)), "yyyy-MM-dd")
            If Err.Number <> 0 Then
                ShowMsgbox "���ڸ�ʽ�������޸ġ�"
                .Row = lngLoop
                .Col = HeadCol.�ƻ�֧������
                .SetFocus
                Exit Function
            End If
            Err.Clear
            
            dblTmp = Val(mshList.TextMatrix(lngLoop, HeadCol.֧�����))
            If dblTmp = 0 Then
                ShowMsgbox "֧�������������!"
                .Row = lngLoop
                .Col = HeadCol.֧�����
                .SetFocus
                Exit Function
            End If
            If CDate(mshList.TextMatrix(lngLoop, HeadCol.�ƶ�����)) > CDate(mshList.TextMatrix(lngLoop, HeadCol.�ƻ�֧������)) Then
                If MsgBox("�ƻ���������С���ƶ��ƻ����ڣ��Ƿ���ԣ�", vbYesNo + vbDefaultButton2 + vbQuestion, Me.Caption) <> vbYes Then
                    mshList.Row = lngLoop
                    mshList.Col = HeadCol.�ƻ�֧������
                    IsValid = False
                    Exit Function
                End If
            End If
            strTmp = Trim(mshList.TextMatrix(lngLoop, HeadCol.�ƻ�֧������))
            For lngRow = lngLoop + 1 To mshList.Rows - 1
                If strTmp = Trim(mshList.TextMatrix(lngRow, HeadCol.�ƻ�֧������)) Then
                    ShowMsgbox "��" & lngLoop & "�����" & lngRow & "�еļƻ�֧��������ͬ��," & vbCrLf & "��ϲ��üƻ�!"
                    .Row = lngLoop
                    .Col = HeadCol.�ƻ�֧������
                    .SetFocus
                    Exit Function
                End If
            Next
        Next
    End With
    IsValid = True
End Function
Private Function Save() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��������
    '--�����:
    '--������:
    '--��  ��:�ɹ�����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------


    Dim lngLoop As Long, strSQL As String, lngNewDate As Boolean
    Dim str�ƻ��� As String
    Dim str�ƶ����� As String
    Dim lngSN As Long
    Dim rsTmp As ADODB.Recordset
    
    Save = False
    
    On Error GoTo ErrHand:
        
    gcnOracle.BeginTrans
    
    strSQL = "ZL_����ƻ�_DELETE (" & mlngID & ")"  '�Ѹ���Ĳ��֣�������δɾ��
    
    zldatabase.ExecuteProcedure strSQL, Me.Caption
    
    With mshList
        For lngLoop = 1 To mshList.Rows - 1
            If Trim(.TextMatrix(lngLoop, HeadCol.�ƻ�֧������)) <> "" And .TextMatrix(lngLoop, HeadCol.ִ��) <> "��" Then
                '���̲���
                '   ID_IN,�ƻ����_IN,�ƻ����_IN,�ƻ�����_IN,�ƻ���_IN,�ƶ�����_IN
                str�ƻ��� = Trim(.TextMatrix(lngLoop, HeadCol.�ƻ���))
                str�ƻ��� = "'" & IIf(str�ƻ��� = "", gstrUserName, str�ƻ���) & "'"
                
                str�ƶ����� = Trim(.TextMatrix(lngLoop, HeadCol.�ƶ�����))
                str�ƶ����� = IIf(str�ƶ����� = "", Format(zldatabase.Currentdate, "yyyy-mm-dd"), str�ƶ�����)
                str�ƶ����� = "to_date('" & str�ƶ����� & "','yyyy-mm-dd')"
                
                '�ƻ����
                If lngSN < lngLoop Or lngSN = 0 Then
                    gstrSQL = "Select a.Rec, b.Sn " & _
                              "From " & _
                              "  (Select 1 ID, Count(1) Rec From Ӧ����¼ Where ID = [1] And ��¼���� = -1 And �ƻ���� = [2]) A," & _
                              "  (Select 1 ID, Max(�ƻ����) Sn From Ӧ����¼ Where ID = [1] And ��¼���� = -1) B " & _
                              "Where a.Id = b.Id "
                    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "�ƻ����", mlngID, lngLoop)
                    If Not rsTmp.EOF Then
                        If rsTmp!rec > 0 Then
                            '�ƻ���Ŵ��ڣ�ȡ�����ţ�����1
                            lngSN = rsTmp!sn + 1
                        Else
                            '�ƻ���Ų����ڣ�ʹ��lngLoop
                            lngSN = lngLoop
                        End If
                    Else
                        lngSN = lngLoop
                    End If
                    rsTmp.Close
                Else
                    lngSN = lngSN + 1
                End If
                
                strSQL = "ZL_����ƻ�_INSERT (" & _
                    mlngID & "," & _
                    lngSN & "," & _
                    Val(mshList.TextMatrix(lngLoop, HeadCol.֧�����)) & ",TO_DATE('" & _
                    Format(mshList.TextMatrix(lngLoop, HeadCol.�ƻ�֧������), "yyyy-MM-dd") & "','yyyy-MM-dd')," & _
                    str�ƻ��� & "," & _
                    str�ƶ����� & ")"
                    
                zldatabase.ExecuteProcedure strSQL, Me.Caption
            End If
        Next
    End With
    gcnOracle.CommitTrans
    Save = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function

Private Sub mshList_AfterDeleteRow()
    cmdSave.Enabled = True
    cmdReset.Enabled = True
    mblnChange = True
    
End Sub

Private Sub mshList_EditChange(curText As String)
    cmdSave.Enabled = True
    cmdReset.Enabled = True
    mblnChange = True
    mshList.TextMatrix(mshList.Row, 3) = UserInfo.����
    If mshList.TextMatrix(mshList.Row, 4) = "" Then mshList.TextMatrix(mshList.Row, 4) = Format(zldatabase.Currentdate, "yyyy-MM-dd")
End Sub

Private Sub mshList_EnterCell(Row As Long, Col As Long)
    With mshList
        If Trim(.TextMatrix(Row, 2)) <> "" Then
            .Active = False
        Else
            Select Case Col
                Case 0, 4
                    .TxtCheck = True
                    .TextMask = "-0123456789"
                    .MaxLength = 16
                Case 1
                    .TxtCheck = True
                    .TextMask = ".0123456789"
                    .MaxLength = 10
            End Select
            .Active = True
        End If
    End With
End Sub

Private Sub mshList_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    With mshList
        .Text = UCase(Trim(.Text))
        strKey = UCase(Trim(.Text))
        If InStr(1, strKey, "'") <> 0 Then
            MsgBox "���зǳ���,���ܼ���!", vbInformation + vbDefaultButton1, gstrSysName
            Cancel = True
            Exit Sub
        End If
        If .ColData(.Col) = 0 Then
            Exit Sub
        End If
        Select Case .Col
            Case HeadCol.֧�����
                If strKey <> "" Then
                        If Not IsNumeric(strKey) And strKey <> "" Then
                            ShowMsgbox "֧��������Ϊ������,�����䣡"
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        If Val(strKey) <= 0 Then
                            ShowMsgbox "֧�������������,�����䣡"
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                        If Val(strKey) < 0.001 Then
                                MsgBox "֧�����������0.001,�����䣡"
                                Cancel = True
                                .TxtSetFocus
                                Exit Sub
                            End If
                            
                        If Val(strKey) >= 10 ^ 11 - 1 Then
                            MsgBox "������������С��" & (10 ^ 11 - 1)
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        .Text = strKey
                Else
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    Else
                        .TxtVisible = True
                        .Text = " "
                        Exit Sub
                    End If
                End If
            Case HeadCol.�ƻ�֧������
                
                If strKey <> "" Then
                    
                        If IsNumeric(strKey) Then
                            strKey = TranNumToDate(Val(strKey))
                        End If
                        
                        If Not IsDate(strKey) Then
                            ShowMsgbox "�ƻ�֧�����ڱ���Ϊ���ڸ�ʽ(��:20030303 �� 2003-03-03),�����䣡"
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        .Text = strKey
                Else
                    If .TxtVisible = True Then
                        .Text = ""
                        Exit Sub
                    Else
                        .TxtVisible = True
                        .Text = ""
                        Exit Sub
                    End If
                End If
        End Select
    End With

End Sub

Private Sub mshList_LeaveCell(Row As Long, Col As Long)
    Dim ingLoop As Integer
    Dim curTemp As Currency
    Dim strTemp As String, strDate As String
    
    If Col = 1 Then
        On Error Resume Next
        txtInfo(4).Text = 0
        For ingLoop = 1 To mshList.Rows - 1
            curTemp = CDbl(Format(mshList.TextMatrix(ingLoop, 1), "0.00"))
            If Err.Number = 0 Then
                mshList.TextMatrix(ingLoop, 1) = Format(mshList.TextMatrix(ingLoop, 1), "0.00")
                txtInfo(4).Text = Format(CDbl(txtInfo(4).Text) + curTemp, "0.00")
            End If
            Err.Clear
        Next
        txtInfo(4).Text = Format(CDbl(txtInfo(3).Text) - CDbl(txtInfo(4).Text), "0.00")
        On Error GoTo 0
    End If
    If Col = 0 Or Col = 4 Then
        On Error Resume Next
        mshList.TextMatrix(Row, Col) = Format(mshList.TextMatrix(Row, Col), "yyyy-MM-dd")
        On Error GoTo 0
    End If
End Sub

