VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.3#0"; "ZL9BillEdit.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSet���� 
   AutoRedraw      =   -1  'True
   Caption         =   " "
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7650
   Icon            =   "frmSet����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   7650
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab tabSel 
      Height          =   4155
      Left            =   60
      TabIndex        =   3
      Top             =   735
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7329
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmSet����.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "mshBill"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "��ѡ����"
      TabPicture(1)   =   "frmSet����.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkQ"
      Tab(1).Control(1)=   "chkBig"
      Tab(1).Control(2)=   "chkʵʱ(0)"
      Tab(1).Control(3)=   "chkʵʱ(1)"
      Tab(1).Control(4)=   "chkSingleD"
      Tab(1).Control(5)=   "chkSingleK"
      Tab(1).Control(6)=   "chk������"
      Tab(1).Control(7)=   "txtEdit"
      Tab(1).Control(8)=   "lblEdit(3)"
      Tab(1).ControlCount=   9
      Begin VB.CheckBox chkQ 
         Caption         =   "��������ʹ��סԺ����(&4)"
         Height          =   195
         Left            =   -72150
         TabIndex        =   13
         Top             =   1965
         Width           =   2415
      End
      Begin VB.CheckBox chkBig 
         Caption         =   "�����ʹ��סԺ����(&3)"
         Height          =   195
         Left            =   -74640
         TabIndex        =   12
         Top             =   1965
         Width           =   2415
      End
      Begin VB.CheckBox chkʵʱ 
         Caption         =   "������ϸʱʵ�ϴ�(&M)"
         Height          =   285
         Index           =   0
         Left            =   -74640
         TabIndex        =   11
         Top             =   1050
         Width           =   2085
      End
      Begin VB.CheckBox chkʵʱ 
         Caption         =   "סԺ��ϸʱʵ�ϴ�(&Z)"
         Height          =   285
         Index           =   1
         Left            =   -72150
         TabIndex        =   10
         Top             =   1050
         Width           =   2085
      End
      Begin VB.CheckBox chkSingleD 
         Caption         =   "�����г�Ժ��������ʾ(&1)"
         Height          =   240
         Left            =   -74640
         TabIndex        =   9
         Top             =   1530
         Width           =   2475
      End
      Begin VB.CheckBox chkSingleK 
         Caption         =   "��������Ժ��������ʾ(&2)"
         Height          =   240
         Left            =   -72150
         TabIndex        =   8
         Top             =   1530
         Width           =   2475
      End
      Begin VB.CheckBox chk������ 
         Caption         =   "������(&K)"
         Height          =   255
         Left            =   -72150
         TabIndex        =   6
         Top             =   645
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   -73590
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "1"
         Top             =   600
         Width           =   360
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   3645
         Left            =   60
         TabIndex        =   4
         Top             =   405
         Width           =   7425
         _ExtentX        =   13097
         _ExtentY        =   6429
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
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
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ǰ����(&D)     �Ŵ���"
         Height          =   180
         Index           =   3
         Left            =   -74640
         TabIndex        =   7
         Top             =   660
         Width           =   1980
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6435
      TabIndex        =   2
      Top             =   5025
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5175
      TabIndex        =   1
      Top             =   5025
      Width           =   1100
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   60
      Picture         =   "frmSet����.frx":0044
      Top             =   150
      Width           =   480
   End
   Begin VB.Label lbl 
      Caption         =   "�����豸�Ĵ��ںż����ô����Ƿ�Ĭ��Ϊ������,�������շ���������ص�ҽ����Ŀ���Ӧ"
      Height          =   315
      Left            =   540
      TabIndex        =   0
      Top             =   285
      Width           =   7125
   End
End
Attribute VB_Name = "frmSet����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnReturn As Boolean
Private mlngҽ������ As Long
Private mlng���� As Long
Private Enum mColHead
    �շ���� = 0
    ������Ŀ
    ������Ŀ
End Enum
Private Sub chk������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbTab
    End If
End Sub


Private Sub chkʵʱ_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnReturn = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngRow As Long
    
    If Trim(txtEdit) = "" Then Exit Sub
    SaveRegInFor g����ģ��, "����", "�˿ں�", Me.txtEdit
    SaveRegInFor g����ģ��, "����", "������", Me.chk������.Value
    If Val(txtEdit) = 0 Then
        gintComPort_���� = 1
    Else
        gintComPort_���� = Val(txtEdit)
    End If
    gblnKFQCom_���� = IIf(chk������.Value = 1, True, False)
    gintComPort = txtEdit.Text
        
    'ɾ���Ѿ�����
    gcnOracle.BeginTrans
    On Error GoTo ErrHand
    
    gstrSQL = "zl_���ղ���_Delete(" & mlng���� & ",NUll)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    With mshBill
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, mColHead.�շ����) <> "" Then
                '������������
                gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",null,'" & .TextMatrix(lngRow, mColHead.�շ����) & "' ,'" & .TextMatrix(lngRow, mColHead.������Ŀ) & ";" & .TextMatrix(lngRow, mColHead.������Ŀ) & "'," & lngRow + 2 & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
        Next
    End With
    '����
    
    gstrSQL = "zl_���ղ���_Update(" & mlng���� & ",NULL,'������ϸʱʵ�ϴ�' ,'" & IIf(chkʵʱ(0).Value = 1, "1", "0") & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)

    gstrSQL = "zl_���ղ���_Update(" & mlng���� & ",NULL,'סԺ��ϸʱʵ�ϴ�' ,'" & IIf(chkʵʱ(1).Value = 1, "1", "0") & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)

    If mlng���� = 82 Then
        gstrSQL = "zl_���ղ���_Update(" & mlng���� & ",NULL,'�����ֳ�Ժ��ʾ' ,'" & IIf(chkSingleD.Value = 1, "1", "0") & "')"
    Else
        gstrSQL = "zl_���ղ���_Update(" & mlng���� & ",NULL,'�����ֳ�Ժ��ʾ' ,'" & IIf(chkSingleK.Value = 1, "1", "0") & "')"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)

    gstrSQL = "zl_���ղ���_Update(" & mlng���� & ",NULL,'�����ʹ��סԺ����' ,'" & IIf(chkBig.Value = 1, "1", "0") & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    If mlng���� = 82 Then
        gstrSQL = "zl_���ղ���_Update(" & mlng���� & ",NULL,'��������ʹ��סԺ����' ,'" & IIf(chkQ.Value = 1, "1", "0") & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    
    gcnOracle.CommitTrans
    mblnReturn = True
    Unload Me
    Exit Sub
ErrHand:
    gcnOracle.RollbackTrans
    Resume
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Dim strReg As String
    mblnReturn = False
    
    Call GetRegInFor(g����ģ��, "����", "�˿ں�", strReg)
    If Val(strReg) = 0 Then
        txtEdit.Text = 1
    Else
        txtEdit.Text = Val(strReg)
    End If
    
    Call GetRegInFor(g����ģ��, "����", "������", strReg)
    If Val(strReg) = 1 Then
        Me.chk������.Value = 1
    Else
        Me.chk������.Value = 0
    End If
    
    Call GetRegInFor(g����ģ��, "����", "�����е�����", strReg)
    If Val(strReg) = 1 Then
        Me.chkSingleD.Value = 1
    Else
        Me.chkSingleD.Value = 0
    End If
    
    Call GetRegInFor(g����ģ��, "����", "������������", strReg)
    If Val(strReg) = 1 Then
        Me.chkSingleK.Value = 1
    Else
        Me.chkSingleK.Value = 0
    End If
    
    Call GetRegInFor(g����ģ��, "����", "�����ʹ��סԺ����", strReg)
    If Val(strReg) = 1 Then
        Me.chkBig.Value = 1
    Else
        Me.chkBig.Value = 0
    End If
    
    If mlng���� = 82 Then
        Call GetRegInFor(g����ģ��, "����", "��������ʹ��סԺ����", strReg)
        If Val(strReg) = 1 Then
            Me.chkQ.Value = 1
        Else
            Me.chkQ.Value = 0
        End If
    End If
    RestoreWinState Me, App.ProductName
    
    '��ʼ����
    Call iniData
End Sub

Public Function ShowMe(ByVal lng���� As Long, ByVal lngҽ������ As Long) As Boolean
    mlngҽ������ = lngҽ������
    mlng���� = lng����
    
    Me.Show 1
    ShowMe = mblnReturn
End Function

Private Sub Form_Resize()
    On Error Resume Next
    
    With cmdCancel
        .Top = ScaleHeight - .Height - 100
        .Left = ScaleWidth - .Width - 50
    End With
    With cmdOK
        .Top = cmdCancel.Top
        .Left = cmdCancel.Left - 50 - .Width
    End With
    
    With tabSel
        .Width = ScaleWidth - 50
        .Height = cmdOK.Top - .Top - 100
    End With
    
    With mshBill
        .Top = tabSel.Top - 300
        .Left = tabSel.Left + 50
        .Height = cmdOK.Top - 1400
        .Width = tabSel.Width - 200
    End With
    mshBill.ZOrder
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub mshBill_EnterCell(Row As Long, Col As Long)
    With mshBill
        Select Case Col
            Case mColHead.������Ŀ
                mshBill.Clear
                mshBill.AddItem "����"
                mshBill.AddItem "��ҩ��"
                mshBill.AddItem "��ҩ��"
                mshBill.AddItem "��ҩ��"
                mshBill.AddItem "����"
                mshBill.AddItem "����"
                mshBill.AddItem "���Ʒ�"
                mshBill.AddItem "�������Ʒ�"
                mshBill.AddItem "Ѫ��"
                
            Case mColHead.������Ŀ
                mshBill.Clear
                mshBill.AddItem "A�в�ҩ��"
                mshBill.AddItem "B�г�ҩ��"
                mshBill.AddItem "C��ҩ��"
                mshBill.AddItem "D����"
                mshBill.AddItem "E������"
                mshBill.AddItem "F�����"
                mshBill.AddItem "G������"
                mshBill.AddItem "H�����"
                mshBill.AddItem "I���Ʒ�"
                mshBill.AddItem "J�����"
                mshBill.AddItem "K��λ��"
                mshBill.AddItem "L�����"
                mshBill.AddItem "X��Ѫ��"
                mshBill.AddItem "M��������"
        End Select
    End With
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtEdit, KeyAscii, m����ʽ
End Sub

Private Function iniData() As Boolean
    '��ʼ����
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim lngRow As Long
    Dim strTmp As String
    
    '����ҳͷ
    Err = 0
    On Error Resume Next
    strSQL = "Select * from ��������Ŀ¼ where ����=" & mlng����
    zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption
    If rsTmp.EOF Then
        tabSel.Caption = "��"
    Else
        tabSel.Caption = Nvl(rsTmp!����)
    End If
    rsTmp.Close
  
    If mlng���� = type_���������� Then
        Me.chk������.Value = 1
    Else
        Me.chk������.Value = 0
    End If
    
    '���ñ���ͷ
    Call initGrid
    strSQL = "" & _
        "   Select A.���,b.����ֵ From �շ���� a,(Select * From ���ղ��� where ����=" & mlng���� & ") b " & _
        "   Where A.���=b.������(+) " & _
        "   order by A.���� "
    zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption
    With mshBill
        .ClearBill
        If rsTmp.RecordCount = 0 Then
            .Rows = 2
        Else
            .Rows = rsTmp.RecordCount + 1
        End If
        lngRow = 1
        Do While Not rsTmp.EOF
            .TextMatrix(lngRow, mColHead.�շ����) = Nvl(rsTmp!���)
            strTmp = Nvl(rsTmp!����ֵ)
            If InStr(1, strTmp, ";") <> 0 Then
                .TextMatrix(lngRow, mColHead.������Ŀ) = Split(strTmp, ";")(0)
                .TextMatrix(lngRow, mColHead.������Ŀ) = Split(strTmp, ";")(1)
            End If
            lngRow = lngRow + 1
            rsTmp.MoveNext
        Loop
        
        strSQL = "Select ������,����ֵ From ���ղ��� " & _
                " Where ������ in('������ϸʱʵ�ϴ�','סԺ��ϸʱʵ�ϴ�','ҽ����ϸʱʵ�ϴ�'," & _
                "'�����ֳ�Ժ��ʾ','�����ʹ��סԺ����','��������ʹ��סԺ����') and ����=" & mlng����
        zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption
        chkʵʱ(0).Value = 1
        chkʵʱ(1).Value = 1
        chkSingleD.Value = 1
        chkSingleK.Value = 1
        chkBig.Value = 1
        chkQ.Value = 0
        chkQ.Visible = False
        Do While Not rsTmp.EOF
            Select Case Nvl(rsTmp!������)
            Case "������ϸʱʵ�ϴ�"
                chkʵʱ(0).Value = IIf(Val(Nvl(rsTmp!����ֵ)) = 1, 1, 0)
            Case "סԺ��ϸʱʵ�ϴ�"
                chkʵʱ(1).Value = IIf(Val(Nvl(rsTmp!����ֵ)) = 1, 1, 0)
            Case "�����ֳ�Ժ��ʾ"
                If mlng���� = 82 Then
                    chkSingleD.Value = IIf(Val(Nvl(rsTmp!����ֵ)) = 1, 1, 0)
                Else
                    chkSingleK.Value = IIf(Val(Nvl(rsTmp!����ֵ)) = 1, 1, 0)
                End If
            Case "�����ʹ��סԺ����"
                chkBig.Value = IIf(Val(Nvl(rsTmp!����ֵ)) = 1, 1, 0)
            Case "��������ʹ��סԺ����"
                If mlng���� = 82 Then
                    chkQ.Visible = True
                    chkQ.Value = IIf(Val(Nvl(rsTmp!����ֵ)) = 1, 1, 0)
                Else
                    chkQ.Visible = False
                End If
            End Select
            rsTmp.MoveNext
        Loop
    End With
End Function
Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = 3
        
        .msfObj.FixedCols = 1
        .AllowAddRow = False
        
        .TextMatrix(0, mColHead.�շ����) = "�շ����"
        .TextMatrix(0, mColHead.������Ŀ) = "������Ŀ"
        .TextMatrix(0, mColHead.������Ŀ) = "������Ŀ"
        
        
        .ColWidth(mColHead.�շ����) = 1500
        .ColWidth(mColHead.������Ŀ) = 2000
        .ColWidth(mColHead.������Ŀ) = 2000
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��

        .ColData(mColHead.�շ����) = 5
        .ColData(mColHead.������Ŀ) = 3
        .ColData(mColHead.������Ŀ) = 3
        
        .ColAlignment(mColHead.�շ����) = flexAlignLeftCenter
        .ColAlignment(mColHead.������Ŀ) = flexAlignLeftCenter
        .ColAlignment(mColHead.������Ŀ) = flexAlignLeftCenter
        .PrimaryCol = mColHead.������Ŀ
        .LocateCol = mColHead.������Ŀ
    End With
End Sub



