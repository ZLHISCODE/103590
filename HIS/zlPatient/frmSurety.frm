VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSurety 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������Ϣ����"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   Icon            =   "frmSurety.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7560
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�(&X)"
      Height          =   350
      Left            =   6120
      TabIndex        =   9
      ToolTipText     =   "(F9)�˳�"
      Top             =   1320
      Width           =   1100
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "ɾ��(&D)"
      Height          =   350
      Left            =   2570
      TabIndex        =   8
      ToolTipText     =   "ֻ����ɾ�����һ��������¼"
      Top             =   1320
      Width           =   1100
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "�޸�(&M)"
      Height          =   350
      Left            =   1470
      TabIndex        =   7
      ToolTipText     =   "ֻ�����޸����һ��������¼"
      Top             =   1320
      Width           =   1100
   End
   Begin VB.CommandButton cmdAdd 
      Cancel          =   -1  'True
      Caption         =   "����(&A)"
      Height          =   350
      Left            =   360
      TabIndex        =   6
      ToolTipText     =   "�������һ��������¼���ڻ�û����������ʱ����������"
      Top             =   1320
      Width           =   1100
   End
   Begin VB.Frame fraEdit 
      Caption         =   "��Ϣ����"
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   7335
      Begin VB.TextBox txtReason 
         Height          =   300
         Left            =   5040
         MaxLength       =   50
         TabIndex        =   5
         Top             =   720
         Width           =   2010
      End
      Begin VB.CheckBox chk��ʱ���� 
         Caption         =   "��ʱ����"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   720
         Width           =   1050
      End
      Begin VB.CheckBox chkUnlimit 
         Caption         =   "���޶��"
         Height          =   255
         Left            =   2760
         TabIndex        =   4
         ToolTipText     =   "���޵�����ʱ�������õ���ʱ��"
         Top             =   720
         Width           =   1050
      End
      Begin VB.TextBox txt������ 
         Height          =   300
         Left            =   840
         MaxLength       =   100
         TabIndex        =   0
         Top             =   360
         Width           =   1005
      End
      Begin VB.TextBox txt������ 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   9
         TabIndex        =   1
         Top             =   360
         Width           =   1005
      End
      Begin MSComCtl2.DTPicker dtp����ʱ�� 
         Height          =   300
         Left            =   5040
         TabIndex        =   2
         Top             =   360
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   275709955
         CurrentDate     =   38915.6041666667
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ԭ��"
         Height          =   180
         Left            =   4140
         TabIndex        =   16
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lbl����ʱ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         Height          =   180
         Left            =   4140
         TabIndex        =   15
         ToolTipText     =   "��Ժ���˲���ʹ��ʱ�޵���"
         Top             =   450
         Width           =   720
      End
      Begin VB.Label lbl������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   240
         TabIndex        =   13
         Top             =   450
         Width           =   540
      End
      Begin VB.Label lbl������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   2160
         TabIndex        =   12
         Top             =   450
         Width           =   540
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh 
      Height          =   2265
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   3995
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   -2147483645
      FixedCols       =   0
      RowHeightMin    =   250
      BackColorBkg    =   16777215
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   14
      Top             =   4080
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9472
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3775
            MinWidth        =   3775
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSurety"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mlng����ID As Long
Public mbln��Ժ���� As Boolean
Public mstrPrivs As String
Private mlng��ҳID As Long      '���ﲡ�˻��Ժ����Ϊ0,��Ժ����Ϊ��ǰסԺ�Ǽǵ���ҳID

Private Sub chkUnlimit_Click()
     '���޵����������ʱ����
    If chkUnlimit.Value = 1 And IsNull(dtp����ʱ��.Value) Then
        dtp����ʱ��.Value = DateAdd("d", 3, dtp����ʱ��.MinDate)
    End If
    chk��ʱ����.Enabled = Not (chkUnlimit.Value = 1)
    txt������.Enabled = Not (chkUnlimit.Value = 1)
    
    If chkUnlimit.Value = 1 Then
        txt������.Text = "999999999":  txt������.BackColor = vbInactiveCaptionText
    Else
        txt������.Text = "": txt������.BackColor = vbWhite
    End If
End Sub

Private Sub chk��ʱ����_Click()
    If chk��ʱ����.Value = 1 Then
        dtp����ʱ��.CheckBox = True: dtp����ʱ��.CustomFormat = "yyyy-MM-dd HH:mm"
        dtp����ʱ��.Value = Null
        chkUnlimit.Value = 0        'ֵ�ı�ʱ����ʽ����click�¼�
    End If
    chkUnlimit.Enabled = Not (chk��ʱ����.Value = 1) And mbln��Ժ����
    dtp����ʱ��.Enabled = Not (chk��ʱ����.Value = 1) And mbln��Ժ����
End Sub

Private Sub cmdDel_Click()
    Dim strSQL As String
    Dim str�Ǽ�ʱ�� As String
    Dim strɾ����־ As String
    
    '����21368 by lesfeng 2010-08-02
    strɾ����־ = Trim(msh.TextMatrix(msh.Row, GetColNum("ɾ����־")))
    If strɾ����־ = "ɾ��" Then
        MsgBox "����������¼�Ѿ�Ϊɾ����ǣ����ܽ���ɾ����ǲ�����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("ȷʵҪ���б�Ǵ���������¼Ϊɾ�������?" & vbCrLf & vbCrLf & "ע��,ɾ����Ǻ󣬵�ǰ�������᲻�ָܻ�!" _
        , vbYesNo + vbDefaultButton2 + vbInformation, gstrSysName) = vbNo Then Exit Sub
    
    On Error GoTo errH
    
    If Trim(msh.TextMatrix(msh.Row, GetColNum("�Ǽ�ʱ��"))) = "" Then
        str�Ǽ�ʱ�� = "NULL"
    Else
        str�Ǽ�ʱ�� = zlStr.To_Date(Trim(msh.TextMatrix(msh.Row, GetColNum("�Ǽ�ʱ��"))))
    End If
    '����21368 by lesfeng 2010-08-02
    strSQL = "zl_���˵�����¼_delete(" & mlng����ID & "," & mlng��ҳID & ",NULL," & str�Ǽ�ʱ�� & ",'" & UserInfo.��� & "','" & UserInfo.���� & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    stbThis.Panels(1).Text = "ɾ�������ɹ�!"
    Call LoadSurety
    
    If cmdExit.Enabled Then cmdExit.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdModify_Click()
    Dim strSQL As String, str������ As String, str����ʱ�� As String
    Dim str�Ǽ�ʱ�� As String
    Dim strɾ����־ As String
    'ֻ���޸ĵ�ǰѡ�в�����Ч�ĵ�����¼
    
    
    If cmdModify.Caption = "�޸�(&M)" Then
    '��ȡ�޸���Ϣ
        If msh.TextMatrix(msh.Row, GetColNum("������")) = "" Then
            stbThis.Panels(1).Text = "û�п����޸ĵĵ�����Ϣ!"
            Exit Sub
        End If
        '����21368 by lesfeng 2010-08-02
        strɾ����־ = Trim(msh.TextMatrix(msh.Row, GetColNum("ɾ����־")))
        If strɾ����־ = "ɾ��" Then
            MsgBox "����������¼�Ѿ�Ϊɾ����ǣ����ܽ����޸Ĳ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        cmdModify.Caption = "����(&S)"
        cmdAdd.Enabled = False
        cmdDel.Enabled = False
        cmdExit.Caption = "ȡ��(&C)"
        fraEdit.Enabled = True
        
        With msh
            txt������.Text = Trim(.TextMatrix(.Row, GetColNum("������")))
            If .TextMatrix(.Row, GetColNum("������")) = "����" Then
                chkUnlimit.Value = 1    'ֵ��ͬʱ��ʽ����click�¼�
                txt������.Text = "999999999"
            Else
                chkUnlimit.Value = 0
                txt������.Text = Val(.TextMatrix(.Row, GetColNum("������")))
            End If
            
            If IsDate(.TextMatrix(.Row, GetColNum("����ʱ��"))) Then
                dtp����ʱ��.CheckBox = True: dtp����ʱ��.CustomFormat = "yyyy-MM-dd HH:mm"
                dtp����ʱ��.Value = CDate(.TextMatrix(.Row, GetColNum("����ʱ��")))
            Else
                dtp����ʱ��.CheckBox = True: dtp����ʱ��.CustomFormat = "yyyy-MM-dd HH:mm" '������ɼ��������ִ�л����
                dtp����ʱ��.Value = Null
            End If
            
            chk��ʱ����.Value = IIf(.TextMatrix(.Row, GetColNum("��ʱ����")) = "��", 1, 0)
            If txt������.Enabled Then txt������.SetFocus
            txt������.Tag = Trim(.TextMatrix(msh.Row, GetColNum("�Ǽ�ʱ��")))
        End With
    Else
    '�����޸Ľ��
        '1.���ݼ��
        If Not Check������Ϣ Then Exit Sub
        
        
        '�Ȼָ����水ť״̬
        cmdModify.Caption = "�޸�(&M)"
        cmdAdd.Enabled = True
        cmdDel.Enabled = True
        cmdExit.Caption = "�˳�(&X)"
        fraEdit.Enabled = True      'SetCanEdit���ٴ�����
        
        str������ = Replace(Trim(txt������.Text), "'", "''")
        str����ʱ�� = "null"
        If Not IsNull(dtp����ʱ��.Value) Then str����ʱ�� = zlStr.To_Date(dtp����ʱ��.Value)
        str�Ǽ�ʱ�� = zlStr.To_Date(txt������.Tag)
        
        '���ȼ��
        If Not CheckLen(txt������, 64) Then Exit Sub
        
        '2.���ݱ���
        On Error GoTo errH
        strSQL = "zl_���˵�����¼_update(" & mlng����ID & "," & mlng��ҳID & ",'" & str������ & "'," & _
            Val(txt������.Text) & "," & chk��ʱ����.Value & ",'" & Trim(txtReason.Text) & "',NULL," & str����ʱ�� & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & str�Ǽ�ʱ�� & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                
        '3.����ˢ��
        stbThis.Panels(1).Text = "�޸Ľ���ѱ���!"
        Call LoadSurety
        Call Init������Ϣ
        If cmdExit.Enabled Then cmdExit.SetFocus
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Init������Ϣ()
    Dim datsys As Date

    txt������.Text = ""
    chkUnlimit.Enabled = mbln��Ժ����
    chkUnlimit.Value = 0            '���ֵ�б仯,����ʽ����click�¼�
    txt������.Text = ""
    txtReason.Text = ""
    
    dtp����ʱ��.Enabled = mbln��Ժ����
    dtp����ʱ��.CheckBox = True: dtp����ʱ��.CustomFormat = "yyyy-MM-dd HH:mm" '����checkbox�ɼ���
    If dtp����ʱ��.Enabled Then
        datsys = zlDatabase.Currentdate
        dtp����ʱ��.MinDate = datsys
        dtp����ʱ��.Value = DateAdd("d", 3, datsys)
    End If
    dtp����ʱ��.Value = Null
    
    chk��ʱ����.Enabled = True
    chk��ʱ����.Value = 0
    chkUnlimit.TabStop = True
End Sub

Private Sub dtp����ʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
    ElseIf KeyAscii = vbKeySpace Then
        If dtp����ʱ��.CheckBox Then
            KeyAscii = 0
            If IsNull(dtp����ʱ��.Value) Then
                dtp����ʱ��.Value = DateAdd("d", 3, zlDatabase.Currentdate)
            Else
                dtp����ʱ��.Value = Null
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
        
    Dim strSQL  As String
    Dim rsTmp As New ADODB.Recordset
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call LoadSurety
    Call Init������Ϣ
    
    'Call GetSuretyBalance   '��ʼmlng��ҳid
    
    If InStr(mstrPrivs, "������Ϣ����") <= 0 Then
        cmdAdd.Visible = False
    End If
    
    If InStr(mstrPrivs, "������Ϣ����") <= 0 Then
        cmdModify.Visible = False
    End If
    
    If InStr(mstrPrivs, "������Ϣɾ��") <= 0 Then
        cmdDel.Visible = False
    End If
    
    If InStr(mstrPrivs, "������Ϣ����") <= 0 And InStr(mstrPrivs, "������Ϣ����") And InStr(mstrPrivs, "������Ϣɾ��") <= 0 Then
        fraEdit.Enabled = False
        Me.Caption = "������Ϣ�鿴(��ǰ�û���" & UserInfo.���� & ")"
    End If
    
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Integer
    For i = 0 To msh.Cols - 1
        If msh.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
    GetColNum = -1
End Function

Private Sub SetHeader()
    Dim strHead As String, i As Long
    '����21368 by lesfeng 2010-08-02
    strHead = ",4,300|���,4,1000|������,4,800|������,7,1250|��ʱ����,4,850|����ԭ��,4,1800|�Ǽ�ʱ��,1,1800|����ʱ��,1,1800|ɾ����־,4,850|����Ա����,4,1050|����Ա���,4,1050|ɾ������Ա����,4,1050|ɾ������Ա���,4,1050|ɾ��ʱ��,1,1800"
    With msh
        .Redraw = False
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(msh, App.ProductName & "\" & Me.Name)
        
        .ForeColor = &H80000003
        .RowHeight(0) = 320
        .Redraw = True
    End With
End Sub

Private Sub GetSuretyBalance()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = _
        " Select To_char(������,'99999999990.00') as ������,Decode(��ǰ����ID,null,0,��ҳID) as ��ҳID" & _
        " From ������Ϣ Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    If rsTmp.RecordCount > 0 Then
        stbThis.Panels(2).Text = "��Ч������:" & IIf(IsNull(rsTmp!������), "��", Val(Trim("" & rsTmp!������)))
        mlng��ҳID = Val("" & rsTmp!��ҳID)
    Else
        stbThis.Panels(2).Text = "��Ч������:��"
        mlng��ҳID = 0
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadSurety()
    Dim rsTmp As ADODB.Recordset, Curdate As Date
    Dim strSQL As String, i As Integer, lngRow As Integer, RowPageid As Integer
    Dim strɾ����־ As String
    
    On Error GoTo errH
    Curdate = zlDatabase.Currentdate
    '����21368 by lesfeng 2010-08-02
    'ɾ����־,4,850|����Ա����,4,1050|����Ա���,4,1050|ɾ������Ա����,4,1050|ɾ������Ա���,4,1050|ɾ��ʱ��,1,1800"
    strSQL = _
        "SELECT '',Decode(��ҳid, NULL, '����', '��' || ��ҳid || '��סԺ') ���, ������," & vbNewLine & _
        "       Decode(������, 999999999, '����', To_Char(������, '999999990.00')) AS ������," & vbNewLine & _
        "       Decode(��������, 1, '��', ' ') AS ��ʱ����, ����ԭ��, To_Char(�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') �Ǽ�ʱ��," & vbNewLine & _
        "       To_Char(����ʱ��, 'yyyy-mm-dd hh24:mi:ss') ����ʱ��,decode(ɾ����־,1,'',-1,'ɾ��','') as ɾ����־," & vbNewLine & _
        "       ����Ա����,����Ա���,ɾ������Ա����,ɾ������Ա���,ɾ��ʱ��" & vbNewLine & _
        "FROM ���˵�����¼" & vbNewLine & _
        "WHERE ����id = [1]" & vbNewLine & _
        "ORDER BY �Ǽ�ʱ�� DESC"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    If rsTmp.RecordCount > 0 Then
        Set msh.DataSource = rsTmp
    Else
        msh.Clear
        msh.Rows = 2
    End If
    Call SetHeader
    GetSuretyBalance
    For lngRow = 1 To msh.Rows - 1
        If UBound(Split(Trim(msh.TextMatrix(lngRow, GetColNum("���"))), "��סԺ")) > 0 Then 'ȡ��ѡ������ҳID
            RowPageid = Val(Split(Split(Trim(msh.TextMatrix(lngRow, GetColNum("���"))), "��סԺ")(0), "��")(1))
        Else
            RowPageid = 0
        End If
        '����21368 by lesfeng 2010-08-02
        strɾ����־ = Trim(msh.TextMatrix(lngRow, GetColNum("ɾ����־")))
        
        If mlng��ҳID = RowPageid And (Trim(msh.TextMatrix(lngRow, GetColNum("����ʱ��"))) = "" Or Trim(msh.TextMatrix(lngRow, GetColNum("����ʱ��"))) > Curdate) Then
            msh.Row = lngRow
            For i = 0 To msh.Cols - 1
                msh.Col = i
                '����21368 by lesfeng 2010-08-02
                If strɾ����־ = "" Then
                    msh.CellForeColor = &HC00000
                Else
                    msh.CellForeColor = &HFF&
                End If
            Next
        Else
             For i = 0 To msh.Cols - 1
                msh.Col = i
                '����21368 by lesfeng 2010-08-02
                If strɾ����־ = "" Then
                Else
                    msh.CellForeColor = &HFF&
                End If
            Next
        End If
        
    Next lngRow
    msh.Row = 1
    msh.Col = 0: msh.ColSel = msh.Cols - 1
    Call msh_EnterCell
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Check������Ϣ() As Boolean
    Check������Ϣ = True
        
    If Trim(txt������.Text) = "" Then
        stbThis.Panels(1).Text = "�����뵣��������,�����˲���Ϊ��!"
        If txt������.Enabled Then txt������.SetFocus
        Check������Ϣ = False
        Exit Function
    End If
    
    If Not IsNumeric(txt������.Text) Then
        stbThis.Panels(1).Text = "��������ȷ�ĵ�����,������Ҫ������ֵ!"
        If txt������.Enabled Then txt������.SetFocus
        Check������Ϣ = False
        Exit Function
    ElseIf Val(txt������.Text) = 0 Then
        stbThis.Panels(1).Text = "�����뵣����,�������Ϊ��!"
        If txt������.Enabled Then txt������.SetFocus
        Check������Ϣ = False
        Exit Function
    End If
    
    If chk��ʱ����.Value = 1 Then
        If Not IsNull(dtp����ʱ��.Value) Or chkUnlimit.Value = 1 Then
            stbThis.Panels(1).Text = "��ʱ�������������õ���ʱ�޻��޵�����!"
            If chk��ʱ����.Enabled Then chk��ʱ����.SetFocus
            Check������Ϣ = False
            Exit Function
        End If
    End If
    
    If zlCommFun.ActualLen(Trim(txtReason.Text)) > 50 Then
        stbThis.Panels(1).Text = "����ԭ�������������� 25 �����ֻ� 50 ���ַ���"
        txtReason.SetFocus
        Check������Ϣ = False
        Exit Function
    End If
    
End Function

Private Sub cmdAdd_Click()
    Dim str������ As String, str����ʱ�� As String
    Dim strSQL As String, i As Integer, Curdate As Date, blnδ���� As Boolean, bln��ʱ As Boolean, RowPageid As Integer
    Dim strɾ����־ As String
    
    '1.���ݼ��
    If Not Check������Ϣ Then Exit Sub
    
    Curdate = zlDatabase.Currentdate
    
    For i = 1 To msh.Rows - 1 '�жϱ���סԺδ���ڵĵ�����¼��������ʾ
         If Trim(msh.TextMatrix(i, GetColNum("���"))) <> "" Then
            If UBound(Split(Trim(msh.TextMatrix(i, GetColNum("���"))), "��סԺ")) > 0 Then 'ȡ��ѡ������ҳID
                RowPageid = Val(Split(Split(Trim(msh.TextMatrix(i, GetColNum("���"))), "��סԺ")(0), "��")(1))
            Else
                RowPageid = 0
            End If
            If mlng��ҳID = RowPageid Then
                '����21368 by lesfeng 2010-08-02
                strɾ����־ = Trim(msh.TextMatrix(i, GetColNum("ɾ����־")))
               If (Trim(Nvl(msh.TextMatrix(i, GetColNum("����ʱ��")))) = "" Or Nvl(msh.TextMatrix(i, GetColNum("����ʱ��"))) > Curdate) And strɾ����־ = "" Then
                   bln��ʱ = Nvl(msh.TextMatrix(i, GetColNum("��ʱ����"))) = "��"
                   blnδ���� = True: Exit For
               End If
            End If
        End If
    Next
    
    If blnδ���� Then
        If MsgBox("����δ���ڵ�" & IIf(bln��ʱ, "��ʱ", "") & "������¼����������" & IIf(bln��ʱ, "��֮ǰ����ʱ�����Զ�ʧЧ", "�ۼƵ���") & "���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
        
    str������ = Replace(Trim(txt������.Text), "'", "''")
    str����ʱ�� = "null"
    If Not IsNull(dtp����ʱ��.Value) Then str����ʱ�� = "To_Date('" & Format(dtp����ʱ��.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    '���ȼ��
    If Not CheckLen(txt������, 64) Then Exit Sub
    
    '2.���ݱ���
    On Error GoTo errH
    
    strSQL = "zl_���˵�����¼_insert(" & mlng����ID & "," & mlng��ҳID & ",'" & str������ & "'," & _
        Val(txt������.Text) & "," & chk��ʱ����.Value & ",'" & Trim(txtReason.Text) & "',Null," & str����ʱ�� & ",'" & UserInfo.��� & "','" & UserInfo.���� & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    '3.����ˢ��
    stbThis.Panels(1).Text = "������Ϣ�ѱ���!"
    Call LoadSurety
    Call Init������Ϣ
    
    If cmdExit.Enabled Then cmdExit.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdExit_Click()
    
    If cmdExit.Caption = "ȡ��(&C)" Then
        cmdModify.Caption = "�޸�(&M)"
        cmdAdd.Enabled = True
        cmdDel.Enabled = True
        cmdExit.Caption = "�˳�(&X)"
        fraEdit.Enabled = True      'SetCanEdit���ٴ�����
       
        'ˢ������,���ǲ�������
        stbThis.Panels(1).Text = ""
        Call LoadSurety
        Call Init������Ϣ
    Else
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF9 Then
        Call cmdExit_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdModify.Caption = "����(&S)" Then
        If MsgBox("��ǰ�޸ĵ���Ϣδ����,ȷʵҪ�˳���?", vbYesNo + vbDefaultButton2 + vbInformation, gstrSysName) = vbNo Then Cancel = 1
    End If
    Call SaveWinState(Me, App.ProductName)
End Sub
Private Sub msh_EnterCell()
    Dim str����ʱ�� As String
    Dim datsys As Date, RowPageid As Integer
    Dim strɾ����־ As String
   '��ǰ����ҳ�벡����ҳ��ͬʱ�������޸�ɾ��,�ѹ��ڲ������޸�ɾ��
    datsys = zlDatabase.Currentdate
    
    '����21368 by lesfeng 2010-08-02
    strɾ����־ = Trim(msh.TextMatrix(msh.Row, GetColNum("ɾ����־")))
    
    If cmdModify.Caption = "�޸�(&M)" Then
        If mlng��ҳID = 0 And Trim(msh.TextMatrix(msh.Row, GetColNum("���"))) = "����" Then
            '����21368 by lesfeng 2010-08-02
            If strɾ����־ = "" Then
                cmdModify.Enabled = True
                cmdDel.Enabled = True
                stbThis.Panels(1).Text = "��ǰ������¼��Ч"
            Else
                cmdModify.Enabled = False
                cmdDel.Enabled = False
                stbThis.Panels(1).Text = "��ǰ������¼�Ѿ����ɾ��"
            End If
        Else
            If UBound(Split(Trim(msh.TextMatrix(msh.Row, GetColNum("���"))), "��סԺ")) > 0 Then 'ȡ��ѡ������ҳID
                RowPageid = Val(Split(Split(Trim(msh.TextMatrix(msh.Row, GetColNum("���"))), "��סԺ")(0), "��")(1))
            Else
                RowPageid = 0
            End If
            If mlng��ҳID <> RowPageid Then
                cmdModify.Enabled = False
                cmdDel.Enabled = False
                stbThis.Panels(1).Text = "��ǰ������¼�Ǳ���סԺ������"
            Else
                str����ʱ�� = Trim(msh.TextMatrix(msh.Row, GetColNum("����ʱ��")))
            
                If str����ʱ�� <> "" Then
                    If CDate(str����ʱ��) < datsys Then
                         cmdModify.Enabled = False
                         cmdDel.Enabled = False
                        '����21368 by lesfeng 2010-08-02
                         If strɾ����־ = "" Then
                            stbThis.Panels(1).Text = "��ǰ������¼�ѹ���"
                        Else
                            stbThis.Panels(1).Text = "��ǰ������¼�Ѿ����ɾ��"
                        End If
                    Else
                        '����21368 by lesfeng 2010-08-02
                        If strɾ����־ = "" Then
                            cmdModify.Enabled = True
                            cmdDel.Enabled = True
                            stbThis.Panels(1).Text = "��ǰ������¼��Ч"
                        Else
                            cmdModify.Enabled = False
                            cmdDel.Enabled = False
                            stbThis.Panels(1).Text = "��ǰ������¼�Ѿ����ɾ��"
                        End If
                    End If
                Else
                    '����21368 by lesfeng 2010-08-02
                    If strɾ����־ = "" Then
                        cmdModify.Enabled = True
                        cmdDel.Enabled = True
                        stbThis.Panels(1).Text = "��ǰ������¼��Ч"
                    Else
                        cmdModify.Enabled = False
                        cmdDel.Enabled = False
                        stbThis.Panels(1).Text = "��ǰ������¼�Ѿ����ɾ��"
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub txtReason_GotFocus()
    zlControl.TxtSelAll txtReason
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtReason_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
    Else
        If InStr("'|?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txtReason, KeyAscii
    End If
End Sub

Private Sub txtReason_LostFocus()
    If gstrIme <> "���Զ�����" Then Call OpenIme
End Sub

Private Sub txt������_GotFocus()
    zlControl.TxtSelAll txt������
End Sub

Private Sub txt������_KeyPress(KeyAscii As Integer)
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then
        If KeyAscii = vbKeyReturn Then
            chkUnlimit.TabStop = (txt������.Text = "")
            SendKeys "{Tab}"
        Else
            KeyAscii = 0
        End If
    ElseIf KeyAscii = Asc(".") And InStr(txt������.Text, ".") > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt������_LostFocus()
    If IsNumeric(txt������.Text) Then
        If txt������.Text = "999999999" Then
            stbThis.Panels(1).Text = "�����������ֵ����ֵ��ʾ���޵�����"
            If txt������.Enabled Then txt������.SetFocus
        Else
            txt������.Text = Format(txt������.Text, "0.00")
        End If
    Else
        txt������.Text = ""
    End If
    
    Call zlCommFun.OpenIme
End Sub

Private Sub txt������_GotFocus()
    zlControl.TxtSelAll txt������
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt������, KeyAscii
    End If
End Sub

Private Sub txt������_LostFocus()
    If gstrIme <> "���Զ�����" Then Call OpenIme
End Sub
