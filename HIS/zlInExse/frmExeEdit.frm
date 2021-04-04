VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmExeEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ִ�еǼ�"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   Icon            =   "frmExeEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdSample 
      Caption         =   "����(&M)"
      Height          =   350
      Left            =   1440
      TabIndex        =   5
      Top             =   2370
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2910
      TabIndex        =   3
      Top             =   2370
      Width           =   1100
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4050
      TabIndex        =   4
      Top             =   2370
      Width           =   1100
   End
   Begin VB.CommandButton cmdSetup 
      Caption         =   "Ʊ������(&S)"
      Height          =   350
      Left            =   105
      TabIndex        =   6
      Top             =   2370
      Width           =   1305
   End
   Begin MSComCtl2.DTPicker dtpִ��ʱ�� 
      Height          =   315
      Left            =   3255
      TabIndex        =   2
      Top             =   1830
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   107216899
      CurrentDate     =   37447
   End
   Begin VB.ComboBox cboִ���� 
      Height          =   300
      Left            =   735
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1830
      Width           =   1395
   End
   Begin VB.TextBox txt���� 
      Height          =   1440
      Left            =   135
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   300
      Width           =   5160
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -105
      TabIndex        =   10
      Top             =   2130
      Width           =   5760
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ִ��ʱ��"
      Height          =   180
      Left            =   2505
      TabIndex        =   9
      Top             =   1890
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ִ����"
      Height          =   180
      Left            =   150
      TabIndex        =   8
      Top             =   1890
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ִ�����:"
      Height          =   180
      Left            =   165
      TabIndex        =   7
      Top             =   75
      Width           =   810
   End
   Begin VB.Menu mnuSample 
      Caption         =   "����(&S)"
      Visible         =   0   'False
      Begin VB.Menu mnuSampleAdd 
         Caption         =   "���浱ǰ����(&S)"
      End
      Begin VB.Menu mnuSampleDel 
         Caption         =   "ɾ�����н���(&D)"
         Begin VB.Menu mnuSampleItemDel 
            Caption         =   "<�޽�������>"
            Index           =   0
         End
      End
      Begin VB.Menu mnuSample_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSampleItem 
         Caption         =   "<�޽�������>"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmExeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
'��
Public mlngDeptID As Long '��ǰִ�п���
Public mblnView As Boolean
Public mstrDate As String
'��/��
Public mstrOper As String '�����ִ��,��Ϊִ����
Public mstrLog As String '��ǰ��¼����
'��
Public mvDate As Date 'ִ��ʱ��

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSql As String
        
    If Not mblnView Then
        If Not zlcontrol.TxtCheckInput(txt����, "����", , True) Then Exit Sub
        
        If cboִ����.ListIndex = -1 Then
            MsgBox "��ȷ��ִ�еǼ��ˣ�", vbInformation, gstrSysName
            cboִ����.SetFocus: Exit Sub
        End If
        
        mstrLog = txt����.Text
        mstrOper = zlStr.NeedName(cboִ����.Text)
        mvDate = dtpִ��ʱ��.Value
        
        gblnOK = True
    End If
    
    Unload Me
End Sub

Private Sub cmdSample_Click()
    If LoadSample Then
        PopupMenu mnuSample, 2, cmdSample.Left, cmdSample.Top + cmdSample.Height, mnuSampleAdd
    End If
End Sub

Private Sub cmdSetup_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1142", Me)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyCode = vbKeyEscape Then
        If mblnView Then Unload Me
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    gblnOK = False
        
    If Not mblnView Then
        Call InitOper
        txt����.Text = mstrLog
        '�µǼǻ��޸Ķ��Ե�ǰ��Ա���,��ǰʱ��ִ�еǼ�
        dtpִ��ʱ��.Value = zlDatabase.Currentdate
        cboִ����.Enabled = Not gbln����ִ��
        cboִ����.ListIndex = cbo.FindIndex(cboִ����, UserInfo.ID)
        If cboִ����.ListIndex = -1 And Not cboִ����.Enabled Then
            MsgBox "�㲻���ڵ�ǰִ�п��ң����㲻����������Ա��ݵǼ�ִ�������", vbInformation, gstrSysName
            Unload Me: Exit Sub
        ElseIf cboִ����.ListIndex = -1 Then
            cboִ����.ListIndex = cbo.FindIndex(cboִ����, mstrOper, True)
        End If
    Else
        Caption = "�鿴�Ǽ�"
        
        txt����.Text = mstrLog
        dtpִ��ʱ��.Value = Format(mstrDate, "yyyy-MM-dd hh:mm:ss")
        cboִ����.AddItem mstrOper
        cboִ����.ListIndex = cboִ����.NewIndex
        
        cmdSetup.Visible = False
        cmdSample.Visible = False
        cmdCancel.Visible = False
        dtpִ��ʱ��.Enabled = False
        cboִ����.Enabled = False
        txt����.Locked = True
        
        cmdOK.Left = cmdOK.Left + cmdOK.Width / 2
    End If
End Sub

Private Sub InitOper()
    Dim strSql As String, i As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select Distinct A.ID,A.���,A.����,A.����" & _
        " From ��Ա�� A,������Ա B" & _
        " Where A.ID=B.��ԱID And B.����ID=[1] And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
        " Order by A.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngDeptID)
    
    For i = 1 To rsTmp.RecordCount
        cboִ����.AddItem rsTmp!���� & "-" & rsTmp!����
        cboִ����.ItemData(cboִ����.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Next
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnView = False
End Sub

Private Sub mnuSampleAdd_Click()
    Dim i As Integer, intCount As Integer
    
    If Trim(txt����.Text) = "" Then
        MsgBox "�������������ݣ�", vbInformation, gstrSysName
        txt����.SetFocus
        Exit Sub
    End If
    
    For i = 0 To mnuSampleItem.UBound
        If mnuSampleItem(i).Tag = txt����.Text Then
            MsgBox "�ý����Ѿ�����Ϊ�����䣡", vbInformation, gstrSysName
            txt����.SetFocus
            Exit Sub
        End If
    Next
        
    intCount = 0
    If mnuSampleItem(0).Caption <> "<�޽�������>" Then
        intCount = mnuSampleItem.UBound + 1
    End If
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��������", "Count", intCount + 1)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��������", "Item" & intCount, txt����.Text)
End Sub

Private Sub mnuSampleItem_Click(Index As Integer)
    If mnuSampleItem(Index).Caption = "<�޽�������>" Then Exit Sub
    
    txt����.Text = mnuSampleItem(Index).Tag
    txt����.SetFocus
End Sub

Private Sub mnuSampleItemDel_Click(Index As Integer)
    Dim i As Integer, intCount As Integer
    Dim intDel As Integer, strText As String
    
    If mnuSampleItem(Index).Caption = "<�޽�������>" Then Exit Sub
    
    intCount = 0: intDel = Index
    If mnuSampleItem(0).Caption <> "<�޽�������>" Then
        intCount = mnuSampleItem.UBound + 1
    End If
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��������", "Count", intCount - 1)
    
    For i = intDel To intCount - 2
        strText = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��������", "Item" & i + 1, "")
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��������", "Item" & i, strText)
    Next
    
    Call DeleteSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��������", "Item" & intCount - 1)
End Sub

Private Sub txt����_GotFocus()
    zlcontrol.TxtSelAll txt����
End Sub

Private Function LoadSample() As Boolean
    Dim intCount As Integer
    Dim i As Integer
    Dim objMenu As Object
    
    '�����������
    For Each objMenu In mnuSampleItem
        objMenu.Tag = ""
        If objMenu.Index <> 0 Then
            Unload objMenu
        Else
            objMenu.Caption = "<�޽�������>"
        End If
    Next
    For Each objMenu In mnuSampleItemDel
        objMenu.Tag = ""
        If objMenu.Index <> 0 Then
            Unload objMenu
        Else
            objMenu.Caption = "<�޽�������>"
        End If
    Next
    
    intCount = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��������", "Count", 0)
    For i = 0 To intCount - 1
        If i <> 0 Then
            Load mnuSampleItem(i)
            Load mnuSampleItemDel(i)
        End If
        mnuSampleItem(i).Visible = True
        mnuSampleItemDel(i).Visible = True
        mnuSampleItem(i).Tag = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��������", "Item" & i, "")
        mnuSampleItemDel(i).Tag = mnuSampleItem(i).Tag
        
        If zlCommFun.ActualLen(mnuSampleItem(i).Tag) > 20 Then
            mnuSampleItem(i).Caption = Left(mnuSampleItem(i).Tag, 20) & " ..."
        Else
            mnuSampleItem(i).Caption = mnuSampleItem(i).Tag
        End If
        mnuSampleItemDel(i).Caption = mnuSampleItem(i).Caption
    Next
    LoadSample = True
End Function

Private Sub txt����_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And mblnView Then
        lngTXTProc = GetWindowLong(txt����.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt����.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt����_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And mblnView Then
        Call SetWindowLong(txt����.hWnd, GWL_WNDPROC, lngTXTProc)
    End If
End Sub

