VERSION 5.00
Begin VB.Form frmMBSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ø��������"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboMachine 
      Height          =   300
      Left            =   1470
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Width           =   2505
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3435
      TabIndex        =   9
      Top             =   1950
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2265
      TabIndex        =   8
      Top             =   1950
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   -45
      TabIndex        =   10
      Top             =   1680
      Width           =   4785
   End
   Begin VB.ComboBox cboPosi 
      Height          =   300
      Left            =   1470
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1215
      Width           =   1215
   End
   Begin VB.TextBox txtNO 
      Height          =   300
      Left            =   1470
      MaxLength       =   20
      TabIndex        =   5
      Top             =   870
      Width           =   1185
   End
   Begin VB.TextBox txtItem 
      Height          =   300
      Left            =   1470
      TabIndex        =   3
      Top             =   525
      Width           =   2505
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "��������(&M)"
      Height          =   180
      Left            =   435
      TabIndex        =   0
      Top             =   255
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "��ʼλ��(&S)"
      Height          =   180
      Left            =   435
      TabIndex        =   6
      Top             =   1260
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��ʼ�걾��(&H)"
      Height          =   180
      Left            =   255
      TabIndex        =   4
      Top             =   930
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "������Ŀ(&I)"
      Height          =   180
      Left            =   435
      TabIndex        =   2
      Top             =   585
      Width           =   990
   End
End
Attribute VB_Name = "frmMBSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mstrItem As String
Public Function ShowMe(ByVal frmMain As Object) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, j As Integer
    Dim lngDeviceID As Long, strItem As String
    
    mblnOK = False
    mstrItem = ""
    
    On Error GoTo DBError
    
    '��������
    gstrSql = "Select * From ��������"
    OpenRecord rsTmp, gstrSql, Me.Caption
    If rsTmp.EOF Then
        MsgBox "û�г�ʼ�����������޷����ã�", vbCritical, Me.Caption
        Unload Me
        Exit Function
    End If
    
    With cboMachine
        .Clear
        Do While Not rsTmp.EOF
            .AddItem "(" & rsTmp("����") & ")" & rsTmp("����")
            .ItemData(.ListCount - 1) = rsTmp("ID")
            
            rsTmp.MoveNext
        Loop
    End With
    lngDeviceID = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "ø������", -1))
    If lngDeviceID = -1 Then
        cboMachine.ListIndex = 0
    Else
        cboMachine.ListIndex = FindComboItem(cboMachine, lngDeviceID)
    End If
    
    On Error Resume Next
    '������Ŀ
    strItem = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "ø������Ŀ", "")
    If Len(strItem) = 0 Then
        mstrItem = ""
        txtItem = "": txtItem.Tag = ""
    Else
        mstrItem = Split(strItem, "|")(0)
        txtItem = mstrItem: txtItem.Tag = Split(strItem, "|")(1)
    End If
    '�걾��
    txtNO = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "ø���Ǳ걾��", "")
    '��ʼλ��
    With cboPosi
        .Clear
        For i = 1 To 8
            For j = 1 To 12
                .AddItem Chr(64 + i) & Format(j, "0#")
            Next j
        Next i
    End With
    cboPosi.Text = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "ø������ʼλ��", "A01")
    
    
    Me.Show vbModal, frmMain
    
    ShowMe = mblnOK
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cboMachine_Click()
    mstrItem = "": txtItem = ""
End Sub

Private Sub cboMachine_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cboPosi_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Len(Trim(txtItem)) = 0 Then
        MsgBox "��ָ����ǰø���ǵļ�����Ŀ", , gstrSysName
        txtItem.SetFocus
        Exit Sub
    End If
    If Len(Trim(txtNO)) = 0 Then
        MsgBox "�������ʼ�ı걾��", , gstrSysName
        txtNO.SetFocus
        Exit Sub
    End If
    
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName, "ø������", cboMachine.ItemData(cboMachine.ListIndex))
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName, "ø������Ŀ", txtItem & "|" & txtItem.Tag)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName, "ø���Ǳ걾��", txtNO)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName, "ø������ʼλ��", cboPosi.Text)
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtItem_GotFocus()
    zlControl.TxtSelAll txtItem
End Sub

Private Sub txtItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtItem_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
    
    If mstrItem = txtItem Then Exit Sub
        
    strSQL = "SELECT Distinct ��ĿID As ID,ͨ������,������||'('||Ӣ����||')' As ������Ŀ " & _
        "FROM ����������Ŀ A,����������Ŀ B,������ĿĿ¼ C,������Ŀ���� D " & _
        "WHERE A.��Ŀid=B.ID And A.����ID=[1] AND B.����=C.���� AND C.ID=D.������ĿID " & _
        "AND (Upper(B.Ӣ����) LIKE [2] OR Upper(D.����) LIKE [2] OR Upper(B.������) LIKE [2])"
    
    On Error GoTo errH
    vRect = GetControlRect(txtItem.Hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "������Ŀ", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, txtItem.Height, blnCancel, False, True, cboMachine.ItemData(cboMachine.ListIndex), UCase(txtItem) & "%")
    If Not rsTmp Is Nothing Then
        txtItem.Text = rsTmp!������Ŀ
        txtItem.Tag = rsTmp!ͨ������
        mstrItem = txtItem
    Else
        If Not blnCancel Then
            MsgBox "δ�ҵ���Ӧ�ļ�����Ŀ��", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtNO_GotFocus()
    zlControl.TxtSelAll txtNO
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub
