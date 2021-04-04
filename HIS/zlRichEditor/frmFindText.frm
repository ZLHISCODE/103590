VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFindText 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   2250
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6510
   Icon            =   "frmFindText.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   1650
      Left            =   120
      ScaleHeight     =   1650
      ScaleWidth      =   6150
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   435
      Width           =   6150
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "�ر�(&C)"
         Height          =   350
         Left            =   4665
         TabIndex        =   9
         Top             =   1290
         Width           =   1455
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "������һ��(&F)"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   350
         Left            =   3165
         TabIndex        =   8
         Top             =   1290
         Width           =   1455
      End
      Begin VB.CommandButton cmdReplAll 
         Caption         =   "ȫ���滻(&A)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   1665
         TabIndex        =   7
         Top             =   1290
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdReplace 
         Caption         =   "�滻(&R)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   165
         TabIndex        =   6
         Top             =   1290
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chkMatchWord 
         Caption         =   "ȫ��ƥ��(&W)"
         Height          =   210
         Left            =   3270
         TabIndex        =   3
         Top             =   510
         Width           =   1635
      End
      Begin VB.CheckBox chkMatchCase 
         Caption         =   "��Сдƥ��(&U)"
         Height          =   210
         Left            =   1245
         TabIndex        =   2
         Top             =   510
         Width           =   1635
      End
      Begin VB.ComboBox cboReplace 
         Height          =   300
         Left            =   1245
         TabIndex        =   5
         Top             =   825
         Visible         =   0   'False
         Width           =   4875
      End
      Begin VB.ComboBox cboFind 
         Height          =   300
         Left            =   1245
         TabIndex        =   1
         Top             =   90
         Width           =   4875
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   195
         Picture         =   "frmFindText.frx":000C
         Top             =   495
         Width           =   240
      End
      Begin VB.Label lblReplace 
         AutoSize        =   -1  'True
         Caption         =   "�滻Ϊ(&I)"
         Height          =   180
         Left            =   195
         TabIndex        =   4
         Top             =   900
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "��������(&N)"
         Height          =   180
         Left            =   195
         TabIndex        =   0
         Top             =   165
         Width           =   990
      End
   End
   Begin MSComctlLib.TabStrip tspFunc 
      Height          =   2205
      Left            =   15
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   30
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   3889
      TabWidthStyle   =   2
      TabFixedWidth   =   1940
      TabFixedHeight  =   616
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����(&D)"
            Key             =   "����"
            Object.Tag             =   "����"
            Object.ToolTipText     =   "���Ҷ�λ�ĵ��е��ı�"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�滻(&H)"
            Key             =   "�滻"
            Object.Tag             =   "�滻"
            Object.ToolTipText     =   "���Ҳ��滻�ĵ��е��ı�"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFindText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objEditor As Editor
Dim intCount As Integer, intFlags As Integer, blnNewFind As Boolean

Private Function GetShowTop() As Long
    '���ݻ�ȡ���ҵ����ı��ģ�����ʵ��ı����嶥��λ�ã��Ա㲻���ǲ��ҵ��ַ�
    Dim lLeft As Long, lTOp As Long, lRight As Long
    Dim pt As POINTAPI, lFormTop As Long
    pt.X = 0
    pt.Y = 0
    ClientToScreen objEditor.Hwnd, pt
    '��ȡ��ʼλ������
    objEditor.Range(objEditor.SelStart, objEditor.SelStart + 1).GetPoint cprGPStart + cprGPLeft + cprGPBottom, lLeft, lTOp
    
    GetShowTop = pt.Y * Screen.TwipsPerPixelY + lTOp
    If GetShowTop < 0 Then GetShowTop = 1000
    If GetShowTop > Screen.Height - Me.Height Then GetShowTop = Screen.Height / 2 - Me.Height
End Function

Private Sub AddItemList(cboSelf As ComboBox)
    '����ǰ��������ּ��뵽�����б���
    Dim blnExist As Boolean
    If cboSelf.Text = "" Then Exit Sub
    blnExist = False
    For intCount = 0 To cboSelf.ListCount - 1
        If cboSelf.List(intCount) = cboSelf.Text Then blnExist = True: Exit For
    Next
    If blnExist = False Then Call cboSelf.AddItem(cboSelf.Text, 0)
End Sub

Private Sub cboFind_Change()
    Me.cmdFind.Enabled = (Me.cboFind.Text <> "")
    Me.cmdReplace.Enabled = (Me.cboFind.Text <> "")
    Me.cmdReplAll.Enabled = (Me.cboFind.Text <> "")
    blnNewFind = True
End Sub

Private Sub cboFind_Click()
    Call cboFind_Change
End Sub

Private Sub chkMatchCase_Click()
    blnNewFind = True
End Sub

Private Sub chkMatchWord_Click()
    blnNewFind = True
End Sub

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Public Sub FindNext(Editor As Editor)
    Set objEditor = Editor
    If Me.cboFind.Text = "" Then Me.cboFind.Text = objEditor.SelText
    cmdFind_Click
End Sub

Private Sub cmdFind_Click()
    Dim bFirst As Boolean
    Call AddItemList(Me.cboFind)
    
    objEditor.InProcessing = True
    intFlags = 0
    If Me.chkMatchCase.Value = vbChecked Then intFlags = intFlags + tomMatchCase
    If Me.chkMatchWord.Value = vbChecked Then intFlags = intFlags + tomMatchWord
    
ReFind:
    If objEditor.FindText(Me.cboFind.Text, intFlags) = True Then
        If objEditor.Selection.Font.Hidden Then GoTo ReFind
        Dim Position As POINTAPI
        ClientToScreen objEditor.Hwnd, Position
        Me.Move Me.Left, GetShowTop
        Call Form_Resize
        blnNewFind = False
        objEditor.InProcessing = False
        Exit Sub
    End If
    If Me.cboFind.Text = "" Then objEditor.InProcessing = False: Exit Sub
    If blnNewFind Then
        Call MsgBox("ָ�����ҵ����ݲ����ڣ�", vbExclamation, Me.Caption)
        If Me.cboFind.Visible And Me.cboFind.Enabled Then Me.cboFind.SetFocus
        objEditor.InProcessing = False
        Exit Sub
    End If
    If MsgBox("�ѵ����βλ�ã��Ƿ�ӿ�ʼλ�����²��ң�", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
        objEditor.InProcessing = False
        Exit Sub
    End If
    blnNewFind = True
    objEditor.SelStart = 1
    GoTo ReFind
End Sub

Private Sub cmdReplace_Click()
    Dim lngStart As Long
    Call AddItemList(Me.cboFind)
    Call AddItemList(Me.cboReplace)

    intFlags = 0
    If Me.chkMatchCase.Value = vbChecked Then intFlags = intFlags + tomMatchCase
    If Me.chkMatchWord.Value = vbChecked Then intFlags = intFlags + tomMatchWord
    
    If Len(objEditor.SelText) = Len(Me.cboFind.Text) Then
        lngStart = objEditor.SelStart
        objEditor.SelStart = lngStart: objEditor.SelLength = 0
        If objEditor.FindText(Me.cboFind.Text, intFlags) = True Then
            If lngStart = objEditor.SelStart And Len(objEditor.SelText) = Len(Me.cboFind.Text) Then
                objEditor.SelText = Me.cboReplace.Text
            Else
                objEditor.SelStart = lngStart: objEditor.SelLength = 0
            End If
        Else
            objEditor.SelStart = lngStart: objEditor.SelLength = 0
        End If
    End If
ReFind:
    If objEditor.FindText(Me.cboFind.Text, intFlags) = True Then
        If objEditor.Selection.Font.Protected <> False Or objEditor.Selection.Font.Hidden <> False Then GoTo ReFind
        Dim Position As POINTAPI
        ClientToScreen objEditor.Hwnd, Position
        Me.Move Me.Left, GetShowTop
        Call Form_Resize
        blnNewFind = False: Exit Sub
    End If
    If blnNewFind Then Call MsgBox("ָ�����ҵ����ݲ����ڣ�", vbExclamation, Me.Caption): Me.cboFind.SetFocus: Exit Sub
    If MsgBox("�ѵ����βλ�ã��Ƿ�ӿ�ʼλ�����²��ң�", vbQuestion + vbYesNo, Me.Caption) = vbNo Then Exit Sub
    blnNewFind = True
    objEditor.SelStart = 1
    GoTo ReFind

End Sub

Private Sub cmdReplAll_Click()
    Dim lngTimes As Long
    Call AddItemList(Me.cboFind)
    Call AddItemList(Me.cboReplace)
    
    intFlags = 0
    If Me.chkMatchCase.Value = vbChecked Then intFlags = intFlags + tomMatchCase
    If Me.chkMatchWord.Value = vbChecked Then intFlags = intFlags + tomMatchWord
    
    If objEditor.Range(0, 1).Font.Hidden Then
        objEditor.SelStart = 1
    Else
        objEditor.SelStart = 0
    End If
    lngTimes = 0
    Do
ReFind:
        If objEditor.FindText(Me.cboFind.Text, intFlags) = False Then Exit Do
        If objEditor.Selection.Font.Protected = False And objEditor.Selection.Font.Hidden = False Then
            objEditor.SelText = Me.cboReplace.Text
        Else
            GoTo ReFind
        End If
        lngTimes = lngTimes + 1
    Loop
    If lngTimes = 0 Then
        Call MsgBox("ָ�����ҵ����ݲ����ڣ�", vbExclamation, Me.Caption)
    Else
        Call MsgBox("�������������ִ��" & lngTimes & "���滻��", vbExclamation, Me.Caption)
    End If
    Me.cboFind.SetFocus
End Sub

Private Sub Form_Activate()
    If objEditor.SelLength <> 0 Then
        Dim Position As POINTAPI
        ClientToScreen objEditor.Hwnd, Position
        Me.Move Me.Left, GetShowTop
        Call Form_Resize
        blnNewFind = False: Exit Sub
    End If
End Sub

Private Sub Form_Resize()
    If Me.Top < 0 Then Me.Top = 1000
    If Me.Top > Screen.Height - Me.Height Then Me.Top = Screen.Height / 2 - Me.Height
    If Me.Left > Screen.Width Then Me.Left = Screen.Width - Me.Width
End Sub

Private Sub tspFunc_Click()
    Me.lblReplace.Visible = (Me.tspFunc.SelectedItem.Key = "�滻")
    Me.cboReplace.Visible = (Me.tspFunc.SelectedItem.Key = "�滻")
    Me.cmdReplace.Visible = (Me.tspFunc.SelectedItem.Key = "�滻")
    Me.cmdReplAll.Visible = (Me.tspFunc.SelectedItem.Key = "�滻")
    Me.Caption = Me.tspFunc.SelectedItem.Key
    If Me.Visible Then Me.cboFind.SetFocus
End Sub

Public Function ShowMe(Editor As Editor, Optional intShowWhat As Integer) As Boolean
    '���ܣ���ʾ�����滻�Ի���ִ�в����滻���滻ʱ�����Ա��������ص����ݽ����滻
    '������
    '   Editor,Ҫ�����滻���ĵ��༭������
    '   intShowWhat,��ʾ�ͽ�ֹ�Ĺ���:
    '    0,������ʾ���Ҵ���
    '    1,������ʾ�滻����
    '   -1,��ʾ���Ҵ��������滻����
    
    Dim i As Long, strFind As String, lS As Long, lE As Long
    If Editor.AuditMode  Or Editor.ReadOnly Then intShowWhat = -1
    
    Set objEditor = Editor
    lS = objEditor.Selection.StartPos
    lE = objEditor.Selection.EndPos
    lE = IIf(lE > lS + 100, lS + 100, lE)
    For i = lS To lE - 1
        If objEditor.Range(i, i + 1).Font.Hidden = False Then
            strFind = strFind & objEditor.Range(i, i + 1)
        End If
    Next
    Me.cboFind.Text = strFind
    If Me.cboFind.Text <> "" Then blnNewFind = False
    If intShowWhat = 1 Then
        Me.tspFunc.Tabs("�滻").Selected = True
        Call tspFunc_Click
    ElseIf intShowWhat = -1 Then
        Call Me.tspFunc.Tabs.Remove("�滻")
    End If
    Me.Show vbModal, objEditor.Parent
    
    If Me.cboFind.ListCount = 0 Then
        ShowMe = False
    Else
        ShowMe = True
    End If
    Unload Me
End Function

