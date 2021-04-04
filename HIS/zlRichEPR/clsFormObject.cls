VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents txt           As VB.TextBox
Attribute txt.VB_VarHelpID = -1
Public WithEvents Cmb           As VB.ComboBox
Attribute Cmb.VB_VarHelpID = -1
Public WithEvents lst           As VB.ListBox
Attribute lst.VB_VarHelpID = -1
Public WithEvents cmd           As VB.CommandButton
Attribute cmd.VB_VarHelpID = -1
Public WithEvents opt           As VB.OptionButton
Attribute opt.VB_VarHelpID = -1
Public WithEvents chk           As VB.CheckBox
Attribute chk.VB_VarHelpID = -1
Public WithEvents frm           As VB.Form
Attribute frm.VB_VarHelpID = -1

Private Sub Cmb_DblClick()
On Error GoTo ErrH
    Clipboard.SetText Cmb.Text
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

'========================================================================================
'=�����б���õ�����ʱѡ������
'========================================================================================
Private Sub Cmb_GotFocus()
On Error GoTo ErrH
    If Not Cmb.Locked Then Cmb.BackColor = &HC0FFC0
    With Cmb
        .Text = Trim(.Text)
        .SelStart = 0
        .SelLength = Len(Trim(.Text))
    End With
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

'========================================================================================
'=�����б����ı��򰴼��ƶ�
'========================================================================================
Private Sub Cmb_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrH
    If KeyCode = 13 Then
        KeyEnter KeyCode
    End If
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

'========================================================================================
'=�����б���ʧ����ʱ�������
'========================================================================================
Private Sub Cmb_LostFocus()
On Error GoTo ErrH
    If Not Cmb.Locked Then Cmb.BackColor = vbWhite
    If Cmb.STYLE <> 2 Then
        Cmb.Text = ChkStrUniCode(Cmb.Text, Val(Cmb.Tag))
    End If
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub Chk_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrH
    KeyEnter KeyCode
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub Cmd_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrH
    KeyUpDown KeyCode
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub Frm_Unload(Cancel As Integer)
    Dim mCtl                As Control
    Dim i                   As Integer
On Error Resume Next
    For Each mCtl In frm.Controls
        Select Case TypeName(mCtl)
            
            Case "Recordset"
                Set mCtl = Nothing
            Case "VSFlexGrid"
                SaveFlexState mCtl, frm.Caption
        End Select
    Next
    DoEvents
End Sub

Private Sub lst_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrH
    KeyUpDownEnter KeyCode
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub opt_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrH
    KeyEnter KeyCode
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub txt_Change()
    Dim lngStart        As Long
    Dim lngLength       As Long
On Error GoTo ErrH
    If InStr(1, txt.Name, "_NotCheck") = 0 Then
        lngLength = Len(txt.Text)
        lngStart = txt.SelStart
        
        txt.Text = ConvertString(txt.Text)
        txt.SelStart = lngStart - (lngLength - Len(txt.Text))
    End If
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

'========================================================================================
'=�����ı��򰴼��ƶ�
'========================================================================================
Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrH
    With txt
        If .ScrollBars = 2 Then
            KeyEnter KeyCode
        Else
            KeyUpDownEnter KeyCode
        End If
    End With
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
On Error GoTo ErrH
    If InStr(1, txt.Name, "_NotCheck") = 0 Then
        If InStr(1, "��'&", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

'========================================================================================
'=�ı���ʧ����ʱ�������
'========================================================================================
Private Sub txt_LostFocus()
On Error GoTo ErrH
    If Not txt.Locked Then txt.BackColor = vbWhite '�ı䱳����ɫ
    With txt
        If (txt.Locked) Then Exit Sub
         
        If .ScrollBars = 2 Then
            If InStr(1, .Name, "_NotCheck") = 0 Then .Text = ChkStrUniCode(.Text, .MaxLength)
        Else
            If InStr(1, .Name, "_NotCheck") = 0 Then .Text = Trim(ChkStrUniCode(.Text, .MaxLength))
        End If
        If .Tag = "MAX" Then
            .Text = UCase(.Text)
        ElseIf .Tag = "MIN" Then
            .Text = LCase(.Text)
        ElseIf .Tag <> "" Then
            .Text = Format(.Text, .Tag)
        End If
    End With
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

'========================================================================================
'=�ı���õ�����ʱѡ������
'========================================================================================
Private Sub txt_GotFocus()
On Error GoTo ErrH
    With txt
        If Not txt.Locked Then .BackColor = &HC0FFC0
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub Txt_DblClick()
On Error GoTo ErrH
    '���ı�д��ճ����
    Clipboard.SetText txt.Text
    If txt.ScrollBars = 2 Then
        txt.Text = Big_Note(txt.Text, txt.ToolTipText & " �༭����", txt.Locked, InStr(1, txt.Name, "_NotCheck") > 0)
    End If
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

'��ⳤ���Ƿ񳬹�����(�ֽ���)
Private Function ChkStrUniCode(mStr As String, mLen As Long) As String
    Dim strL        As String
On Error GoTo ErrH
    mStr = ConvertString(mStr)
    If mLen <= 0 Then
        ChkStrUniCode = mStr
        Exit Function
    Else
        strL = StrConv(mStr, vbFromUnicode)
        strL = LeftB(strL, mLen)
        ChkStrUniCode = StrConv(strL, vbUnicode)
    End If
    Exit Function
ErrH:
    Err.Clear
    ChkStrUniCode = ""
    Exit Function
End Function

'�ϡ��¡��س���
Private Sub KeyUpDownEnter(KeyCode As Integer)
On Error GoTo ErrH
    If KeyCode = 13 Then 'Enter
        SendKeys "{Tab}"
    ElseIf KeyCode = 40 Then 'Down
        SendKeys "{Tab}"
    ElseIf KeyCode = 38 Then 'Up
        SendKeys "+{Tab}"
    End If
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

'�ϡ��¼�
Private Sub KeyUpDown(KeyCode As Integer)
On Error GoTo ErrH
    If KeyCode = 40 Then 'Down
        SendKeys "{Tab}"
    ElseIf KeyCode = 38 Then 'Up
        SendKeys "+{Tab}"
    End If
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

'�س���
Private Sub KeyEnter(KeyCode As Integer)
On Error GoTo ErrH
    If KeyCode = 13 Then 'Enter
        SendKeys "{Tab}"
    End If
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub