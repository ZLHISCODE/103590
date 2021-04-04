VERSION 5.00
Begin VB.Form frmIdentify�ɶ� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ���������ʶ��"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIdentify�ɶ�.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtCard 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1965
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   675
      Width           =   3765
   End
   Begin VB.TextBox txtPwd 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1965
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1125
      Width           =   3765
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   405
      Left            =   2745
      TabIndex        =   2
      Top             =   2220
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   405
      Left            =   4200
      TabIndex        =   3
      Top             =   2220
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   75
      Left            =   -180
      TabIndex        =   4
      Top             =   2025
      Width           =   6660
   End
   Begin VB.Label lblNote 
      Caption         =   "������ȷˢ��֮������������롣"
      Height          =   255
      Left            =   900
      TabIndex        =   8
      Top             =   165
      Width           =   3645
   End
   Begin VB.Label lblCard 
      AutoSize        =   -1  'True
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1290
      TabIndex        =   7
      Top             =   735
      Width           =   510
   End
   Begin VB.Label lblPwd 
      AutoSize        =   -1  'True
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1290
      TabIndex        =   6
      Top             =   1185
      Width           =   510
   End
   Begin VB.Label lblPatiInfo 
      AutoSize        =   -1  'True
      Caption         =   "������Ϣ"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   195
      TabIndex        =   5
      Top             =   1740
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   255
      Picture         =   "frmIdentify�ɶ�.frx":030A
      Top             =   345
      Width           =   480
   End
End
Attribute VB_Name = "frmIdentify�ɶ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrPatiInfo As String
Public mcur��� As Currency
'200308z012:סԺ����ʹ��
Public mcurסԺ���� As Currency
Public mcurסԺ�޶� As Currency
Public mcur�������� As Currency

Private mstrҽ���� As String
Private mstr���� As String

Private mintTimes As Integer
Private mintCardLen As Integer

Private Sub cmdCancel_Click()
    mstrPatiInfo = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If mstrҽ���� = "" And mstr���� = "" Then
        MsgBox "δ��ȷ��ˢ��,����ͨ����֤��", vbInformation, gstrSysName
        Exit Sub
    End If
    Me.Hide
End Sub

Private Sub Form_Load()
    mstrPatiInfo = ""
    mcur��� = 0
    mcurסԺ���� = 0
    mcurסԺ�޶� = 0
    mcur�������� = 0
    
    mintTimes = 0
    Me.lblPatiInfo.Caption = ""
    mintCardLen = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("CardNoLength"), 26)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrҽ���� = "": mstr���� = ""
End Sub

Private Sub txtCard_GotFocus()
    zlControl.TxtSelAll txtCard
    If gblnLED And txtCard.Text = "" Then
        zl9LedVoice.Speak "#5"
    End If
End Sub

Private Sub txtCard_KeyPress(KeyAscii As Integer)
'���ܣ�ˢ���������������롢����
    Dim strҽ���� As String, str���� As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call ExecuteZ015(txtCard.Text, strҽ����, str����)
        If strҽ���� = "" And str���� = "" Then
            MsgBox "ˢ������ʧ�ܣ������ԣ�", vbInformation, gstrSysName
            txtCard.Text = "": txtCard.SetFocus: Exit Sub
        Else
            mstrҽ���� = strҽ����
            mstr���� = str����
            txtPwd.SetFocus: Exit Sub
        End If
    End If
    If txtCard.SelLength = Len(txtCard.Text) Then txtCard.Text = ""
    
    If Len(txtCard.Text) + 1 = mintCardLen Then
        txtCard.Text = txtCard.Text & Chr(KeyAscii)
        KeyAscii = 0
        Call ExecuteZ015(txtCard.Text, strҽ����, str����)
        If strҽ���� = "" And str���� = "" Then
            MsgBox "ˢ������ʧ�ܣ������ԣ�", vbInformation, gstrSysName
            txtCard.Text = "": txtCard.SetFocus: Exit Sub
        Else
            mstrҽ���� = strҽ����
            mstr���� = str����
            txtPwd.SetFocus: Exit Sub
        End If
    End If
    
    Me.cmdOK.Enabled = False
    Me.lblPatiInfo.Caption = ""
    Me.txtPwd.Text = ""
End Sub

Private Sub txtPwd_GotFocus()
    zlControl.TxtSelAll txtPwd
    If gblnLED And txtPwd.Text = "" Then
        zl9LedVoice.Speak "#0"
    End If
End Sub

Private Sub txtPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtPwd.Text = "" Then Exit Sub
        KeyAscii = 0: Call zlControl.TxtSelAll(txtPwd)
        If mstrҽ���� = "" And mstr���� = "" Then Exit Sub
        
        Call ThisIdentify: Exit Sub
    End If
    
    If txtPwd.SelLength = Len(txtPwd.Text) Then txtPwd.Text = ""
    
    If Len(txtPwd.Text) + 1 = txtPwd.MaxLength Then
        txtPwd.Text = txtPwd.Text & Chr(KeyAscii)
        KeyAscii = 0: Call zlControl.TxtSelAll(txtPwd)
        If mstrҽ���� = "" And mstr���� = "" Then Exit Sub
        
        Call ThisIdentify: Exit Sub
    End If
End Sub

Private Sub ThisIdentify()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    Dim strSelfNo As String, strSelfPwd As String, strSerial As String, strKH As String
    Dim strSwapNo As String         '����˳���
    
    strSelfNo = mstrҽ����
    strKH = mstr����
    strSelfPwd = TrimStr(txtPwd.Text)

    mintTimes = mintTimes + 1
    strSQL = "select ���ű�_id.nextval||'1' from dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    With rsTmp
        strSwapNo = .Fields(0).Value
        strSerial = getSerial(strSelfNo)
        
        'New:���ױ��,�ͻ������,����˳���,����,����Ա���,����ǼǺ�,ҽ����,ҽԺ����,����ʱ��,��������,֧�����,����
        strSQL = "z001('z001','" & UserInfo.վ�� & "','" & strSwapNo & "','" & strSelfPwd & "','" & UserInfo.��� & "'," & _
            "'" & strSerial & "','" & strSelfNo & "','" & Trim(gstrҽԺ����) & "','" & DateStr & "','" & strSwapNo & "','" & IIf(Me.Tag = 0, "11", "31") & "','" & strKH & "')"
        gcnSybase.Execute strSQL, , adCmdStoredProc
        
        If .State = adStateOpen Then .Close
        .Open "select code from zjycl  where jysxh='" & strSwapNo & "' and jybh='z001' order by jyend desc", gcnSybase, adOpenStatic, adLockReadOnly
        If Trim(.Fields(0).Value) <> "0000" Then
            MsgBox "����""z001""���ִ���""" & !CODE & """:" & vbCrLf & String(2, "��") & GetErrInfo(!CODE, TYPE_�ɶ���) & String(2, vbTab), vbInformation, gstrSysName
            If mintTimes > 6 Then
                MsgBox "�޷�ʶ�������ݣ���ȷ����Ŀ��������������", vbExclamation, gstrSysName
                mstrPatiInfo = "": Me.Hide: Exit Sub
            End If
            
            Me.lblNote.Caption = "�޷�ʶ����ݣ�������ˢ����"
            Me.txtPwd.Text = ""
            Me.cmdOK.Enabled = False
            Me.txtPwd.SetFocus
            mstrPatiInfo = ""
        Else
            strSQL = "select * from grjbxx where grbm='" & strSelfNo & "'"
            If .State = adStateOpen Then .Close
            .CursorLocation = adUseClient
            .Open strSQL, gcnSybase, adOpenKeyset
            If Not .EOF Then
                'New:0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����)
                mstrPatiInfo = strKH & ";" & strSelfNo & ";" & strSelfPwd & ";" & _
                        TrimStr(.Fields("xm").Value) & ";" & _
                        IIf(TrimStr(Nvl(.Fields("xb").Value)) = "1", "��", "Ů") & ";" & _
                        TrimStr(Nvl(.Fields("csrq").Value)) & ";" & _
                        TrimStr(Nvl(.Fields("sfz").Value)) & ";" & _
                        TrimStr(Nvl(.Fields("dwmc").Value)) & "(" & Trim(Nvl(.Fields("dwbm").Value)) & ")"
                mcur��� = IIf(IsNull(!grzhlnye), 0, !grzhlnye) + IIf(IsNull(!grzhbnye), 0, !grzhbnye)
                '200308z012:סԺ����ʹ��
                If Val(Me.Tag) <> 0 Then
                    mcurסԺ���� = IIf(IsNull(!zyjs), 0, !zyjs)
                    mcur�������� = IIf(IsNull(!tcbxbl), 0, !tcbxbl)
                    mcurסԺ�޶� = IIf(IsNull(!zyxe), 0, !zyxe)
                End If
                
                Me.lblNote.Caption = "�Ѿ���ȷ������ʶ��"
                Me.lblPatiInfo.Caption = "����:" & Trim(.Fields("xm").Value) & "  " & IIf(Trim(Nvl(.Fields("xb").Value)) = "1", "��", "Ů") & "  " & Trim(Nvl(.Fields("csrq").Value)) & ",��ȷ�ϣ�"
                
                '��������2005-10-14�� �����֤�ɹ��󣬽��������ʾ��
                If gblnLED Then
                   zl9LedVoice.Speak "#26 " & mcur���
                End If
                
                Me.cmdOK.Enabled = True
                Me.cmdOK.SetFocus
            Else
                Me.lblNote.Caption = "�޷�ʶ����ݣ�������ˢ����"
                Me.txtPwd.Text = ""
                Me.cmdOK.Enabled = False
                Me.txtPwd.SetFocus
                mstrPatiInfo = ""
            End If
        End If
    End With
End Sub
