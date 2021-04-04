VERSION 5.00
Object = "{CA73588D-282F-4592-9369-A61CC244FADA}#15.3#0"; "Codejock.SyntaxEdit.v15.3.1.ocx"
Begin VB.Form frmDocShiftTypeEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ҽ�����Ӱಡ������-����"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8505
   Icon            =   "frmDocShiftTypeEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   8505
   StartUpPosition =   1  '����������
   Begin XtremeSyntaxEdit.SyntaxEdit synSQL 
      Height          =   2895
      Left            =   960
      TabIndex        =   3
      Top             =   1200
      Width           =   7215
      _Version        =   983043
      _ExtentX        =   12726
      _ExtentY        =   5106
      _StockProps     =   84
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      EnableSyntaxColorization=   -1  'True
      ShowLineNumbers =   0   'False
      ShowSelectionMargin=   0   'False
      ShowScrollBarVert=   -1  'True
      ShowScrollBarHorz=   -1  'True
      EnableVirtualSpace=   0   'False
      EnableAutoIndent=   -1  'True
      ShowWhiteSpace  =   0   'False
      ShowCollapsibleNodes=   -1  'True
      AutoCompleteWndWidth=   160
      EnableEditAccelerators=   -1  'True
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "��֤(&C)"
      Height          =   350
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   1100
   End
   Begin VB.Frame fraLine2 
      Height          =   30
      Left            =   0
      TabIndex        =   12
      Top             =   5160
      Width           =   9375
   End
   Begin VB.Frame fraLine1 
      Height          =   30
      Left            =   0
      TabIndex        =   11
      Top             =   960
      Width           =   9375
   End
   Begin VB.TextBox txtBegin 
      Height          =   300
      Left            =   960
      TabIndex        =   2
      ToolTipText     =   "��ʹ�ñ���[ʱ���ʽ]"
      Top             =   525
      Width           =   4575
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox txtSName 
      Height          =   300
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5760
      TabIndex        =   5
      Top             =   5400
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7080
      TabIndex        =   7
      Top             =   5400
      Width           =   1100
   End
   Begin VB.Label lblxing 
      AutoSize        =   -1  'True
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   5640
      TabIndex        =   15
      Top             =   585
      Width           =   90
   End
   Begin VB.Label lblDescript 
      AutoSize        =   -1  'True
      Caption         =   "��ʹ��[��Ŀ����]����"
      Height          =   180
      Left            =   5760
      TabIndex        =   14
      Top             =   585
      Width           =   1800
   End
   Begin VB.Label lblExplain 
      Caption         =   $"frmDocShiftTypeEdit.frx":5C02
      Height          =   615
      Left            =   960
      TabIndex        =   13
      Top             =   4320
      Width           =   7215
   End
   Begin VB.Label lblSQL 
      AutoSize        =   -1  'True
      Caption         =   "��ȡSQL"
      Height          =   180
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label lblBegin 
      AutoSize        =   -1  'True
      Caption         =   "��ʼ����"
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   585
      Width           =   720
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   2640
      TabIndex        =   8
      Top             =   165
      Width           =   360
   End
   Begin VB.Label lblSName 
      AutoSize        =   -1  'True
      Caption         =   "���"
      Height          =   180
      Left            =   480
      TabIndex        =   6
      Top             =   165
      Width           =   360
   End
End
Attribute VB_Name = "frmDocShiftTypeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte
Private mstrSName As String
Private mblnOK As Boolean

Public Function ShowMe(ByVal bytType As Byte, ByRef strSName As String) As Boolean
'bytType:1-������2-�޸�
    
    mbytType = bytType
    mstrSName = strSName
    Me.Show 1
    If mblnOK Then strSName = mstrSName
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdCheck_Click()
    If CheckSQL Then
        MsgBox "��֤�ɹ���", vbInformation, Me.Caption
    End If
End Sub

Private Sub cmdOK_Click()
    Dim strSql As String
    
    If CheckData = False Then Exit Sub
    
    strSql = SynSQL.Text
    strSql = Replace(strSql, "'", "''")
    On Error GoTo errH
    gstrSql = "Zl_ҽ�����Ӱಡ������_Edit(" & IIf(mbytType = 1, 1, 2) & ",'" & txtSName.Text & "','" & _
        mstrSName & "','" & txtName.Text & "','" & txtBegin.Text & "','" & SynSQL.Text & "')"
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    
    mblnOK = True
    mstrSName = txtSName.Text
    Unload Me
    Exit Sub
errH:
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim rsTemp As ADODB.Recordset
    
    With SynSQL
        '���ÿؼ�����ʾ��ɫ����Ϊ��SQL
        .SyntaxSet "[Schemes]" & vbCrLf & "SQL" & vbCrLf & "[Themes]" & vbCrLf & "Default" & vbCrLf & "Alternative" & vbCrLf
        .SyntaxScheme = GetSqlColor
    End With
    mblnOK = False
    Select Case mbytType
        Case 1
            Me.Caption = "ҽ�����Ӱಡ������-����"
        Case 2
            Me.Caption = "ҽ�����Ӱಡ������-�޸�"
            txtSName.Text = mstrSName
            Set rsTemp = rsPatiType(mstrSName)
            If rsTemp.RecordCount = 1 Then
                txtName.Text = rsTemp!����
                txtBegin.Text = rsTemp!��ʼ���� & ""
                SynSQL.Text = rsTemp!��ȡSQL & ""
            End If
    End Select
End Sub

Private Function CheckData() As Boolean
'����ǰ�������
    
    If txtSName.Text = "" Then
        MsgBox "��Ʋ���Ϊ�գ����飡"
        Call zlcontrol.ControlSetFocus(txtSName)
        Exit Function
    ElseIf zlstr.ActualLen(txtSName.Text) > 10 Then
        MsgBox "��Ʋ��ܳ���5�����֣����飡"
        Call zlcontrol.ControlSetFocus(txtSName)
        Exit Function
    End If

    If txtName.Text = "" Then
        MsgBox "���Ʋ���Ϊ�գ����飡"
        Call zlcontrol.ControlSetFocus(txtName)
        Exit Function
    ElseIf zlstr.ActualLen(txtName.Text) > 20 Then
        MsgBox "���Ʋ��ܳ���10�����֣����飡"
        Call zlcontrol.ControlSetFocus(txtName)
        Exit Function
    End If
    
    If zlstr.ActualLen(txtBegin.Text) > 50 Then
        MsgBox "��ʼ�������ܳ���25�����֣����飡"
        Call zlcontrol.ControlSetFocus(txtBegin)
        Exit Function
    End If
    
    If Trim(SynSQL.Text) <> "" Then
        If CheckSQL = False Then Exit Function
    End If
    CheckData = True
End Function

Private Function CheckSQL() As Boolean
'У��SQL����ȷ��
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
        
    strSql = Trim(UCase(SynSQL.Text))
    If Trim(SynSQL.Text) = "" Then
        MsgBox "��ȡSQL����Ϊ�գ����飡", vbInformation, "��֤SQL"
        Call zlcontrol.ControlSetFocus(SynSQL)
        Exit Function
    ElseIf zlstr.ActualLen(strSql) > 4000 Then
        MsgBox "��ȡSQL���ܳ���4000�ַ������飡", vbInformation, "��֤SQL"
        Call zlcontrol.ControlSetFocus(SynSQL)
        Exit Function
    End If
    strSql = Replace(strSql, " ", "")
    If InStr(strSql, "A.����ID<>-1ANDA.��ҳID<>-1") = 0 And InStr(strSql, "A.��ҳID<>-1ANDA.����ID<>-1") = 0 Then
        MsgBox "��ȡSQL�б������[a.����ID<>-1 And a.��ҳID<>-1]����"
        Call zlcontrol.ControlSetFocus(SynSQL)
        Exit Function
    End If
    On Error GoTo errH
    gstrSql = "Select ��ȡsql From ҽ�����Ӱಡ������ Where ��� <> [1] And ��ȡSQL is not null order by ˳��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "��ȡ����������Ϣ", mstrSName)
    If rsTemp.RecordCount > 0 Then
        strSql = rsTemp!��ȡSQL
        strSql = SynSQL.Text & vbNewLine & "Union All " & strSql
        strSql = UCase(strSql)
        strSql = Replace(strSql, "[��ʼʱ��]", zlstr.To_Date(Now))
        strSql = Replace(strSql, "[����ʱ��]", zlstr.To_Date(Now))
        strSql = Replace(strSql, "[����ID]", "[1]")
        On Error Resume Next
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����������Ϣ", 0)
        If err.Number = 0 Then
            CheckSQL = True
            Exit Function
        Else
            MsgBox "��ȡSQL��д����ȷ�����飡" & vbNewLine & err.Description, vbInformation, "��֤SQL"
            Call zlcontrol.ControlSetFocus(SynSQL)
            Exit Function
        End If
    Else
        '������ݿ���û��һ�����ݣ�������ֶεļ��
        strSql = Trim(UCase(SynSQL.Text))
        If Not strSql Like "SELECT*����ID,��ҳID,����,�Ա�,����,����,��ʶ��,��Ժʱ��,��Ժ��ʽ,��Ժ����ID*" Then
            MsgBox "��ȡSQL��д����ȷ�����飡", vbInformation, "��֤SQL"
            Call zlcontrol.ControlSetFocus(SynSQL)
            Exit Function
        End If
    End If
    CheckSQL = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, Me.Caption
End Function

Private Sub synSQL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyV And Shift = 2 Then
        SynSQL.Paste
    ElseIf KeyCode = vbKeyZ And Shift = 2 Then
        SynSQL.Undo
    ElseIf KeyCode = vbKeyY And Shift = 2 Then
        SynSQL.Redo
    ElseIf KeyCode = vbKeyC And Shift = 2 Then
        SynSQL.Copy
    ElseIf KeyCode = vbKeyA And Shift = 2 Then
        SynSQL.SelectAll
    End If
End Sub

Private Sub txtBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(KeyAscii = Asc("'"), 0, KeyAscii)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(KeyAscii = Asc("'"), 0, KeyAscii)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub txtSName_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(KeyAscii = Asc("'"), 0, KeyAscii)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub
