VERSION 5.00
Begin VB.Form frmLink 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6735
   Icon            =   "frmLink.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   6735
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   960
      MaxLength       =   20
      TabIndex        =   19
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   360
      Left            =   5400
      TabIndex        =   21
      Top             =   4440
      Width           =   1110
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      Height          =   360
      Left            =   4200
      TabIndex        =   20
      Top             =   4440
      Width           =   1110
   End
   Begin VB.Frame fraLink 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.OptionButton optLink 
         Caption         =   "���Ӵ�(&L)"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optLink 
         Caption         =   "Web Services(&W)"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtConnectStr 
         ForeColor       =   &H80000006&
         Height          =   285
         Left            =   480
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   720
         Width           =   4455
      End
      Begin VB.CommandButton cmdBuild 
         Caption         =   "����(&U)"
         Height          =   360
         Left            =   5040
         TabIndex        =   3
         Top             =   720
         Width           =   990
      End
      Begin VB.Frame fraWS 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   480
         TabIndex        =   5
         Top             =   1320
         Width           =   5775
         Begin VB.TextBox txtConfirm 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1200
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   14
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox txtPass 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1200
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   12
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox txtUser 
            Height          =   285
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   10
            Top             =   600
            Width           =   1575
         End
         Begin VB.CommandButton cmdWSTest 
            Caption         =   "����(&T)"
            Height          =   360
            Left            =   4560
            TabIndex        =   8
            Top             =   240
            Width           =   990
         End
         Begin VB.TextBox txtURL 
            Height          =   285
            Left            =   1200
            TabIndex        =   7
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label lblLink 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ȷ�����룺"
            Height          =   180
            Index           =   3
            Left            =   240
            TabIndex        =   13
            Top             =   1350
            Width           =   900
         End
         Begin VB.Label lblLink 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��    �룺"
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   11
            Top             =   990
            Width           =   900
         End
         Begin VB.Label lblLink 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��    ����"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   9
            Top             =   630
            Width           =   900
         End
         Begin VB.Label lblLink 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����ַ��"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   270
            Width           =   900
         End
      End
      Begin VB.OptionButton optLink 
         Caption         =   "����Ŀ¼(&D)"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox txtDirectory 
         ForeColor       =   &H80000006&
         Height          =   285
         Left            =   480
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3600
         Width           =   4455
      End
      Begin VB.CommandButton cmdBrowser 
         Caption         =   "���(&B)"
         Height          =   360
         Left            =   5040
         TabIndex        =   17
         Top             =   3600
         Width           =   990
      End
   End
   Begin VB.Label lblLink 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      Height          =   180
      Index           =   4
      Left            =   240
      TabIndex        =   18
      Top             =   4470
      Width           =   650
   End
End
Attribute VB_Name = "frmLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (LpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDlist Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Type BROWSEINFO
    hOwner As Long
    pidlroot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lparam As Long
    iImage As Long
End Type

Private mobjDataLink As MSDASC.DataLinks
Private mcnTmp As New ADODB.Connection

Public Sub ShowMe(ByVal frmOwner As Form, Optional ByVal lngID As Long)
    Dim rsTmp As ADODB.Recordset
    
    If lngID > 0 Then
        gstrSQL = "Select * From ҩ���豸���� Where ID = [1] "
        On Error GoTo errHandle
        Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "��ȡҩ���豸������Ϣ", lngID)
        If Not rsTmp.EOF Then
            txtName.Text = rsTmp!����
            txtName.Tag = lngID
            optLink(rsTmp!��������).Value = True
            Select Case rsTmp!��������
                Case enuLinkType.DB
                    txtConnectStr.Text = rsTmp!��������
                Case enuLinkType.WEBServices
                    txtURL.Text = GetConnectStrEle(rsTmp!��������, enuLinkType.WEBServices, "URL")
                    txtUser.Text = GetConnectStrEle(rsTmp!��������, enuLinkType.WEBServices, "USER")
                    txtPass.Text = GetConnectStrEle(rsTmp!��������, enuLinkType.WEBServices, "PWD")
                    txtConfirm.Text = txtPass.Text
                Case Else
                    txtDirectory.Text = rsTmp!��������
            End Select
        End If
        rsTmp.Close
        
        Me.Tag = 1      '�޸�״̬
    Else
        Call optLink_Click(0)
        txtName.Tag = 0
        Me.Tag = 0      '����״̬
    End If
    
    Show vbModal, frmOwner
    
    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Function GetFolder(ByVal hWnd As Long, Optional Title As String) As String
    Dim typBI As BROWSEINFO
    Dim lngPID As Long
    Dim strFolder As String
    
    strFolder = Space(255)
    With typBI
       If IsNumeric(hWnd) Then .hOwner = hWnd
       .ulFlags = BIF_RETURNONLYFSDIRS
       .pidlroot = 0
       If Title <> "" Then
          .lpszTitle = Title & Chr$(0)
       Else
          .lpszTitle = "ѡ��Ŀ¼" & Chr$(0)
        End If
    End With

    lngPID = SHBrowseForFolder(typBI)
    
    If SHGetPathFromIDlist(ByVal lngPID, ByVal strFolder) Then
        GetFolder = Left(strFolder, InStr(strFolder, Chr$(0)) - 1)
    Else
        GetFolder = ""
    End If
End Function

Private Sub cmdBrowser_Click()
    Dim strPath As String
    strPath = GetFolder(Me.hWnd, "����ļ���")
    If strPath <> "" Then
        txtDirectory.Text = strPath
    End If
End Sub

Private Sub cmdBuild_Click()
    
    cmdBuild.Enabled = False
    
    On Error GoTo errHandle
    If mobjDataLink Is Nothing Then
        Set mobjDataLink = New MSDASC.DataLinks
    End If
    If mcnTmp Is Nothing Then
        Set mcnTmp = mobjDataLink.PromptNew
    Else
        mcnTmp.ConnectionString = txtConnectStr.Text
        mobjDataLink.PromptEdit mcnTmp
    End If
    
    If Not mcnTmp Is Nothing Then txtConnectStr.Text = mcnTmp.ConnectionString
    
errHandle:
    If Err.Number <> 0 Then
        mcnTmp.ConnectionString = ""
        txtConnectStr.Text = ""
    End If
    cmdBuild.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim intType As Integer
    Dim strContent As String
    
    '���
    If optLink(0).Value Then
        
        If Trim(txtConnectStr.Text) = "" Then
            MsgBox "δ���ɡ����Ӵ������ܾ����棡", vbInformation, GSTR_INTERFACE_NAME
            Exit Sub
        End If
        
        intType = 0
        strContent = txtConnectStr.Text
        
    ElseIf optLink(1).Value Then
        
        If Trim(txtURL.Text) = "" Then
            MsgBox "������Ҫ��д�������ַ�����ܱ��棡", vbInformation, GSTR_INTERFACE_NAME
            Exit Sub
        End If
        If Trim(txtPass.Text) & Trim(txtConfirm.Text) <> "" Then
            If txtPass.Text <> txtConfirm.Text Then
                MsgBox "ȷ������¼�벻��ȷ��", vbInformation, GSTR_INTERFACE_NAME
                Exit Sub
            End If
        End If
        If TestURL(txtURL.Text) = False Then
            MsgBox "�����ַ���Ӳ���ʧ�ܣ�", vbInformation, GSTR_INTERFACE_NAME
            Exit Sub
        End If
        
        intType = 1
        strContent = "URL=" & txtURL.Text & ";USER=" & Trim(txtUser.Text) & ";" & "PWD=" & Trim(txtPass.Text)
        
    Else
        
        If Trim(txtDirectory.Text) = "" Then
            MsgBox "δ��д������Ŀ¼�����ܾ����棡", vbInformation, GSTR_INTERFACE_NAME
            Exit Sub
        End If
        
        intType = 2
        strContent = txtDirectory.Text
    
    End If
    
    If Me.Tag = 1 Then
        '�޸�
        gstrSQL = "ZL_ҩ���豸����_UPDATE("
        gstrSQL = gstrSQL & txtName.Tag & ","
        gstrSQL = gstrSQL & "'" & txtName.Text & "',"
        gstrSQL = gstrSQL & intType & ","
        gstrSQL = gstrSQL & "'" & strContent & "')"
    Else
        '����
        gstrSQL = "ZL_ҩ���豸����_INSERT("
        gstrSQL = gstrSQL & "'" & txtName.Text & "',"
        gstrSQL = gstrSQL & intType & ","
        gstrSQL = gstrSQL & "'" & strContent & "')"
    End If
    
    On Error GoTo errHandle
    Call gobjComLib.zldatabase.ExecuteProcedure(gstrSQL, "�Զ���ϵͳ��������-" & IIf(Me.Tag = 1, "�޸�", "����"))
    
    Unload Me
    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub cmdWSTest_Click()
    If TestURL(txtURL.Text) = False Then
        MsgBox "�����ַ���Ӳ���ʧ�ܣ�" & vbNewLine & gstrMessage, vbInformation, GSTR_INTERFACE_NAME
    Else
        gstrMessage = ""
        MsgBox "���Ӳ��Գɹ���", vbInformation, GSTR_INTERFACE_NAME
    End If
End Sub

Private Sub Form_Activate()
    Dim i As Integer
    For i = 0 To optLink.Count - 1
        If optLink(i).Value Then
            Call optLink_Click(i)
            Exit For
        End If
    Next
End Sub

Private Sub Form_Load()
    Call txtName_Change
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjDataLink = Nothing
End Sub

Private Sub optLink_Click(Index As Integer)
    cmdBuild.Enabled = Index = 0
    cmdWSTest.Enabled = Index = 1
    cmdBrowser.Enabled = Index = 2
    
    txtURL.Enabled = Index = 1
    txtUser.Enabled = Index = 1
    txtPass.Enabled = Index = 1
    txtConfirm.Enabled = Index = 1
    
    Select Case Index
        Case enuLinkType.DB
            If cmdBuild.Visible Then
                If Trim(txtConnectStr.Text) = "" Then
                    cmdBuild.Caption = "����(&U)"
                Else
                    cmdBuild.Caption = "�޸�(&M)"
                End If
                cmdBuild.SetFocus
            End If
        Case enuLinkType.WEBServices
            If txtURL.Visible Then txtURL.SetFocus
        Case enuLinkType.Directory
            If cmdBrowser.Visible Then cmdBrowser.SetFocus
    End Select
End Sub

Private Sub txtConfirm_KeyPress(KeyAscii As Integer)
    If InStr("`~@#$%^&*()=+[]\{}|;':"",./<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtConnectStr_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not (KeyCode = 35 Or KeyCode = 36 Or KeyCode = 37 Or KeyCode = 39) Then KeyCode = 0
End Sub

Private Sub txtConnectStr_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtDirectory_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not (KeyCode = 35 Or KeyCode = 36 Or KeyCode = 37 Or KeyCode = 39) Then KeyCode = 0
End Sub

Private Sub txtDirectory_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtName_Change()
    cmdSave.Enabled = Trim(txtName.Text) <> ""
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If InStr("`~@#$%^&*()=+[]\{}|;':"",./<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If InStr("`~@#$%^&*()=+[]\{}|;':"",./<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If InStr("`~@#$%^&*()=+[]\{}|;':"",./<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
