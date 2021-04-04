VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmUpdateDB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�������ݿ�"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4380
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "����"
      Height          =   350
      Left            =   1560
      TabIndex        =   3
      Top             =   1440
      Width           =   1100
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   "���..."
      Height          =   350
      Left            =   3240
      TabIndex        =   2
      Top             =   840
      Width           =   1100
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   1920
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox TxtPath 
      Height          =   350
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3075
   End
   Begin VB.Label Label2 
      Caption         =   "��ǰ�汾�ţ�"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblCurrentVersion 
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "�����ű�·��"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "frmUpdateDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_cnAccess As ADODB.Connection


Private Sub cmdNavigate_Click()
    Dim strFileName As String
    dlgOpen.Filter = "�����ű���*.SQL)|*.SQL|ȫ��(*.*)|*.*"
    dlgOpen.ShowOpen
    strFileName = dlgOpen.Filename
    TxtPath.Text = strFileName
End Sub

Private Sub cmdUpdate_Click()
    '�򿪽ű��ļ�
    Dim strSQL As String
    Dim strTemp As String
    Dim lngFileNo As Long
    Dim blnError As Boolean
    
    If m_cnAccess Is Nothing Then
        MsgBox "û����ȷ��Access���ݿ����ӣ��޷�������", vbInformation, gstrSysName
        Unload Me
    End If
    
    If TxtPath.Text <> "" Then
        On Error GoTo err1
        lngFileNo = FreeFile()
        Open TxtPath.Text For Input As lngFileNo
    
        'ѭ����ȡ�ű��ļ��е�SQL��䲢ִ��
        Do While Not EOF(lngFileNo)     'ѭ�����ļ�β
            Line Input #lngFileNo, strTemp
            If left(Trim(strTemp), 2) <> "--" Then
                strSQL = strSQL & " " & Trim(strTemp)
                If Right(Trim(strTemp), 1) = ";" Then
                    On Error Resume Next
                    m_cnAccess.Execute strSQL
                    If err <> 0 Then
                        Call WriteLog(1, err.Number, err.Description)
                        err.Clear
                        blnError = True
                    End If
                    strSQL = ""
                    On Error GoTo err1
                End If
            End If
        Loop
        Close lngFileNo
        '����������ˢ�´�������ʾ�İ汾��
        Form_Load
        If blnError Then
            MsgBox "������ɣ����������д���������ϸ��Ϣ����<������־>��", vbInformation, gstrSysName
        Else
            MsgBox "������ɡ���ǰ�汾���ǣ�" & lblCurrentVersion.Caption, vbInformation, gstrSysName
        End If
    End If
    Exit Sub
    '������
err1:
    MsgBox "����ʧ�ܣ��ű��ļ�����" & strSQL, vbInformation, gstrSysName
    Close lngFileNo
End Sub


Private Sub Form_Load()
    Dim rsResult As Recordset
    On Error Resume Next
    Set rsResult = m_cnAccess.Execute("select �汾�� from �汾��")
    If Not rsResult.EOF Then
        lblCurrentVersion.Caption = rsResult!�汾��
    End If
End Sub
