VERSION 5.00
Begin VB.Form frmInputFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   2385
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5595
   Icon            =   "frmInputFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4290
      TabIndex        =   3
      Top             =   1935
      Width           =   1100
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "������һ��(&N)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2715
      TabIndex        =   2
      Top             =   1935
      Width           =   1530
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   195
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1935
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   60
      Left            =   0
      TabIndex        =   4
      Top             =   1785
      Width           =   5565
   End
   Begin VB.ComboBox cboSource 
      Height          =   300
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   930
      Width           =   3435
   End
   Begin VB.Label lblComment 
      Caption         =   "    ����ϣ�����ҵĻ����ִʵ��ִʡ�ƴ���������롣����ڶ����������������һ����ֱ���ҵ���ϣ�����ҵ���Ŀ��"
      Height          =   525
      Left            =   885
      TabIndex        =   7
      Top             =   135
      Width           =   4500
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "(�����ҵ�10������ǰΪ��1��)"
      Height          =   180
      Left            =   870
      TabIndex        =   6
      Top             =   1455
      Width           =   2430
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   195
      Picture         =   "frmInputFind.frx":058A
      Top             =   150
      Width           =   480
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      Caption         =   "��������(&F)"
      Height          =   180
      Left            =   885
      TabIndex        =   0
      Top             =   1005
      Width           =   990
   End
End
Attribute VB_Name = "frmInputFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsFind As New ADODB.Recordset
Private strFind As String
Private intCount As Integer
Private mfrmMain As Object
Private mbytMode As Byte

Public Function ShowFind(ByVal frmMain As Object, Optional ByVal bytMode As Byte = 1) As Boolean
    
    Set mfrmMain = frmMain
    mbytMode = bytMode
    
    Me.Show 1, frmMain
    
End Function

Private Sub cboSource_Click()
    If Trim(cboSource.Text) <> "" Then
        cmdFind.Enabled = True
    Else
        cmdFind.Enabled = False
    End If
End Sub

Private Sub cboSource_GotFocus()
    cboSource.SelStart = 0
    cboSource.SelLength = 100
End Sub

Private Sub cboSource_KeyPress(KeyAscii As Integer)
    
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
        Exit Sub
    End If
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboSource_KeyUp(KeyCode As Integer, Shift As Integer)
    If Trim(cboSource.Text) <> "" Then
        cmdFind.Enabled = True
    Else
        cmdFind.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim strTemp As String
    Dim strKey As String
    
    If Trim(cboSource.Text) = "" Then
        MsgBox "��������ҵ�����", vbExclamation, gstrSysName
        cboSource.SetFocus
        Exit Sub
    End If
    
    If InStr(cboSource.Text, "'") > 0 Then
        MsgBox "����������зǷ��ַ� '", vbExclamation, gstrSysName
        cboSource.SetFocus
        Exit Sub
    End If
    
    strTemp = ""
    For intCount = 0 To cboSource.ListCount
        strTemp = strTemp & ";" & cboSource.List(intCount)
    Next
    
    If InStr(1, strTemp, ";" & Trim(cboSource.Text)) = 0 Then
        cboSource.AddItem Trim(cboSource.Text), 0
    End If
    
    gstrSQL = "select �ִ�,������,���뷨 " & _
            " from zlwordbasic " & _
            " where �Ƿ���=" & mbytMode & " and (������ like '%" & Trim(cboSource.Text) & "%'" & _
                    "or �ִ� like '%" & Trim(cboSource.Text) & "%')  ORDER BY ���뷨,������"
    
    Err = 0
    On Error GoTo errHand
    
    With rsFind
        If strFind <> gstrSQL Or .State <> adStateOpen Then
            If .State = adStateOpen Then .Close
            rsFind.Open gstrSQL, gcnOracle
            
            If .EOF Then
                MsgBox "�����ڲ��ҵ����ݣ�", vbExclamation, gstrSysName
                
                .Close
                cmdFind.Enabled = False
                lblNote.Caption = ""
                cboSource.SetFocus
                
                Exit Sub
            End If
            strFind = gstrSQL
        Else
            .MoveNext
            If .EOF Then
                MsgBox "�Ѳ��ҵ����һ����Ŀ��", vbExclamation, gstrSysName
                
                .Close
                cboSource.Text = ""
                cmdFind.Enabled = False
                lblNote.Caption = ""
                cboSource.SetFocus
                
                Exit Sub
            End If
        End If
        
        lblNote.Caption = "(�����ҵ�" & .RecordCount & "������ǰΪ��" & .AbsolutePosition & "��)"
        
        If rsFind("���뷨").Value = 1 Then
            strKey = "P" & Left(rsFind("������").Value, 1)
        Else
            strKey = "W" & Left(rsFind("������").Value, 1)
        End If
        
        Call mfrmMain.LocationItem(strKey, rsFind("�ִ�").Value, rsFind("������").Value)
        
    End With
    
    Exit Sub
    
errHand:
    MsgBox "���һ����ִ�ʧ�ܣ�" & vbNewLine & Err.Description, vbInformation, gstrSysName
End Sub

Private Sub cmdHelp_Click()
    ShowHelp Me.hwnd, "zl9svrtools\" & Me.name    '
End Sub

Private Sub Form_Load()
    Dim lngLoop As Long
    Dim strSectoin  As String
    Dim str����  As String
    
    cboSource.Clear
    
    strSectoin = "ʵ�ù���\�����빤��\" & gstrUserName & "\��ʷ����"
    
    For lngLoop = 0 To CLng(Val(GetSetting("ZLSOFT", strSectoin, "��������", "0")))
        str���� = GetSetting("ZLSOFT", strSectoin, "��������" & lngLoop, "")
        If str���� <> "" Then cboSource.AddItem str����
    Next
    
    strFind = ""
    lblNote.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim lngLoop As Long
    Dim strSectoin As String
    
    '�����˳��������ĸ���
    strSectoin = "ʵ�ù���\�����빤��\" & gstrUserName & "\��ʷ����"
    
    On Error Resume Next '���û�иü�ֵ���ͻ����
    DeleteSetting "ZLSOFT", strSectoin 'ɾ����ǰ������
    On Error GoTo 0
    
    Call SaveSetting("ZLSOFT", strSectoin, "��������", IIf(cboSource.ListCount > 10, 10, cboSource.ListCount))
    
    For lngLoop = 0 To cboSource.ListCount
        
        If lngLoop > 10 Then Exit For
        
        Call SaveSetting("ZLSOFT", strSectoin, "��������" & lngLoop, cboSource.List(lngLoop))

    Next
    
End Sub
