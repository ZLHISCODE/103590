VERSION 5.00
Begin VB.Form frmPrintPageSet 
   BorderStyle     =   0  'None
   Caption         =   "��ӡҳ������"
   ClientHeight    =   3690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdȡ�� 
      Caption         =   "ȡ ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txt˵�� 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmPrintPageSet.frx":0000
      Top             =   120
      Width           =   8055
   End
   Begin VB.TextBox txtѡ��ҳ�� 
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   3000
      Width           =   4095
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txt����ҳ 
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txt��ʼҳ 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "ѡ���ӡҳ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "����ҳ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "��ʼҳ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   855
   End
End
Attribute VB_Name = "frmPrintPageSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mstrPrintPage As String
Private mintPageCount As Integer

Public Function ShowMe(intPageCount As Integer) As String
    mintPageCount = intPageCount
    Me.Show 1
    ShowMe = mstrPrintPage
End Function

Private Sub cmdȡ��_Click()
    mstrPrintPage = ""
    Unload Me
    
End Sub

Private Sub cmdȷ��_Click()
    Dim strTemp As String
    Dim Temp As String
    Dim i As Integer
    If Trim(txtѡ��ҳ��.Text) <> "" Then
        strTemp = Trim(txtѡ��ҳ��.Text)
        For i = 1 To Len(strTemp)
            Temp = Mid(strTemp, i, i + 1)
            If (Asc(Temp) >= 48 And Asc(Temp) <= 57) Or Asc(Temp) = 44 Then
                
            Else
                MsgBox "�������ݰ����Ƿ�����,ֻ���������ֺ�Ӣ�Ķ���(1,2,3,4,5,6,7,8,9,0,)", vbOKOnly, gstrSysName
                mstrPrintPage = ""
                Exit Sub
            End If
        Next
        mstrPrintPage = Trim(txtѡ��ҳ��.Text)
    Else
        If Trim(txt����ҳ.Text) = "" Then
            MsgBox "����ҳ����������,����������", vbOKOnly, gstrSysName
            Exit Sub
        End If
        If Trim(txt��ʼҳ.Text) = "" Then
            MsgBox "��ʼҳ����������,����������", vbOKOnly, gstrSysName
            Exit Sub
        End If
        If CInt(Trim(txt����ҳ.Text)) > mintPageCount Then
            MsgBox "����ҳ�Ų��ܴ�����ҳ��,����������", vbOKOnly, gstrSysName
            Exit Sub
        End If
        If CInt(Trim(txt��ʼҳ.Text)) <= 0 Or CInt(Trim(txt��ʼҳ.Text)) > mintPageCount Then
            MsgBox "��ʼҳ�ű������0С�ڵ�����ҳ��,����������", vbOKOnly, gstrSysName
            Exit Sub
        End If
        
        If CInt(Trim(txt����ҳ.Text)) <= 0 Or CInt(Trim(txt��ʼҳ.Text)) <= 0 Then
            MsgBox "��ʼҳ�ͽ���ҳ���������0,����������", vbOKOnly, gstrSysName
            mstrPrintPage = ""
            txt����ҳ.Text = ""
            txt��ʼҳ.Text = ""
            Exit Sub
        End If
        If CInt(Trim(txt����ҳ.Text)) = CInt(Trim(txt��ʼҳ.Text)) Then
            mstrPrintPage = CInt(Trim(txt��ʼҳ.Text))
        ElseIf CInt(Trim(txt����ҳ.Text)) < CInt(Trim(txt��ʼҳ.Text)) Then
            MsgBox "����ҳ������ڵ�����ʼҳ,����������", vbOKOnly, gstrSysName
            mstrPrintPage = ""
            txt����ҳ.Text = ""
            txt��ʼҳ.Text = ""
        Else
            mstrPrintPage = ""
            For i = CInt(Trim(txt��ʼҳ.Text)) To CInt(Trim(txt����ҳ.Text))
                mstrPrintPage = mstrPrintPage & "," & i
            Next
        End If
    End If
    If Mid(mstrPrintPage, 1, 1) = "," Then
        mstrPrintPage = Mid(mstrPrintPage, 2)
    End If
    If Mid(mstrPrintPage, Len(mstrPrintPage), 1) = "," Then
        mstrPrintPage = Mid(mstrPrintPage, Len(mstrPrintPage), 1)
    End If
    Unload Me
End Sub

'Private Sub txt����ҳ_Change()
'    If CInt(Trim(txt����ҳ.Text)) > 0 Then
'
'    Else
'        txt����ҳ.Text = ""
'    End If
'End Sub

Private Sub txt����ҳ_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
        Exit Sub
    Else
        MsgBox "�������������,����������", vbOKOnly, gstrSysName
        Exit Sub
    End If
End Sub

Private Sub txt��ʼҳ_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
        Exit Sub
    Else
        MsgBox "�������������,����������", vbOKOnly, gstrSysName
        Exit Sub
    End If
    
End Sub



Private Sub txtѡ��ҳ��_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 44 Or KeyAscii = 8 Then
        Exit Sub
    Else
        MsgBox "�������������,����������", vbOKOnly, gstrSysName
        Exit Sub
    End If
End Sub
