VERSION 5.00
Begin VB.Form frmParaSetup 
   Caption         =   "��������"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7365
   Icon            =   "frmParaSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5025
   ScaleWidth      =   7365
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame2 
      Caption         =   "��Ϣ���շ�ʽ"
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7095
      Begin VB.OptionButton optInputDataType 
         Caption         =   "�ļ�"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   2655
      End
      Begin VB.Frame frmFileInput 
         Height          =   1935
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   6735
         Begin VB.TextBox txtFileBackupDir 
            Height          =   375
            Left            =   1320
            TabIndex        =   14
            Top             =   1350
            Width           =   5200
         End
         Begin VB.TextBox txtFileSuffix 
            Height          =   375
            Left            =   1320
            TabIndex        =   12
            Top             =   825
            Width           =   5200
         End
         Begin VB.TextBox txtFileDir 
            Height          =   375
            Left            =   1320
            TabIndex        =   10
            Top             =   300
            Width           =   5200
         End
         Begin VB.Label Label4 
            Caption         =   "����Ŀ¼��"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1410
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "�ļ���׺"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   885
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "�ļ�Ŀ¼��"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.OptionButton optInputDataType 
         Caption         =   "socket��Ĭ�ϣ�"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   370
      Left            =   6120
      TabIndex        =   2
      Top             =   4440
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Height          =   370
      Left            =   4440
      TabIndex        =   1
      Top             =   4440
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "��������"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   7095
      Begin VB.TextBox txtTimeOut 
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "��ʱ��                    ��"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmParaSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '�������
    gintTimeOutMax = Val(txtTimeOut.Text)
    SaveSetting "ZLSOFT", gstrRegPath, "��ʱ", gintTimeOutMax
    
    If optInputDataType(0).Value = True Then
        gintInputDataType = 0
    Else
        gintInputDataType = 1
    End If
    SaveSetting "ZLSOFT", gstrRegPath, "������Ϣ��ʽ", gintInputDataType
    
    gstrFileDir = txtFileDir.Text
    SaveSetting "ZLSOFT", gstrRegPath, "�ļ���ϢĿ¼", gstrFileDir
    
    gstrFileSuffix = txtFileSuffix.Text
    SaveSetting "ZLSOFT", gstrRegPath, "�ļ���Ϣ��׺", gstrFileSuffix
    
    gstrFileBackupDir = txtFileBackupDir.Text
    SaveSetting "ZLSOFT", gstrRegPath, "�ļ���Ϣ����Ŀ¼", gstrFileBackupDir
    

    Unload Me
End Sub

Private Sub Form_Load()
    
    '��ע����ȡ��ʱ����
    txtTimeOut.Text = gintTimeOutMax
    If gintInputDataType = 1 Then
        optInputDataType(1).Value = True
        frmFileInput.Enabled = True
    Else
        optInputDataType(0).Value = True
        frmFileInput.Enabled = False
    End If
    
    txtFileDir.Text = gstrFileDir
    txtFileSuffix.Text = gstrFileSuffix
    txtFileBackupDir.Text = gstrFileBackupDir
        
End Sub

Public Sub zlSohwMe(frmParent As Form)
    Me.Show 1, frmParent
End Sub

Private Sub optInputDataType_Click(Index As Integer)
    If Index = 0 Then
        frmFileInput.Enabled = False
    Else
        frmFileInput.Enabled = True
    End If
End Sub
