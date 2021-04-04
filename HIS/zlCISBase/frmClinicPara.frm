VERSION 5.00
Begin VB.Form frmClinicPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   Icon            =   "frmClinicPara.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3210
      TabIndex        =   8
      Top             =   3615
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2010
      TabIndex        =   7
      Top             =   3615
      Width           =   1100
   End
   Begin VB.Frame fraAddMode 
      Caption         =   " 1����Ŀ���Ӳ���ģʽ"
      Height          =   1365
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4155
      Begin VB.OptionButton opt����ģʽ 
         Caption         =   "��������(�����رձ༭)"
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   2580
      End
      Begin VB.OptionButton opt����ģʽ 
         Caption         =   "��������(������Զ�������Ŀ)"
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   3105
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " 2����Ŀ����Ӧ�÷�Χ����"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   1695
      Width           =   4155
      Begin VB.CheckBox chkӦ�÷�Χ 
         Caption         =   "����Ӧ����ͬ��������Ŀ"
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
      Begin VB.CheckBox chkӦ�÷�Χ 
         Caption         =   "����Ӧ����ͬ����������Ŀ"
         Height          =   300
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   2535
      End
      Begin VB.CheckBox chkӦ�÷�Χ 
         Caption         =   "����Ӧ����ͬ���������Ŀ"
         Height          =   300
         Index           =   2
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmClinicPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrPrivs As String

Public Sub ShowMe(ByVal frmParent As Object, ByVal strPrivs As String)
    mstrPrivs = strPrivs
    
    Me.Show 1, frmParent
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    Dim strӦ�÷�Χ As String
    
    If Me.opt����ģʽ(0).Value = True Then
        Call zlDatabase.SetPara("������Ŀ��������", 0, glngSys, 1054)
    Else
        Call zlDatabase.SetPara("������Ŀ��������", 1, glngSys, 1054)
    End If
    
    strӦ�÷�Χ = IIf(chkӦ�÷�Χ(0).Value = 1, "1", "0")
    strӦ�÷�Χ = strӦ�÷�Χ & IIf(chkӦ�÷�Χ(1).Value = 1, "1", "0")
    strӦ�÷�Χ = strӦ�÷�Χ & IIf(chkӦ�÷�Χ(2).Value = 1, "1", "0")
    
    Call zlDatabase.SetPara("��ĿӦ�÷�Χ", strӦ�÷�Χ, glngSys, 1054)
    
    Unload Me
End Sub

Private Sub Form_Load()
    '�����û�Ȩ�ޣ�װ��ؼ�
    Dim lngValues As Long
    Dim strӦ�÷�Χ As String
    Dim blnSetPara As Boolean
    
    blnSetPara = zlStr.IsHavePrivs(mstrPrivs, "��������")
    
    lngValues = Val(zlDatabase.GetPara("������Ŀ��������", glngSys, 1054, 0, Array(Me.opt����ģʽ(0), Me.opt����ģʽ(1)), blnSetPara))
    strӦ�÷�Χ = zlDatabase.GetPara("��ĿӦ�÷�Χ", glngSys, 1054, "000", Array(chkӦ�÷�Χ(0), chkӦ�÷�Χ(1), chkӦ�÷�Χ(2)), blnSetPara)
    
    If lngValues = 0 Then
        Me.opt����ģʽ(0).Value = True: Me.opt����ģʽ(1).Value = False
    Else
        Me.opt����ģʽ(0).Value = False: Me.opt����ģʽ(1).Value = True
    End If
    
    If Val(Mid(strӦ�÷�Χ, 1, 1)) = 1 Then
        chkӦ�÷�Χ(0).Value = 1
    End If
    
    If Val(Mid(strӦ�÷�Χ, 2, 1)) = 1 Then
        chkӦ�÷�Χ(1).Value = 1
    End If
    
    If Val(Mid(strӦ�÷�Χ, 3, 1)) = 1 Then
        chkӦ�÷�Χ(2).Value = 1
    End If
End Sub

