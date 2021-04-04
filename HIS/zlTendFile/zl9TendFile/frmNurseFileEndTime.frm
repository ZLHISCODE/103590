VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmNurseFileEndTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ʱ��"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4155
   Icon            =   "frmNurseFileEndTime.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSMask.MaskEdBox txt����ʱ�� 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy-MM-dd HH:mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   0
      EndProperty
      Height          =   300
      Left            =   1350
      TabIndex        =   2
      Top             =   630
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   19
      Format          =   "yyyy-MM-dd HH:mm:ss"
      Mask            =   "####-##-## ##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1410
      TabIndex        =   4
      Top             =   1440
      Width           =   1155
   End
   Begin VB.CommandButton cmdCanCel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2640
      TabIndex        =   5
      Top             =   1440
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   -120
      TabIndex        =   3
      Top             =   1110
      Width           =   4365
   End
   Begin VB.Label lbl����ʱ�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����ʱ��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   1
      Top             =   690
      Width           =   720
   End
   Begin VB.Label lbl��ǰ�ļ� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ǰ�ļ�:"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   330
      Width           =   810
   End
End
Attribute VB_Name = "frmNurseFileEndTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngFile As Long
Private mblnOK As Boolean

Public Function ShowEditor(ByVal lngFile As Long) As Boolean
    On Error Resume Next
    mlngFile = lngFile
    mblnOK = False
    Me.Show 1
    ShowEditor = mblnOK
End Function

Private Sub cmdCanCel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHand
    
    If Not IsDate(txt����ʱ��.Text) Then
        MsgBox "������Ϸ���ʱ�㣡", vbInformation, gstrSysName
        txt����ʱ��.SetFocus
        Exit Sub
    End If
    If txt����ʱ��.Text < txt����ʱ��.Tag Then
        MsgBox "�ļ��Ľ���ʱ�䲻��С�ڻ������ݵ������ʱ��[" & txt����ʱ��.Tag & "]��", vbInformation, gstrSysName
        txt����ʱ��.SetFocus
        Exit Sub
    End If
    
    gstrSQL = "ZL_���˻����ļ�_ENDTIME(" & mlngFile & ",to_date('" & txt����ʱ��.Text & "','yyyy-MM-dd hh24:mi:ss'))"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�趨��ǰ�ļ��Ľ���ʱ��")
    
    mblnOK = True
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Dim str����ʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    '��ȡ�ò�����ָ���ļ���ʽ��ͬ���ļ�,�趨�ϲ���ӡ(ֻ�ܰ�ʱ����Ⱥ�˳������趨)
    
    gstrSQL = " Select �ļ�����,����ʱ�� From ���˻����ļ� Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ļ�����", mlngFile)
    txt����ʱ��.Text = Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    Me.lbl��ǰ�ļ�.Caption = "��ǰ�ļ���" & rsTemp!�ļ�����
    
    gstrSQL = " Select max(����ʱ��) AS ����ʱ�� from ���˻������� B Where B.�ļ�ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����޸�ʱ��", mlngFile)
    txt����ʱ��.Tag = Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    
    cmdOK.Enabled = (txt����ʱ��.Tag <> "")
End Sub

Private Sub txt����ʱ��_GotFocus()
    txt����ʱ��.SelStart = 0
    txt����ʱ��.SelLength = 20
End Sub
