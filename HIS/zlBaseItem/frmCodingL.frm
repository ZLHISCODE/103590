VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCodingL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ӳ��¼�����"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   Icon            =   "frmCodingL.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Caption         =   "����"
      Height          =   1605
      Left            =   180
      TabIndex        =   2
      Top             =   150
      Width           =   2715
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   390
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1170
         Width           =   765
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   1155
         TabIndex        =   3
         Top             =   1170
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Text1"
         BuddyDispid     =   196610
         OrigLeft        =   1530
         OrigTop         =   900
         OrigRight       =   1770
         OrigBottom      =   1215
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lbl��ʾ 
         Caption         =   "    ���������ĳ��ȣ�Ĭ��Ϊ��ǰ���ȣ��ұ�����ڵ�ǰ���ȡ�"
         Height          =   795
         Left            =   420
         TabIndex        =   5
         Top             =   330
         Width           =   1995
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3150
      TabIndex        =   1
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3150
      TabIndex        =   0
      Top             =   630
      Width           =   1100
   End
End
Attribute VB_Name = "frmCodingL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOK As Boolean

Private Sub cmdCancel_Click()
    mblnOK = False
    Me.Hide
End Sub



Private Sub cmdOK_Click()
    mblnOK = True
    Me.Hide
End Sub

Public Function GetLength(ByVal intValue As Integer, ByVal intMax As Integer, ByVal str���� As String, Optional ByVal strMsg As String) As Integer
'����:��������ô��ڽ���ͨѶ�ĳ���
'����:intValue ��С����
'     intMax   ��󳤶�
'����ֵ:�õ��ĳ���
    UpDown1.Min = intValue
    UpDown1.Max = intMax
    UpDown1.Value = intValue
    
    If str���� <> "" Then
        lbl��ʾ.Caption = "������" & "��" & Mid(str����, InStr(1, str����, "��") + 1, Len(str����) - InStr(1, str����, "��")) & "��" & "Ŀǰ���¼����볤�ȣ������������޸ĵĳ��ȣ��ó��ȱ�������г��ȴ�"
    End If
    
    If strMsg <> "" Then Frame1.Caption = strMsg
    Me.Show vbModal
    GetLength = IIF(mblnOK, UpDown1.Value, 0)
    Unload Me
End Function
