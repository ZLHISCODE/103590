VERSION 5.00
Begin VB.Form frmMessageBox 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�������"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5415
   Icon            =   "frmMessageBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5415
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picInformation 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2805
      Index           =   2
      Left            =   0
      ScaleHeight     =   2805
      ScaleWidth      =   5415
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   5420
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "��Ӱ��"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   23
         Left            =   2325
         TabIndex        =   34
         Top             =   255
         Width           =   765
      End
      Begin VB.Label lblWarning 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ܻ�      �ļ��嵥��"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   22
         Left            =   1785
         TabIndex        =   33
         Top             =   1215
         Width           =   2160
      End
      Begin VB.Label lblWarningCustom 
         BackStyle       =   0  'Transparent
         Caption         =   "�Ƿ�ȷ��ɾ�����ļ���"
         ForeColor       =   &H000000FF&
         Height          =   750
         Index           =   2
         Left            =   1425
         TabIndex        =   17
         Top             =   2010
         Width           =   3915
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "3.ɾ�����ܻ�ʹ��ǰ�ļ��嵥�𻵣����������"
         Height          =   390
         Index           =   10
         Left            =   1245
         TabIndex        =   16
         Top             =   1215
         Width           =   4125
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "2.ɾ��������������"
         Height          =   390
         Index           =   9
         Left            =   1245
         TabIndex        =   15
         Top             =   735
         Width           =   3705
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "1.ɾ���ļ���Ӱ��ͻ����Ѵ����ļ�"
         Height          =   390
         Index           =   8
         Left            =   1245
         TabIndex        =   14
         Top             =   255
         Width           =   3765
      End
      Begin VB.Image imgMessage 
         Height          =   720
         Index           =   2
         Left            =   255
         Picture         =   "frmMessageBox.frx":0CCA
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.PictureBox picInformation 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2805
      Index           =   1
      Left            =   0
      ScaleHeight     =   2805
      ScaleWidth      =   5415
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   5420
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "���õ��ļ�              ɾ����"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   21
         Left            =   1785
         TabIndex        =   32
         Top             =   1155
         Width           =   3240
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "������ǰ�ļ��嵥"
         Height          =   330
         Index           =   20
         Left            =   1425
         TabIndex        =   31
         Top             =   1395
         Width           =   2685
      End
      Begin VB.Label lblWarning 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ܻ�ԭ"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   18
         Left            =   2160
         TabIndex        =   30
         Top             =   690
         Width           =   720
      End
      Begin VB.Label lblWarning 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Զ�������ļ�"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   3
         Left            =   3765
         TabIndex        =   29
         Top             =   225
         Width           =   1260
      End
      Begin VB.Image imgMessage 
         Height          =   720
         Index           =   1
         Left            =   255
         Picture         =   "frmMessageBox.frx":280C
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "1.�����ļ��󣬿ͻ�������ʱ��"
         Height          =   390
         Index           =   7
         Left            =   1245
         TabIndex        =   12
         Top             =   225
         Width           =   4125
      End
      Begin VB.Label lblWarning 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2.�����ļ�        ��ֻ���������"
         Height          =   180
         Index           =   6
         Left            =   1245
         TabIndex        =   11
         Top             =   690
         Width           =   2880
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "3.�Ѿ�          ��Ҫ�����ñ���      ��������"
         Height          =   240
         Index           =   5
         Left            =   1245
         TabIndex        =   10
         Top             =   1155
         Width           =   4125
      End
      Begin VB.Label lblWarningCustom 
         BackStyle       =   0  'Transparent
         Caption         =   "�Ƿ�ȷ�����õ�ǰ�ļ���"
         ForeColor       =   &H000000FF&
         Height          =   735
         Index           =   1
         Left            =   1410
         TabIndex        =   9
         Top             =   2115
         Width           =   3915
      End
   End
   Begin VB.PictureBox picInformation 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2805
      Index           =   0
      Left            =   0
      ScaleHeight     =   2805
      ScaleWidth      =   5415
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5420
      Begin VB.Label lblWarning 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������Ҫ�����ϴ������ļ�"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   19
         Left            =   1425
         TabIndex        =   28
         Top             =   1440
         Width           =   2340
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "���᲻����������"
         Height          =   225
         Index           =   14
         Left            =   1425
         TabIndex        =   22
         Top             =   1695
         Width           =   4155
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "��ӵĵ����������ᱣ��"
         Height          =   255
         Index           =   13
         Left            =   1440
         TabIndex        =   21
         Top             =   1110
         Width           =   3435
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "��            ʱ����Ҫ"
         Height          =   240
         Index           =   12
         Left            =   1620
         TabIndex        =   20
         Top             =   480
         Width           =   4485
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "��  �Զ�ע�����        ʹ�øù�������"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   11
         Left            =   1425
         TabIndex        =   19
         Top             =   480
         Width           =   3945
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "�޸��ļ��嵥��            ������ʹ����"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   1785
         TabIndex        =   18
         Top             =   225
         Width           =   4410
      End
      Begin VB.Label lblWarningCustom 
         BackStyle       =   0  'Transparent
         Caption         =   "�Ƿ�ȷ��������ǰ�����ļ��嵥��"
         ForeColor       =   &H000000FF&
         Height          =   645
         Index           =   0
         Left            =   1440
         TabIndex        =   7
         Top             =   2190
         Width           =   3900
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "3.                          ������ͻ�������   "
         Height          =   225
         Index           =   2
         Left            =   1245
         TabIndex        =   6
         Top             =   1440
         Width           =   4300
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "2.ʹ�ñ�׼�ļ��嵥��������ǰ�ļ��嵥���û�"
         Height          =   255
         Index           =   1
         Left            =   1245
         TabIndex        =   5
         Top             =   825
         Width           =   4125
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "1.�ֶ�              ����ɿͻ���        "
         Height          =   255
         Index           =   0
         Left            =   1245
         TabIndex        =   4
         Top             =   225
         Width           =   4300
      End
      Begin VB.Image imgMessage 
         Height          =   720
         Index           =   0
         Left            =   255
         Picture         =   "frmMessageBox.frx":434E
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.PictureBox picInformation 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2805
      Index           =   3
      Left            =   0
      ScaleHeight     =   2805
      ScaleWidth      =   5415
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   5420
      Begin VB.Image imgMessage 
         Height          =   720
         Index           =   3
         Left            =   330
         Picture         =   "frmMessageBox.frx":5E90
         Top             =   510
         Width           =   720
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "1.ɾ���ļ���Ӱ��ͻ����Ѵ����ļ�"
         Height          =   390
         Index           =   17
         Left            =   1470
         TabIndex        =   27
         Top             =   315
         Width           =   3765
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "2.ɾ��������������"
         Height          =   390
         Index           =   16
         Left            =   1470
         TabIndex        =   26
         Top             =   780
         Width           =   3705
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "3.ɾ�����ܻ�ʹ��ǰ�ļ��嵥�𻵣������"
         Height          =   390
         Index           =   15
         Left            =   1470
         TabIndex        =   25
         Top             =   1245
         Width           =   3750
      End
      Begin VB.Label lblWarningCustom 
         BackStyle       =   0  'Transparent
         Caption         =   "�Ƿ�ȷ��ɾ�����ļ���"
         Height          =   750
         Index           =   3
         Left            =   1455
         TabIndex        =   24
         Top             =   2010
         Width           =   3915
      End
   End
   Begin VB.PictureBox picOpretion 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   5415
      TabIndex        =   1
      Top             =   2805
      Width           =   5420
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&Q)"
         Height          =   450
         Left            =   3915
         TabIndex        =   3
         Top             =   105
         Width           =   1350
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&A)"
         Height          =   450
         Left            =   2415
         TabIndex        =   2
         Top             =   105
         Width           =   1350
      End
   End
End
Attribute VB_Name = "frmMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private blnInformation As Boolean

Private Enum MessageMode
    MM_�����ļ��嵥�޸� = 0
    MM_�����ļ� = 1
    MM_ɾ���ļ� = 2
End Enum

Public Function ShowMe(intMode As Integer, Optional strCaption As String = "", Optional strMessage As String = "") As Boolean
    If strCaption <> "" Then Me.Caption = strCaption

    Select Case intMode
        Case MM_�����ļ��嵥�޸�
            picInformation(MM_�����ļ��嵥�޸�).Visible = True
            If strMessage <> "" Then lblWarningCustom(MM_�����ļ��嵥�޸�).Caption = strMessage
        Case MM_�����ļ�
            picInformation(MM_�����ļ�).Visible = True
            If strMessage <> "" Then lblWarningCustom(MM_�����ļ�).Caption = strMessage
        Case MM_ɾ���ļ�
            picInformation(MM_ɾ���ļ�).Visible = True
            If strMessage <> "" Then lblWarningCustom(MM_ɾ���ļ�).Caption = strMessage
    End Select
    
    Call picInformation(intMode).Move(0, 0, Me.ScaleWidth, Me.ScaleHeight - picOpretion.Height)
    
    Me.Show 1, frmMDIMain
    ShowMe = blnInformation
End Function

Private Sub cmdCancel_Click()
    blnInformation = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    blnInformation = True
    Unload Me
End Sub

