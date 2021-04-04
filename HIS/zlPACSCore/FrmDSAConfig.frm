VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmDSAConfig 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "���ּ�Ӱ����"
   ClientHeight    =   1410
   ClientLeft      =   30
   ClientTop       =   300
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      Height          =   330
      Left            =   4092
      TabIndex        =   6
      Top             =   828
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   330
      Left            =   4092
      TabIndex        =   5
      Top             =   312
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   828
      Left            =   360
      TabIndex        =   0
      Top             =   264
      Width           =   3396
      Begin VB.TextBox txtFrame 
         Height          =   300
         Left            =   1992
         TabIndex        =   2
         Top             =   276
         Width           =   552
      End
      Begin MSComCtl2.UpDown UpFrame 
         Height          =   300
         Left            =   2568
         TabIndex        =   3
         Top             =   276
         Width           =   252
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "֡"
         Height          =   180
         Left            =   2904
         TabIndex        =   4
         Top             =   336
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��׼ͼ��Ϊ�����еĵ�"
         Height          =   180
         Left            =   156
         TabIndex        =   1
         Top             =   336
         Width           =   1800
      End
   End
End
Attribute VB_Name = "FrmDSAConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintMaxFrame As Integer
Private mthisForm As frmViewer

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '��ͼ��������Ӱ
    Dim imgsTmp As New DicomImages
    Dim intViewerIndex As Integer
    Dim ww As Long
    Dim wl As Long
    
    '�����ж������Ƿ�Ϸ�
    If Val(txtFrame.Text) < 1 Then
        MsgBox "����ͼ��֡������С��1", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(txtFrame.Text) > mintMaxFrame Then
        MsgBox "����ͼ��֡�����ܴ���ͼ������֡��:" & mintMaxFrame, vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Ѽ�Ӱ��ͼ����Ϊ��ʱͼ����ӵ�Viewer��
    '���������ʾ��Viewer��������2�������ú�����ʾ2��
    If mthisForm.intCountX < 2 Then
        mthisForm.intCountX = 2
        Call subChangeSeriesLayout(mthisForm)
    End If
    
    '��Ҫ��Ӱ��ͼ����Ϊ��ʱͼ����ӵ�Viewer��
    imgsTmp.Add mthisForm.viewer(mthisForm.intSelectedSerial).Images(mthisForm.SelectedImageIndex)
    intViewerIndex = funShowTempImages(mthisForm, imgsTmp, 0)
    
    '������Ӱ
    mthisForm.viewer(intViewerIndex).Images(1).Mask = 1
    mthisForm.viewer(intViewerIndex).Images(1).MaskFrame = Val(txtFrame.Text)
    
    '��������λ
    If funAutoWinWL(imgsTmp(1), 0, 0, imgsTmp(1).sizex, imgsTmp(1).sizey, ww, wl) Then
        mthisForm.viewer(intViewerIndex).Images(1).width = ww
        mthisForm.viewer(intViewerIndex).Images(1).Level = wl
    End If
    
    '��ʾ��Ϊ��Ƭ����һ��ͼ��
    If Val(txtFrame.Text) = mthisForm.viewer(intViewerIndex).Images(1).FrameCount Then
        mthisForm.viewer(intViewerIndex).Images(1).Frame = 1
    Else
        mthisForm.viewer(intViewerIndex).Images(1).Frame = Val(txtFrame.Text) + 1
    End If
    
    '��ʾ��Ӱ��ɣ���˵���鿴����
    MsgBox "��ʹ�õ�Ӱ���Ź��ܲ鿴��ӰЧ��!", vbInformation, gstrSysName
    
    Unload Me
End Sub

Private Sub txtFrame_Change()
    If Val(txtFrame) < 1 Then
        MsgBox "����ͼ��֡������С��1", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(txtFrame) > mintMaxFrame Then
        MsgBox "����ͼ��֡�����ܴ���ͼ������֡��:" & mintMaxFrame, vbInformation, gstrSysName
        Exit Sub
    End If
    
    UpFrame.Value = Val(txtFrame)

End Sub

Private Sub UpFrame_Change()
    txtFrame = UpFrame.Value
End Sub

Public Sub zlShowMe(intMaxFrame As Integer, intCurrentFrame As Integer, thisForm As frmViewer)
    mintMaxFrame = intMaxFrame
    UpFrame.Max = mintMaxFrame
    UpFrame.Min = 1
    UpFrame.Value = intCurrentFrame
    txtFrame = intCurrentFrame
    Set mthisForm = thisForm
    Me.Show 1, mthisForm
End Sub
