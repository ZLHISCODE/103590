VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRptMan 
   BackColor       =   &H80000005&
   Caption         =   "�������"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5445
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmRptMan.frx":0000
   ScaleHeight     =   4275
   ScaleWidth      =   5445
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdEnter 
      Caption         =   "���ڽ��뱨����(&E)�� "
      Height          =   350
      Left            =   915
      TabIndex        =   1
      Top             =   3585
      Width           =   2190
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   120
      Top             =   1710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptMan.frx":04F9
            Key             =   "K0501"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptMan.frx":228B
            Key             =   "K0502"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptMan.frx":401D
            Key             =   "K0505"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   200
      TabIndex        =   2
      Top             =   100
      Width           =   960
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  'Transparent
      Height          =   3330
      Left            =   945
      TabIndex        =   0
      Top             =   600
      Width           =   4140
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   255
      Picture         =   "frmRptMan.frx":8E1F
      Top             =   600
      Width           =   480
   End
End
Attribute VB_Name = "frmRptMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstr��� As String    '������
Private mfrmProcMain As frmProcMain

Private Sub Form_Load()
    Select Case mstr���
        Case "0501"
            lblTitle.Caption = "�������"
            cmdEnter.Caption = "���ڽ��뱨����(&E)�� "
            lblMain.Caption = "���ϵͳ����Ʊ�ݸ�ʽ��������ݵĶ����޸ġ�" & _
                vbCrLf & vbCrLf & "�����������Ĳ��ԣ����ص�ͼԪ���Ʒ�ʽ��ͼ��Ԫ�ص�ѡ��棩����ȷ����Ʊ���뱨�������������ص���Ʊ�ݵ�ֽ������(��С������)�������ʽ�����塢��ɫ�����У�����������Ԥ����ӡ��" & _
                vbCrLf & vbCrLf & "��ǿSQL��ʵ��Ʊ������������ݵĸı䣬�Զ������д��ȷ�ԡ�" & _
                vbCrLf & vbCrLf & "������С��������������ߡ�ѡ�����ܣ�������������ϵͳ�Ļ��������������µĶ��ֱ���ʵ���������ݷ�����"
        Case "0502"
            lblTitle.Caption = "��������"
            cmdEnter.Caption = "���ڽ��뺯������(&E)�� "
            lblMain.Caption = "��ɸ�ϵͳ���ݴ��ݺ����Ĺ������������ı���������򵼵Ķ��塢�޸������á�" & _
                vbCrLf & vbCrLf & "���ݴ��ݺ����Ǳ������Ӧ��ϵͳ���໥��ѡ�������ݵ���Ҫ��ʽ��ʹ����Ӧ�ó�Ϊһ�����������壻�϶��Ӧ���ڲ��������Զ�ƾ֤���ɱ�Ч�����ͱ��������ȡ��Ӧ��ϵͳ�ķ������ݡ�" & _
                vbCrLf & vbCrLf & "���ϵͳװ��ʱ���Ѿ�װ����ṩһЩ����������ȡ�ĺ�������ͨ�û������Ȩ�󼴿�ʹ�ã�" & _
                vbCrLf & vbCrLf & "��Ҫʱ��ϵͳ�������û��ɸ���Ӧ����Ҫ�������µĺ������޸����к�����ʵ�ֶ�Ӧ��ϵͳ���������ݵ���ȡ��"
        Case "0505"
            lblTitle.Caption = "���̹���"
            cmdEnter.Caption = "���ڽ�����̹���(&E)"
            lblMain.Caption = "��ɸ�ϵͳ�������������У����Զ�����̵��޸ļ�����" & _
            vbCrLf & vbCrLf & "��ǰ���ݿ���ű��Աȣ��Զ��Ѽ��������Ĺ��̡�" & _
            vbCrLf & vbCrLf & "�Ա�����ǰ����̵ı仯�����ó���Ӧ�Ĳ���Աȱ��棬������Ա��ֱ�Ӹ��ݲ��첿�����������ж��޸���Ӧ�Ĺ��̡�"
    End Select
    Me.Caption = lblTitle.Caption
    imgMain.Picture = ils32.ListImages("K" & mstr���).Picture
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With imgMain
        .Top = 700
        .Left = ScaleLeft + 200
    End With
    
    With lblMain
        .Top = imgMain.Top
        .Height = ScaleHeight - .Top * 2
        .Left = imgMain.Left * 2 + imgMain.Width
        .Width = ScaleWidth - .Left - imgMain.Left
    End With

    Dim intCount As Integer, intRows As Integer, aryRow() As String
    intRows = 1
    aryRow() = Split(lblMain.Caption, vbCrLf)
    For intCount = 0 To UBound(aryRow)
        intRows = intRows + TextWidth(aryRow(intCount)) \ (lblMain.Width - 90) + 1
    Next
    If intRows * TextHeight("A") < lblMain.Height + TextHeight("A") Then
        cmdEnter.Top = lblMain.Top + intRows * TextHeight("A")
    Else
        cmdEnter.Top = lblMain.Top + lblMain.Height + TextHeight("A")
    End If
    cmdEnter.Left = lblMain.Left
    
End Sub

Private Sub cmdEnter_Click()
    Dim frmMain As frmConnectionsManager
    
    Select Case mstr���
        Case "0501"
            If gobjReport Is Nothing Then
                Set gobjReport = CreateObject("zl9Report.clsReport")
            End If
            Set frmMain = New frmConnectionsManager
            gobjReport.ReportMan gcnOracle, frmMDIMain, gstrLoginUserName, frmMain
        Case "0502"
            If gobjFunction Is Nothing Then
                Set gobjFunction = CreateObject("zl9Function.clsFunction")
            End If
            
            '�����������ݲ�֧��������
            gobjFunction.funcmanage gcnOldOra, frmMDIMain
        Case "0505"
            If mfrmProcMain Is Nothing Then
                Set mfrmProcMain = New frmProcMain
            End If
            Call mfrmProcMain.ShowMe(frmMDIMain)
    End Select
End Sub

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�

End Sub


