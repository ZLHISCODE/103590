VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOutStatus 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1230
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5535
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmOutStatus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   1230
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame1 
      Height          =   1230
      Left            =   30
      MousePointer    =   11  'Hourglass
      TabIndex        =   0
      Top             =   -60
      Width           =   5475
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   150
         Left            =   150
         TabIndex        =   3
         Top             =   990
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   265
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����� 10%"
         Height          =   180
         Left            =   1080
         MousePointer    =   11  'Hourglass
         TabIndex        =   2
         Top             =   690
         Width           =   900
      End
      Begin VB.Label lblTitle 
         Caption         =   "�����������ӡ����"
         Height          =   210
         Left            =   255
         MousePointer    =   11  'Hourglass
         TabIndex        =   1
         Top             =   285
         Width           =   4965
      End
   End
End
Attribute VB_Name = "frmOutStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�����������������ӡ��
Public mintBegin As Integer, mintEnd As Integer

Private Sub Form_Activate()
    On Error Resume Next
    
'    Printer.Width = gsngPageWidth
'    Printer.Height = gsngPageHeight
'    Printer.Orientation = gintOri
    Set gobjOutTo = Printer
    
    Dim i As Integer
    Dim j As Integer
    
    Dim intTotal As Integer
    '������ʱ������ű���
    Dim sngTemp As Single
    
    intTotal = mintEnd - mintBegin + 2
    
    sngTemp = gsngScale
    gsngScale = 1
    
    Dim intTemp As Integer
    '��һ��������ŵ�ǰ��ҳ��
    intTemp = gintPage
    Printer.EndDoc '�˴�һִ�о�ʹ�����ɺ���
    Printer.Orientation = gintOri
    
    Dim lngMaxPage As Long, lngStartPage As Long
    lngMaxPage = frmTendFileReader.GetPages
    lngStartPage = frmTendFileReader.GetStartPage
    
    Dim intTempCopies As Integer
    Dim blnCopies As Boolean      '��ӡ�������óɹ�û��
    
    blnCopies = True
    intTempCopies = Printer.Copies
    Printer.Copies = gintCopies
    If Printer.Copies <> gintCopies Then blnCopies = False
    '������߲���ȣ���ʾ�ô�ӡ����������֧�ַ���
    
    '����ֽ�ţ��Զ���ֽ�ŵ����ñ���ŵ����
    If gintSize = 256 Then
        Call SetCustonPager(gsngPageWidth, gsngPageHeight)
    Else
        Printer.PaperSize = gintSize
    End If
    
    On Error GoTo errHandle
    For i = mintBegin To mintEnd
        lblNote.Caption = "�������" & CInt((i - mintBegin + 1) / intTotal * 100) & "%"
        ProgressBar1.Value = (i - mintBegin + 1) / intTotal * 100
        Me.Refresh
        ProgressBar1.Refresh
        '��ӡ֮ǰ���ú�ҳ�룬Ϊ����ʹҳ����ҳü��ȷ
        gintPage = i
        PrintPage i
        If Not frmTendFileReader.PrintPage Then Exit For
        If i <> mintEnd Then
            If Not frmTendFileReader.NextPage Then Exit For
            Printer.NewPage
        End If
    Next
    gintPage = intTemp
    lblNote.Caption = "�������100%"
    ProgressBar1.Value = 100
    Me.Refresh
    ProgressBar1.Refresh
    Printer.EndDoc
    If gintSize = 256 Then
        Call SetCustonPager(gsngPageWidth, gsngPageHeight)
    Else
        Printer.PaperSize = gintSize
    End If
    gsngScale = sngTemp
    
    On Error Resume Next
    Printer.Copies = intTempCopies
    Err.Clear
    Me.Hide
    Exit Sub
errHandle:
    MsgBox "��ӡ�����жϡ�", vbCritical + vbOKOnly, gstrSysName
    Me.Hide
End Sub
