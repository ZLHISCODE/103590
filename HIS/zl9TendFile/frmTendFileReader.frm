VERSION 5.00
Begin VB.Form frmTendFileReader 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����ļ���ӡ����������"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8565
   Icon            =   "frmTendFileReader.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Visible         =   0   'False
   Begin zl9TendFile.usrTendFileReader usrTendFileReader 
      Height          =   4875
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8505
      _extentx        =   15002
      _extenty        =   8599
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1785
      Left            =   6330
      ScaleHeight     =   1755
      ScaleWidth      =   2085
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "��ʱ��ͼ��,ǧ���ɾ"
      Top             =   1905
      Visible         =   0   'False
      Width           =   2115
   End
End
Attribute VB_Name = "frmTendFileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public blnReady As Boolean          'TRUE��ʾ��׼���ô�ӡ,FALSE��ʾ�������޷���ӡ

Private Sub Form_Load()
    Dim strPage As String
    
    strPage = frmAsk.mstrPrintPages          '���ʾ����ָ��ҳ��ʾ�ش�
    blnReady = usrTendFileReader.ShowMe(Me, glng�ļ�ID, glng����ID, glng��ҳID, gintӤ��, strPage)
    On Error Resume Next
    Me.Hide
    If Err <> 0 Then Err.Clear
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With usrTendFileReader
        .Left = 0
        .Top = 0
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With
End Sub

Public Function GetFixedProperty(ByVal strName As String) As Variant
    GetFixedProperty = usrTendFileReader.GetFixedProperty(strName)
End Function

Public Function GetFixedCol(ByVal strName As String) As Long
'�������ƻ�ȡ�̶�����Ϣ
    GetFixedCol = usrTendFileReader.GetFixedCol(strName)
End Function

Public Function NextPage() As Boolean
    If usrTendFileReader.isEndPage Then Exit Function
    NextPage = usrTendFileReader.NextPage
End Function

Public Function AppointPage(ByVal intPage As Integer) As Boolean
'ָ����ӡҳ,��Ҫ����ż��ӡʱʹ��
    AppointPage = usrTendFileReader.AppointPage(intPage)
End Function

Public Function GetPages() As Integer
    GetPages = usrTendFileReader.GetPages
End Function

Public Function GetStartPage() As Integer
    GetStartPage = usrTendFileReader.GetStartPage
End Function

Public Sub ShowPage(ByVal intPage As Integer)
    Call usrTendFileReader.ShowPage(intPage)
End Sub

Public Function PrintPage(Optional blnOddEvenPrint As Boolean = False, Optional ArrSQL As Variant) As Boolean
    '��ӡ��ǰҳ�����ݲ���¼��ӡ���
    PrintPage = usrTendFileReader.PrintPage(blnOddEvenPrint, ArrSQL)
End Function

Public Function GetFileName() As String
    GetFileName = usrTendFileReader.GetFileName
End Function

Public Function blnOddEvenPagePrint() As Boolean
    blnOddEvenPagePrint = usrTendFileReader.blnOddEvenPagePrint
End Function

Public Function GetCollectCols(ByVal lngRaw As Long) As String
    GetCollectCols = usrTendFileReader.GetCollectCols(lngRaw)
End Function

Public Function PrintHead() As Boolean
    PrintHead = usrTendFileReader.PrintHead
End Function

Public Function PrintFoot() As Boolean
    PrintFoot = usrTendFileReader.PrintFoot
End Function

Public Function GetBuffer() As Object
    Set GetBuffer = picBuffer
End Function


Public Function blnShowNullCollet() As Boolean
    blnShowNullCollet = usrTendFileReader.blnShowNullCollet
End Function
