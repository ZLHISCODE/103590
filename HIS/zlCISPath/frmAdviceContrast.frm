VERSION 5.00
Begin VB.Form frmAdviceContrast 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ҽ���ԱȲ鿴"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12480
   Icon            =   "frmAdviceContrast.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   12480
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox PicInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   12480
      TabIndex        =   10
      Top             =   0
      Width           =   12480
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ·����Ŀ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1080
         TabIndex        =   12
         Top             =   75
         Width           =   1170
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "    "
         Height          =   360
         Left            =   1320
         TabIndex        =   11
         Top             =   360
         Width           =   10245
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   120
         Picture         =   "frmAdviceContrast.frx":6633E
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   12480
      TabIndex        =   4
      Top             =   7215
      Width           =   12480
      Begin VB.CommandButton cmdQuit 
         Cancel          =   -1  'True
         Caption         =   "�˳�(&Q)"
         Height          =   350
         Left            =   11280
         TabIndex        =   9
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "��һ��&N)"
         Height          =   350
         Index           =   1
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "��һ��(&P)"
         Height          =   350
         Index           =   0
         Left            =   8880
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1100
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   16080
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   13920
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Frame fraBottom 
      BorderStyle     =   0  'None
      Caption         =   "�����(��һ�汾)"
      Height          =   3285
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   12375
      Begin zlCISPath.UCAdviceList UCAdviceOld 
         Height          =   2895
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   12135
         _ExtentX        =   22251
         _ExtentY        =   5106
      End
      Begin VB.Label lblOldInfo 
         AutoSize        =   -1  'True
         Caption         =   "�����(��һ�汾)"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1440
      End
   End
   Begin VB.Frame fraTop 
      BorderStyle     =   0  'None
      Caption         =   "�����(��ǰ�汾)"
      Height          =   2925
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   12255
      Begin zlCISPath.UCAdviceList UCAdviceNew 
         Height          =   2415
         Left            =   0
         TabIndex        =   2
         Top             =   360
         Width           =   12015
         _ExtentX        =   21616
         _ExtentY        =   5106
      End
      Begin VB.Label lblNewInfo 
         AutoSize        =   -1  'True
         Caption         =   "δ���(��ǰ�汾)"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   1440
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   0
      X2              =   11880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   2
      X1              =   0
      X2              =   11880
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "frmAdviceContrast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event MovePathItemFocus(ByVal lngItemID As Long) '������ĿID,ʹ��������ʾ����������Ŀ�����ƶ�����һ�£���������ǰ��Ŀ����
Private mlngNewId   As Long
Private mlngOldId   As Long
Private mcolItemID  As Collection                       'item�� �°���ĿID:�ϰ���ĿID:�±�λ��
Private mintMode    As Integer                          '1-���0-סԺ

Private Enum Move_Index
    MovePrev = 0
    MoveNext = 1
End Enum

Private Sub cmdMove_Click(Index As Integer)
    Dim lngCurrIndex As Long
    Dim lngTmp As Long

    lngCurrIndex = Split(mcolItemID("_" & mlngNewId), ":")(2)
    Select Case Index

    Case MovePrev
        If lngCurrIndex = 2 Then
            cmdMove(Index).Enabled = False
        Else
            cmdMove(Index).Enabled = True
        End If
        cmdMove(MoveNext).Enabled = True

    Case MoveNext
        If lngCurrIndex = mcolItemID.count - 1 Then
            cmdMove(Index).Enabled = False
        Else
            cmdMove(Index).Enabled = True
        End If
        cmdMove(MovePrev).Enabled = True
    End Select
    
    lngTmp = IIf(Index = MovePrev, -1, 1)
    mlngNewId = Split(mcolItemID(lngCurrIndex + lngTmp), ":")(0)
    mlngOldId = Split(mcolItemID(lngCurrIndex + lngTmp), ":")(1)
    
    '��������
    Call LoadData
      
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub LoadData()
'����: �������ݴ��ڲ����·����Ŀҽ���嵥
    Dim strSql As String
    Dim lngCurrIndex As Long
    If mintMode = 1 Then
         strSql = "Select a.Id, a.���id, a.���, a.��Ч, a.������Ŀid, a.�շ�ϸĿid, a.ҽ������, a.��������, a.�ܸ�����, a.�걾��λ, a.��鷽��, a.ҽ������, a.ִ��Ƶ��, a.Ƶ�ʴ���," & vbNewLine & _
                 "       a.Ƶ�ʼ��, a.�����λ, a.ִ������, a.ִ�п���id, a.ʱ�䷽��, a.�Ƿ�ȱʡ, a.�Ƿ�ѡ, a.�䷽id, a.�����Ŀid,a.ִ�б�� " & vbNewLine & _
                 "From ����·��ҽ������ A, ����·��ҽ�� B" & vbNewLine & _
                 "Where a.Id = b.ҽ������id And b.·����Ŀid =[3] "
    Else
        strSql = "Select a.Id, a.���id, a.���, a.��Ч, a.������Ŀid, a.�շ�ϸĿid, a.ҽ������, a.��������, a.�ܸ�����, a.�걾��λ, a.��鷽��, a.ҽ������, a.ִ��Ƶ��, a.Ƶ�ʴ���," & vbNewLine & _
                 "       a.Ƶ�ʼ��, a.�����λ, a.ִ������, a.ִ�п���id, a.ʱ�䷽��, a.�Ƿ�ȱʡ, a.�Ƿ�ѡ, a.�䷽id, a.�����Ŀid,a.ִ�б�� " & vbNewLine & _
                 "From ·��ҽ������ A, �ٴ�·��ҽ�� B" & vbNewLine & _
                 "Where a.Id = b.ҽ������id And b.·����Ŀid =[3] "
    End If
    
    UCAdviceNew.ShowAdvice 0, strSql, , , True, mlngNewId
    UCAdviceOld.ShowAdvice 0, strSql, , , True, mlngOldId

    lngCurrIndex = Split(mcolItemID("_" & mlngNewId), ":")(2)

    If lngCurrIndex = 1 Then
        cmdMove(MovePrev).Enabled = False
    End If
    If lngCurrIndex = mcolItemID.count Then
        cmdMove(MoveNext).Enabled = False
    End If
    
    'ʹ��������ʾ����������Ŀ�����ƶ�����һ�£���������ǰ��Ŀ����
    RaiseEvent MovePathItemFocus(mlngNewId)
End Sub

Private Sub Form_Load()
    Call LoadData
End Sub

Private Sub Form_Resize()

    On Error Resume Next
  
    fraTop.Move 120, picInfo.Height, Me.ScaleWidth - 240, (Me.ScaleHeight - picBottom.Height - picInfo.Height) / 2
    fraBottom.Move 120, fraTop.Top + fraTop.Height, Me.ScaleWidth - 240, (Me.ScaleHeight - picBottom.Height - picInfo.Height) / 2
    
    With lblNewInfo
        .Top = 50: .Left = 0
    End With

    With lblOldInfo
         .Top = 50: .Left = 0
    End With

    With UCAdviceNew
        .Left = 0: .Top = lblNewInfo.Height + 50
        .Width = fraTop.Width: .Height = fraTop.Height - lblNewInfo.Height
    End With
    With UCAdviceOld
        .Left = 0: .Top = lblOldInfo.Height + 50
        .Width = UCAdviceNew.Width: .Height = UCAdviceNew.Height
    End With
End Sub

Public Sub ShowMe(frmParent As Object, ByVal lngNewId As Long, ByVal lngOldId As Long, ByVal colItemID As Collection, Optional ByVal intMode As Integer)
    mlngNewId = lngNewId
    mlngOldId = lngOldId
    mintMode = intMode
    Set mcolItemID = colItemID
    Me.Show 1, frmParent
End Sub

Public Sub SetNoteInfo(ByVal strInfo As String)
'����:���õ�ǰѡ����Ŀ����
    lblNote.Caption = strInfo
End Sub




