VERSION 5.00
Begin VB.Form frmDeviceSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�豸����"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4395
   StartUpPosition =   1  '����������
   Begin VB.Frame fra 
      Caption         =   "�豸ѡ��"
      Height          =   3360
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton cmdDevice 
         Caption         =   "��������豸(&P)"
         Height          =   345
         Index           =   6
         Left            =   375
         TabIndex        =   8
         Top             =   2745
         Width           =   1965
      End
      Begin VB.CommandButton cmdDevice 
         Caption         =   "���㿨�豸(&J)"
         Height          =   350
         Index           =   5
         Left            =   360
         TabIndex        =   7
         Top             =   2280
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.CommandButton cmdDevice 
         Caption         =   "IC���豸(&C)"
         Height          =   350
         Index           =   4
         Left            =   360
         TabIndex        =   6
         Top             =   1800
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.CommandButton cmdDevice 
         Caption         =   "LED�豸(&L)"
         Height          =   350
         Index           =   1
         Left            =   360
         TabIndex        =   0
         Top             =   840
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.CommandButton cmdDevice 
         Caption         =   "���֤ʶ���豸(&I)"
         Height          =   350
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.CommandButton cmdDevice 
         Caption         =   "˰�ش�ӡ�豸(&M)"
         Height          =   350
         Index           =   2
         Left            =   360
         TabIndex        =   2
         Top             =   1320
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.CommandButton cmdDevice 
         Caption         =   "˰�ش�ӡ�豸(&Z)"
         Height          =   350
         Index           =   3
         Left            =   360
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
         Width           =   1965
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&X)"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmDeviceSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Private mintOnlyOne As Integer
Private mblnDoing As Boolean

Private zl9LedVoice As Object
Private zl9IDCard As Object
Private zl9TaxBill As Object
Private zl9ESign As Object
Private zl9ICCard As Object
Private zl9SquareCard As Object
Private zl9keyboard As Object

Private Const SPACEHEIGHT = 300
Private Enum Kind
    C0IDCard = 0
    C1LED = 1
    C2OutTax = 2
    C3InTax = 3
    C4ICCard = 4
    C5SquareCard = 5
    C6Keyboard = 6
End Enum

Private Sub SetHeight()
    Dim i As Long, lngTmp As Long, j As Long
    
    mblnDoing = True
    For i = 0 To cmdDevice.UBound
        If cmdDevice(i).Visible Then
            If i > lngTmp Then lngTmp = i
            mintOnlyOne = i
            j = j + 1
        End If
    Next
    lngTmp = fra.Top + fra.Height - (cmdDevice(lngTmp).Top + cmdDevice(lngTmp).Height) - SPACEHEIGHT
    fra.Height = fra.Height - lngTmp
    Me.Height = Me.Height - lngTmp
    
    If j <> 1 Then mintOnlyOne = -1
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If mintOnlyOne >= 0 Then
        Me.Hide
        Call cmdDevice_Click(mintOnlyOne)
        Call cmdOK_Click
    End If
End Sub

Private Sub Form_Resize()
    If mblnDoing Then Exit Sub
    
    Call SetHeight
    mblnDoing = False
End Sub


Public Sub ShowMe(frmParent As Object, lngSys As Long, lngModule As Long)
    'ȱʡ�����豸��ť���ɼ�,����ģ��Ż��Ҳ�����ģ���ʱ,���а�ť�ɼ�
    Dim i As Long
    
    mintOnlyOne = -1
    mblnDoing = False
    cmdDevice(Kind.C6Keyboard).Visible = True   '���еط����������豸����
    cmdDevice(Kind.C5SquareCard).Visible = True    '���еط����н��㿨���豸����
    If lngSys = 100 Then
        Select Case lngModule
        Case 1101   '������Ϣ����
           cmdDevice(Kind.C0IDCard).Visible = True
           cmdDevice(Kind.C4ICCard).Visible = True
        Case 1102   '���￨
           cmdDevice(Kind.C0IDCard).Visible = True
        Case 1103   'Ԥ����
            cmdDevice(Kind.C0IDCard).Visible = True
            cmdDevice(Kind.C1LED).Visible = True
            cmdDevice(Kind.C4ICCard).Visible = True
        Case 1111   '�Һ�
            cmdDevice(Kind.C0IDCard).Visible = True
            cmdDevice(Kind.C1LED).Visible = True
        Case 1121   '�����շ�
            cmdDevice(Kind.C0IDCard).Visible = True
            cmdDevice(Kind.C1LED).Visible = True
            cmdDevice(Kind.C2OutTax).Visible = True
            cmdDevice(Kind.C4ICCard).Visible = True
        Case 1120   '���ﻮ��
            cmdDevice(Kind.C0IDCard).Visible = True
            cmdDevice(Kind.C4ICCard).Visible = True
        Case 1122   '�������
            cmdDevice(Kind.C0IDCard).Visible = True
            cmdDevice(Kind.C4ICCard).Visible = True
        Case 1131   '��Ժ�Ǽ�
           cmdDevice(Kind.C0IDCard).Visible = True
           cmdDevice(Kind.C1LED).Visible = True
           cmdDevice(Kind.C4ICCard).Visible = True
        Case 1137   '����
            cmdDevice(Kind.C0IDCard).Visible = True
            cmdDevice(Kind.C1LED).Visible = True
            cmdDevice(Kind.C3InTax).Visible = True
            cmdDevice(Kind.C3InTax).Top = cmdDevice(Kind.C2OutTax).Top
            cmdDevice(Kind.C4ICCard).Visible = True
            cmdDevice(Kind.C5SquareCard).Visible = True
                   
        Case 1260, 1263 '����ҽ��վ,ҽ������վ
            cmdDevice(Kind.C0IDCard).Visible = True
            cmdDevice(Kind.C4ICCard).Visible = True
        Case 1536   '�����ѯ
            cmdDevice(Kind.C0IDCard).Visible = True
            cmdDevice(Kind.C4ICCard).Visible = True
        Case 1341   'ҩƷ������ҩ
            cmdDevice(Kind.C0IDCard).Visible = True
            cmdDevice(Kind.C4ICCard).Visible = True
        Case 1342   'ҩƷ���ŷ�ҩ
            cmdDevice(Kind.C4ICCard).Visible = True
        Case 1723   '���ķ��Ź���
            cmdDevice(Kind.C4ICCard).Visible = True
        Case 1804   '�����ҹ���վ
            cmdDevice(Kind.C0IDCard).Visible = True
            cmdDevice(Kind.C4ICCard).Visible = True
        Case 1503   '���㿨����
            cmdDevice(Kind.C5SquareCard).Visible = True
        Case Else
            For i = 0 To cmdDevice.UBound
                cmdDevice(i).Visible = True
            Next
        End Select
    ElseIf lngSys = 2200 Then
        Select Case lngModule
        Case 1935, 1938     '������Ѫ����,��Ѫ��Ӧ��¼
            cmdDevice(Kind.C0IDCard).Visible = True
            cmdDevice(Kind.C4ICCard).Visible = True
        End Select
    ElseIf lngSys = 2100 Then
        Select Case lngModule
        Case 2106, 2121, 2122, 2123, 2124, 2125, 2126        '�ͻ���ϵ����,������Ĺ���,���ֿ�ִ��,������Ǽ�,��켯�еǼ�,����ܼ챨��,�����ù���
            cmdDevice(Kind.C0IDCard).Visible = True
            cmdDevice(Kind.C4ICCard).Visible = True
        Case Else
            For i = 0 To cmdDevice.UBound
                cmdDevice(i).Visible = True
            Next
        End Select
    End If
    Me.Show vbModal, frmParent
End Sub

Private Sub cmdDevice_Click(Index As Integer)
    On Error GoTo errH
    Select Case Index
        Case Kind.C1LED
            If zl9LedVoice Is Nothing Then
                Set zl9LedVoice = CreateObject("zl9LedVoice.ClsLedVoice")
            End If
            zl9LedVoice.VoiceSetting
        Case Kind.C0IDCard
            If zl9IDCard Is Nothing Then
                Set zl9IDCard = CreateObject("zlIDCard.clsIDCard")
            End If
            zl9IDCard.ParameterSet
        Case Kind.C2OutTax, Kind.C3InTax
            If zl9TaxBill Is Nothing Then
                Set zl9TaxBill = CreateObject("zl9TaxBill.ClsTaxBill")
            End If
            Call zl9TaxBill.zlTaxBillSet(gcnOracle, IIf(Index = Kind.C2OutTax, 1, 2))
        Case Kind.C4ICCard
            If zl9ICCard Is Nothing Then
                Set zl9ICCard = CreateObject("zlICCard.clsICCard")
            End If
            zl9ICCard.Set_Card
        Case Kind.C5SquareCard
        
            If zl9SquareCard Is Nothing Then
                Set zl9SquareCard = CreateObject("zl9CardSquare.clsCardSquare")
                ';zlInitComponents(ByVal frmMain As Object, _
                ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
                ByVal cnOracle As ADODB.Connection, _
                Optional blnDeviceSet As Boolean = False, _
                Optional strExpand As String) As Boolean
                If zl9SquareCard.zlInitComponents(Me, 0, 0, gstrDBUser, gcnOracle, True) = False Then
                        Set zl9SquareCard = Nothing: Exit Sub
                End If
            End If
            Call zl9SquareCard.zlCardDevSet(Me, 0)
        Case Kind.C6Keyboard
            If zl9keyboard Is Nothing Then
                Set zl9keyboard = CreateObject("zl9keyboard.clskeyboard")
            End If
            Call zl9keyboard.zlCardDevSet(Me)
    End Select
    Exit Sub
errH:
    MsgBox Err.Description, vbInformation, gstrSysName
End Sub


