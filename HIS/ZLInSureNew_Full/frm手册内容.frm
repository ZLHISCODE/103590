VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm�ֲ����� 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ֲ�����"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   Icon            =   "frm�ֲ�����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   7980
      TabIndex        =   2
      Top             =   150
      Width           =   1100
   End
   Begin VB.CommandButton cmd��ӡ 
      Caption         =   "��ӡ(&P)"
      Height          =   350
      Left            =   6720
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   150
      Width           =   1100
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   5475
      Index           =   0
      Left            =   30
      ScaleHeight     =   5415
      ScaleWidth      =   9315
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   660
      Width           =   9375
      Begin VB.VScrollBar vsScroll 
         Height          =   5385
         Index           =   0
         Left            =   9090
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   225
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshHead 
         Height          =   1155
         Index           =   0
         Left            =   90
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   870
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   2037
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   -2147483628
         BackColorBkg    =   -2147483628
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         ScrollBars      =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDeal 
         Height          =   1155
         Index           =   0
         Left            =   90
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2010
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   2037
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColorFixed  =   -2147483628
         BackColorBkg    =   -2147483628
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         ScrollBars      =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
      Begin VB.Label lbl��λ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��λ��Ԫ���ǡ���"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   7530
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�������ⲡ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   30
         TabIndex        =   4
         Top             =   180
         Width           =   9045
      End
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   5475
      Index           =   1
      Left            =   30
      ScaleHeight     =   5415
      ScaleWidth      =   9315
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   660
      Width           =   9375
      Begin VB.VScrollBar vsScroll 
         Height          =   5385
         Index           =   1
         Left            =   9090
         TabIndex        =   14
         Top             =   0
         Width           =   225
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshHead 
         Height          =   1155
         Index           =   1
         Left            =   90
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   870
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   2037
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   -2147483628
         BackColorBkg    =   -2147483628
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         ScrollBars      =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDeal 
         Height          =   1155
         Index           =   1
         Left            =   90
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2010
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   2037
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColorFixed  =   -2147483628
         BackColorBkg    =   -2147483628
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         ScrollBars      =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
      Begin VB.Label lbl��λ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��λ��Ԫ���ǡ���"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   7560
         TabIndex        =   11
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   30
         TabIndex        =   12
         Top             =   180
         Width           =   9045
      End
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������Ҫ�ڲ���ҽ���ֲ�����д�����ݣ�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   900
      TabIndex        =   0
      Top             =   240
      Width           =   3420
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frm�ֲ�����.frx":1272
      Stretch         =   -1  'True
      Top             =   90
      Width           =   480
   End
End
Attribute VB_Name = "frm�ֲ�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Private mblnStartup As Boolean
Private mdbl��ֵ As Double
Private mblnOutPatient As Boolean       '����
Private mrsHead As New ADODB.Recordset
Private mrsDeal As New ADODB.Recordset

Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����

Private Enum ҳ��
    ���� = 0
    סԺ
End Enum

Private Sub cmd��ӡ_Click()
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim intIndex As Integer
    Dim bytMode As Byte
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    
    bytMode = 1
    intIndex = IIf(mblnOutPatient, ����, סԺ)
    
    Set objPrint = New zlPrintGrds
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = Trim(lblTitle(intIndex).Caption)
        
    objRow.Add lbl��λ(intIndex).Caption
    objPrint.UnderAppRows.Add objRow
    
    Set objPrint.Grds = New Collection
    objPrint.Grds.Add mshHead(intIndex)
    objPrint.Grds.Add mshDeal(intIndex)
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewGrds objPrint, 1
          Case 2
              zlPrintOrViewGrds objPrint, 2
          Case 3
              zlPrintOrViewGrds objPrint, 3
      End Select
    Else
        zlPrintOrViewGrds objPrint, bytMode
    End If
End Sub

Private Sub cmdȷ��_Click()
    Unload Me
End Sub

Public Sub ShowBalance(ByVal rsHead As ADODB.Recordset, ByVal rsDeal As ADODB.Recordset, Optional ByVal bln���� As Boolean = True)
    mblnOutPatient = bln����
    Set mrsHead = rsHead
    Set mrsDeal = rsDeal
    Me.Show 1
End Sub

Private Sub Form_Activate()
    If Not mblnStartup Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim dbl�߶� As Double
    Dim objHead As MSHFlexGrid
    Dim objDeal As MSHFlexGrid
    
    '��������
    mdbl��ֵ = 0
    picBack(IIf(mblnOutPatient, ����, סԺ)).Visible = True
    picBack(IIf(mblnOutPatient, ����, סԺ)).ZOrder
    Set objHead = IIf(mblnOutPatient, mshHead(����), mshHead(סԺ))
    Set objDeal = IIf(mblnOutPatient, mshDeal(����), mshDeal(סԺ))
    Call InitStruct
    Call LoadData
    Call SetRowHeight(objHead)
    Call SetRowHeight(objDeal)
    
    '����λ��
    objHead.Height = objHead.Rows * 700
    objDeal.Height = objDeal.Rows * 700
    objDeal.Top = objHead.Top + objHead.Height
    mblnStartup = True
    If mblnOutPatient Then Exit Sub
    
    vsScroll(סԺ).Visible = (objDeal.Top + objDeal.Height > picBack(סԺ).Height)
    With vsScroll(סԺ)
        .Value = 0
        .Min = 0
        .Max = (mshDeal(סԺ).Top + mshDeal(סԺ).Height) / picBack(סԺ).Height
        .LargeChange = 1
    End With
    
    dbl�߶� = mshDeal(סԺ).Top + mshDeal(סԺ).Height
    '�����ֵ
    mdbl��ֵ = dbl�߶� / (vsScroll(סԺ).Max + 1)
End Sub

Private Sub InitStruct()
    Dim arrHead, arrDeal
    Dim intCol As Long, intCols As Integer
    Dim strHead As String, strDeal As String
    Dim objHead As MSHFlexGrid
    Dim objDeal As MSHFlexGrid
    Const strHead_���� As String = "ҽԺ����,2500|��������,2000|ҽԺ����,1500|�������,2900"
    Const strHead_סԺ As String = "ҽԺ����,2400|��Ժ-��Ժ���ڣ��ꡢ�¡��գ�,2500|ҽԺ����,1200|����" & vbCrLf & "���,500|��Ժ����,800|��;תԺ" & vbCrLf & "ת������,1500"
    Const strDeal_���� As String = "�����ܶ�,1200|ͳ��֧��,1200|���/" & vbCrLf & "����Ա֧��,1200|�����Ը�,1600|�����Է�,1200|ͳ��ⶥ��" & vbCrLf & "ҽ���ڽ��,1200|���ڡ�" & vbCrLf & "����ǩ��,1300"
    Const strDeal_סԺ As String = "�����ܶ�,1200|ͳ��֧��,1200|���/" & vbCrLf & "����Ա֧��,1200|�����Ը�,1600|�����Է�,1200|ͳ��ⶥ��" & vbCrLf & "ҽ���ڽ��,1200|���ڡ�" & vbCrLf & "����ǩ��,1300"
    
    If mblnOutPatient Then
        strHead = strHead_����
        strDeal = strDeal_����
        Set objHead = mshHead(����)
        Set objDeal = mshDeal(����)
    Else
        strHead = strHead_סԺ
        strDeal = strDeal_סԺ
        Set objHead = mshHead(סԺ)
        Set objDeal = mshDeal(סԺ)
    End If
    
    '���ñ�ͷ
    arrHead = Split(strHead, "|")
    intCols = UBound(arrHead)
    objHead.Cols = intCols + 1
    For intCol = 0 To intCols
        objHead.TextMatrix(0, intCol) = Split(arrHead(intCol), ",")(0)
        objHead.ColWidth(intCol) = Split(arrHead(intCol), ",")(1)
        objHead.ColAlignmentFixed(intCol) = 4
        objHead.ColAlignment(intCol) = IIf(intCol = intCols, 7, 1)
    Next
    '���ô������
    arrDeal = Split(strDeal, "|")
    intCols = UBound(arrDeal)
    objDeal.Cols = intCols + 1
    For intCol = 0 To intCols
        objDeal.TextMatrix(0, intCol) = Split(arrDeal(intCol), ",")(0)
        objDeal.ColWidth(intCol) = Split(arrDeal(intCol), ",")(1)
        objDeal.ColAlignmentFixed(intCol) = 4
        objDeal.ColAlignment(intCol) = IIf(intCol = 3, 1, 7)
    Next
End Sub

Private Sub LoadData()
    Dim objMsh As MSHFlexGrid
    '���ݼ�¼����������ʾ
    If mblnOutPatient Then
        '����ֻ������һ����¼
        Set objMsh = mshDeal(����)
        With mshHead(����)
            .TextMatrix(1, 0) = Nvl(mrsHead!ҽԺ����)
            .TextMatrix(1, 1) = Nvl(mrsHead!��������)
            .TextMatrix(1, 2) = Nvl(mrsHead!ҽԺ����)
            .TextMatrix(1, 3) = Nvl(mrsHead!�������)
        End With
    Else
        Set objMsh = mshDeal(סԺ)
        With mshHead(סԺ)
            If mrsHead.RecordCount <> 0 Then mrsHead.MoveFirst
            Do While Not mrsHead.EOF
                If mrsHead.AbsolutePosition > 1 Then .Rows = .Rows + 1
                .TextMatrix(mrsHead.AbsolutePosition, 0) = Nvl(mrsHead!ҽԺ����)
                .TextMatrix(mrsHead.AbsolutePosition, 1) = Nvl(mrsHead!��Ժ����) & "-" & Nvl(mrsHead!ת������)
                .TextMatrix(mrsHead.AbsolutePosition, 2) = Nvl(mrsHead!ҽԺ����)
                .TextMatrix(mrsHead.AbsolutePosition, 3) = Nvl(mrsHead!�������)
                .TextMatrix(mrsHead.AbsolutePosition, 4) = Nvl(mrsHead!��Ժ����)
                .TextMatrix(mrsHead.AbsolutePosition, 5) = Nvl(mrsHead!ת������)
                mrsHead.MoveNext
            Loop
        End With
    End If
    
    'ͳһ��������Ϣд�루��ʽһ����
    With objMsh
        If mrsDeal.RecordCount <> 0 Then mrsDeal.MoveFirst
        Do While Not mrsDeal.EOF
            If mrsDeal.AbsolutePosition > 1 Then .Rows = .Rows + 1
            .TextMatrix(mrsDeal.AbsolutePosition, 0) = Format(Nvl(mrsDeal!�����ܶ�, 0), "#0.00;-#0.00;0.00;")
            .TextMatrix(mrsDeal.AbsolutePosition, 1) = Format(Nvl(mrsDeal!ͳ��֧��, 0), "#0.00;-#0.00;0.00;")
            .TextMatrix(mrsDeal.AbsolutePosition, 2) = Format(Nvl(mrsDeal!���֧��, 0), "#0.00;-#0.00;0.00;")
            .TextMatrix(mrsDeal.AbsolutePosition, 3) = "�Ը�1��" & Format(Nvl(mrsDeal!�����Ը�, 0), "#0.00;-#0.00;0.00;") & _
                vbCrLf & "�Ը�2��" & Format(Nvl(mrsDeal!�����Ը�, 0), "#0.00;-#0.00;0.00;")
            .TextMatrix(mrsDeal.AbsolutePosition, 4) = Format(Nvl(mrsDeal!�����Է�, 0), "#0.00;-#0.00;0.00;")
            .TextMatrix(mrsDeal.AbsolutePosition, 5) = Format(Nvl(mrsDeal!ͳ��ⶥ��ҽ���ڽ��, 0), "#0.00;-#0.00;0.00;")
            .TextMatrix(mrsDeal.AbsolutePosition, 6) = Nvl(mrsDeal!��������)
            mrsDeal.MoveNext
        Loop
    End With
End Sub

Private Sub SetRowHeight(ByVal objMsh As MSHFlexGrid)
    Dim intRow As Integer, intRows As Integer
    intRows = objMsh.Rows - 1
    For intRow = 0 To intRows
        objMsh.RowHeight(intRow) = 700
    Next
End Sub

Private Sub vsScroll_Change(Index As Integer)
    Static intValue As Integer          '�ϴε�ֵ
    Dim intCur As Integer               '��ǰ��ֵ
    Dim dbl��ֵ As Double
    If Index = ���� Then Exit Sub
    
    intCur = vsScroll(Index).Value
    dbl��ֵ = mdbl��ֵ * (intValue - intCur)
    intValue = intCur
    
    '�ƶ����пؼ�
    picBack(Index).AutoRedraw = False
    lblTitle(Index).Top = lblTitle(Index).Top + dbl��ֵ
    lbl��λ(Index).Top = lbl��λ(Index).Top + dbl��ֵ
    mshHead(Index).Top = mshHead(Index).Top + dbl��ֵ
    mshDeal(Index).Top = mshDeal(Index).Top + dbl��ֵ
    picBack(Index).AutoRedraw = True
End Sub
