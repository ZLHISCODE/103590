VERSION 5.00
Begin VB.Form frmCaseTendBodyDelLine 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ͼ�������߳�"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   Icon            =   "frmCaseTendBodyDelLine.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ListBox lst 
      Height          =   3000
      Left            =   795
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   1230
      Width           =   2730
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   3855
      TabIndex        =   5
      Top             =   3420
      Width           =   1100
   End
   Begin VB.Frame fraBottom 
      Height          =   9075
      Left            =   3720
      TabIndex        =   3
      Top             =   -150
      Width           =   30
   End
   Begin VB.CommandButton cmdCanc 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3855
      TabIndex        =   2
      Top             =   675
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3840
      TabIndex        =   1
      Top             =   180
      Width           =   1100
   End
   Begin VB.Label lblAsk 
      Caption         =   "���ϣ���߳�������������(����ȥ������Ҫ�߳�����Ŀ)"
      Height          =   375
      Left            =   705
      TabIndex        =   4
      Top             =   180
      Width           =   2985
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmCaseTendBodyDelLine.frx":000C
      Top             =   150
      Width           =   480
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "ʱ�䣺2001��11��16�� 4��8ʱ"
      Height          =   180
      Left            =   825
      TabIndex        =   0
      Top             =   630
      Width           =   2430
   End
End
Attribute VB_Name = "frmCaseTendBodyDelLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintNowCol As Integer
Private mblnChanged As Boolean
Public mfrmParent As Object
Private mvar������ As Long
Private mlng����ȼ� As Long
Private mintBaby As Integer
Private mArrTmp As Variant

Public Function ShowEdit(ByVal frmParent As Object, ByVal intNowCol As Integer, ByVal lng����ȼ� As Long, ByVal intBaby As Integer, Optional Marr���� As Variant) As Boolean
    
    mblnChanged = False

    mintNowCol = intNowCol
    mintBaby = intBaby
'
    mlng����ȼ� = lng����ȼ�
    mArrTmp = Marr����
    
    Set mfrmParent = frmParent
    
    Call InitData
        
    Me.Show 1
    
    ShowEdit = mblnChanged
    
End Function

Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrHandle
    Dim aryValue() As String
    Dim intRewrite As Integer
    Dim aryPara() As String
    Dim intCount As Integer
    Dim intValue As Integer
    Dim arrCode() As String
    Dim strTmp As String
    
    With mfrmParent.GetmshScale
        '����ע��˵��
        If lst.Selected(mvar������) Then
            aryValue = Split(.TextMatrix(0, mintNowCol + .FixedCols), ";")
            intRewrite = Val(aryValue(0))
            Select Case intRewrite
            Case 0
                aryValue(0) = 0
            Case 1
                aryValue(0) = 4
            Case 2
                aryValue(0) = 0
            Case 3
                aryValue(0) = 4
            Case 4
                aryValue(0) = 4
            End Select
            .TextMatrix(0, mintNowCol + .FixedCols) = Join(aryValue, ";")
            .TextMatrix(2, mintNowCol + .FixedCols) = ""
        End If
        
        '������������
        For intCount = 0 To mvar������ - 1
            If lst.Selected(intCount) Then
                
                If lst.ItemData(intCount) = 3 And mfrmParent.������� = True Then
                    intRewrite = mfrmParent.Getvsf.ColData(mintNowCol + mfrmParent.Getvsf.FixedCols)
                    Select Case intRewrite
                    Case 0
                        intValue = 0
                    Case 1
                        intValue = 4
                    Case 2
                        intValue = 0
                    Case 3
                        intValue = 4
                    Case 4
                        intValue = 4
                    End Select
                    mfrmParent.Getvsf.ColData(mintNowCol + mfrmParent.Getvsf.FixedCols) = intValue
                    mfrmParent.Getvsf.TextMatrix(1, mintNowCol + mfrmParent.Getvsf.FixedCols) = ""
                    
                    If UBound(mArrTmp) >= mintNowCol Then
                        strTmp = mArrTmp(mintNowCol)
                        strTmp = strTmp & String(2 - UBound(Split(strTmp, "-")), "-")
                        arrCode = Split(strTmp, "-")
                        arrCode(0) = ""
                        arrCode(1) = ""
                        arrCode(2) = ""
                        mArrTmp(mintNowCol) = Join(arrCode, "-")
                    End If
                
                    mfrmParent.Marr���� = mArrTmp
                Else
                    aryValue = Split(.TextMatrix(0, mintNowCol + .FixedCols), ";")
                    intRewrite = Val(aryValue(intCount + 1))
                    Select Case intRewrite
                    Case 0
                        aryValue(intCount + 1) = 0
                    Case 1
                        aryValue(intCount + 1) = 4
                    Case 2
                        aryValue(intCount + 1) = 0
                    Case 3
                        aryValue(intCount + 1) = 4
                    Case 4
                        aryValue(intCount + 1) = 4
                    End Select
                    .TextMatrix(0, mintNowCol + .FixedCols) = Join(aryValue, ";")
                    
                    aryValue = Split(.TextMatrix(1, mintNowCol + .FixedCols), ";")
                    aryValue(intCount + 1) = ""
                    .TextMatrix(1, mintNowCol + .FixedCols) = Join(aryValue, ";")
                End If
            End If
        Next
    End With
    
    '�����ϼ��������ͼ�δ���
    Call mfrmParent.DrawPaper
    Call mfrmParent.DrawGraph
    
    mblnChanged = True
    
    Unload Me
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitData() As Boolean
    Dim aryValue() As String
    Dim dtNow As Date
    Dim lngHourBegin As Long

    Call RefreshList
    
    lngHourBegin = Val(zlDatabase.GetPara(67, glngSys, , 4))

    aryValue = Split(mfrmParent.GetPicScale.Tag, ";")
    dtNow = Int(CDate(aryValue(0))) + ((mintNowCol - 1) * 4 + lngHourBegin) / 24
    If Val(Format(dtNow, "hh")) + 4 > 23 Then
        Me.lblTime = "ʱ�䣺" & Format(dtNow, "yyyy��MM��DD��") & "" & Val(Format(dtNow, "hh")) & "ʱ" & vbCrLf & "     ��" & Format(DateAdd("d", 1, dtNow), "yyyy��MM��DD��") & Val(Format(dtNow, "hh")) + 4 - 24 & "ʱ"
    Else
        Me.lblTime = "ʱ�䣺" & Format(dtNow, "yyyy��MM��DD��") & " " & Val(Format(dtNow, "hh")) & "��" & Val(Format(dtNow, "hh")) + 4 & "ʱ"
    End If
End Function

Private Sub RefreshList()
    Dim i As Long
    Dim rsTmp As New ADODB.Recordset
    
    
    On Error GoTo ErrHandle
    
    '�õ����е�������Ŀ
    gstrSQL = "SELECT A.��Ŀ���,A.��¼�� FROM ���¼�¼��Ŀ A,�����¼��Ŀ B " & _
            "WHERE A.��¼�� =1 And A.��Ŀ���=B.��Ŀ��� and Nvl(b.Ӧ�÷�ʽ,0)=1 And Nvl(b.���ò���,0) In (0,[2]) " & _
                    "AND B.����ȼ�>=[1] " & _
            "ORDER BY A.�������"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ȼ�, IIf(mintBaby = 0, 1, 2))
    With rsTmp
        If .RecordCount > 0 Then
            .MoveFirst
            mvar������ = .RecordCount
            For i = 1 To .RecordCount
                lst.AddItem IIf(IsNull(rsTmp!��¼��), "", rsTmp!��¼��)
                lst.ItemData(lst.NewIndex) = Val(NVL(rsTmp!��Ŀ���))
                .MoveNext
            Next
        End If
    End With
    
    '����Ϊ���ʱҲ��ʾ
    If mfrmParent.������� = True Then
        mvar������ = mvar������ + 1
        lst.AddItem "����"
        lst.ItemData(lst.NewIndex) = 3
    End If
    
    lst.AddItem "˵  ��"
    lst.ItemData(lst.NewIndex) = 0
    For i = 0 To lst.ListCount - 1
        lst.Selected(i) = True
    Next
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

