VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmErrInfor 
   Caption         =   "��������"
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9135
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmErrInfor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   9135
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picDown 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   9135
      TabIndex        =   2
      Top             =   5790
      Width           =   9135
      Begin VB.Frame fraSplit 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   30
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   10695
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Ԥ��(&V)"
         Height          =   350
         Left            =   3945
         TabIndex        =   8
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "�����&Excel"
         Height          =   350
         Left            =   2505
         TabIndex        =   7
         Top             =   135
         Width           =   1425
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "��ӡ(&P)"
         Height          =   350
         Left            =   1365
         TabIndex        =   6
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "��ӡ����(&S)"
         Height          =   350
         Left            =   -15
         TabIndex        =   5
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "��������(&O)"
         Height          =   350
         Left            =   7455
         TabIndex        =   4
         Top             =   135
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "����(&C)"
         Height          =   350
         Left            =   6300
         TabIndex        =   3
         Top             =   135
         Width           =   1100
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsErr 
      Height          =   4680
      Left            =   60
      TabIndex        =   0
      Top             =   1020
      Width           =   8985
      _cx             =   15849
      _cy             =   8255
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   12632256
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmErrInfor.frx":06EA
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lblNote 
      Caption         =   "   ������ʱ,�������浥����Ϣ�к����쳣��Ϣ,��ȷ������쳣��Ϣ�Ƿ���ȷ,�����ȷ,�������������ʡ���ť,�������������ء���ť��"
      Height          =   600
      Left            =   1170
      TabIndex        =   1
      Top             =   345
      Width           =   7920
   End
   Begin VB.Image imgNotes 
      Height          =   720
      Left            =   165
      Picture         =   "frmErrInfor.frx":0764
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmErrInfor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mcllData As Collection
Private mblnOK As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdExcel_Click()
    Call zlPrint(3)
End Sub

Private Sub cmdOK_Click()
    mblnOK = True: Unload Me
End Sub

Private Sub cmdPreview_Click()
    Call zlPrint(2)
End Sub

Private Sub cmdPrint_Click()
    Call zlPrint(1)
End Sub

Private Sub cmdPrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With lblNote
        .Width = Me.ScaleWidth - .Left
        vsErr.Height = Me.ScaleHeight - picDown.Height - vsErr.Top - 100
        vsErr.Width = Me.ScaleWidth - vsErr.Left * 2 - 50
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'mblnOK = False
    SaveWinState Me, App.ProductName
End Sub

Private Sub picDown_Resize()
    Err = 0: On Error Resume Next
    With picDown
        fraSplit.Width = .ScaleWidth + 100
        cmdOK.Left = .ScaleWidth - cmdOK.Width - 100
        cmdCancel.Left = cmdOK.Left - cmdCancel.Width - 50
    End With
End Sub
Public Sub zlPrint(ByVal bytMode As Byte)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����б���Ϣ
    '���:bytMode=1-��ӡ,2-Ԥ��,3-�����Excel
    '����:���˺�
    '����:2013-09-13 10:23:30
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim i As Long, lngRow As Long, strTemp As String
    Dim rsTemp As ADODB.Recordset
    Err = 0: On Error GoTo ErrHand:
    '���������Ϣ
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstr��λ���� & "�շ�Ա���������������"
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    Set objPrint.Body = vsErr
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub
Public Function ShowErrInfor(ByVal frmMain As Object, _
    ByVal cllErrData As Collection, Optional ByVal blnOlnyOK As Boolean = False) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ������Ϣ
    '���:cllErrData-��������:Array(����, NO, ��¼״̬, ������, ��Ԥ��)
    '                          ����=1(�������ȷ;2.�쳣����)
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-22 11:37:28
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mcllData = cllErrData: mblnOK = False
    If mcllData Is Nothing Then Exit Function
    cmdCancel.Visible = Not blnOlnyOK
    If blnOlnyOK Then
        cmdOK.Caption = "ȷ��(&O)": cmdOK.Width = 1100
    End If
    Call LoadErrInfor
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    ShowErrInfor = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub LoadErrInfor()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ش�����Ϣ
    '����:���˺�
    '����:2013-09-11 17:34:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, int���� As Integer, lngRow As Long
    Dim strErrNO As String, varData As Variant
    Dim c As Long
    
    On Error GoTo errHandle
    With vsErr
        .Clear: .Redraw = flexRDNone
        .AutoResize = False
        .Rows = 1: .Cols = 4: lngRow = 0: .FixedRows = 0
        .MergeCells = flexMergeRestrictRows
        For c = 0 To .Cols - 1
            .ColWidth(c) = 1700
            .TextMatrix(lngRow, c) = "--���µ��ݽ���������ý�һ�£���˲� --"
        Next
        .MergeRow(lngRow) = True
        .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = &H8000000F
        .Rows = .Rows + 1: lngRow = lngRow + 1
        For c = 0 To .Cols - 1
             .TextMatrix(lngRow, c) = Switch(c = 0, "���ݺ�", c = 1, "���ý��", c = 2, "������", True, "���")
        Next
        .Cell(flexcpAlignment, lngRow, 0, lngRow, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = &H8000000F
        
        For i = 1 To mcllData.Count
            int���� = Val(mcllData(i)(0))
            If int���� = 2 Then '����=1(�������ȷ;2.�쳣����)
                'Array(����, NO, ��¼״̬, ������, ��Ԥ��)
                strErrNO = strErrNO & "," & Trim(mcllData(i)(1))
            Else
                .Rows = .Rows + 1: lngRow = lngRow + 1
                .TextMatrix(lngRow, 0) = Trim(mcllData(i)(1)) & IIf(Val(mcllData(i)(2)) = 2, "(��)", "")
                .TextMatrix(lngRow, 1) = Format(Val(mcllData(i)(3)), "###0.00;-###0.00;;")
                .TextMatrix(lngRow, 2) = Format(Val(mcllData(i)(4)), "###0.00;-###0.00;;")
                .TextMatrix(lngRow, 3) = Format(Val(mcllData(i)(3)) - Val(mcllData(i)(4)), "###0.00;-###0.00;;")
            End If
        Next
        If strErrNO <> "" Then
            strErrNO = Mid(strErrNO, 2)
            varData = Split(strErrNO, ",")
            .Rows = .Rows + 1: lngRow = lngRow + 1
            For c = 0 To .Cols - 1
                 .TextMatrix(lngRow, c) = "--���µ���Ϊ�쳣�շѵ���,��˲�--"
            Next
            .MergeRow(lngRow) = True
            .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = &H8000000F
            
            .Rows = .Rows + 1: lngRow = lngRow + 1
            c = 0
            For i = 0 To UBound(varData)
                If c > .Cols - 1 Then
                    c = 0
                    .Cell(flexcpAlignment, lngRow, 0, lngRow, .Cols - 1) = flexAlignCenterCenter
                    .Rows = .Rows + 1: lngRow = lngRow + 1
                End If
                .TextMatrix(lngRow, c) = varData(i)
                c = c + 1
            Next
        End If
        .Cell(flexcpAlignment, lngRow, 0, lngRow, .Cols - 1) = flexAlignCenterCenter
        .Redraw = flexRDBuffered
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

