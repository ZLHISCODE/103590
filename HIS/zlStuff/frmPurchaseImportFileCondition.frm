VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frmPurchaseImportFileCondition 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8745
   Icon            =   "frmPurchaseImportFileCondition.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�"
      Height          =   300
      Left            =   7200
      TabIndex        =   5
      Top             =   240
      Width           =   885
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����"
      Height          =   300
      Left            =   6120
      TabIndex        =   4
      Top             =   240
      Width           =   885
   End
   Begin VB.OptionButton optFullImport 
      Caption         =   "��ȫ����"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.OptionButton optPartImport 
      Caption         =   "����ȫ����"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfError 
      Height          =   4485
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   7215
      _cx             =   12726
      _cy             =   7911
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
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
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   17
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPurchaseImportFileCondition.frx":000C
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
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
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   0   'False
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
   Begin VB.Label lblImportMethod 
      AutoSize        =   -1  'True
      Caption         =   "���뷽ʽ"
      Height          =   180
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmPurchaseImportFileCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MCONFIXECOLOR As Long = &H8000000F  '�����޸��б���ɫ
Private strPara As String   '����ֵ������Ϊ���뷽ʽ/���ı���|����|�ɱ���|�ɱ����|��Ʊ���|����*�ɱ���=�ɱ����|��Ʊ���=�ɱ����|���ɱ���=HIS�ɱ���|Ч��|�������|���Ч��|��������|�洢�ⷿ|����ⷿ|��Ʒ����(0-����ȫ����1-��ȫ����/0-��ʾ1-��ֹ|....)
Private mlngModal As Long '��ǰģ���

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strTemp As String
    Dim intRow As Integer
    
    With vsfError
        If optFullImport.Value = True Then
            strTemp = "1/"
        Else
            strTemp = "0/"
        End If
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) = "��ֹ" Then
                strTemp = strTemp & "1|"
            Else
                strTemp = strTemp & "0|"
            End If
        Next
    End With
    If strTemp <> "" Then
        strTemp = Mid(strTemp, 1, LenB(StrConv(strTemp, vbFromUnicode)) - 1)
    Else
        strTemp = "0/0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"
    End If
    Call zlDatabase.SetPara("�����ļ���鷽ʽ", strTemp, glngSys, mlngModal)
    Unload Me
End Sub

Private Sub Form_Load()
    Call InitControlPosition
    Call InitVSF
    Call LoadData
End Sub

Public Sub ShowMe(ByVal frmPar As Form, ByVal lngModal As Long)
    mlngModal = lngModal
    Me.Show vbModal, frmPar
End Sub

Private Sub InitControlPosition()
    '�ؼ�λ��
    lblImportMethod.Move 70, 100
    optPartImport.Move lblImportMethod.Left + lblImportMethod.Width + 300, 100
    optFullImport.Move optPartImport.Left + optPartImport.Width + 150, 100
    cmdExit.Move Me.Width - cmdExit.Width - 100, lblImportMethod.Top - 50
    cmdSave.Move cmdExit.Left - cmdSave.Width - 100, lblImportMethod.Top - 50
    vsfError.Move lblImportMethod.Left, lblImportMethod.Top + lblImportMethod.Height + 150, Me.Width - 170, Me.Height - vsfError.Top - 150
End Sub

Private Sub InitVSF()
    '��ʼ��vsf�ؼ�
    With vsfError
        .Editable = flexEDNone
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExNone
        .ExtendLastCol = True '���һ�������
        .ColComboList(0) = "��ֹ|��ʾ"
        .WordWrap = True
        .AutoSize 2, 2, False, 0 = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .ScrollBars = flexScrollBarVertical '�����������ȡ����
        .Cell(flexcpBackColor, 1, 1, .Rows - 1, 2) = MCONFIXECOLOR '�����޸�����ɫ
    End With
End Sub

Private Sub optFullImport_Click()
    Dim intRow As Integer
    
    With vsfError
        If optFullImport.Value = True Then
            For intRow = 1 To .Rows - 1
                .TextMatrix(intRow, 0) = "��ֹ"
            Next
            .Cell(flexcpBackColor, 1, 0, .Rows - 1, 2) = MCONFIXECOLOR '�����޸�����ɫ
            .Editable = flexEDNone
        End If
    End With
End Sub

Private Sub optPartImport_Click()
    If optPartImport.Value = True Then
        vsfError.Cell(flexcpBackColor, 1, 0, vsfError.Rows - 1, 0) = &H80000005    '���޸�����ɫ
    End If
End Sub

Private Sub vsfError_CellChanged(ByVal Row As Long, ByVal Col As Long)
    With vsfError
        If Col = 0 Then
            If .TextMatrix(Row, Col) = "��ֹ" Then
                .Cell(flexcpFontBold, Row, 0, Row, 0) = True
            Else
                .Cell(flexcpFontBold, Row, 0, Row, 0) = False
            End If
        End If
    End With
End Sub

Private Sub vsfError_EnterCell()
    With vsfError
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = MCONFIXECOLOR Then
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub LoadData()
    Dim strPara As String
    Dim intRow As Integer
    Dim intCol As Integer
    Dim arryPara As Variant
    Dim arryTempPara As Variant
    Dim strTemp As String
    Dim strImportMethod As String
    '��������
    If mlngModal = 1712 Then
        strPara = zlDatabase.GetPara("�����ļ���鷽ʽ", glngSys, mlngModal, "0/0|0|0|0|0|0|0|0|0|0|0|0|0|0|0")
    Else
        strPara = zlDatabase.GetPara("�����ļ���鷽ʽ", glngSys, mlngModal, "0/0|0|0|0|0|0|0|0|0|0|0|0|0|0|0")
    End If
    
    arryPara = Split(strPara, "|")
    With vsfError
        For intRow = 0 To UBound(arryPara)
            strTemp = arryPara(intRow)
            If intRow = 0 Then
                strImportMethod = Split(strTemp, "/")(0)
                If strImportMethod = "0" Then
                    optFullImport.Value = False
                    optPartImport.Value = True
                Else
                    optFullImport.Value = True
                    optPartImport.Value = False
                End If
                strTemp = Split(strTemp, "/")(1)
                strTemp = Split(strTemp, ",")(0)
                If strTemp = "0" Then
                    .TextMatrix(intRow + 1, 0) = "��ʾ"
                Else
                    .TextMatrix(intRow + 1, 0) = "��ֹ"
                    .Cell(flexcpFontBold, intRow + 1, 0) = True
                End If
            End If
            If strTemp = "0" Then
                .TextMatrix(intRow + 1, 0) = "��ʾ"
            Else
                .TextMatrix(intRow + 1, 0) = "��ֹ"
                .Cell(flexcpFontBold, intRow + 1, 0) = True
            End If
        Next
    End With
End Sub

