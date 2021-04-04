VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmCertifyRecord 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ʵ����Ϣ�����¼"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11625
   Icon            =   "frmCertifyRecord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   11625
   StartUpPosition =   2  '��Ļ����
   Begin VSFlex8Ctl.VSFlexGrid vsfRecord 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      _cx             =   20558
      _cy             =   5530
      Appearance      =   0
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
      BackColorSel    =   16777215
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   9
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   325
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
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
End
Attribute VB_Name = "frmCertifyRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngʵ��id As Long
Dim mrsRescord As New ADODB.Recordset

Private Enum VSFRECORD_INFO
    COL_ʵ��ID = 0
    COL_����
    COL_�����Ŀ
    COL_ԭ��Ϣ
    COL_����Ϣ
    COL_�����
    COL_���ʱ��
    COL_���ԭ��
End Enum
Private Sub InitVsfGridHeader()
'���ܣ���ʼ���б�
    Dim strHeader As String
    
    strHeader = "ʵ��ID;����,1000,1;�����Ŀ,1500,1;ԭ��Ϣ,3500,1;����Ϣ,3500,1;�����,1000,4;���ʱ��,1500,1;���ԭ��,3500,1"
    Call grid.Init(vsfRecord, strHeader)
End Sub

Private Sub Form_Load()
    Call InitVsfGridHeader
    LoadCertifyRecord mrsRescord
End Sub

Private Sub LoadCertifyRecord(ByVal rsTmp As ADODB.Recordset)
'���ܣ����ز���ʵ����Ϣ�����¼
    Dim i As Long, j As Long
    
    With vsfRecord
        j = .FixedRows
        If Not rsTmp.EOF Then
            For i = 0 To rsTmp.RecordCount - 1
                .AddItem "", j
                .TextMatrix(j, COL_ʵ��ID) = "" & rsTmp!ʵ��ID
                .TextMatrix(j, COL_����) = "" & rsTmp!����
                .TextMatrix(j, COL_�����Ŀ) = "" & rsTmp!�����Ŀ
                If Trim("" & rsTmp!�����Ŀ) = "�����˳�������" Or Trim("" & rsTmp!�����Ŀ) = "��������" Then
                    .TextMatrix(j, COL_ԭ��Ϣ) = Format("" & rsTmp!ԭ��Ϣ, "YYYY-MM-DD HH:MM")
                    .TextMatrix(j, COL_����Ϣ) = Format("" & rsTmp!����Ϣ, "YYYY-MM-DD HH:MM")
                ElseIf Trim("" & rsTmp!�����Ŀ) = "���������֤����" Or Trim("" & rsTmp!�����Ŀ) = "���֤����" Then
                    .TextMatrix(j, COL_ԭ��Ϣ) = decode(Val("" & rsTmp!ԭ��Ϣ), 1, "�������֤", 2, "�۰�̨��ס֤", 3, "����˾���֤", 0, "", -1, "", "" & rsTmp!ԭ��Ϣ)
                    .TextMatrix(j, COL_����Ϣ) = decode(Val("" & rsTmp!����Ϣ), 1, "�������֤", 2, "�۰�̨��ס֤", 3, "����˾���֤", 0, "", -1, "", "" & rsTmp!����Ϣ)
                Else
                    .TextMatrix(j, COL_ԭ��Ϣ) = "" & rsTmp!ԭ��Ϣ
                    .TextMatrix(j, COL_����Ϣ) = "" & rsTmp!����Ϣ
                End If
                .TextMatrix(j, COL_���ʱ��) = Format("" & rsTmp!���ʱ��, "YYYY-MM-DD HH:MM")
                .TextMatrix(j, COL_�����) = "" & rsTmp!�����
                .TextMatrix(j, COL_���ԭ��) = "" & rsTmp!���ԭ��
                rsTmp.MoveNext
                j = j + 1
            Next
        End If
    End With
End Sub

Public Function ShowMe(frmParent As Object, ByVal lngʵ��ID As Long) As Boolean
    mlngʵ��id = lngʵ��ID
    If mlngʵ��id = 0 Then
        MsgBox "��ѡ��һ�����ˣ�", vbInformation, gstrSysName
    Else
        Set mrsRescord = GetCertifyRecord(mlngʵ��id)
        If mrsRescord.EOF Then
            MsgBox "�ò���û��ʵ����Ϣ�����¼��", vbInformation, gstrSysName
        Else
            If Not frmParent Is Nothing Then
                Me.Show , frmParent
            End If
        End If
    End If
End Function
