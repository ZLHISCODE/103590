VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm���ŷ�ҩ���������嵥 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���������嵥"
   ClientHeight    =   5760
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   7620
   Icon            =   "frm���ŷ�ҩ���������嵥.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7620
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkȫ����ѡ 
      Caption         =   "ȫ����ѡ"
      Height          =   252
      Left            =   120
      TabIndex        =   5
      Top             =   5329
      Value           =   1  'Checked
      Width           =   1332
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ����ҩ"
      Height          =   350
      Left            =   6360
      TabIndex        =   4
      Top             =   5280
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "������ҩ"
      Height          =   350
      Left            =   5160
      TabIndex        =   3
      Top             =   5280
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   30
      Index           =   0
      Left            =   -240
      TabIndex        =   2
      Top             =   600
      Width           =   8292
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   4452
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7332
      _cx             =   12933
      _cy             =   7853
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      Caption         =   "[��ѡ]-����õ��ݷ�ҩ         [����ѡ]-������õ��ݷ�ҩ"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   720
      TabIndex        =   6
      Top             =   360
      Width           =   5040
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      Caption         =   "�����µ��ݴ���δ����ġ��������롿��¼����Ҫ�ֶ����д���"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   732
      TabIndex        =   1
      Top             =   120
      Width           =   5220
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   120
      Picture         =   "frm���ŷ�ҩ���������嵥.frx":6852
      Top             =   135
      Width           =   480
   End
End
Attribute VB_Name = "frm���ŷ�ҩ���������嵥"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsData As Recordset        '���ڽ��ܸ�����ļ�¼��
Private mbln������ As Boolean

Private mintģʽ As Integer         '0-���ŷ�ҩ;1-������ҩ

Private mstrArray As String   '���ڼ�¼�����ݵ���

Private mIntCol��� As Integer
Private mIntCol����ID As Integer
Private mIntCol�շ�ID As Integer
Private mIntColNO As Integer
Private mIntColҩƷ���� As Integer
Private mintcol���� As Integer
Private mintcol���� As Integer
Private mIntCol������������ As Integer
Private mIntCol���� As Integer
Private mIntCol�Ա� As Integer
Private mIntCol���� As Integer
Private mIntCol��ҩ���� As Integer
Private mIntCol���� As Integer
Private mIntCol���˿��� As Integer
Private Const mconIntCol���� As Integer = 14

Public Sub ShowCard(FrmMain As Form, ByRef rsData As ADODB.Recordset, ByRef bln������ As Boolean, Optional ByVal intģʽ As Integer)
    Set mrsData = rsData
    mbln������ = False
    mintģʽ = intģʽ
    
    Me.Show vbModal, FrmMain
    '��������
    Set rsData = mrsData
    bln������ = mbln������
End Sub

Private Sub InitList()
    '��ʼ����ͷ
    mIntCol��� = 0
    mIntColNO = 1
    mIntColҩƷ���� = 2
    mintcol���� = 3
    mintcol���� = 4
    mIntCol������������ = 5
    mIntCol���� = 6
    mIntCol�Ա� = 7
    mIntCol���� = 8
    mIntCol��ҩ���� = 9
    mIntCol���� = 10
    mIntCol���˿��� = 11
    mIntCol�շ�ID = 12
    mIntCol����ID = 13
    
    With vsfList
        .Cols = mconIntCol����
        .rows = 1
        
        .SelectionMode = flexSelectionByRow
        .AllowSelection = False
        .ColDataType(mIntCol���) = flexDTBoolean
        
        VsfGridColFormat vsfList, mIntCol���, "����", 400, flexAlignCenterCenter, "���"
        VsfGridColFormat vsfList, mIntColNO, "NO", 900, flexAlignLeftCenter, "NO"
        VsfGridColFormat vsfList, mIntColҩƷ����, "ҩƷ����", 1500, flexAlignLeftCenter, "ҩƷ����"
        VsfGridColFormat vsfList, mintcol����, "����", 1300, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfList, mintcol����, "����", 800, flexAlignRightCenter, "����"
        VsfGridColFormat vsfList, mIntCol������������, "������������", 1200, flexAlignRightCenter, "������������"
        VsfGridColFormat vsfList, mIntCol����, "����", 800, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfList, mIntCol�Ա�, "�Ա�", 500, flexAlignLeftCenter, "�Ա�"
        VsfGridColFormat vsfList, mIntCol����, "����", 500, flexAlignRightCenter, "����"
        
        If mintģʽ = 0 Then
            VsfGridColFormat vsfList, mIntCol��ҩ����, "��ҩ����", 1000, flexAlignRightCenter, "��ҩ����"
            VsfGridColFormat vsfList, mIntCol����, "����", 500, flexAlignRightCenter, "����"
            VsfGridColFormat vsfList, mIntCol���˿���, "���˿���", 1000, flexAlignLeftCenter, "���˿���"
        End If
        
        VsfGridColFormat vsfList, mIntCol�շ�ID, "�շ�ID", 0, flexAlignLeftCenter, "�շ�ID"
        VsfGridColFormat vsfList, mIntCol����ID, "����ID", 0, flexAlignLeftCenter, "����ID"
        
        .ColHidden(mIntCol�շ�ID) = True
        .ColHidden(mIntCol����ID) = True
        
        If mintģʽ = 1 Then
            .ColHidden(mIntCol��ҩ����) = True
            .ColHidden(mIntCol����) = True
            .ColHidden(mIntCol���˿���) = True
        End If
        
    End With
End Sub

Private Sub LoadList(ByVal rsData As ADODB.Recordset)
    Dim lngRow As Long
    Dim lng����id As Long
    
    '---------��������---------
    If rsData.RecordCount > 0 Then
        rsData.MoveFirst
        
        With vsfList
            .Redraw = flexRDNone
            
            Do While Not rsData.EOF
                '��ӿ����������У���ֹ��ͬ�ķ���ID�еġ������������������ºϲ�
                If lng����id <> rsData!����ID Then
                    lngRow = lngRow + 2
                    .rows = lngRow + 1
                    '������������
                    .RowHidden(lngRow - 1) = True
                Else
                    lngRow = lngRow + 1
                    .rows = lngRow + 1
                End If
                
                .TextMatrix(lngRow, mIntCol���) = True    'Ĭ�϶���ѡ
                .TextMatrix(lngRow, mIntColNO) = zlCommFun.NVL(rsData!NO, "")
                .TextMatrix(lngRow, mIntColҩƷ����) = zlCommFun.NVL(rsData!ҩƷ����, "")
                .TextMatrix(lngRow, mintcol����) = zlCommFun.NVL(rsData!����, "")
                .TextMatrix(lngRow, mintcol����) = rsData!����
                .TextMatrix(lngRow, mIntCol������������) = rsData!������������
                .TextMatrix(lngRow, mIntCol����) = rsData!����
                .TextMatrix(lngRow, mIntCol�Ա�) = zlCommFun.NVL(rsData!�Ա�, "")
                .TextMatrix(lngRow, mIntCol����) = zlCommFun.NVL(rsData!����, "")
                
                If mintģʽ = 0 Then
                    .TextMatrix(lngRow, mIntCol��ҩ����) = zlCommFun.NVL(rsData!��ҩ����, "")
                    .TextMatrix(lngRow, mIntCol����) = zlCommFun.NVL(rsData!����, "")
                    .TextMatrix(lngRow, mIntCol���˿���) = zlCommFun.NVL(rsData!���˿���, "")
                End If
                
                .TextMatrix(lngRow, mIntCol�շ�ID) = rsData!�շ�ID
                .TextMatrix(lngRow, mIntCol����ID) = rsData!����ID
                   
                lng����id = rsData!����ID
                
                rsData.MoveNext
            Loop
            
            .RowHeight(-1) = 300
            
            .MergeCells = flexMergeRestrictColumns
            .MergeCol(mIntCol������������) = True
            
            .Redraw = flexRDDirect
        End With
    End If
    
End Sub

Private Sub chkȫ����ѡ_Click()
    Dim i As Integer
    
    If chkȫ����ѡ.Value = 1 Then
        vsfList.Cell(flexcpText, 1, mIntCol���, vsfList.rows - 1, mIntCol���) = True
    ElseIf chkȫ����ѡ.Value = 0 Then
        vsfList.Cell(flexcpText, 1, mIntCol���, vsfList.rows - 1, mIntCol���) = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    'ͳ�Ʋ�����ҩ������,������ִ��״̬
    Dim i As Integer
    
    If MsgBox("ֻ���ѹ�ѡ�ĵ��ݲŻᱻ��ҩ�������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    With vsfList
        For i = 1 To .rows - 1
            If .TextMatrix(i, mIntCol�շ�ID) <> "" Then
                If .TextMatrix(i, mIntCol���) = False Then
                    mrsData.Filter = "�շ�ID =" & Val(.TextMatrix(i, mIntCol�շ�ID))
                    If mintģʽ = 0 Then
                        mrsData!ִ��״̬ = 3
                    ElseIf mintģʽ = 1 Then
                        mrsData!��־ = 0
                    End If
                    
                    mrsData.Update
                End If
            End If
        Next
    End With
    
    mbln������ = True
    
    Unload Me
End Sub

Private Sub Form_Load()
    Call InitList
    Call LoadList(mrsData)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mbln������ = False Then
        If MsgBox("�Ƿ�ȡ�����η�ҩ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = 1
    End If
End Sub

Private Sub vsfList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer
    Dim blnHaveCheck As Integer      '������һ�й�ѡ
    Dim blnNoCheck As Integer       '������һ��û�й�ѡ
    Dim strNo As String
    
    With vsfList
        '��������ҩ��ģ����Ϊ���ܲ�ַ�ҩ������ͬ���ݵĹ�ѡ״̬��Ҫ��������
        If mintģʽ = 1 Then
            strNo = .TextMatrix(Row, mIntColNO)
            
            For i = 1 To .rows - 1
                If .TextMatrix(i, mIntColNO) = strNo Then
                    .TextMatrix(i, mIntCol���) = .TextMatrix(Row, mIntCol���)
                End If
            Next
        End If
        
        For i = 1 To .rows - 1
            If .TextMatrix(i, mIntCol�շ�ID) <> "" Then
                If .TextMatrix(i, mIntCol���) = True Then
                    blnHaveCheck = blnHaveCheck + 1
                Else
                    blnNoCheck = blnNoCheck + 1
                End If
            End If
        Next
        
        If blnHaveCheck > 0 And blnNoCheck > 0 Then
            chkȫ����ѡ.Value = 2
        ElseIf blnHaveCheck = 0 And blnNoCheck > 0 Then
            chkȫ����ѡ.Value = 0
        ElseIf blnHaveCheck > 0 And blnNoCheck = 0 Then
            chkȫ����ѡ.Value = 1
        End If
        
    End With
End Sub

Private Sub vsfList_EnterCell()
    vsfList.Editable = flexEDNone
    
    If vsfList.ColSel = mIntCol��� Then
        vsfList.Editable = flexEDKbd
    End If
End Sub
