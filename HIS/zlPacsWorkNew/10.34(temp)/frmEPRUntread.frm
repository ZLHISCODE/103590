VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmEPRUntread 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�汾����"
   ClientHeight    =   3555
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5700
   Icon            =   "frmEPRUntread.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdUntread 
      Caption         =   "����(&U)"
      Height          =   375
      Left            =   2790
      TabIndex        =   3
      Top             =   2955
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   375
      Left            =   4095
      TabIndex        =   2
      Top             =   2955
      Width           =   1230
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgThis 
      Height          =   2085
      Left            =   285
      TabIndex        =   1
      Top             =   720
      Width           =   5055
      _cx             =   8916
      _cy             =   3678
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   255
      Picture         =   "frmEPRUntread.frx":058A
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "�ò��������޶�������£������𲽻����Գ����Բ������޶���ǩ����"
      Height          =   360
      Left            =   840
      TabIndex        =   0
      Top             =   195
      Width           =   4500
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmEPRUntread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mblnOK As Boolean
Private mSignLevel As EPRSignLevelEnum

'��ʱ����

Dim lngCount As Long
Dim edtType As EditTypeEnum

Public Function ShowMe(ByVal lngID As Long, _
    ByVal EditType As EditTypeEnum, _
    ByRef lngVersion As Long, _
    ByRef lngSignKey As Long, _
    ByRef fParent As Object) As Boolean
    
    '���ܣ���ʾ�����İ汾�޶��仯��������û�����ִ�л���
    '������lngId-���Ӳ�����¼id
    '      lngVersion-��Ҫ���˵İ汾��
    '      lngSignKey-��Ҫ���˵�ǩ��Keyֵ
    '���أ��ɹ����
    
    '----------------------
    '��ȡ�汾�޶��仯
    
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    
    edtType = EditType
    err = 0: On Error GoTo errHand
    strSql = "Select 0 As Id, -null As ������, 1 As �汾, '�����༭' As ����, l.������ As ��Ա," & _
        "        To_Char(l.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ʱ��,-1 as ���� " & _
        " From ���Ӳ�����¼ l" & _
        " Where L.ID = [1]" & _
        " Union All" & _
        " Select c.Id, c.������, c.��ʼ�� As �汾," & _
        "        Decode(l.��������, 4, Decode(c.Ҫ�ر�ʾ, 3, '��ʿ��', '��ʿ')," & _
        "                Decode(c.Ҫ�ر�ʾ, 3, '����ҽʦ', 2, '����ҽʦ', '����ҽʦ')) || Decode(c.��ʼ��, 1, 'ǩ��', '�޶�') As ����," & _
        "        c.�����ı� As ��Ա, Rtrim(Substr(c.��������, Instr(c.��������, ';', 1, 4) + 1)) As ʱ��,������ as ���� " & _
        " From ���Ӳ�����¼ l, ���Ӳ������� c" & _
        " Where L.ID = c.�ļ�ID And L.ID = [1] And c.�������� = 8" & _
        " Union All" & _
        " Select c.Id,  -null as ������, l.���汾 As �汾, '�����޶���' As ����, l.������ As ��Ա," & _
        "        To_Char(l.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ʱ��,c.������ as ���� " & _
        " From ���Ӳ�����¼ l," & _
        "      (Select Max(c.��ʼ��) As ��ʼ��, Max(Id + 1) As Id,Max(������+1) as ������ From ���Ӳ������� c Where c.�ļ�id = [1] And c.�������� = 8) c" & _
        " Where L.ID = [1] And L.���汾 > c.��ʼ��" & _
        " Order By ���� Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngID)
    With Me.vfgThis
        .Clear
        Set .DataSource = rsTemp
        .ColWidth(0) = 0: .ColHidden(0) = True
        .ColWidth(1) = 0: .ColHidden(1) = True
        .ColWidth(6) = 0: .ColHidden(6) = True
        For lngCount = .FixedCols To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
        For lngCount = .FixedRows To .Rows - 1
            If InStr(1, .TextMatrix(lngCount, 5), ";") > 0 Then .TextMatrix(lngCount, 5) = Left(.TextMatrix(lngCount, 5), 19)
        Next
        If EditType = cprET_�������༭ Then
            If .Rows <= .FixedRows + 1 Then Me.cmdUntread.Enabled = False
        Else
            If .Rows <= .FixedRows + 2 Then Me.cmdUntread.Enabled = False
        End If
    End With
    
    mSignLevel = GetUserSignLevel(UserInfo.ID)
    If mSignLevel <= cprSL_�հ� Then Me.cmdUntread.Enabled = False
    
    Me.Show vbModal, fParent
    If mblnOK = False Then ShowMe = False: Unload Me: Exit Function
    
    '----------------------
    '����
    lngVersion = Val(vfgThis.TextMatrix(1, 2))
    lngSignKey = Val(vfgThis.TextMatrix(1, 1))
    
    ShowMe = True: Unload Me: Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ShowMe = False
End Function

Private Sub cmdCancel_Click()
    mblnOK = False: Me.Hide
End Sub

Private Sub cmdUntread_Click()
    mblnOK = True: Me.Hide
End Sub

Private Sub Form_Load()
On Error Resume Next
    '������ʾ����ǰ��
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '�������ö�
err.Clear
End Sub

Private Sub vfgThis_RowColChange()
    Dim blnEnable As Boolean
    If mSignLevel <= cprSL_�հ� Then Me.cmdUntread.Enabled = False: Exit Sub
    blnEnable = True
    If edtType = cprET_�������༭ Then
        If vfgThis.Rows <= vfgThis.FixedRows + 1 Then blnEnable = False
    Else
        If vfgThis.Rows <= vfgThis.FixedRows + 2 Then blnEnable = False
    End If
    cmdUntread.Enabled = blnEnable And (vfgThis.Row = 1)
End Sub
