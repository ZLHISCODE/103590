VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmRegistList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8955
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRegistList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "�˳�(&E)"
      Height          =   390
      Left            =   7065
      TabIndex        =   7
      ToolTipText     =   "�ȼ�:F2"
      Top             =   5775
      Width           =   1350
   End
   Begin VB.Frame fraInfo 
      Caption         =   "��Դ��Ϣ"
      Height          =   1455
      Left            =   90
      TabIndex        =   0
      Top             =   105
      Width           =   8760
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         Caption         =   "�Ƿ񽨵�:"
         Height          =   240
         Left            =   5040
         TabIndex        =   10
         Top             =   675
         Width           =   1080
      End
      Begin VB.Label lblControlDays 
         AutoSize        =   -1  'True
         Caption         =   "ԤԼ����:"
         Height          =   240
         Left            =   5040
         TabIndex        =   9
         Top             =   1035
         Width           =   1080
      End
      Begin VB.Label lblControl 
         AutoSize        =   -1  'True
         Caption         =   "�Ű�ģʽ:"
         Height          =   240
         Left            =   255
         TabIndex        =   8
         Top             =   1035
         Width           =   1080
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "����:"
         Height          =   240
         Left            =   2370
         TabIndex        =   5
         Top             =   315
         Width           =   600
      End
      Begin VB.Label lblDoc 
         Caption         =   "ҽ��:"
         Height          =   240
         Left            =   255
         TabIndex        =   4
         Top             =   675
         Width           =   2070
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "��Ŀ:"
         Height          =   240
         Left            =   5040
         TabIndex        =   3
         Top             =   315
         Width           =   600
      End
      Begin VB.Label lblDept 
         Caption         =   "����:"
         Height          =   240
         Left            =   2370
         TabIndex        =   2
         Top             =   675
         Width           =   2565
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         Caption         =   "����:"
         Height          =   240
         Left            =   255
         TabIndex        =   1
         Top             =   315
         Width           =   600
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   3960
      Left            =   90
      TabIndex        =   6
      Top             =   1680
      Width           =   8760
      _cx             =   15452
      _cy             =   6985
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   16185078
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRegistList.frx":058A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
Attribute VB_Name = "frmRegistList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng��ԴId As Long
Private mblnUnload As Boolean

Public Sub ShowMe(frmMain As Object, ByVal lng��ԴId As Long)
    mlng��ԴId = lng��ԴId
    Me.Show vbModal, frmMain
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnUnload Then
        mblnUnload = False
        Unload Me
    End If
    If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
End Sub

Private Sub Form_Load()
    Dim strSQL As String, strCurDate As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    strSQL = "Select ����, ����, b.���� As ����, c.���� As ��Ŀ, a.ҽ������, a.�Ƿ񽨲���, a.�Ű෽ʽ, a.ԤԼ����" & vbNewLine & _
            "From �ٴ������Դ A, ���ű� B, �շ���ĿĿ¼ C" & vbNewLine & _
            "Where a.����id = b.Id And a.��Ŀid = c.Id And a.Id = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��ԴId)
    If rsTemp.EOF Then
        MsgBox "�޷�ȷ�Ϻ�Դ��Ϣ,��ȡ�����ʧ��!", vbInformation, gstrSysName
        mblnUnload = True
        Exit Sub
    Else
        lblNO.Caption = "����:" & rsTemp!����
        lblType.Caption = "����:" & rsTemp!����
        lblDept.Caption = "����:" & rsTemp!����
        lblItem.Caption = "��Ŀ:" & rsTemp!��Ŀ
        lblDoc.Caption = "ҽ��:" & rsTemp!ҽ������
        lblPati.Caption = "�Ƿ񽨵�:" & IIf(Val(Nvl(rsTemp!�Ƿ񽨲���)) = 0, "��", "��")
        Select Case Val(Nvl(rsTemp!�Ű෽ʽ))
        Case 0
            lblControl.Caption = "�Ű�ģʽ:�̶��Ű�"
        Case 1
            lblControl.Caption = "�Ű�ģʽ:�����Ű�"
        Case 2
            lblControl.Caption = "�Ű�ģʽ:�����Ű�"
        End Select
        lblControlDays.Caption = "ԤԼ����:" & Val(Nvl(rsTemp!ԤԼ����, gintԤԼ����)) & "��"
    End If
    
    strSQL = "Select ��������,�ϰ�ʱ��,��ʼʱ��,��ֹʱ��,�ѹ���,�޺���,��Լ��,��Լ�� From �ٴ������¼ Where ��ԴID=[1] And �������� >= Trunc(Sysdate) Order By ��������,��ʼʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��ԴId)
    With vsfList
        .Clear 1
        .Rows = 1
        Do While Not rsTemp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Format(rsTemp!��������, "yyyy-mm-dd")
            .TextMatrix(.Rows - 1, 1) = rsTemp!�ϰ�ʱ�� & "(" & Format(rsTemp!��ʼʱ��, "hh:mm") & "-" & Format(rsTemp!��ֹʱ��, "hh:mm") & ")"
            .TextMatrix(.Rows - 1, 2) = Nvl(rsTemp!�ѹ���, 0) & "/" & Nvl(rsTemp!�޺���, "��")
            .TextMatrix(.Rows - 1, 3) = Nvl(rsTemp!��Լ��, 0) & "/" & Nvl(rsTemp!��Լ��, "��")
            rsTemp.MoveNext
        Loop
        .MergeCol(0) = True
        .AutoSize 0, .Cols - 1
    End With
End Sub
