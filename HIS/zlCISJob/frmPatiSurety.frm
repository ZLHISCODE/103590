VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatiSurety 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������Ϣ"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   Icon            =   "frmPatiSurety.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   8310
   Begin VB.PictureBox picUserInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   0
      ScaleHeight     =   960
      ScaleWidth      =   8295
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.Image Image1 
         Height          =   360
         Left            =   255
         Picture         =   "frmPatiSurety.frx":038A
         Top             =   255
         Width           =   360
      End
      Begin VB.Label lblUserInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label1"
         Height          =   180
         Left            =   810
         TabIndex        =   1
         Top             =   360
         Width           =   540
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfSurety 
      Height          =   2055
      Left            =   2115
      TabIndex        =   2
      Top             =   870
      Width           =   2115
      _cx             =   3731
      _cy             =   3625
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483638
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   330
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
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   4080
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatiSurety.frx":0A74
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12224
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmPatiSurety"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Private mlng����ID As Long
Private mlng��ҳID As Long


Public Sub ShowMe(ByVal frmParent As Form, ByVal lng����ID As Long, ByVal lng��ҳID As Long)
'-------------------------------------------------------------------
'���ܣ���ɵ�����Ϣ�������ʾ���ⲿ�������
'����:
'   frmParent:�ⲿ���ô�����
'   lng����ID������ID
'   lng��ҳID����ҳID
'���أ���
'-------------------------------------------------------------------
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    '������Ϣ
    Call GetPatiInfo
    '������Ϣ
    Call LoadSurety
    
    Me.Show 1, frmParent
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
End Sub


Private Sub SetHeader()
'����:���ñ����ͷ��Ϣ
    Dim strHead As String, i As Long
    
    strHead = ",4,300|������,4,800|������,7,1250|��ʱ����,4,850|����ԭ��,4,1800|�Ǽ�ʱ��,1,1800|����ʱ��,1,1800"
    With vsfSurety
        .Redraw = False
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
        Next
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        
        .ForeColor = &H80000003
        .RowHeight(0) = 320
        .Redraw = True
    End With
End Sub

Private Sub GetPatiInfo()
'���ܣ���ʾ���˻�����Ϣ
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String
    
    On Error GoTo errH
    StrSQL = _
        " Select Nvl(b.����, a.����) ����, Nvl(Nvl(b.�Ա�, a.�Ա�),'δ֪') �Ա�, Nvl(b.����, a.����) ����, b.סԺ��" & vbNewLine & _
        " From ������Ϣ a, ������ҳ��b" & vbNewLine & _
        " Where a.����id = b.����id And b.����id = [1] And b.��ҳid = [2]"

    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If rsTmp.RecordCount > 0 Then
        Me.lblUserInfo = "������" & rsTmp!���� & "     " & "�Ա�" & rsTmp!�Ա� & "   " & "���䣺" & rsTmp!���� & "   " & "סԺ�ţ�" & rsTmp!סԺ��
    Else
        Me.lblUserInfo = "������" & "" & "     " & "�Ա�" & "" & "   " & "���䣺" & "" & "   " & "סԺ�ţ�" & ""
    End If
    staThis.Panels(2).Text = "���˵ڡ�" & mlng��ҳID & "����סԺ������Ϣ"
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadSurety()
'���ܣ���ȡ����ĳ��סԺ����Ч������Ϣ
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String
    
    On Error GoTo errH
    
    StrSQL = _
        " Select '', ������, Decode(������, 999999999, '����', To_Char(������, '999999990.00')) As ������, Decode(��������, 1, '��', ' ') As ��ʱ����," & vbNewLine & _
        "       ����ԭ��, To_Char(�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') �Ǽ�ʱ��, To_Char(����ʱ��, 'yyyy-mm-dd hh24:mi:ss') ����ʱ��" & vbNewLine & _
        " From ���˵�����¼" & vbNewLine & _
        " Where ����id = [1] And ��ҳid = [2] And (����ʱ�� Is Null Or ����ʱ�� > Sysdate) And ɾ����־ = 1" & vbNewLine & _
        " Order By �Ǽ�ʱ�� Desc"


    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    
    vsfSurety.Clear
    vsfSurety.Rows = 2
    If rsTmp.RecordCount > 0 Then
        Set vsfSurety.DataSource = rsTmp
    End If
    
    Call SetHeader
    vsfSurety.Row = 1
    vsfSurety.Col = 0: vsfSurety.ColSel = vsfSurety.Cols - 1
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    With Me.picUserInfo
         .Left = 0: .Top = 10
         .Width = Me.ScaleWidth
    End With
    
    With Me.vsfSurety
        .Left = 10: .Top = picUserInfo.Height + picUserInfo.Top + 30
        .Width = picUserInfo.Width - 20
        .Height = Me.ScaleHeight - .Top - staThis.Height - 30
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

