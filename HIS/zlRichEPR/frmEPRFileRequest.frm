VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmEPRFileRequest 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "����Ӧ��Ҫ��"
   ClientHeight    =   3945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vgdRequest 
      Height          =   2445
      Left            =   315
      TabIndex        =   4
      Top             =   2070
      Visible         =   0   'False
      Width           =   5355
      _cx             =   9446
      _cy             =   4313
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
      BackColorSel    =   16764057
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmEPRFileRequest.frx":0000
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
      WordWrap        =   -1  'True
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
   Begin VB.Label lbl˵������ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "˵�����ļ������ڡ���ʱ����д��"
      Height          =   180
      Left            =   45
      TabIndex        =   5
      Top             =   75
      Width           =   5475
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblҪ������ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ڲ�����ƺ�һ��ÿ120Сʱ��дһ�Σ�������ÿ48Сʱ��дһ�Σ���Σ��ÿ24Сʱ��дһ�Ρ�"
      Height          =   360
      Left            =   255
      TabIndex        =   3
      Top             =   1350
      Width           =   5475
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblҪ����� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2)ʱ��Ҫ��:"
      Height          =   180
      Left            =   90
      TabIndex        =   2
      Top             =   1095
      Width           =   990
   End
   Begin VB.Label lblͨ������ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ò����ʺ����п��ҡ�"
      Height          =   180
      Left            =   255
      TabIndex        =   1
      Top             =   720
      Width           =   5475
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblͨ�ñ��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1)���ÿ���:"
      Height          =   180
      Left            =   90
      TabIndex        =   0
      Top             =   465
      Width           =   990
   End
End
Attribute VB_Name = "frmEPRFileRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum zlEnumDClick
    cprEmDClickApplyTo = 1         '˫�����ÿ���
    cprEmDClickRequest = 2         '˫��ʱ��Ҫ��
End Enum

'-----------------------------------------------------
'���幫���¼�
'-----------------------------------------------------
Public Event DblClick(lngWhere As zlEnumDClick)    '����˫���¼�

'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mintKind As Integer       '��������
Private mlngFileID As Long        '�����ļ�ID

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblͨ������.Font.Underline = False
    Me.lblͨ������.ForeColor = Me.lblͨ�ñ���.ForeColor
    Me.lblҪ������.Font.Underline = False
    Me.lblҪ������.ForeColor = Me.lblҪ�����.ForeColor
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    
    With Me.lbl˵������
        .Left = 90: .Width = Me.ScaleWidth - .Left * 2: .Top = 195
    End With
    Me.lblͨ�ñ���.Left = 90: Me.lblͨ�ñ���.Top = Me.lbl˵������.Top + Me.lbl˵������.Height + 195
    With Me.lblͨ������
        .Left = 255: .Width = Me.ScaleWidth - Me.lblͨ������.Left - Me.lblͨ�ñ���.Left
        .Top = Me.lblͨ�ñ���.Top + Me.lblͨ�ñ���.Height + 75
    End With
    With Me.lblҪ�����
        .Left = Me.lblͨ�ñ���.Left
        .Top = Me.lblͨ������.Top + Me.lblͨ������.Height + 195
    End With
    With Me.lblҪ������
        .Left = Me.lblͨ������.Left: .Width = Me.ScaleWidth - Me.lblҪ������.Left - Me.lblͨ�ñ���.Left
        .Top = Me.lblҪ�����.Top + Me.lblҪ�����.Height + 75
    End With
    
    With Me.vgdRequest
        .Left = Me.lblҪ������.Left: .Width = Me.lblҪ������.Width
        .Top = Me.lblҪ������.Top: .Height = Me.ScaleHeight - .Top - 195
        
        If mintKind = 5 Then
            .ColWidth(2) = .Width - .ColWidth(0) - .ColWidth(1) - .ColWidth(3) - 300
        Else
            .ColWidth(1) = .Width - .ColWidth(0) - .ColWidth(2) - 30
        End If
    End With
End Sub

Private Sub lblͨ������_DblClick()
    RaiseEvent DblClick(cprEmDClickApplyTo)
End Sub

Private Sub lblͨ������_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblͨ������.Font.Underline = True
    Me.lblͨ������.ForeColor = RGB(0, 0, 128)
End Sub

Private Sub lblҪ������_DblClick()
    RaiseEvent DblClick(cprEmDClickRequest)
End Sub

Private Sub lblҪ������_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblҪ������.Font.Underline = True
    Me.lblҪ������.ForeColor = RGB(0, 0, 128)
End Sub

Private Sub vgdRequest_DblClick()
    RaiseEvent DblClick(cprEmDClickRequest)
End Sub


'-----------------------------------------------------
'���幫������
'-----------------------------------------------------
Public Sub zlRefresh(ByVal lngFileID As Long)
    '���ܣ�ˢ����ʾ
Dim rsTemp As New ADODB.Recordset
Dim strTemp As String, lngCount As Long
    
    mlngFileID = lngFileID
    '--------------------------------------------
    Me.lblͨ������ = "": Me.lblҪ������.Caption = "": Me.lbl˵������ = "": Me.vgdRequest.Visible = False
    If mlngFileID = 0 Then Call Form_Resize: Exit Sub
    
    '--------------------------------------------
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select ����, ���, ����, ͨ��, ˵�� From �����ļ��б� Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    With rsTemp
        If .RecordCount = 0 Then MsgBox "�ļ���ʧ(���ܱ������û�ɾ��)��", vbInformation, gstrSysName: Exit Sub
        mintKind = !����
        Me.lblͨ������.Tag = IIf(IsNull(!ͨ��), 0, !ͨ��)
        Me.lbl˵������.Caption = "˵��:" & !˵��
    End With
    Select Case Val(Me.lblͨ������.Tag)
    Case 0: Me.lblͨ������.Caption = "�ò����ļ���ʱ��������ʹ�á�"
    Case 1: Me.lblͨ������.Caption = "�ò����ļ��ʺ����п���ʹ�á�"
    Case Else
        gstrSQL = "Select d.����, d.���� From ���ű� d, ����Ӧ�ÿ��� s Where d.Id = s.����id And �ļ�id =[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        Me.lblͨ������.Caption = ""
        With rsTemp
            Do While Not .EOF()
                Me.lblͨ������.Caption = Me.lblͨ������.Caption & "��[" & !���� & "]" & !����
                .MoveNext
            Loop
        End With
        If Me.lblͨ������.Caption = "" Then
            Me.lblͨ������.Caption = "��δ���ñ������ļ������ÿ��ң�"
        Else
            Me.lblͨ������.Caption = "�������ļ������ÿ��Ұ�����" & Mid(Me.lblͨ������, 2) & "��"
        End If
    End Select
    
    '--------------------------------------------
    Select Case mintKind
    Case 1      '���ﲡ��
        Me.lblҪ�����.Caption = "2)ʱ��Ҫ��:"
        gstrSQL = "Select �¼�, ���� From ����ʱ��Ҫ�� Where �ļ�id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        With rsTemp
            If .RecordCount = 0 Then
                Me.lblҪ������.Caption = "��δ���ñ������ļ���ʱ��Ҫ��"
            Else
                Me.lblҪ������.Caption = "�ڲ���" & !�¼� & "ʱ��" & IIf(!���� = 0, "����", "����") & "��д�������ļ���"
            End If
        End With
    Case 2, 4       'סԺ������������
        Me.lblҪ�����.Caption = "2)ʱ��Ҫ��:"
        gstrSQL = "Select �¼�, Nvl(����,0) As ����, Nvl(Ψһ,0) As Ψһ," & _
                "       Nvl(��дʱ��,0) As ��дʱ��, Nvl(����ʱ��,0) As ����ʱ��, Nvl(���ʱ��,0) As ���ʱ��," & _
                "       Nvl(һ������,0) As һ������, Nvl(��������,0) As ��������, Nvl(��Σ����,0) As ��Σ����" & _
                " From ����ʱ��Ҫ�� Where �ļ�id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        With rsTemp
            If .RecordCount = 0 Then
                Me.lblҪ������.Caption = "��δ���ñ������ļ���ʱ��Ҫ��"
            Else
                Me.lblҪ������.Caption = "�ڷ�������" & !�¼� & IIf(!��дʱ�� < 0, "ǰ��", "��") & _
                    IIf(!���� = 0, "����", "����") & IIf(!Ψһ = 0, "ѭ����¼", "��дһ��") & "�������ļ���"
                strTemp = ""
                If !Ψһ <> 0 Then
                    If !��дʱ�� > 0 Then strTemp = strTemp & "��" & !��дʱ�� & "Сʱ��ɲ���"
                    If !����ʱ�� > 0 Then strTemp = strTemp & "��" & !����ʱ�� & "Сʱ�������"
                    If !���ʱ�� > 0 Then strTemp = strTemp & "��" & !���ʱ�� & "Сʱ����������"
                Else
                    If !һ������ > 0 Then strTemp = strTemp & "��һ�㲡��ÿ" & !һ������ & "Сʱ��¼һ��"
                    If !�������� > 0 Then strTemp = strTemp & "�����ز���ÿ" & !�������� & "Сʱ��¼һ��"
                    If !��Σ���� > 0 Then strTemp = strTemp & "����Σ����ÿ" & !��Σ���� & "Сʱ��¼һ��"
                End If
                If strTemp <> "" Then Me.lblҪ������.Caption = Me.lblҪ������.Caption & vbCrLf & "Ҫ��" & Mid(strTemp, 2) & "��"
            End If
        End With
        
        gstrSQL = "Select l.���, l.���� From �����ļ��б� l, ���������ϵ e Where l.Id = e.���id And e.�ļ�id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        strTemp = ""
        With rsTemp
            Do While Not .EOF()
                strTemp = strTemp & "��[" & !��� & "]" & !����
                .MoveNext
            Loop
        End With
        If strTemp <> "" Then Me.lblҪ������.Caption = Me.lblҪ������.Caption & vbCrLf & "��ɱ��������ɲ�����дͬ�ڵ�" & Mid(strTemp, 2) & "��"
    
    Case 3      '�����¼
        Me.lblҪ�����.Caption = "2)ʹ��Ҫ��:"
        gstrSQL = "Select Decode(nvl(f.����, 3), 0, '�ؼ�����', 1, 'һ������', 2, '��������', 3, '��������') As �ȼ�" & _
                " From �����ļ��б� l, ����ҳ���ʽ f" & _
                " Where l.���� = f.���� And l.ҳ�� = f.��� And f.���� = 3 And l.Id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        If rsTemp.EOF Then
            Me.lblҪ������.Caption = ""
        Else
            Me.lblҪ������.Caption = "�����ڡ�" & rsTemp.Fields(0).Value & "�������ϵȼ��Ĳ��ˡ�"
        End If
    
    Case 5      '����֤���뱨��
        Me.lblҪ�����.Caption = "2)�����������ʱ��д���ļ�:"
        Me.vgdRequest.Visible = True
        gstrSQL = "Select '����' As ����,����, ����, p.���没�� From ��������Ŀ¼ i, ��������ǰ�� p Where i.Id = p.����id And p.�ļ�id = [1]"
        gstrSQL = gstrSQL & " Union All Select  '���' As ����,����, ����, p.���没�� From �������Ŀ¼ i, ��������ǰ�� p Where i.Id = p.���id And p.�ļ�id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        Set Me.vgdRequest.DataSource = rsTemp
        With vgdRequest
            For lngCount = 0 To .Cols - 1
                .FixedAlignment(lngCount) = flexAlignCenterCenter
                .ColAlignment(lngCount) = flexAlignLeftCenter
            Next
            
            For lngCount = 1 To .Rows - 1
                If .TextMatrix(lngCount, 0) = "���" Then
                    .Cell(flexcpForeColor, lngCount, 0, .Rows - 1, .Cols - 1) = &HFF0000
                    Exit For
                End If
                
            Next
            
            .MergeCells = flexMergeFree
            .MergeCol(0) = True
            .ColWidth(0) = 510
            .ColWidth(1) = 1000: .ColWidth(3) = 1000
        End With
        
    Case 6      '֪���ļ�
        Me.lblҪ�����.Caption = "2)�����������ƴ�ʩǰ��д���ļ�:"
        Me.vgdRequest.Visible = True
        
        gstrSQL = "Select Distinct i.����, i.����, k.���� As ���" & _
                " From ������Ŀ��� k, ������ĿĿ¼ i, ��������Ӧ�� a" & _
                " Where k.���� = i.��� And i.Id = a.������Ŀid And a.�����ļ�id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        Set Me.vgdRequest.DataSource = rsTemp
        With Me.vgdRequest
            For lngCount = 0 To .Cols - 1
                .FixedAlignment(lngCount) = flexAlignCenterCenter
                .ColAlignment(lngCount) = flexAlignLeftCenter
            Next
            .ColWidth(0) = 1000: .ColWidth(2) = 700
        End With
    
    Case 7      '���Ƶ���
    
    End Select
    
    Call Form_Resize
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

