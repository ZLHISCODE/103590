VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPubColShow 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2415
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   2415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   30
      ScaleHeight     =   405
      ScaleWidth      =   2385
      TabIndex        =   1
      Top             =   6840
      Width           =   2385
      Begin VB.Label lblCancle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȡ��"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   1845
         TabIndex        =   7
         Top             =   120
         Width           =   360
      End
      Begin VB.Label lblOk 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȷ��"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   1275
         TabIndex        =   6
         Top             =   120
         Width           =   360
      End
      Begin VB.Label lblClearAll 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȫ��"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   690
         TabIndex        =   5
         Top             =   120
         Width           =   360
      End
      Begin VB.Label lblSelAll 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȫѡ"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   135
         TabIndex        =   4
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   315
         Left            =   60
         TabIndex        =   2
         Top             =   30
         Width           =   450
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfColShow 
      Height          =   4545
      Left            =   210
      TabIndex        =   0
      Top             =   180
      Width           =   1920
      _cx             =   3387
      _cy             =   8017
      Appearance      =   2
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483635
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   2
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   350
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
      Editable        =   2
      ShowComboButton =   0
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
      Begin VB.Label lblMove 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��7��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   2895
         Visible         =   0   'False
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmPubColShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrMouseDown As String  '������ʱ����ѡ���ı�
Private mstrMouseUp As String    '��굯��ʱ����ѡ���ı�
Private mlngMouseDownRow As Long '�����������һ�а���
Private mlngMouseUpRow As Long '�����������һ�е���
Private mstrTitle  As String
Private mX As Long  'X����
Private mY As Long  'Y����
Private mintMouseDownCheck As Integer       '��������Ƿ�ѡ��,1=ѡ��,0=δѡ��
Private mstrSetPara As String               '��������

Public Function ShowMe(objFrm As Object, objVSF As VSFlexGrid, ByVal X As Long, ByVal Y As Long, _
                    ByVal strPara As String, ByVal lngSysNo As Long, ByVal lngModlNo As Long, _
                    Optional ByVal strHiddenCols As String, Optional ByVal strShwoCols As String) As String
    '����           ����vsf�ؼ���ͷ����ʾ˳��,����ʾ������
    '               ���ô˹���ʱ,��������һ��������������Щ����
    
    
    'VSFlexGrid                     ��������VSF
    'X                              ���������X����
    'Y                              ���������Y����
    'strPara                        ������
    'lngSysNo                       ϵͳ��
    'lngModlNo                      ģ���
    '[strHiddenCols]                �̶���Զ������ʾ����,����ID,,��Щ
    '[strShwoCols]                  �̶���Զ����ʾ����
    
    '����                           ��������֮�����ͷ˳��,����Ĳ���Ҳ�������ʽ
    '                               ��ʽ:�е�keyֵ1,���,�Ƿ���ʾ(1=��ʾ,0=����ʾ);�е�keyֵ2,���,�Ƿ���ʾ(1=��ʾ,0=����ʾ),,,,,,,,
    Dim lngCol As Long
    
    mX = X
    mY = Y
    mstrSetPara = strPara & "," & lngSysNo & "," & lngModlNo
    
    '���strHiddenCols�����ڿ�,����strHiddenCols����û��",",����strHiddenCols�������","
    If strHiddenCols <> "" Then
        If Right(strHiddenCols, 1) <> "," Then
            strHiddenCols = strHiddenCols & ","
        End If
        If Left(strHiddenCols, 1) <> "," Then
            strHiddenCols = "," & strHiddenCols
        End If
    End If
    
    '���strShwoCol�����ڿ�,����strShwoCol����û��",",����strShwoCol�������","
    If strShwoCols <> "" Then
        If Right(strShwoCols, 1) <> "," Then
            strShwoCols = strShwoCols & ","
        End If
        If Left(strShwoCols, 1) <> "," Then
            strShwoCols = "," & strShwoCols
        End If
    End If
    
    '�����б�
    With Me.vsfColShow
        .FixedCols = 0
        .FixedRows = 1
        .Cols = 5
        .Rows = 1
        .ColWidth(0) = 250
        .ColDataType(0) = flexDTBoolean
        .GridLines = flexGridNone
        
        .ColKey(0) = "ѡ��": .TextMatrix(0, .ColIndex("ѡ��")) = "ѡ��"
        .ColKey(1) = "����": .TextMatrix(0, .ColIndex("����")) = "����"
        .ColKey(2) = "ColKey": .TextMatrix(0, .ColIndex("ColKey")) = "ColKey": .ColHidden(.ColIndex("ColKey")) = True
        .ColKey(3) = "�п�": .TextMatrix(0, .ColIndex("�п�")) = "�п�": .ColHidden(.ColIndex("�п�")) = True
        .ColKey(4) = "ǿ����ʾ": .TextMatrix(0, .ColIndex("ǿ����ʾ")) = "ǿ����ʾ": .ColHidden(.ColIndex("ǿ����ʾ")) = True
        .ColWidth(0) = 800
        
        With objVSF
            For lngCol = 0 To .Cols - 1
                '��objVsf��������ӵ�vsfColShow��
                vsfColShow.Rows = vsfColShow.Rows + 1
                vsfColShow.TextMatrix(vsfColShow.Rows - 1, vsfColShow.ColIndex("����")) = IIf(.TextMatrix(0, lngCol) = "", .ColKey(lngCol), .TextMatrix(0, lngCol))
                vsfColShow.TextMatrix(vsfColShow.Rows - 1, vsfColShow.ColIndex("ColKey")) = .ColKey(lngCol)
                vsfColShow.TextMatrix(vsfColShow.Rows - 1, vsfColShow.ColIndex("�п�")) = .ColWidth(lngCol)
                
                '������Ҫ���ػ�������Ϊ�յ���
                If InStr(UCase(strHiddenCols), "," & UCase(.ColKey(lngCol)) & ",") > 0 Or (.TextMatrix(0, lngCol) = "" And .Cell(flexcpPicture, 0, lngCol) Is Nothing) Then
                    vsfColShow.RowHidden(vsfColShow.Rows - 1) = True
                End If
                
                'ʼ����ʾ��Ҫ��ʾ����
                If InStr(UCase(strShwoCols), "," & UCase(.ColKey(lngCol)) & ",") > 0 And (.TextMatrix(0, lngCol) <> "" Or Not .Cell(flexcpPicture, 0, lngCol) Is Nothing) Then
                    vsfColShow.Cell(flexcpForeColor, vsfColShow.Rows - 1, vsfColShow.ColIndex("����")) = vbRed
                    vsfColShow.TextMatrix(vsfColShow.Rows - 1, vsfColShow.ColIndex("ǿ����ʾ")) = 1
                End If
                
                'ѡ��objVsf����ʾ����
                If .ColHidden(lngCol) = False And (.TextMatrix(0, lngCol) <> "" Or Not .Cell(flexcpPicture, 0, lngCol) Is Nothing) Then
                    vsfColShow.Cell(flexcpChecked, vsfColShow.Rows - 1, vsfColShow.ColIndex("ѡ��")) = 1
                Else
                    vsfColShow.Cell(flexcpChecked, vsfColShow.Rows - 1, vsfColShow.ColIndex("ѡ��")) = 0
                End If
            Next
        End With
        
        .Select 1, 1
    End With
    
    Me.Show 1, objFrm
    ShowMe = mstrTitle
    
End Function

Private Sub setRowColor()
    '����������ɫ
    With Me.vsfColShow
        
    End With
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        mstrTitle = ""
        Unload Me
    End If
End Sub

Private Sub Form_Resize()
    Me.Left = mX
    Me.Top = mY
    With Me.vsfColShow
        .Top = 0
        .Left = 0
        .Width = Me.Width
        .Height = Me.Height - Me.picButton.Height - 50
    End With
    With Me.picButton
        .Left = 50
        .Width = Me.Width - 100
        .Height = Me.Height - .Top - 50
    End With
End Sub

Private Sub setVSFList(ByVal strTitle As String)
    '���������б�����
    Dim var_tmp As Variant
    Dim var_tmp1 As Variant
    Dim lngLoop As Long
    
    var_tmp = Split(strTitle, ";")
    With Me.vsfColShow
        .FixedCols = 0
        .FixedRows = 1
        .Cols = 2
        .Rows = 1
        .ColWidth(0) = 250
        .ColDataType(0) = flexDTBoolean
        .GridLines = flexGridNone
        
        .TextMatrix(0, 0) = "ѡ��"
        .TextMatrix(0, 1) = "����"
        .ColWidth(0) = 800
        
        For lngLoop = 0 To UBound(var_tmp)
            var_tmp1 = Split(var_tmp(lngLoop), ",")
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = var_tmp1(0)
            If var_tmp1(1) = 0 Then .RowHidden(.Rows - 1) = True    '�Ƿ���ʾ
            If var_tmp1(2) = 1 Then .Cell(flexcpChecked, .Rows - 1, 0) = 1 '�Ƿ�ѡ��
        Next
        .Select 1, 1
    End With
End Sub

Private Sub saveVSFCols()
          '�����б���Ϣ
          Dim strCols As String
          Dim lngRow As Long
          Dim strPara As String       '������
          Dim lngSysNo As Long        'ϵͳ��
          Dim lngModlNo As Long       'ģ���
          
          '��ȡ������,ϵͳ��,ģ���
1         On Error GoTo saveVSFCols_Error

2         If mstrSetPara <> "" Then
3             strPara = Split(mstrSetPara, ",")(0)
4             lngSysNo = Val(Split(mstrSetPara, ",")(1))
5             lngModlNo = Val(Split(mstrSetPara, ",")(2))
6         End If
          
7         With Me.vsfColShow
8             For lngRow = 1 To .Rows - 1
9                 strCols = strCols & ";" & .TextMatrix(lngRow, .ColIndex("ColKey")) & "," & _
                          .TextMatrix(lngRow, .ColIndex("�п�")) & "," & _
                          IIf(.Cell(flexcpChecked, lngRow, vsfColShow.ColIndex("ѡ��")) = 1, _
                          IIf(.RowHidden(lngRow) = False, 1, 0), 0)
10            Next
              
              '�����ʼ�ձ����ڵ�һ��,��ʼ����ʾ
11            If strCols <> "" Then
12                strCols = Mid(strCols, 2)
13                mstrTitle = strCols
14            End If
15        End With
          '�������
16        Call ComSetPara(Sel_Lis_DB, strPara, mstrTitle, lngSysNo, lngModlNo)


17        Exit Sub
saveVSFCols_Error:
18        Call WriteErrLog("zlPublicHisCommLis", "frmPubColShow", "ִ��(saveVSFCols)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
19        Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrMouseUp = ""
    mlngMouseUpRow = 0
    mstrMouseDown = ""
    mlngMouseDownRow = 0
    mstrSetPara = ""
End Sub

Private Sub lblCancle_Click()
    mstrTitle = ""
    Unload Me
End Sub

Private Sub lblClearAll_Click()
    Call SelorClearAll(0) 'ȫ��
End Sub

Private Sub lblOk_Click()
    Call saveVSFCols
    Unload Me
End Sub

Private Sub lblSelAll_Click()
    Call SelorClearAll(1) 'ȫѡ
End Sub

Private Sub SelorClearAll(ByVal intType As Integer)
    'ȫѡ/ȫ��
    Dim lngRow As Long
    
    With Me.vsfColShow
        For lngRow = 1 To .Rows - 1
            If .RowHidden(lngRow) = False And Val(.TextMatrix(lngRow, .ColIndex("ǿ����ʾ"))) <> 1 Then
                .Cell(flexcpChecked, lngRow, 0) = intType
            End If
        Next
    End With
End Sub

Private Sub vsfColShow_Click()
    '�޷�ȡ����ɫ�еĹ�ѡ
    Dim lngRow As Long
    Dim lngCol As Long
    
    With Me.vsfColShow
        lngRow = .MouseRow
        lngCol = .MouseCol
        If lngRow <= 0 Or lngCol <> .ColIndex("ѡ��") Then Exit Sub
        If .Cell(flexcpForeColor, lngRow, .ColIndex("����")) = vbRed Then
            .Cell(flexcpChecked, lngRow, lngCol, lngRow, lngCol) = 1
        End If
        
    End With
End Sub

Private Sub vsfColShow_DblClick()
    Dim lngCol As Long
    With Me.vsfColShow
        lngCol = .MouseCol
        If lngCol < 0 Then Exit Sub
        If lngCol = 0 Then
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If
    End With
End Sub

'�����������̾�Ϊģ������϶��б���
Private Sub vsfColShow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
        
    With Me.vsfColShow
        lngRow = .MouseRow
        If lngRow < 0 Then Exit Sub
        If Button = 1 Then
            mstrMouseDown = .TextMatrix(lngRow, 1) & "," & .TextMatrix(lngRow, 2) & "," & .TextMatrix(lngRow, 3) & "," & Val(.TextMatrix(lngRow, 4))
            mlngMouseDownRow = lngRow
            mintMouseDownCheck = Val(.Cell(flexcpChecked, mlngMouseDownRow, 0))
        End If
    End With
    
End Sub

Private Sub vsfColShow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        
        If mstrMouseDown <> "" Then
            Me.lblMove.Visible = True
            Me.lblMove.Caption = Split(mstrMouseDown, ",")(0)
        End If

        With Me.lblMove
            .Left = X - .Width / 2
            .Top = Y - .Height / 2
        End With
    End If
End Sub

Private Sub vsfColShow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
    
    
    Me.lblMove.Visible = False
    With Me.vsfColShow
        lngRow = .MouseRow
        If lngRow < 0 Then Exit Sub
        If Button = 1 Then
            mstrMouseUp = .TextMatrix(lngRow, 1)
            mlngMouseUpRow = lngRow
        End If
        
        If mstrMouseDown <> "" And mlngMouseUpRow > 0 And mstrMouseUp <> "" And mlngMouseDownRow > 0 Then
            If mlngMouseUpRow < 1 Or mlngMouseUpRow > .Rows Then Exit Sub
            '����
            If mlngMouseDownRow > mlngMouseUpRow Then
                .AddItem "", mlngMouseUpRow
                .TextMatrix(mlngMouseUpRow, 1) = Split(mstrMouseDown, ",")(0)
                .TextMatrix(mlngMouseUpRow, 2) = Split(mstrMouseDown, ",")(1)
                .TextMatrix(mlngMouseUpRow, 3) = Split(mstrMouseDown, ",")(2)
                .TextMatrix(mlngMouseUpRow, 4) = Split(mstrMouseDown, ",")(3)
                If .TextMatrix(mlngMouseUpRow, 4) = 1 Then
                    .Cell(flexcpForeColor, mlngMouseUpRow, 1) = vbRed
                End If
                .RemoveItem mlngMouseDownRow + 1
                .Cell(flexcpChecked, mlngMouseUpRow, 0) = mintMouseDownCheck
            Else
                .AddItem "", mlngMouseUpRow + 1
                .TextMatrix(mlngMouseUpRow + 1, 1) = Split(mstrMouseDown, ",")(0)
                .TextMatrix(mlngMouseUpRow + 1, 2) = Split(mstrMouseDown, ",")(1)
                .TextMatrix(mlngMouseUpRow + 1, 3) = Split(mstrMouseDown, ",")(2)
                .TextMatrix(mlngMouseUpRow + 1, 4) = Split(mstrMouseDown, ",")(3)
                If Val(.TextMatrix(mlngMouseUpRow + 1, 4)) = 1 Then
                    .Cell(flexcpForeColor, mlngMouseUpRow + 1, 1) = vbRed
                End If
                .RemoveItem mlngMouseDownRow
                .Cell(flexcpChecked, mlngMouseUpRow, 0) = mintMouseDownCheck
            End If
        End If
        
        mstrMouseUp = ""
        mlngMouseUpRow = 0
        mstrMouseDown = ""
        mlngMouseDownRow = 0
    End With
End Sub


