VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmUntread 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�汾����"
   ClientHeight    =   3555
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5700
   Icon            =   "frmUntread.frx":0000
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
      Picture         =   "frmUntread.frx":058A
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
Attribute VB_Name = "frmUntread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mfParent As Object, mstrPrivs As String

Public Function ShowMe(ByVal fParent As Object, ByVal strPrivs As String) As Boolean
'���ܣ���ʾ�����İ汾�޶��仯��������û�����ִ�л���
'���أ��ɹ����
Dim rsTemp As New ADODB.Recordset
    mblnOK = False
    On Error GoTo errHand
1    Set mfParent = fParent: mstrPrivs = strPrivs
2    gstrSQL = "Select Ҫ�ر�ʾ, �����ı�, ��������, ��ֹ��,��������" & vbNewLine & _
            "From (Select Ҫ�ر�ʾ, �����ı�, ��������, ��ֹ��,��������" & vbNewLine & _
            "       From ���Ӳ�������" & vbNewLine & _
            "       Where �ļ�id = [1] And �������� In (6, 7, 8) And nvl(��ֹ��,0)>0 " & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select Distinct 0 Ҫ�ر�ʾ, '�޶�' �����ı�, '|0;;;;' ��������, ��ֹ��-0.1 ��ֹ��,6 ��������" & vbNewLine & _
            "       From ���Ӳ�������" & vbNewLine & _
            "       Where �ļ�id = [1] And �������� Not In (6, 7, 8))" & vbNewLine & _
            "Order By ��ֹ�� Desc"
3    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mfParent.Document.EPRPatiRecInfo.ID)
4    If rsTemp.EOF Then
5        MsgBox "��ǰû�п��Ի��˵�ǩ���汾��", vbInformation, gstrSysName
6        Exit Function
7    Else
8        If rsTemp.RecordCount = 1 Then
9            MsgBox "��ǰû�п��Ի��˵�ǩ���汾��", vbInformation, gstrSysName
10            Exit Function
11        End If
12    End If
    
13    With Me.vfgThis
14        .Clear
15        .Tag = mfParent.Document.EPRPatiRecInfo.ID
16        .Cols = 6: .Rows = rsTemp.RecordCount + 1
17        .ColWidth(0) = 1200: .ColWidth(1) = 1200: .ColWidth(2) = 1800: .ColWidth(3) = 0:    .ColWidth(4) = 0:   .ColWidth(5) = 0
18        .TextMatrix(0, 0) = "ǩ������": .TextMatrix(0, 1) = "ǩ����": .TextMatrix(0, 2) = "ǩ��ʱ��"
19        .TextMatrix(0, 3) = "ǩ���汾":   .TextMatrix(0, 4) = "ǩ����ʽ":   .TextMatrix(0, 5) = "ǩ������"
20        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
21        Do Until rsTemp.EOF
22            .TextMatrix(rsTemp.AbsolutePosition, 0) = Decode(mfParent.Document.EPRPatiRecInfo.��������, 4, Decode(rsTemp!Ҫ�ر�ʾ, 3, "��ʿ��", 1, "��ʿ", "�޶�"), Decode(rsTemp!Ҫ�ر�ʾ, 3, "����ҽʦ", 2, "����ҽʦ", 1, "����ҽʦ", "�޶�"))
23            .TextMatrix(rsTemp.AbsolutePosition, 1) = Nvl(rsTemp!�����ı�)
24            .TextMatrix(rsTemp.AbsolutePosition, 2) = Split(Split(rsTemp!��������, "|")(1), ";")(4)
25            .TextMatrix(rsTemp.AbsolutePosition, 3) = CInt(rsTemp!��ֹ��)
26            .TextMatrix(rsTemp.AbsolutePosition, 4) = Val(Split(Split(rsTemp!��������, "|")(1), ";")(0))
27            .TextMatrix(rsTemp.AbsolutePosition, 5) = CInt(rsTemp!��������)
28            rsTemp.MoveNext
29        Loop
30    End With
    
31    Me.Show vbModal
32    ShowMe = mblnOK
    Exit Function
errHand:
    Call MsgBox("frmUntread:ShowMe�����У�" & Erl(), vbInformation, gstrSysName)
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    ShowMe = False
End Function

Private Sub cmdCancel_Click()
    mblnOK = False: Unload Me
End Sub

Private Sub cmdUntread_Click()
Dim objESign As Object  '����ǩ���ӿڲ���
On Error GoTo errHand
    If vfgThis.TextMatrix(vfgThis.Row, 5) = 6 Then
        If mfParent.Document.EPRPatiRecInfo.������ <> UserInfo.���� And InStr(mstrPrivs, "��������ǩ��") = 0 Then
            MsgBox "��󱣴����뵱ǰ�����߲���ͬһ�ˣ����ܻ��ˣ�", vbInformation, gstrSysName
            Exit Sub
        End If
    ElseIf vfgThis.TextMatrix(vfgThis.Row, 5) = 7 Or vfgThis.TextMatrix(vfgThis.Row, 5) = 8 Then
        If vfgThis.TextMatrix(vfgThis.Row, 1) <> UserInfo.���� And InStr(mstrPrivs, "��������ǩ��") = 0 Then
            MsgBox "��Ҫ���˵�ǩ���뵱ǰ�����߲���ͬһ�ˣ����ܻ��ˣ�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If MsgBox("ע�⣺���˲��������ɻָ����Ƿ������", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Sub
    
    If vfgThis.TextMatrix(vfgThis.Row, 4) = 2 Then
        '����ǩ����֤
        Err.Clear: On Error Resume Next
        If objESign Is Nothing Then
            Set objESign = CreateObject("zl9ESign.clsESign")
            If Err <> 0 Then Err = 0
        End If
        If Not objESign Is Nothing Then
            If objESign.Initialize(gcnOracle, glngSys) Then
                If Not objESign.CheckCertificate(UserInfo.�û���) Then Exit Sub
            Else
                MsgBox "ȡ����ǩ���ļ�ʱ��Ҫ�ٴ���֤ǩ������ϵͳû������ǩ����֤���ģ�����ȡ����", vbOKOnly + vbInformation, gstrSysName
                Exit Sub
            End If
        Else
            MsgBox "ǩ��������ʼ��ʧ�ܣ�", vbOKOnly + vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    On Error GoTo errHand
    gstrSQL = "Zl_���Ӳ�������_Untread(" & vfgThis.Tag & "," & vfgThis.TextMatrix(vfgThis.Row, 3) & "," & IIf(vfgThis.TextMatrix(vfgThis.Row, 0) <> "�޶�", 1, 0) & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "����"
    'ֻ�����У�0��Ϊ�̶��У�1��Ϊǩ����2��Ϊ���ԭʼ��¼
    If vfgThis.Rows = 3 Then mfParent.Document.mReadOnly = 0: mfParent.Document.ET = TabET_�������༭              '���˵�δǩ��״̬�������ٴ�ǩ��
    mblnOK = True: Unload Me
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mfParent = Nothing
End Sub

Private Sub vfgThis_DblClick()
    If cmdUntread.Enabled Then
        Call cmdUntread_Click
    End If
End Sub

Private Sub vfgThis_RowColChange()
    Dim blnEnable As Boolean


    cmdUntread.Enabled = IIf(vfgThis.TextMatrix(vfgThis.Row, 5) = 6, vfgThis.Row = 1, True)
End Sub


