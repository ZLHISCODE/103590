VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmElementEdit 
   BackColor       =   &H00F6F6F6&
   BorderStyle     =   0  'None
   Caption         =   "���ݱ༭��"
   ClientHeight    =   3690
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5490
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtTip 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFEBD7&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   870
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   17
      Text            =   "frmElementEdit.frx":0000
      Top             =   210
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.PictureBox pic�滻��Ŀ 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   90
      ScaleHeight     =   870
      ScaleWidth      =   3525
      TabIndex        =   11
      Top             =   810
      Width           =   3525
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         Height          =   195
         Left            =   315
         TabIndex        =   16
         Top             =   360
         Width           =   510
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   45
         Picture         =   "frmElementEdit.frx":0007
         Top             =   45
         Width           =   240
      End
      Begin VB.Label lblʾ�� 
         BackStyle       =   0  'Transparent
         Caption         =   "ʾ��"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   810
         TabIndex        =   15
         Top             =   585
         Width           =   2940
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ʾ��:"
         Height          =   195
         Left            =   315
         TabIndex        =   14
         Top             =   585
         Width           =   510
      End
      Begin VB.Label lbl���� 
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   810
         TabIndex        =   13
         Top             =   360
         Width           =   2670
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "�Զ��滻��Ŀ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   315
         TabIndex        =   12
         Top             =   90
         Width           =   1635
      End
   End
   Begin VB.TextBox txt����2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   945
      TabIndex        =   4
      Text            =   "99999"
      Top             =   2295
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picTitle 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F5BE9E&
      BorderStyle     =   0  'None
      Height          =   100
      Left            =   45
      MousePointer    =   5  'Size
      ScaleHeight     =   105
      ScaleWidth      =   5325
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   90
      Width           =   5325
      Begin VB.Image imgTitle 
         Height          =   45
         Left            =   1350
         MousePointer    =   5  'Size
         Picture         =   "frmElementEdit.frx":027C
         Top             =   30
         Width           =   2250
      End
   End
   Begin VB.PictureBox picStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   5415
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3375
      Width           =   5415
      Begin VB.Image imgResize 
         Height          =   270
         Left            =   5175
         MousePointer    =   8  'Size NW SE
         Picture         =   "frmElementEdit.frx":02FE
         Top             =   0
         Width           =   225
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "TAB:�Ƶ���һλ�ã��س�:�����޸ġ�"
         Height          =   240
         Left            =   90
         TabIndex        =   7
         Top             =   45
         Width           =   3570
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfg��ѡ 
      Height          =   930
      Left            =   2790
      TabIndex        =   0
      Top             =   270
      Visible         =   0   'False
      Width           =   2550
      _cx             =   4498
      _cy             =   1640
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   240
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmElementEdit.frx":06A0
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
      Ellipsis        =   1
      ExplorerBar     =   7
      PicturesOver    =   -1  'True
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   2
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
      WallPaperAlignment=   1
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.TextBox txt����1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Text            =   "99999"
      Top             =   2295
      Visible         =   0   'False
      Width           =   630
   End
   Begin MSComCtl2.UpDown ud���� 
      Height          =   300
      Left            =   1395
      TabIndex        =   5
      Top             =   2250
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      OrigLeft        =   1065
      OrigTop         =   2295
      OrigRight       =   1320
      OrigBottom      =   2595
      Enabled         =   -1  'True
   End
   Begin VSFlex8Ctl.VSFlexGrid vfg��ѡ 
      Height          =   930
      Left            =   2775
      TabIndex        =   2
      Top             =   1260
      Visible         =   0   'False
      Width           =   2550
      _cx             =   4498
      _cy             =   1640
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16777215
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   240
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmElementEdit.frx":06DD
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
      Ellipsis        =   1
      ExplorerBar     =   7
      PicturesOver    =   -1  'True
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   2
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
      WallPaperAlignment=   1
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.TextBox txt�ı� 
      BorderStyle     =   0  'None
      Height          =   840
      Left            =   2790
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2265
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Image imgOpt1 
      Height          =   195
      Left            =   2160
      Picture         =   "frmElementEdit.frx":071A
      Top             =   2655
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgOpt2 
      Height          =   195
      Left            =   2160
      Picture         =   "frmElementEdit.frx":09A0
      Top             =   2925
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblDot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      Height          =   300
      Left            =   765
      TabIndex        =   10
      Top             =   2295
      Width           =   105
   End
   Begin VB.Shape shpBorder2 
      BorderColor     =   &H00E09060&
      Height          =   375
      Left            =   450
      Top             =   270
      Width           =   330
   End
   Begin VB.Label lbl��λ 
      BackStyle       =   0  'Transparent
      Caption         =   "��λ"
      Height          =   210
      Left            =   1665
      TabIndex        =   6
      Top             =   2340
      Width           =   555
   End
   Begin VB.Shape shpBorder1 
      BorderColor     =   &H00E09060&
      Height          =   375
      Left            =   45
      Top             =   270
      Width           =   330
   End
   Begin VB.Menu mnu_Edit 
      Caption         =   "�༭"
      Visible         =   0   'False
      Begin VB.Menu mnu_Cut 
         Caption         =   "����"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_Copy 
         Caption         =   "����"
      End
      Begin VB.Menu mnu_Delete 
         Caption         =   "ɾ��"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmElementEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private frmParent As Object       '������
Public Element As cEPRElement   '����Ҫ��

'## �Զ����¼�
Public Event pOK()              '��������
Public Event pCancel()          'ȡ���޸�

'## �ֲ���ʱ����
Private lngX As Long, lngY As Long
Private EditType As EditTypeEnum
Private mblnModel As Long

'################################################################################################################
'## ���ܣ�  ��ʾ����Ҫ�ر༭��
'##
'## ������  Ele         :���༭������Ҫ��
'##         (X,Y)       :��ʾλ�ã���Ļ���꣩
'##         blnModel    :������ʾģ̬��0-vbModeless  1-vbModal��
'##         ofrmParent  :������
'##         eEditType   :�༭ģʽ
'################################################################################################################
Public Sub ShowMe(ByRef Ele As cEPRElement, ByVal X As Long, ByVal Y As Long, _
    Optional ByVal blnModel As Long = vbModeless, _
    Optional ByRef ofrmParent As Object, _
    Optional ByVal eEditType As EditTypeEnum)
    
    Dim i As Long, j As Long, T As Variant, strTmp As String
    
    mblnModel = blnModel
    
    Set Me.Element = Ele.Clone(True)
    Set frmParent = ofrmParent
    EditType = eEditType
    
    With Me.Element
        Select Case .Ҫ�ر�ʾ       '0-�ı�,1-����,2-��ѡ,3-��ѡ  5-�ֵ���Ŀ
        Case 0
            If Me.Element.�滻�� = 1 And EditType = cprET_�����ļ����� Then
                lbl���� = Me.Element.Ҫ������
                If ofrmParent.Name = "frmMain" Then
                    lblʾ�� = GetReplaceEleValue(lbl����, ofrmParent.Document.EPRPatiRecInfo.����ID, ofrmParent.Document.EPRPatiRecInfo.��ҳID, ofrmParent.Document.EPRPatiRecInfo.������Դ, ofrmParent.Document.EPRPatiRecInfo.ҽ��id)
                    If lblʾ�� = "" Then
                        lblʾ��.Visible = False
                        Label3.Visible = False
                        Me.Height = 1250
                    Else
                        lblʾ��.Visible = True
                        Label3.Visible = True
                        Me.Height = 1500
                    End If
                ElseIf ofrmParent.Name = "frmTableEditor" Then
                    lblʾ�� = GetReplaceEleValue(lbl����, ofrmParent.frmParent.Document.EPRPatiRecInfo.����ID, ofrmParent.frmParent.Document.EPRPatiRecInfo.��ҳID, ofrmParent.frmParent.Document.EPRPatiRecInfo.������Դ, ofrmParent.Document.EPRPatiRecInfo.ҽ��i)
                    If lblʾ�� = "" Then
                        lblʾ��.Visible = False
                        Label3.Visible = False
                        Me.Height = 1250
                    Else
                        lblʾ��.Visible = True
                        Label3.Visible = True
                        Me.Height = 1500
                    End If
                Else
                    lblʾ�� = "--"
                End If
            Else
                txt�ı�.MaxLength = .Ҫ�س���
                txt�ı� = .�����ı�
                txt�ı�.Visible = True
                txt�ı�.SelStart = 0: txt�ı�.SelLength = Len(.�����ı�)
            End If
        Case 1
            T = Split(.Ҫ��ֵ��, ";")    '��ʽ:  0;100000
            If UBound(T) < 1 Then
                ud����.Min = 0
                ud����.Max = 0
            Else
                ud����.Min = IIf(CLng(T(0)) > CLng(T(1)), CLng(T(1)), CLng(T(0)))
                ud����.Max = IIf(CLng(T(0)) > CLng(T(1)), CLng(T(0)), CLng(T(1)))
            End If
            txt����1.Tag = "��ֵ..."
            i = InStr(1, .�����ı�, ".")
            If i > 0 Then
                txt����1 = Mid(.�����ı�, 1, i - 1)
                txt����1.Visible = True
                txt����1.SelStart = 0: txt����1.SelLength = Len(txt����1)
                txt����2 = Mid(.�����ı�, i + 1)
            Else
                txt����1 = .�����ı�
                txt����2 = ""
            End If
            txt����1.Tag = ""
            txt����1.MaxLength = .Ҫ�س���
            lbl��λ = .Ҫ�ص�λ
            If Trim(.Ҫ�ص�λ) <> "" Then
                lbl��λ.Visible = True
            Else
                lbl��λ.Visible = False
            End If
            If .Ҫ��С�� > 0 Then
                txt����2.MaxLength = .Ҫ��С��
                txt����2.Visible = True
                lblDot.Visible = True
            Else
                txt����2.Visible = False
                lblDot.Visible = False
            End If
        Case 2
            T = Split(.Ҫ��ֵ��, ";")
            vfg��ѡ.Clear
            vfg��ѡ.RowHeightMax = 240
            vfg��ѡ.Cols = 3
            vfg��ѡ.ColWidth(0) = 80
            vfg��ѡ.ColWidth(1) = 200
            vfg��ѡ.Rows = UBound(T) + 1
            For i = 0 To UBound(T)
                vfg��ѡ.Cell(flexcpText, i, 2) = Trim(T(i))
                vfg��ѡ.Cell(flexcpPicture, i, 1) = imgOpt1.Picture
            Next i
            
            If Element.������̬ = 0 Then
                strTmp = Trim(.�����ı�)
            Else
                'չ����ʽ   '������
                strTmp = ""
                i = InStr(1, .�����ı�, "��")
                If i > 0 Then
                    j = InStr(i, .�����ı�, "��")
                    If j > 0 Then
                        strTmp = Trim(Mid(.�����ı�, i + 1, j - i - 1))
                    Else
                        strTmp = Trim(Mid(.�����ı�, i + 1))
                    End If
                Else
                    strTmp = ""
                End If
            End If
'            vfg��ѡ.Cell(flexcpChecked, 1, 0, vfg��ѡ.Rows - 1, 0) = flexUnchecked
            vfg��ѡ.FocusRect = flexFocusNone
            vfg��ѡ.Editable = flexEDKbdMouse
            vfg��ѡ.Row = 0
            vfg��ѡ.Col = 0
            Dim blnFinded As Boolean
            vfg��ѡ.Row = 0
            For i = 0 To vfg��ѡ.Rows - 1
                If strTmp = vfg��ѡ.Cell(flexcpText, i, 2) And blnFinded = False Then
                    'vfg��ѡ.Cell(flexcpChecked, i, 1) = flexChecked
                    vfg��ѡ.Cell(flexcpPicture, i, 1) = imgOpt2.Picture
                    vfg��ѡ.Row = i
                    blnFinded = True
                Else
                    'vfg��ѡ.Cell(flexcpChecked, i, 1) = flexUnchecked
                    vfg��ѡ.Cell(flexcpPicture, i, 1) = imgOpt1.Picture
                End If
            Next
        Case 3
            T = Split(.Ҫ��ֵ��, ";")
            vfg��ѡ.Clear
            vfg��ѡ.RowHeightMax = 240
            vfg��ѡ.Cols = 2
            vfg��ѡ.Rows = UBound(T) + 1
            For i = 0 To UBound(T)
                vfg��ѡ.Cell(flexcpText, i, 1) = T(i)
            Next i
            
            If Element.������̬ = 0 Then
                strTmp = "��" & Trim(.�����ı�) & "��"
            Else
                'չ����ʽ
                strTmp = ""
                i = InStr(1, .�����ı�, "��")
                Do While i > 0
                    j = InStr(i, .�����ı�, " ")
                    If j > 0 Then
                        strTmp = strTmp & "��" & Mid(.�����ı�, i + 1, j - i - 1)
                    Else
                        strTmp = strTmp & "��" & Mid(.�����ı�, i + 1)
                    End If
                    i = InStr(i + 1, .�����ı�, "��")
                Loop
                strTmp = strTmp & "��"
            End If
            vfg��ѡ.Cell(flexcpChecked, 0, 0, vfg��ѡ.Rows - 1, 0) = flexUnchecked
            vfg��ѡ.Editable = flexEDKbdMouse
        
            vfg��ѡ.ColWidth(0) = 240
            For i = 0 To vfg��ѡ.Rows - 1
                If InStr(1, strTmp, "��" & vfg��ѡ.Cell(flexcpText, i, 1) & "��") > 0 Then
                    vfg��ѡ.Cell(flexcpChecked, i, 0) = flexChecked
                Else
                    vfg��ѡ.Cell(flexcpChecked, i, 0) = flexUnchecked
                End If
            Next
            vfg��ѡ.Row = 0

        End Select
    End With
    
    Me.Left = X
    Me.Top = Y
    Me.Width = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmElementEdit", "MainWidth", 2500)
    If Me.Element.Ҫ�ر�ʾ <> 1 And (Me.Element.�滻�� = 0 Or EditType = cprET_�������༭ Or EditType = cprET_���������) Then
        Me.Height = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmElementEdit", "MainHeight", 3690)
    End If
    Call Form_Resize
    If Me.Visible = False Then Me.Show blnModel, frmParent
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form_Deactivate
End Sub

Private Sub mnu_Copy_Click()
    On Error Resume Next
    If glngTXTProc <> 0 Then
        Call SetWindowLong(txt�ı�.hWnd, GWL_WNDPROC, glngTXTProc)
        glngTXTProc = 0
    End If
    gstrCopyPID = CStr(frmParent.Document.EPRPatiRecInfo.����ID)
    Clipboard.SetText txt�ı�.SelText
End Sub

Private Sub mnu_Cut_Click()
    On Error Resume Next
    If glngTXTProc <> 0 Then
        Call SetWindowLong(txt�ı�.hWnd, GWL_WNDPROC, glngTXTProc)
        glngTXTProc = 0
    End If
    gstrCopyPID = CStr(frmParent.Document.EPRPatiRecInfo.����ID)
    If txt�ı�.SelLength <> 0 Then
        Clipboard.SetText txt�ı�.SelText
        txt�ı�.SelText = ""
    End If
End Sub

Private Sub mnu_Delete_Click()
    If glngTXTProc <> 0 Then
        Call SetWindowLong(txt�ı�.hWnd, GWL_WNDPROC, glngTXTProc)
        glngTXTProc = 0
    End If
    If txt�ı�.SelLength <> 0 Then
        txt�ı�.SelText = ""
    End If
End Sub

Private Sub txt����1_GotFocus()
    zlCommFun.OpenIme
    txt����1.SelStart = 0
    txt����1.SelLength = Len(txt����1)
    ud����.BuddyControl = txt����1
End Sub

Private Sub txt����1_KeyPress(KeyAscii As Integer)
    If InStr("1234567890. " & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If KeyAscii = vbKeySpace Or InStr(".", Chr(KeyAscii)) = 1 Then
        KeyAscii = 0
        If txt����2.Visible And txt����2.Enabled Then
            txt����2.SelStart = 0
            txt����2.SelLength = Len(txt����2)
            txt����2.SetFocus
        End If
    End If
End Sub

Private Sub txt����2_Change()
    If txt����1.Tag = "" Then
        If Me.Element.Ҫ��С�� > 0 Then
            Dim lngLen As Long, strR As String
            lngLen = Len(Trim(txt����2))
            If lngLen > Me.Element.Ҫ��С�� Then
                strR = Trim(txt����1.Text) & "." & Trim(txt����2) & String(Me.Element.Ҫ��С�� - Len(Trim(txt����2)), "0")
            Else
                strR = Trim(txt����1.Text) & "." & Left(Trim(txt����2), Me.Element.Ҫ��С��)
            End If
        Else
            strR = Trim(txt����1.Text)
        End If
        Me.Element.�����ı� = IIf(Me.Element.Ҫ��С�� > 0, Format(strR, "0." & String(Me.Element.Ҫ��С��, "0")), strR)
'        Me.Element.�����ı� = Trim(txt����1.Text) & IIf(Me.Element.Ҫ��С�� > 0, "." & Format(Trim(txt����2.Text), String(Me.Element.Ҫ��С��, "0")), "")
    End If
End Sub

Private Sub txt����2_GotFocus()
    zlCommFun.OpenIme
    txt����2.SelStart = 0
    txt����2.SelLength = Len(txt����2)
    ud����.BuddyControl = txt����2
End Sub

Private Sub txt����2_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt�ı�_GotFocus()
    If Me.Element.Ҫ������ = 0 Then
        zlCommFun.OpenIme
    End If
End Sub

Private Sub txt�ı�_KeyPress(KeyAscii As Integer)
    If Me.Element.Ҫ������ = 0 Then
        '��ֵ�͵Ŀ��ƣ�ֻ���������֣�С����͸��ţ���С����ֻ��Ϊ1���������ڿ�ͷ������ֻ���ڿ�ʼ����
        'Asc(".") = vbKeyDelete = 46
        If Len(txt�ı�.Text) = 0 And KeyAscii = 46 Then KeyAscii = 0
        If InStr(1, txt�ı�.Text, ".") <> 0 And KeyAscii = 46 Then
            KeyAscii = 0
        ElseIf InStr(1, txt�ı�.Text, ".") = 0 And KeyAscii = 46 And txt�ı�.SelLength = Len(txt�ı�) And txt�ı�.SelStart = 0 Then
            KeyAscii = 0
        End If
        If txt�ı�.Text = "-" And KeyAscii = 46 Then KeyAscii = 0
        If KeyAscii = vbKeyBack Or KeyAscii = 46 Then Exit Sub
        If KeyAscii = Asc("-") Then
            If txt�ı�.SelStart <> 0 Then KeyAscii = 0
        Else
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
        End If
    End If
End Sub
Private Sub txt�ı�_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txt�ı�.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt�ı�.hWnd, GWL_WNDPROC, AddressOf WndMessage)
        PopupMenu mnu_Edit
        Call SetWindowLong(txt�ı�.hWnd, GWL_WNDPROC, glngTXTProc)
        glngTXTProc = 0
    End If
End Sub

Private Sub txt�ı�_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And glngTXTProc <> 0 Then
        Call SetWindowLong(txt�ı�.hWnd, GWL_WNDPROC, glngTXTProc)
        glngTXTProc = 0
    End If
End Sub

Private Sub vfg��ѡ_DblClick()
    Form_KeyPress vbKeyReturn
End Sub

Private Sub vfg��ѡ_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim i As Long, j As Long, strValue As String
    strValue = ""
    Select Case KeyCode
    Case vbKeySpace
        For i = 0 To vfg��ѡ.Rows - 1
            If i = vfg��ѡ.Row Then
                If vfg��ѡ.Cell(flexcpPicture, i, 1) = imgOpt2.Picture Then
                    vfg��ѡ.Cell(flexcpPicture, i, 1) = imgOpt1.Picture
                Else
                    vfg��ѡ.Cell(flexcpPicture, i, 1) = imgOpt2.Picture
                End If
            Else
                vfg��ѡ.Cell(flexcpPicture, i, 1) = imgOpt1.Picture
            End If
        Next
        If Me.Element.������̬ = 0 Then
            If vfg��ѡ.Cell(flexcpPicture, vfg��ѡ.Row, 1) = imgOpt2.Picture Then
                strValue = Trim(vfg��ѡ.Cell(flexcpText, vfg��ѡ.Row, 2))
            Else
                strValue = ""
            End If
        Else
            For i = 0 To vfg��ѡ.Rows - 1
                If vfg��ѡ.Cell(flexcpPicture, i, 1) = imgOpt2.Picture Then
                    strValue = strValue & IIf(j = 0, "��", "  ��") & Trim(vfg��ѡ.Cell(flexcpText, i, 2))
                    j = j + 1
                Else
                    strValue = strValue & IIf(j = 0, "��", "  ��") & Trim(vfg��ѡ.Cell(flexcpText, i, 2))
                    j = j + 1
                End If
            Next
        End If
        Element.�����ı� = strValue
        KeyCode = 0
    End Select
End Sub

Private Sub vfg��ѡ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If vfg��ѡ.Col = 1 Then
        Dim i As Long, j As Long, strValue As String
        strValue = ""
        
        LockWindowUpdate vfg��ѡ.hWnd
        For i = 0 To vfg��ѡ.Rows - 1
            If i = vfg��ѡ.Row Then
                vfg��ѡ.Cell(flexcpPicture, i, 1) = imgOpt2.Picture
            Else
                vfg��ѡ.Cell(flexcpPicture, i, 1) = imgOpt1.Picture
            End If
        Next
        If Me.Element.������̬ = 0 Then
            If vfg��ѡ.Cell(flexcpPicture, vfg��ѡ.Row, 1) = imgOpt2.Picture Then
                strValue = Trim(vfg��ѡ.Cell(flexcpText, vfg��ѡ.Row, 2))
            Else
                strValue = ""
            End If
        Else
            For i = 0 To vfg��ѡ.Rows - 1
                If vfg��ѡ.Cell(flexcpPicture, i, 1) = imgOpt2.Picture Then
                    strValue = strValue & IIf(j = 0, "��", "  ��") & Trim(vfg��ѡ.Cell(flexcpText, i, 2))
                    j = j + 1
                Else
                    strValue = strValue & IIf(j = 0, "��", "  ��") & Trim(vfg��ѡ.Cell(flexcpText, i, 2))
                    j = j + 1
                End If
            Next
        End If
        Element.�����ı� = strValue
        LockWindowUpdate 0
        UpdateWindow vfg��ѡ.hWnd
    'End If
End Sub
'#####################################################################################
'## �ڲ��ؼ��¼�
'#####################################################################################

Private Sub vfg��ѡ_AfterEdit(ByVal Row As Long, ByVal Col As Long) '������
    Dim i As Long, j As Long, strValue As String
    strValue = ""
    For i = 0 To vfg��ѡ.Rows - 1
        If Me.Element.������̬ = 0 Then
            If vfg��ѡ.Cell(flexcpChecked, i, 0) = flexChecked Then
                strValue = strValue & IIf(j = 0, "", "��") & Trim(vfg��ѡ.Cell(flexcpText, i, 1))
                j = j + 1
            End If
        Else
            If vfg��ѡ.Cell(flexcpChecked, i, 0) = flexChecked Then
                strValue = strValue & IIf(j = 0, "��", "  ��") & Trim(vfg��ѡ.Cell(flexcpText, i, 1))
                j = j + 1
            Else
                strValue = strValue & IIf(j = 0, "��", "  ��") & Trim(vfg��ѡ.Cell(flexcpText, i, 1))
                j = j + 1
            End If
        End If
    Next
    Element.�����ı� = strValue
End Sub

Private Sub vfg��ѡ_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    vfg��ѡ.Col = 0
End Sub

Private Sub vfg��ѡ_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    vfg��ѡ.Col = 0
    Cancel = True
End Sub

Private Sub Form_Activate()
    SetCtlFocus
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
    RaiseEvent pCancel
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmElementEdit", "MainWidth", Me.Width
    If Me.Element.Ҫ�ر�ʾ <> 1 And (Me.Element.�滻�� = 0 Or EditType = cprET_�������༭ Or EditType = cprET_���������) Then SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmElementEdit", "MainHeight", Me.Height
    If mblnModel = vbModal Then
        Unload Me
    Else
        Me.Hide
    End If
    Set frmParent = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Me.Element.Ҫ������ = 0 Then
            '��ֵ��
            Dim T As Variant, dblMax As Double, dblMin As Double
            T = Split(Me.Element.Ҫ��ֵ��, ";")    '��ʽ:  0;100000
            If UBound(T) < 1 Then
                dblMin = 0#
                dblMax = 0#
            Else
                dblMin = IIf(CLng(T(0)) > CLng(T(1)), CLng(T(1)), CLng(T(0)))
                dblMax = IIf(CLng(T(0)) > CLng(T(1)), CLng(T(0)), CLng(T(1)))
            End If
            If Me.Element.Ҫ�ر�ʾ = 0 Then
                '�ı���ʾ
                If Trim(txt�ı�) = "" Then
                    Me.Element.�����ı� = ""
                ElseIf Me.Element.Ҫ��ֵ�� <> ";" And Me.Element.Ҫ��ֵ�� <> "0;0" And Me.Element.Ҫ��ֵ�� <> "" Then
                    If Val(txt�ı�) > dblMax Then
                        txt�ı� = dblMax
                    ElseIf Val(txt�ı�) < dblMin Then
                        txt�ı� = dblMin
                    End If
                    Me.Element.�����ı� = IIf(Me.Element.Ҫ��С�� > 0, Format(txt�ı�, "0." & String(Me.Element.Ҫ��С��, "0")), txt�ı�)
                Else
                    Me.Element.�����ı� = IIf(Me.Element.Ҫ��С�� > 0, Format(txt�ı�, "0." & String(Me.Element.Ҫ��С��, "0")), txt�ı�)
                End If
            ElseIf Me.Element.Ҫ�ر�ʾ = 1 Then
                '���±�ʾ
                If Trim(Me.Element.�����ı�) <> "" And Me.Element.Ҫ��ֵ�� <> ";" And Me.Element.Ҫ��ֵ�� <> "0;0" Then
                    If Val(Me.Element.�����ı�) > dblMax Then
                        Me.Element.�����ı� = dblMax
                    ElseIf Val(Me.Element.�����ı�) < dblMin Then
                        Me.Element.�����ı� = dblMin
                    End If
                Else
                    Me.Element.�����ı� = IIf(Me.Element.Ҫ��С�� > 0, Format(Me.Element.�����ı�, "0." & String(Me.Element.Ҫ��С��, "0")), Me.Element.�����ı�)
                End If
            End If
        End If
        
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmElementEdit", "MainWidth", Me.Width
        If Me.Element.Ҫ�ر�ʾ <> 1 And (Me.Element.�滻�� = 0 Or EditType = cprET_�������༭ Or EditType = cprET_���������) Then SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmElementEdit", "MainHeight", Me.Height
        If mblnModel = vbModal Then
            Unload Me
        Else
            Me.Hide
        End If
        RaiseEvent pOK
    ElseIf KeyAscii = vbKeyEscape Then
        Form_Deactivate
    ElseIf KeyAscii = vbKeySpace Then
        If vfg��ѡ.Visible Then vfg��ѡ_KeyDown KeyAscii, 0
    ElseIf KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyUp Or KeyAscii = vbKeyDown Then
        If vfg��ѡ.Visible Then vfg��ѡ_KeyDown KeyAscii, 0
    End If
End Sub

Private Sub Form_Load()
    Me.KeyPreview = True
    Me.Width = 2000
End Sub

Private Sub Form_Paint()
    Cls
    Line (0, 0)-(ScaleWidth - Screen.TwipsPerPixelX, ScaleHeight - Screen.TwipsPerPixelY), &H996600, B
End Sub

Private Sub Form_Resize()
    Dim lX As Long, lY As Long
    lX = Screen.TwipsPerPixelX
    lY = Screen.TwipsPerPixelY
    
    txt����1.Visible = False
    txt����2.Visible = False
    lblDot.Visible = False
    lbl��λ.Visible = False
    shpBorder2.Visible = False
    txt�ı�.Visible = False
    ud����.Visible = False
    vfg��ѡ.Visible = False
    vfg��ѡ.Visible = False
    pic�滻��Ŀ.Visible = False
    On Error Resume Next
    If Not Me.Element Is Nothing Then
        '�Ƿ���ʾ��ʾ
        If Element.��ʾ = "" Then
            txtTip.Move 0, 0, 0, 0
            txtTip.Text = "����ʾ"
            txtTip.Visible = False
        Else
            txtTip.Move 80, picTitle.Height + 100, ScaleWidth - 160, (UBound(Split(Element.��ʾ, vbCrLf)) + 2) * TextHeight("��")
            txtTip.Text = "��ʾ:  " & Element.��ʾ
            txtTip.Visible = True
        End If
        
        Select Case Me.Element.Ҫ�ر�ʾ
        Case 0
            If Me.Element.�滻�� = 1 And EditType = cprET_�����ļ����� Then
                pic�滻��Ŀ.Move 80, picTitle.Height + IIf(Element.��ʾ = "", 0, txtTip.Height) + 120, ScaleWidth - 160, ScaleHeight - 200 - picStatus.Height - picTitle.Height - IIf(Element.��ʾ = "", 0, txtTip.Height)
                shpBorder1.Move pic�滻��Ŀ.Left - lX, pic�滻��Ŀ.Top - lY, pic�滻��Ŀ.Width + lX * 2, pic�滻��Ŀ.Height + lY * 2
                lblʾ��.Height = Abs(pic�滻��Ŀ.Height - lblʾ��.Top)
                pic�滻��Ŀ.Visible = True
                shpBorder1.Visible = True
            Else
                txt�ı�.Move 80, picTitle.Height + IIf(Element.��ʾ = "", 0, txtTip.Height) + 120, ScaleWidth - 160, ScaleHeight - 200 - picStatus.Height - picTitle.Height - IIf(Element.��ʾ = "", 0, txtTip.Height)
                shpBorder1.Move txt�ı�.Left - lX, txt�ı�.Top - lY, txt�ı�.Width + lX * 2, txt�ı�.Height + lY * 2
                txt�ı�.Visible = True
                shpBorder1.Visible = True
                If txt�ı�.Visible And txt�ı�.Enabled Then txt�ı�.SetFocus
            End If
        Case 1
            Dim lW1 As Long, lW2 As Long, lW3 As Long, lW4 As Long, lW5 As Long
            If Trim(Element.Ҫ�ص�λ) <> "" Then
                lbl��λ.Width = Me.TextWidth(lbl��λ) + lX * 6
                lbl��λ.Move Me.ScaleWidth - lbl��λ.Width + lX * 3, picTitle.Height + IIf(Element.��ʾ = "", 0, txtTip.Height) + 170
                lbl��λ.Visible = True
                lW5 = lbl��λ.Width
            Else
                lbl��λ.Visible = False
                lW5 = 0
            End If
            lW4 = ud����.Width + lX * 4
            ud����.Move Me.ScaleWidth - lW4 - lW5 + lX * 3, picTitle.Height + IIf(Element.��ʾ = "", 0, txtTip.Height) + 120
            ud����.Visible = True
            If Element.Ҫ��С�� > 0 Then
                txt����2.Width = Me.TextWidth(Space(Element.Ҫ��С��)) + lX * 4
                lW3 = txt����2.Width + lX
                txt����2.Move Me.ScaleWidth - lW5 - lW4 - lW3 + lX, picTitle.Height + IIf(Element.��ʾ = "", 0, txtTip.Height) + 170
                shpBorder2.Move txt����2.Left - lX, txt����2.Top - lY - 50, txt����2.Width + lX * 2, txt����2.Height + 50 + lY * 2
                shpBorder2.Visible = True
                txt����2.Visible = True
                lblDot.Width = Me.TextWidth(".") + lX * 2
                lW2 = lblDot.Width
                lblDot.Move txt����2.Left - lW2 + lX * 2, picTitle.Height + IIf(Element.��ʾ = "", 0, txtTip.Height) + 170
                lblDot.BackStyle = 0
                lblDot.Visible = True
            Else
                lW2 = 0
                lW3 = 0
                shpBorder2.Visible = False
                txt����2.Visible = False
                lblDot.Visible = False
            End If
            lW1 = Me.TextWidth(txt����1.Text) + lX * 2
            lW1 = IIf(lW1 < 400, 400, lW1)
            
            If Me.Width < lW1 + lW2 + lW3 + lW4 + lW5 Then Me.Width = lW1 + lW2 + lW3 + lW4 + lW5
            Me.Height = txt����1.Height + lY * 3 + picStatus.Height + picTitle.Height + IIf(Element.��ʾ = "", 0, txtTip.Height) + 180
            
            txt����1.Move 80, picTitle.Height + IIf(Element.��ʾ = "", 0, txtTip.Height) + 170, Me.ScaleWidth - lW5 - lW4 - lW3 - lW2 - lX * 4
            shpBorder1.Move txt����1.Left - lX, txt����1.Top - lY - 50, txt����1.Width + lX * 2, txt����1.Height + 50 + lY * 2
            txt����1.Visible = True
            shpBorder1.Visible = True
            If txt����1.Visible And txt����1.Enabled Then txt����1.SelStart = 0: txt����1.SelLength = Len(txt����1): txt����1.SetFocus
        Case 2
            vfg��ѡ.Move 80, picTitle.Height + IIf(Element.��ʾ = "", 0, txtTip.Height) + 120, ScaleWidth - 160, ScaleHeight - 200 - picStatus.Height - picTitle.Height - IIf(Element.��ʾ = "", 0, txtTip.Height)
            shpBorder1.Move vfg��ѡ.Left - lX, vfg��ѡ.Top - lY, vfg��ѡ.Width + lX * 3, vfg��ѡ.Height + lY * 2
            vfg��ѡ.Visible = True
            shpBorder1.Visible = True
            If vfg��ѡ.Visible And vfg��ѡ.Enabled Then vfg��ѡ.SetFocus
        Case 3
            vfg��ѡ.Move 80, picTitle.Height + IIf(Element.��ʾ = "", 0, txtTip.Height) + 120, ScaleWidth - 160, ScaleHeight - 200 - picStatus.Height - picTitle.Height - IIf(Element.��ʾ = "", 0, txtTip.Height)
            shpBorder1.Move vfg��ѡ.Left - lX, vfg��ѡ.Top - lY, vfg��ѡ.Width + lX * 3, vfg��ѡ.Height + lY * 2
            vfg��ѡ.Visible = True
            shpBorder1.Visible = True
            If vfg��ѡ.Visible And vfg��ѡ.Enabled Then vfg��ѡ.SetFocus
        End Select
    End If
    
    picTitle.Move 60, 60, ScaleWidth - 120
    picStatus.Move lX, ScaleHeight - picStatus.Height - lY, ScaleWidth - lX * 2
    
    If Me.Top + Me.Height > Screen.Height - 800 Then Me.Top = Me.Top - Me.Height - 200
    If Me.Left + Me.Width > Screen.Width Then Me.Left = Me.Left - Me.Width
End Sub

Private Sub imgResize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgResize.Tag = "Down"
    lngX = X
    lngY = Y
End Sub

Private Sub imgResize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imgResize.Tag = "Down" Then
        If Me.Width + X - lngX >= 1000 And Me.Width + X - lngX <= 12000 Then
            Me.Width = Me.Width + X - lngX
        End If
        If Me.Height + Y - lngY >= 1000 And Me.Height + Y - lngY <= 9000 Then
            Me.Height = Me.Height + Y - lngY
        End If
        DoEvents
    End If
End Sub

Private Sub imgResize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgResize.Tag = ""
    Call SetCtlFocus
End Sub

Private Sub imgTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgTitle.Tag = "Down"
    lngX = X
    lngY = Y
End Sub

Private Sub imgTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imgTitle.Tag = "Down" Then
        Me.Move Me.Left + X - lngX, Me.Top + Y - lngY
    Else
        If X > 0 And X < picTitle.ScaleWidth And Y > 0 And Y < picTitle.ScaleHeight Then
            SetCapture picTitle.hWnd
            picTitle.Cls
            picTitle.BackColor = &HC2EEFF
            picTitle.Line (0, 0)-(picTitle.ScaleWidth - Screen.TwipsPerPixelX, picTitle.ScaleHeight - Screen.TwipsPerPixelY), &H800000, B
            lblInfo.Caption = "���������ק�����ƶ��༭��"
        Else
            ReleaseCapture
            picTitle.Cls
            picTitle.BackColor = &HF5BE9E
            lblInfo.Caption = "Esc:ȡ���༭���س�:�����޸ġ�"
        End If
    End If
End Sub

Private Sub imgTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgTitle.Tag = ""
    Call SetCtlFocus
End Sub

Private Sub picStatus_Resize()
    imgResize.Move picStatus.ScaleWidth - imgResize.Width, 0
End Sub

Private Sub picTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picTitle.Tag = "Down"
    lngX = X
    lngY = Y
End Sub

Private Sub picTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picTitle.Tag = "Down" Then
        Me.Move Me.Left + X - lngX, Me.Top + Y - lngY
    Else
        If X > 0 And X < picTitle.ScaleWidth And Y > 0 And Y < picTitle.ScaleHeight Then
            SetCapture picTitle.hWnd
            picTitle.Cls
            picTitle.BackColor = &HC2EEFF
            picTitle.Line (0, 0)-(picTitle.ScaleWidth - Screen.TwipsPerPixelX, picTitle.ScaleHeight - Screen.TwipsPerPixelY), &H800000, B
            lblInfo.Caption = "���������ק�����ƶ��༭��"
            If picTitle.Tag = "Down" Then
                Me.Move Me.Left + X - lngX, Me.Top + Y - lngY
            End If
        Else
            ReleaseCapture
            picTitle.Cls
            picTitle.BackColor = &HF5BE9E
            lblInfo.Caption = "Esc:ȡ���༭���س�:�����޸ġ�"
        End If
    End If
End Sub

Private Sub picTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picTitle.Tag = ""
    Call SetCtlFocus
End Sub

Private Sub picTitle_Resize()
    imgTitle.Move (picTitle.ScaleWidth - imgTitle.Width) / 2, 30
End Sub

Private Sub txt����1_Change()
    If txt����1.Tag = "" Then
        Me.Element.�����ı� = Trim(txt����1.Text) & IIf(Me.Element.Ҫ��С�� > 0, "." & Format(Trim(txt����2.Text), String(Me.Element.Ҫ��С��, "0")), "")
    End If
End Sub

Private Sub txt�ı�_Change()
    Me.Element.�����ı� = Trim(txt�ı�.Text)
End Sub

'#####################################################################################
'## �ֲ�����
'#####################################################################################

Private Sub SetCtlFocus()
    '���ÿؼ�����
    If txt����1.Visible And txt����1.Enabled Then
        txt����1.SetFocus
    ElseIf txt����2.Visible And txt����2.Enabled Then
        txt����2.SetFocus
    ElseIf txt�ı�.Visible And txt�ı�.Enabled Then
        txt�ı�.SetFocus
    ElseIf vfg��ѡ.Visible And vfg��ѡ.Enabled Then
        vfg��ѡ.SetFocus
    ElseIf vfg��ѡ.Visible And vfg��ѡ.Enabled Then
        vfg��ѡ.SetFocus
    End If
End Sub

Private Sub vfg��ѡ_DblClick()
Dim lngCol As Long
    lngCol = vfg��ѡ.MouseCol
    If lngCol = 1 Then
        Form_KeyPress vbKeyReturn
    End If
End Sub
