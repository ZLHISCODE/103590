VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmImgSrv 
   BorderStyle     =   0  'None
   Caption         =   "Ӱ����շ���"
   ClientHeight    =   6390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraReceiveSet 
      ForeColor       =   &H00FF0000&
      Height          =   6315
      Left            =   90
      TabIndex        =   6
      Top             =   0
      Width           =   11310
      Begin VB.ComboBox cboDevice 
         Height          =   300
         ItemData        =   "frmImgSrv.frx":0000
         Left            =   1380
         List            =   "frmImgSrv.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   345
         Width           =   1815
      End
      Begin VB.ComboBox cboEncode 
         Height          =   300
         ItemData        =   "frmImgSrv.frx":0004
         Left            =   1380
         List            =   "frmImgSrv.frx":0014
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Frame FraAuto 
         Caption         =   "�Զ�ƥ������"
         Height          =   2130
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   10995
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1140
            Left            =   5670
            TabIndex        =   16
            Top             =   690
            Width           =   4065
            Begin VB.OptionButton optMatch 
               Caption         =   "�� ""ҽ��ID"" ƥ��"
               Height          =   195
               Index           =   2
               Left            =   690
               TabIndex        =   19
               ToolTipText     =   "��ҽ��ID�����˺ͽ��յ�Ӱ�����ƥ��"
               Top             =   780
               Width           =   3300
            End
            Begin VB.OptionButton optMatch 
               Caption         =   "�� ""����"" ƥ��"
               Height          =   195
               Index           =   0
               Left            =   690
               TabIndex        =   18
               ToolTipText     =   "�����Ž����˺ͽ��յ�Ӱ�����ƥ��"
               Top             =   240
               Width           =   3300
            End
            Begin VB.OptionButton optMatch 
               Caption         =   "�� ""���˱�ʶ��(����/סԺ��)"" ƥ��"
               Height          =   195
               Index           =   1
               Left            =   690
               TabIndex        =   17
               ToolTipText     =   "�����˱�ʶ�Ž����˺ͽ��յ�Ӱ�����ƥ��"
               Top             =   510
               Width           =   3300
            End
            Begin VB.Label lblDataItem 
               Caption         =   "���ݿ���Ŀ"
               Height          =   885
               Left            =   90
               TabIndex        =   20
               Top             =   150
               Width           =   225
            End
            Begin VB.Line Line5 
               X1              =   345
               X2              =   510
               Y1              =   585
               Y2              =   585
            End
            Begin VB.Line Line6 
               X1              =   510
               X2              =   510
               Y1              =   315
               Y2              =   890
            End
            Begin VB.Line Line7 
               X1              =   510
               X2              =   630
               Y1              =   315
               Y2              =   315
            End
            Begin VB.Line Line8 
               X1              =   525
               X2              =   630
               Y1              =   870
               Y2              =   870
            End
         End
         Begin VB.OptionButton optImgMatch 
            Caption         =   "Accession Number"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   10
            Top             =   1155
            Width           =   1740
         End
         Begin VB.OptionButton optImgMatch 
            Caption         =   "Patient Name"
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   9
            Top             =   1425
            Width           =   1740
         End
         Begin VB.OptionButton optImgMatch 
            Caption         =   "Patient ID"
            Height          =   255
            Index           =   0
            Left            =   1080
            TabIndex        =   8
            Top             =   885
            Width           =   1740
         End
         Begin VB.ComboBox cboMatchOther 
            Height          =   300
            ItemData        =   "frmImgSrv.frx":0045
            Left            =   8550
            List            =   "frmImgSrv.frx":004F
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   330
            Width           =   1785
         End
         Begin VB.CheckBox chkMatchStudyUID 
            Caption         =   "���� ""���UID"" ƥ��"
            Height          =   300
            Left            =   120
            TabIndex        =   2
            Top             =   330
            Width           =   2100
         End
         Begin VB.CheckBox chkImageType 
            Caption         =   "����ͼ�����Ͳ������"
            Height          =   300
            Left            =   4170
            TabIndex        =   3
            Top             =   330
            Width           =   2130
         End
         Begin VB.Line Line4 
            X1              =   930
            X2              =   1035
            Y1              =   1545
            Y2              =   1545
         End
         Begin VB.Line Line3 
            X1              =   915
            X2              =   1035
            Y1              =   990
            Y2              =   990
         End
         Begin VB.Line Line2 
            X1              =   915
            X2              =   915
            Y1              =   990
            Y2              =   1565
         End
         Begin VB.Line Line1 
            X1              =   735
            X2              =   900
            Y1              =   1260
            Y2              =   1260
         End
         Begin VB.Label lblImgItem 
            Caption         =   "ͼ����Ŀ"
            Height          =   690
            Left            =   480
            TabIndex        =   11
            Top             =   930
            Width           =   225
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "����ƥ��(&A)"
            Height          =   180
            Left            =   7545
            TabIndex        =   4
            ToolTipText     =   "�ò������[���ݿ���Ŀ]��""���˱�ʶ��""/""����""ƥ����Ч"
            Top             =   390
            Width           =   990
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgList 
         Height          =   1290
         Left            =   5010
         TabIndex        =   14
         Top             =   225
         Width           =   6150
         _cx             =   10848
         _cy             =   2275
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
         Rows            =   3
         Cols            =   2
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
         AutoResize      =   0   'False
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
      Begin VB.Line Line12 
         X1              =   4545
         X2              =   4710
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line11 
         X1              =   4710
         X2              =   4710
         Y1              =   570
         Y2              =   1145
      End
      Begin VB.Line Line10 
         X1              =   4710
         X2              =   4830
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Line Line9 
         X1              =   4725
         X2              =   4830
         Y1              =   1125
         Y2              =   1125
      End
      Begin VB.Label lblRoute 
         Caption         =   "�Զ�ת������"
         Height          =   1080
         Left            =   4335
         TabIndex        =   15
         Top             =   330
         Width           =   225
      End
      Begin VB.Label LblCmp 
         AutoSize        =   -1  'True
         Caption         =   "ѹ����ʽ(&Y)"
         Height          =   180
         Left            =   240
         TabIndex        =   1
         Top             =   1260
         Width           =   990
      End
      Begin VB.Label lblSave 
         AutoSize        =   -1  'True
         Caption         =   "�洢�豸(&F)"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   405
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmImgSrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngSrvID As Long

'�Զ�·�ɵĲ����ͳ���
Private str�Զ�·��Ŀ�ĵ� As String
Private str�Զ�·��ѹ����ʽ As String
Private str�Զ�·��Ŀ¼�ṹ As String
Private Const AR��ѹ�� = "��ѹ��"
Private Const AR����ѹ�� = "����ǰ��ʽѹ��"
Private Const AR��鼶�� = "��鼶��(Ĭ��)"
Private Const AR���м��� = "���м���(3D)"


Public Sub ShowRefresh(ByVal SrvID As Long)
    mlngSrvID = SrvID
    If mlngSrvID = 0 Then
        fraReceiveSet.Caption = "�Ϸ��б�����ѡ������δ���棬���ܽ������ã�"
        fraReceiveSet.Enabled = False
    Else
        fraReceiveSet.Caption = ""
        fraReceiveSet.Enabled = True
    End If
    RefreshPara
End Sub

Private Sub RefreshPara()
Dim rsTemp As New ADODB.Recordset, i As Integer
    gstrSQL = "select ����ID,�������� ,����ֵ from Ӱ��DICOM������� where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����", mlngSrvID)
    InitvfgList
    cboDevice.ListIndex = -1
    cboEncode.ListIndex = -1
    chkImageType.value = False
    chkMatchStudyUID.value = False
    cboMatchOther.ListIndex = 0
    str�Զ�·��Ŀ�ĵ� = ""
    str�Զ�·��ѹ����ʽ = ""
    str�Զ�·��Ŀ¼�ṹ = ""
    
    Do Until rsTemp.EOF
        Select Case rsTemp!��������
            Case "�洢�豸"
                Call SeekIndexWithNo(cboDevice, Nvl(rsTemp!����ֵ), True)
            Case "ѹ����ʽ"
                Call SeekIndex(cboEncode, Nvl(rsTemp!����ֵ), True)
            Case "�����UIDƥ��"
                chkMatchStudyUID.value = rsTemp!����ֵ
            Case "�����Ͳ������"
                chkImageType.value = rsTemp!����ֵ
            Case "ƥ��ͼ����Ŀ"
                optImgMatch(Nvl(rsTemp!����ֵ, 0)) = True
            Case "ƥ�����ݿ���Ŀ"
                optMatch(Nvl(rsTemp!����ֵ, 0)) = True
            Case "��Ϣת��" '�����ʽ "Ŀ�ĵ�1|Ŀ�ĵ�2---" ��Ϣ��UDP��Ϣ,��������Ϊ����վ����������ʵ���Զ�����,����鿴ʱ����ȡ
                Call FillBlRoute("��Ϣת��", Nvl(rsTemp!����ֵ), "", "")
            Case "�Զ�·��"
                str�Զ�·��Ŀ�ĵ� = Nvl(rsTemp!����ֵ)
            Case "�Զ�·��ѹ����ʽ"
                str�Զ�·��ѹ����ʽ = Nvl(rsTemp!����ֵ)
            Case "�Զ�·��Ŀ¼�ṹ"
                str�Զ�·��Ŀ¼�ṹ = Nvl(rsTemp!����ֵ)
            Case "�洢���˷�ʽ"
                Call SeekIndexWithNo(cboMatchOther, Nvl(rsTemp!����ֵ, 0), True)
        End Select
        rsTemp.MoveNext
    Loop
    
    '��д�Զ�·�ɲ���
    If str�Զ�·��Ŀ�ĵ� <> "" Then
        Call FillBlRoute("�Զ�·��", str�Զ�·��Ŀ�ĵ�, str�Զ�·��ѹ����ʽ, str�Զ�·��Ŀ¼�ṹ)
    End If
    
    '����ͼ�����Ͳ�����С����������ֻ�����ĳЩCTʹ��
    gstrSQL = "select Ӱ����� from Ӱ��DICOM����� A,Ӱ���豸Ŀ¼ B WHERE A.����ID=[1] AND A.�豸��=B.�豸��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����", mlngSrvID)
    If Not rsTemp.EOF Then
        If UCase(rsTemp!Ӱ�����) <> "CT" Then
            chkImageType.value = 0
            chkImageType.Visible = False
        End If
    End If
End Sub

Private Sub FillBlRoute(ByVal strType As String, ByVal strData As String, ByVal strPara1 As String, ByVal strPara2 As String)
'------------------------------------------------
'���ܣ���д�Զ�·�ɻ�����Ϣת������Ϣ����Ϣת����ʽ���ڻ�û��ʹ��
'������ strType--����---���Զ�·�ɡ����ߡ���Ϣת����
'       strData--�������ݣ������Զ�·����Ŀ�ĵ�
'       strPara1--��������1�������Զ�·����ѹ����ʽ,
'       strPara2--��������2�������Զ�·����Ŀ¼�ṹ
'���أ��ޣ�ֱ����д�ؼ�
'------------------------------------------------
    Dim i As Integer, j As Integer
    Dim blnWritePara As Boolean
    '�����ʽ "·�ɷ�ʽ,Ŀ�ĵ�|---" ·�ɷ�ʽ��Ϊ ·��/��Ϣ ,·�ɼ���DICOM����,��Ϣ��UDP��Ϣ,��������Ϊ����վ����������ʵ���Զ�����,����鿴ʱ����ȡ
    
    If strData = "" Then Exit Sub
    
    '�������
    If strType = "�Զ�·��" Then
        If UBound(Split(strData, "|")) = UBound(Split(strPara1, "|")) And UBound(Split(strData, "|")) = UBound(Split(strPara2, "|")) Then
            blnWritePara = True
        Else
            blnWritePara = False
        End If
    End If
    
    With vfgList
        For i = 0 To UBound(Split(strData, "|"))
            .TextMatrix(.Rows - 1, 0) = strType
            If strType = "�Զ�·��" Then '�Զ�·�ɱ�����豸��,ͨ��ѭ����Cbo������ȡ��
                For j = 0 To UBound(Split(.ColComboList(1), "|"))
                    If InStr(Split(.ColComboList(1), "|")(j), Split(strData, "|")(i)) > 0 Then
                        .TextMatrix(.Rows - 1, 1) = Split(.ColComboList(1), "|")(j)
                        If blnWritePara = True Then '�в��������ղ�������д
                            .TextMatrix(.Rows - 1, 2) = IIf(Split(strPara1, "|")(i) = 1, AR��ѹ��, AR����ѹ��)
                            .TextMatrix(.Rows - 1, 3) = IIf(Split(strPara2, "|")(i) = 1, AR���м���, AR��鼶��)
                        Else    'û�в���������дĬ��ֵ
                            .TextMatrix(.Rows - 1, 2) = AR����ѹ��
                            .TextMatrix(.Rows - 1, 3) = AR��鼶��
                        End If
                    End If
                Next
            Else
                .TextMatrix(.Rows - 2, 1) = Split(strData, "|")(i)
            End If
            .Rows = .Rows + 1
        Next
        .TextMatrix(.Rows - 1, 0) = "�Զ�·��"
    End With
End Sub

Public Sub SavePara()
    Dim strData As String
    Dim i As Integer, strData1 As String
    Dim arrSQL() As Variant
    Dim blnInTrans As Boolean       '�Ƿ���������֮��
    
    On Error GoTo ErrHandle
    If cboDevice.Text = "" Then
        MsgBoxD Me, "��ѡ��洢�豸��", vbInformation, gstrSysName: cboDevice.SetFocus: Exit Sub
    End If
    
    If cboEncode.Text = "" Then
        MsgBoxD Me, "��ѡ��ѹ����ʽ��", vbInformation, gstrSysName: cboEncode.SetFocus: Exit Sub
    End If
    
    arrSQL = Array()
    
    If cboDevice.ListIndex <> -1 Then
        gstrSQL = "Zl_Ӱ��DICOM�������_SAVE(" & mlngSrvID & ",'�洢�豸','" & NeedNo(cboDevice.Text) & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
    End If
    
    If cboEncode.ListIndex <> -1 Then
        gstrSQL = "Zl_Ӱ��DICOM�������_SAVE(" & mlngSrvID & ",'ѹ����ʽ','" & NeedName(cboEncode.Text) & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
    End If
    
    gstrSQL = "Zl_Ӱ��DICOM�������_SAVE(" & mlngSrvID & ",'�����UIDƥ��','" & chkMatchStudyUID.value & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL
    
    gstrSQL = "Zl_Ӱ��DICOM�������_SAVE(" & mlngSrvID & ",'�����Ͳ������','" & chkImageType.value & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL
    
    gstrSQL = "Zl_Ӱ��DICOM�������_SAVE(" & mlngSrvID & ",'�洢���˷�ʽ','" & NeedNo(cboMatchOther.Text) & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL
    
    strData = 0
    For i = 0 To optImgMatch.UBound
        If optImgMatch(i).value = True Then
            strData = i
            Exit For
        End If
    Next
    If strData = "" Then strData = "0"
    gstrSQL = "Zl_Ӱ��DICOM�������_SAVE(" & mlngSrvID & ",'ƥ��ͼ����Ŀ','" & strData & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL

    strData = 0
    For i = 0 To optMatch.UBound
        If optMatch(i).value = True Then
            strData = i
            Exit For
        End If
    Next
    If strData = "" Then strData = "0"
    gstrSQL = "Zl_Ӱ��DICOM�������_SAVE(" & mlngSrvID & ",'ƥ�����ݿ���Ŀ','" & strData & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL
    
    gstrSQL = "Zl_Ӱ��DICOM�������_Delete(" & mlngSrvID & ",'�Զ�·��')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL

    gstrSQL = "Zl_Ӱ��DICOM�������_Delete(" & mlngSrvID & ",'��Ϣת��')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL
    
    With vfgList
        strData = ""
        str�Զ�·��ѹ����ʽ = ""
        str�Զ�·��Ŀ¼�ṹ = ""
        For i = 1 To vfgList.Rows - 1
            If Trim(vfgList.TextMatrix(i, 1)) <> "" And vfgList.RowHidden(i) = False Then
                If vfgList.TextMatrix(i, 0) = "�Զ�·��" Then
                    If InStr(strData, NeedNo(vfgList.TextMatrix(i, 1))) = 0 Then '�ظ��Ĳ�����
                        strData = strData & "|" & NeedNo(vfgList.TextMatrix(i, 1))
                        str�Զ�·��ѹ����ʽ = str�Զ�·��ѹ����ʽ & "|" & IIf(vfgList.TextMatrix(i, 2) = AR��ѹ��, 1, 0)
                        str�Զ�·��Ŀ¼�ṹ = str�Զ�·��Ŀ¼�ṹ & "|" & IIf(vfgList.TextMatrix(i, 3) = AR���м���, 1, 0)
                    End If
                Else
                    If InStr(strData1, vfgList.TextMatrix(i, 1)) = 0 Then '�ظ��Ĳ�����
                        strData1 = strData1 & "|" & vfgList.TextMatrix(i, 1)
                    End If
                End If
            End If
        Next
    End With
    strData = Mid(strData, 2)
    gstrSQL = "Zl_Ӱ��DICOM�������_SAVE(" & mlngSrvID & ",'�Զ�·��','" & strData & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL
    
    strData1 = Mid(strData1, 2)
    gstrSQL = "Zl_Ӱ��DICOM�������_SAVE(" & mlngSrvID & ",'��Ϣת��','" & strData1 & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL
    
    str�Զ�·��ѹ����ʽ = Mid(str�Զ�·��ѹ����ʽ, 2)
    gstrSQL = "Zl_Ӱ��DICOM�������_SAVE(" & mlngSrvID & ",'�Զ�·��ѹ����ʽ','" & str�Զ�·��ѹ����ʽ & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL
    
    str�Զ�·��Ŀ¼�ṹ = Mid(str�Զ�·��Ŀ¼�ṹ, 2)
    gstrSQL = "Zl_Ӱ��DICOM�������_SAVE(" & mlngSrvID & ",'�Զ�·��Ŀ¼�ṹ','" & str�Զ�·��Ŀ¼�ṹ & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL
    
    gcnOracle.BeginTrans        '��ʼ�������
    blnInTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "����ͼ����ղ���")
    Next i
    gcnOracle.CommitTrans
    blnInTrans = False
    
    RefreshPara
   Exit Sub
ErrHandle:
    If blnInTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Call InitvfgList
End Sub
Private Sub InitvfgList()
Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHandle
    With vfgList
        .Clear
        .FixedRows = 1
        .Rows = 2
        .Cols = 4
        .ColWidth(0) = 800
        .ColWidth(1) = 800
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .TextMatrix(0, 0) = "ת������"
        .TextMatrix(0, 1) = "ת��Ŀ�ĵ�"
        .TextMatrix(0, 2) = "ѹ����ʽ"
        .TextMatrix(0, 3) = "Ŀ¼�ṹ"
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .TextMatrix(1, 0) = "�Զ�·��"
        .TextMatrix(1, 2) = AR����ѹ��
        .TextMatrix(1, 3) = AR��鼶��
    End With
    gstrSQL = "select �豸��,�豸�� from Ӱ���豸Ŀ¼ where ����=1 and NVL(״̬,0)=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�洢�豸")
    cboDevice.Clear
    Dim strList As String
    Do Until rsTemp.EOF
        cboDevice.AddItem rsTemp!�豸�� & "-" & rsTemp!�豸��
        strList = strList & "|" & rsTemp!�豸�� & "-" & rsTemp!�豸��
        rsTemp.MoveNext
    Loop
    vfgList.ColComboList(1) = strList
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vfgList_Click()
    With vfgList
        If .Col = 0 Or .Col = 2 Or .Col = 3 Then
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub vfgList_DblClick()
    With vfgList
        If .Col = 0 Then
            If .TextMatrix(.Row, .Col) = "�Զ�·��" Then
                .TextMatrix(.Row, .Col) = "��Ϣת��"
            Else
                .TextMatrix(.Row, .Col) = "�Զ�·��"
            End If
            .TextMatrix(.Row, 1) = ""
            If .TextMatrix(.Row, 0) = "�Զ�·��" Then
                .TextMatrix(.Row, 2) = AR����ѹ��
                .TextMatrix(.Row, 3) = AR��鼶��
            Else
                .TextMatrix(.Row, 2) = ""
                .TextMatrix(.Row, 3) = ""
            End If
        End If
        
        If .Col = 2 Then
            If .TextMatrix(.Row, 2) = AR����ѹ�� Then
                .TextMatrix(.Row, 2) = AR��ѹ��
            Else
                .TextMatrix(.Row, 2) = AR����ѹ��
            End If
        End If
        
        If .Col = 3 Then
            If .TextMatrix(.Row, 3) = AR��鼶�� Then
                .TextMatrix(.Row, 3) = AR���м���
            Else
                .TextMatrix(.Row, 3) = AR��鼶��
            End If
        End If
    End With
End Sub

Private Sub vfgList_KeyDown(KeyCode As Integer, Shift As Integer)
    '�س�������һ��
    If KeyCode = vbKeyReturn Then
        vfgList.Rows = vfgList.Rows + 1
        vfgList.TextMatrix(vfgList.Rows - 1, 0) = "�Զ�·��"
        vfgList.TextMatrix(vfgList.Rows - 1, 2) = AR����ѹ��
        vfgList.TextMatrix(vfgList.Rows - 1, 3) = AR��鼶��
    End If
    'deleteɾ�����һ��
    If KeyCode = vbKeyDelete And vfgList.Row >= 1 Then
        If MsgBoxD(Me, "�Ƿ�ɾ�����У�", vbYesNo) = vbYes Then
            vfgList.RowHidden(vfgList.Row) = True
        End If
    End If
End Sub

Private Sub vfgList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = 0 Then
        KeyAscii = 0
    ElseIf Col = 1 And vfgList.TextMatrix(vfgList.Row, 0) = "�Զ�·��" Then
        KeyAscii = 0
    End If
End Sub
