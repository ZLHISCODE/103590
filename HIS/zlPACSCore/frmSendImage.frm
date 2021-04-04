VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSendImage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���������б�"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   Icon            =   "frmSendImage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdVerify 
      Caption         =   "�������(&K)"
      Height          =   350
      Left            =   1560
      TabIndex        =   3
      Top             =   4200
      Width           =   1185
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFHostList 
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6735
      _cx             =   11880
      _cy             =   6800
      Appearance      =   1
      BorderStyle     =   1
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   1100
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&C)"
      Height          =   350
      Left            =   5520
      TabIndex        =   1
      Top             =   4200
      Width           =   1185
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   210
      TabIndex        =   0
      Top             =   4200
      Width           =   1185
   End
End
Attribute VB_Name = "frmSendImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public f As frmViewer

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSend_Click()
    Dim v As DicomViewer
    Dim im As DicomImage
    Dim imgs As New DicomImages
    Dim strHostNode As String, strHostAE As String
    Dim strLocalAE As String
    Dim lngHostPort As Long
    Dim lngRowNum As Long
    Dim lngResult As Long
    Dim i As Integer
    Dim j As Integer
    Dim iImageIndex As Integer
    Dim iImageCount As Integer
    
    On Error GoTo err
    
    lngResult = 1
    iImageCount = 0
    lngRowNum = Me.VSFHostList.RowSel
    If lngRowNum = 0 Then Exit Sub  'û�н������������˳�
    
    strHostAE = Me.VSFHostList.TextMatrix(lngRowNum, 2)
    strLocalAE = Me.VSFHostList.TextMatrix(lngRowNum, 3)
    strHostNode = Me.VSFHostList.TextMatrix(lngRowNum, 4)
    lngHostPort = Me.VSFHostList.TextMatrix(lngRowNum, 5)
   
    '���䱻ѡ���ͼ��
    For i = 1 To ZLShowSeriesInfos.Count
        iImageIndex = 1
        For j = 1 To ZLShowSeriesInfos(i).ImageInfos.Count
            If ZLShowSeriesInfos(i).ImageInfos(j).blnSelected = True Then
                Set im = Nothing
                '�����ж�ͼ���Ƿ��Ѿ�װ�أ�����Ѿ�װ�أ����ҵ����ͼ����ʾ���������û��װ�أ���װ�ظ�ͼ��
                If ZLShowSeriesInfos(i).ImageInfos(j).blnDisplayed = False Then
                    Call funcAddAImageA(f.Viewer(i), j)
                End If
                
                '����ͼ�������
                While f.Viewer(i).Images(iImageIndex).Tag < j And iImageIndex < f.Viewer(i).Images.Count
                    iImageIndex = iImageIndex + 1
                Wend
                
                If iImageIndex <= f.Viewer(i).Images.Count Then
                    If f.Viewer(i).Images(iImageIndex).Tag = j Then
                        Set im = f.Viewer(i).Images(iImageIndex)
                    End If
                End If
                
                If Not im Is Nothing Then
                    On Error Resume Next
                    lngResult = im.Send(strHostNode, lngHostPort, strLocalAE, strHostAE)    '����ͼ��ѡ������
                    iImageCount = iImageCount + 1
                    On Error GoTo 0
                    If lngResult <> 0 Then
                        MsgBox "����ͼ����������������״̬��", vbExclamation, gstrSysName
                        Exit Sub
                    End If
                    
                End If
            End If
        Next j
    Next i
   
    If iImageCount = 0 Then
        MsgBox "û��ѡ��ͼ���޷����͡�����ѡ��ͼ����ٷ��͡�", vbInformation, gstrSysName
    Else
        MsgBox "ͼ������ɣ������� " & iImageCount & " ��ͼ��", vbInformation, gstrSysName
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdVerify_Click()
    Dim strHostNode As String, strHostAE As String
    Dim lngHostPort As Long
    Dim strLocalAE As String
    Dim lngRowNum As Long
    Dim lngResult As Long
    Dim dgDicomGlobal As New DicomGlobal

    On Error GoTo err
    
    lngResult = 1
    lngRowNum = Me.VSFHostList.RowSel
    If lngRowNum = 0 Then Exit Sub  'û�н������������˳�
    
    strHostAE = Me.VSFHostList.TextMatrix(lngRowNum, 2)
    strLocalAE = Me.VSFHostList.TextMatrix(lngRowNum, 3)
    strHostNode = Me.VSFHostList.TextMatrix(lngRowNum, 4)
    lngHostPort = Me.VSFHostList.TextMatrix(lngRowNum, 5)
    
    On Error Resume Next
    lngResult = dgDicomGlobal.Echo(strHostNode, lngHostPort, strLocalAE, strHostAE)
    On Error GoTo 0
    If lngResult = 0 Then
        MsgBox "DICOM���ӳɹ���", vbInformation, gstrSysName
    Else
        MsgBox "DICOM���Ӳ��ɹ�����������״̬���������Ϊ��" & lngResult, vbExclamation, gstrSysName
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    If gcnOracle Is Nothing Then
        Me.cmdSend.Enabled = False
        Me.cmdVerify.Enabled = False
    Else
        strSQL = "select �豸��,�豸��,�豸AE,����AE,IP��ַ,�˿ں� from Ӱ���豸Ŀ¼ where ���� = 2 and NVL(״̬,0)=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        
        Set Me.VSFHostList.DataSource = rsTemp
        
        If rsTemp.EOF = True Then
            MsgBox "û���ҵ������������뵽��Ӱ���豸Ŀ¼��ģ�������ý���������", vbOKOnly, "��Ƭվ��ʾ"
        End If
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
