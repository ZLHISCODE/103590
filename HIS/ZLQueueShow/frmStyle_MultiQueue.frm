VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmStyle_MultiQueue 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11715
   Icon            =   "frmStyle_MultiQueue.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer tmrRemarkInfo 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8040
      Top             =   120
   End
   Begin VB.Timer tmrTime 
      Interval        =   60000
      Left            =   6240
      Top             =   120
   End
   Begin VB.Timer tmrRefreshInterval 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7200
      Top             =   120
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfCallingList 
      Height          =   1215
      Left            =   3240
      TabIndex        =   1
      Top             =   1800
      Width           =   4815
      _cx             =   8493
      _cy             =   2143
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2141904383
      ForeColor       =   12582912
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2141904383
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   0
      Rows            =   15
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
      ComboSearch     =   0
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
   Begin VSFlex8Ctl.VSFlexGrid vsfQueueList 
      Height          =   1215
      Left            =   3120
      TabIndex        =   2
      Top             =   3240
      Width           =   4815
      _cx             =   8493
      _cy             =   2143
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2141904383
      ForeColor       =   12582912
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2141904383
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   0
      Rows            =   15
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
      ComboSearch     =   0
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
   Begin VB.Image imgLOGO 
      Height          =   720
      Left            =   240
      Picture         =   "frmStyle_MultiQueue.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   840
   End
   Begin VB.Label lblClinicName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000DA3F&
      Height          =   435
      Left            =   840
      TabIndex        =   9
      Top             =   2640
      Width           =   240
   End
   Begin VB.Label lblPatientName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0002F6FC&
      Height          =   435
      Left            =   480
      TabIndex        =   8
      Top             =   2040
      Width           =   240
   End
   Begin VB.Label lblWeek 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����һ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   9600
      TabIndex        =   7
      Top             =   0
      Width           =   990
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2013��12��17 14:19"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   9120
      TabIndex        =   6
      Top             =   480
      Width           =   1890
   End
   Begin VB.Label lblHospitalName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����е�һ����ҽԺ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   1080
      TabIndex        =   5
      Top             =   120
      Width           =   4860
   End
   Begin VB.Label lblRemarkInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   360
      TabIndex        =   4
      Top             =   6480
      Width           =   240
   End
   Begin VB.Label lblCallContext 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��δ�е��ŵĻ������ĵȴ���"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6720
      TabIndex        =   3
      Top             =   6600
      Width           =   2340
   End
   Begin VB.Image imgBack 
      Height          =   7335
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmStyle_MultiQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements ISty

'��Ҫʵ�ֵĽӿڷ������£�
'
'
'��lcd��ʾ����
'public sub ISty_Show(byval lngWindowNo as long)
'lngWindowNo:���ڱ�ţ����ݴ��ڱ�Ŷ�ȡ������Ϣ����������ʾ
'
'end sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private mlngWindowNo As Long            '���ڱ��
Private mlngRefreshInterval As Long     '��ѯʱ����
Private mlngInterval As Long            '�ۼ�ʱ����
Private mstrStyleTylePath As String     '������ʽͼƬ·��
Private mstrQueryQueueNames As String
Private mstrQueueListQueryNames As String   '���Ŷ��б�����ʾ�Ķ�������
Private mstrClinicNames As String       '�ٴ��Ŷ�ҵ���µ���������

Private mLcdCommonParameter As TLcdCommonParameter

Private Type TPageObj
    tpBackImage  As TRect        '����(Ƥ��)
    tpTopArea    As TRect        '����
    tpMiddleArea As TRect        '�в�
    tpBottomArea As TRect        '�ײ�
    
    tpHospitalLOGO   As TRect    'ҽԺͼ��
    tpHospitalName   As TRect    'ҽԺ����
    tpWeek           As TRect    '����
    tpDate           As TRect    '����
    
    tpCurCallingInf As TRect     '��ǰ������Ϣ
    tpCurCalledList As TRect     '�����б�
    tpCurQueuedList As TRect     '׼�������б�
    
    tplngCallingListMaxRows As Long     '�����б������ʾ��������
    tplngQueueListMaxRows As Long       '�����׼�������б������ʾ����������������ͷ��
    tplngQueueListShowNum As Long       '׼�������б�����׼����������ʾ����
End Type

Private mtpPageObj As TPageObj

Private Sub GetSkinObj(ByVal strSkinName As String)
'��ȡ��ʽ�����ļ����Խ���ؼ�λ�ý��г�ʼ��
    
    Call SetIniFile(strSkinName)
    
    With mtpPageObj
        '����ͼ��С
        .tpBackImage.lngWidth = Val(ReadValue("Ƥ���ֱ���", "��"))
        .tpBackImage.lngHeight = Val(ReadValue("Ƥ���ֱ���", "��"))
        
        '��������
        .tpTopArea.lngLeft = Val(ReadValue("��������", "��"))
        .tpTopArea.lngTop = Val(ReadValue("��������", "��"))
        .tpTopArea.lngWidth = Val(ReadValue("��������", "��"))
        .tpTopArea.lngHeight = Val(ReadValue("��������", "��"))
        
        'ҽԺͼ��
        .tpHospitalLOGO.lngLeft = Val(ReadValue("ҽԺͼ��", "��"))
        .tpHospitalLOGO.lngTop = Val(ReadValue("ҽԺͼ��", "��"))
        .tpHospitalLOGO.lngWidth = Val(ReadValue("ҽԺͼ��", "��"))
        .tpHospitalLOGO.lngHeight = Val(ReadValue("ҽԺͼ��", "��"))
        
        'ҽԺ����
        .tpHospitalName.lngLeft = Val(ReadValue("ҽԺ����", "��"))
        .tpHospitalName.lngTop = Val(ReadValue("ҽԺ����", "��"))
        .tpHospitalName.lngWidth = Val(ReadValue("ҽԺ����", "��"))
        .tpHospitalName.lngHeight = Val(ReadValue("ҽԺ����", "��"))
        
        '����
        .tpWeek.lngLeft = Val(ReadValue("����", "��"))
        .tpWeek.lngTop = Val(ReadValue("����", "��"))
        .tpWeek.lngWidth = Val(ReadValue("����", "��"))
        .tpWeek.lngHeight = Val(ReadValue("����", "��"))
        
        '����
        .tpDate.lngLeft = Val(ReadValue("����", "��"))
        .tpDate.lngTop = Val(ReadValue("����", "��"))
        .tpDate.lngWidth = Val(ReadValue("����", "��"))
        .tpDate.lngHeight = Val(ReadValue("����", "��"))
            
        '�в�����
        .tpMiddleArea.lngLeft = Val(ReadValue("�в�����", "��"))
        .tpMiddleArea.lngTop = Val(ReadValue("�в�����", "��"))
        .tpMiddleArea.lngWidth = Val(ReadValue("�в�����", "��"))
        .tpMiddleArea.lngHeight = Val(ReadValue("�в�����", "��"))
        
        '�ײ�����
        .tpBottomArea.lngLeft = Val(ReadValue("�ײ�����", "��"))
        .tpBottomArea.lngTop = Val(ReadValue("�ײ�����", "��"))
        .tpBottomArea.lngWidth = Val(ReadValue("�ײ�����", "��"))
        .tpBottomArea.lngHeight = Val(ReadValue("�ײ�����", "��"))
        
        '������Ϣ����
        .tpCurCallingInf.lngLeft = Val(ReadValue("������Ϣ����", "��"))
        .tpCurCallingInf.lngTop = Val(ReadValue("������Ϣ����", "��"))
        .tpCurCallingInf.lngWidth = Val(ReadValue("������Ϣ����", "��"))
        .tpCurCallingInf.lngHeight = Val(ReadValue("������Ϣ����", "��"))
        
        '�����б�����
        .tpCurCalledList.lngLeft = Val(ReadValue("�����б�����", "��"))
        .tpCurCalledList.lngTop = Val(ReadValue("�����б�����", "��"))
        .tpCurCalledList.lngWidth = Val(ReadValue("�����б�����", "��"))
        .tpCurCalledList.lngHeight = Val(ReadValue("�����б�����", "��"))
        
        .tplngCallingListMaxRows = Val(ReadValue("�����б�����", "������"))
        
        '׼�������б�����
        .tpCurQueuedList.lngLeft = Val(ReadValue("׼�������б�����", "��"))
        .tpCurQueuedList.lngTop = Val(ReadValue("׼�������б�����", "��"))
        .tpCurQueuedList.lngWidth = Val(ReadValue("׼�������б�����", "��"))
        .tpCurQueuedList.lngHeight = Val(ReadValue("׼�������б�����", "��"))

        .tplngQueueListMaxRows = Val(ReadValue("׼�������б�����", "������"))
        .tplngQueueListShowNum = Val(ReadValue("׼�������б�����", "��ʾ����"))
    End With
End Sub

Public Sub ISty_RefreshQueueData(Optional ByVal lngQueueId As Long)
'ˢ������
    Call LoadListData(lngQueueId)
    Call SetStyleFont
    
    '����ˢ�º󽫼�ʱ����0
    mlngInterval = 0
End Sub

'��lcd��ʾ����
Public Sub ISty_Show(ByVal lngWindowNo As Long)
'lngWindowNo:���ڱ�ţ����ݴ��ڱ�Ŷ�ȡ������Ϣ����������ʾ
    mlngWindowNo = lngWindowNo
    
    Call InitMonitor    '��ʼ������������
    
    If Not InitLocalPars Then Exit Sub
    
    Call LoadListData
    
    Call SetStyleFont

    Call Show
End Sub


Public Function ISty_ShowCfg(ByVal lngWindowNo As Long, objOwner As Object) As Boolean
'�򿪶�Ӧ����ʽ���ô���
    Dim objConfig As frmStyle_CommonCfg
    
    Set objConfig = New frmStyle_CommonCfg
            
    ISty_ShowCfg = objConfig.OpenShowConfig(lngWindowNo, TShowStyle.ssMultiQueue, Me)
End Function


Public Function ISty_MsgProcess(ByVal lngWindowNo As Long, _
    ByVal strMsgKey As String, ByVal strXmlContext As String, rsData As ADODB.Recordset) As Boolean
'��Ϣ���մ���

    Dim strValue As String

On Error GoTo ErrorHand
    
    '�ж���Ϣ�еĶ��������Ƿ���Ҫ���д���Ķ�������
    rsData.Filter = "node_name='queue_name'"
    If rsData.RecordCount <= 0 Then
        Debug.Print "��Ϣ��Ч����⵽δ������Ч�Ķ������ƣ���ֹ��Ϣ����"
        Exit Function
    End If
    
    strValue = Nvl(rsData!node_value)
    
    If InStr(mLcdCommonParameter.strQueryQueueNames, strValue) <= 0 Then
        Debug.Print "����Ϣ�������в����ڵ�ǰҵ����Χ��������Ϣ����"
        Exit Function
    End If
    
    '���ݽ��յ�����Ϣ���д���......
    Select Case strMsgKey
        Case G_STR_MSG_QUEUE_001, G_STR_MSG_QUEUE_002, G_STR_MSG_QUEUE_003
            Call ISty_RefreshQueueData
    End Select

    Exit Function
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Function

Public Function ISty_WindNo() As Long
'��ȡ��ǰ��ʽ���ڵı��
    ISty_WindNo = mlngWindowNo
End Function


Private Function InitLocalPars() As Boolean
    Dim i As Integer
    Dim lngCurLCDNo As Long
    Dim strBusinessType As String
    Dim strLCDLocation As String

On Error GoTo ErrorHand
    If gobjFile.FolderExists(App.Path & "\Skin\�������ʽ") Then
        mstrStyleTylePath = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "Ƥ����ʽ", App.Path & "\Skin\�������ʽ\�������ʽ����") & ".jpg"
    Else
        mstrStyleTylePath = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "Ƥ����ʽ", App.Path & "\zlQueueShow\Skin\�������ʽ\�������ʽ����") & ".jpg"
    End If
    
    imgBack.Picture = LoadPicture(mstrStyleTylePath)
    
    Call GetSkinObj(Replace(mstrStyleTylePath, ".jpg", ".ini"))
    
    '��ʾ�����
    lngCurLCDNo = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "��ʾ�����", 1)) - 1
    If lngCurLCDNo < 0 Then lngCurLCDNo = 0
        
    '��ʾģʽ,0-ȫ����1-�Զ���
    If Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "��ʾģʽ", 0)) = 0 Then
        Call SetFullScreenWindow(Me, lngCurLCDNo)
    Else
        strLCDLocation = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "�Զ���λ��")
        
        If strLCDLocation <> "" Then
            mLcdCommonParameter.recPos.lngLeft = Mid(Split(strLCDLocation, "|")(0), 3) * Screen.TwipsPerPixelX
            mLcdCommonParameter.recPos.lngTop = Mid(Split(strLCDLocation, "|")(1), 3) * Screen.TwipsPerPixelY
            mLcdCommonParameter.recPos.lngWidth = Mid(Split(strLCDLocation, "|")(2), 3) * Screen.TwipsPerPixelX
            mLcdCommonParameter.recPos.lngHeight = Mid(Split(strLCDLocation, "|")(3), 3) * Screen.TwipsPerPixelY
        End If
        
        Call SetCustomWindow(Me, lngCurLCDNo, mLcdCommonParameter.recPos)
    End If
    '���ݹ�������
    mLcdCommonParameter.strFilter = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "��������", "")
    '�Ŷ��б�����ʾ�Ķ�����
    mstrQueryQueueNames = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "��ʾ����", "")
    
    mLcdCommonParameter.blnConvertQueueName = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "ת����������", 0)) = 1

    If GetQueueNames(mstrQueryQueueNames) <> "" Then
        If UBound(Split(mLcdCommonParameter.strQueryQueueNames, ",")) + 1 <= 4 Then
            vsfCallingList.Cols = 1
            
            mLcdCommonParameter.lngCallingRows = UBound(Split(mLcdCommonParameter.strQueryQueueNames, ",")) + 1
        Else
            vsfCallingList.Cols = 2
            
            If (UBound(Split(mLcdCommonParameter.strQueryQueueNames, ",")) + 1) Mod 2 = 0 Then
                mLcdCommonParameter.lngCallingRows = (UBound(Split(mLcdCommonParameter.strQueryQueueNames, ",")) + 1) \ 2
            Else
                mLcdCommonParameter.lngCallingRows = (UBound(Split(mLcdCommonParameter.strQueryQueueNames, ",")) + 1) \ 2 + 1
            End If
        End If
        
        mLcdCommonParameter.lngQueueRows = UBound(Split(mLcdCommonParameter.strQueryQueueNames, ",")) + 2
    Else
        mLcdCommonParameter.lngCallingRows = mtpPageObj.tplngCallingListMaxRows
        mLcdCommonParameter.lngQueueRows = mtpPageObj.tplngQueueListMaxRows
        vsfCallingList.Cols = 2
    End If
    
    '������ʾ�������б���ͼ
    If vsfCallingList.Cols Mod 2 = 0 Then
        vsfCallingList.WallPaper = CutPicture(mstrStyleTylePath, picTemp, mtpPageObj.tpCurCalledList.lngLeft * Screen.TwipsPerPixelX, mtpPageObj.tpCurCalledList.lngTop * Screen.TwipsPerPixelY, mtpPageObj.tpCurCalledList.lngWidth * Screen.TwipsPerPixelX, mtpPageObj.tpCurCalledList.lngHeight * Screen.TwipsPerPixelY * mLcdCommonParameter.lngCallingRows / mtpPageObj.tplngCallingListMaxRows)
    Else
        vsfCallingList.WallPaper = CutPicture(mstrStyleTylePath, picTemp, mtpPageObj.tpCurCalledList.lngLeft * Screen.TwipsPerPixelX, mtpPageObj.tpCurCalledList.lngTop * Screen.TwipsPerPixelY, (mtpPageObj.tpCurCalledList.lngWidth / 2 - 2) * Screen.TwipsPerPixelX, mtpPageObj.tpCurCalledList.lngHeight * Screen.TwipsPerPixelY * mLcdCommonParameter.lngCallingRows / mtpPageObj.tplngCallingListMaxRows)
    End If

    vsfQueueList.WallPaper = CutPicture(mstrStyleTylePath, picTemp, mtpPageObj.tpCurQueuedList.lngLeft * Screen.TwipsPerPixelX, mtpPageObj.tpCurQueuedList.lngTop * Screen.TwipsPerPixelY, mtpPageObj.tpCurQueuedList.lngWidth * Screen.TwipsPerPixelX, mtpPageObj.tpCurQueuedList.lngHeight * Screen.TwipsPerPixelY * mLcdCommonParameter.lngQueueRows / mtpPageObj.tplngQueueListMaxRows)
    
    '����ҽԺLOGO
    Call LoadPictureInfo(imgLOGO, GetSetting("ZLSOFT", G_STR_REGPATH, "ҽԺLOGO"))
    
    'ҽԺ����
    lblHospitalName.Caption = GetSetting("ZLSOFT", G_STR_REGPATH, "ҽԺ����", "�����е�һ����ҽԺ")
    '�׶��ı�
    lblRemarkInfo.Caption = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "�׶��ı�", "")

    '�Ƿ������ʾ
    mLcdCommonParameter.blnScrollDisplay = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "������ʾ", "0") = 1
    '�Ŷ��б���ѯ���
    mlngRefreshInterval = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "��ѯ���", 30))
    
    mLcdCommonParameter.blnFontAutoSizeToList = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "�б���������Ӧ", True)
    
    tmrRefreshInterval.Enabled = True
    tmrRemarkInfo.Enabled = True
    
    InitLocalPars = True
Exit Function
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Function

Private Function GetQueueNames(ByVal strQueryQueueNames As String) As String
'�����Ŷӷ�ʽ��ȡ��������
    Dim i As Integer
    Dim lngPreDeptID As Long
    Dim lngCurDeptID As Long
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    
    mLcdCommonParameter.strQueryQueueNames = ""
    lngPreDeptID = 0
    lngCurDeptID = 0

    If mLcdCommonParameter.blnConvertQueueName Then    'ת�����ϰ汾��ʽ�Ķ�������
        For i = 0 To UBound(Split(strQueryQueueNames, ","))
            lngCurDeptID = Split(Split(Split(strQueryQueueNames, ",")(i), "|")(1), "_")(0)
    
            Select Case glngBusinessType
                Case TBusinessType.btClinical   '�������ƴ洢����"������ID1,����ID2,����ID3"�� ��"63,64,65"
                    mLcdCommonParameter.strQueryQueueNames = mLcdCommonParameter.strQueryQueueNames & "," & lngCurDeptID
                    
                    If InStr(Split(mstrQueryQueueNames, ",")(i), "���Ҷ���") > 0 Then
                        mstrClinicNames = mstrClinicNames & "," & lngCurDeptID & "-"
                        mstrQueueListQueryNames = mstrQueueListQueryNames & "," & Split(Split(strQueryQueueNames, ",")(i), "|")(0)
                    Else
                        mstrClinicNames = mstrClinicNames & "," & lngCurDeptID & "-" & Split(Split(strQueryQueueNames, ",")(i), ":")(1)
                        mstrQueueListQueryNames = mstrQueueListQueryNames & "," & Split(Split(strQueryQueueNames, ",")(i), ":")(1)
                    End If
                    
                Case TBusinessType.btPacs       '�������ƴ洢����
                    If InStr(Split(strQueryQueueNames, ",")(i), "���Ҷ���") > 0 Then     '�������Ŷӣ�"����1-������1,����2-������2,����3-������3"�� ��"050202-�����,050203-CT�����"
                        mLcdCommonParameter.strQueryQueueNames = mLcdCommonParameter.strQueryQueueNames & "," & Split(Split(Split(Split(strQueryQueueNames, ",")(i), "|")(1), ":")(0), "_")(1) & "-" & Split(Split(strQueryQueueNames, ",")(i), "|")(0)
                        mstrQueueListQueryNames = mstrQueueListQueryNames & "," & Split(Split(strQueryQueueNames, ",")(i), "|")(0)
                    Else                                            '��ִ�м��Ŷ�
                        mLcdCommonParameter.strQueryQueueNames = mLcdCommonParameter.strQueryQueueNames & "," & lngCurDeptID & ":" & Split(Split(Split(strQueryQueueNames, ",")(i), "|")(1), ":")(1)
                        mstrQueueListQueryNames = mstrQueueListQueryNames & "," & Split(Split(strQueryQueueNames, ",")(i), ":")(1)
                    End If
                    
                Case TBusinessType.btPeis       '�������ƴ洢����"վ����:ִ�н�"����"վ��һ:ִ�м�1,վ���:ִ�м�1"
                    '����ִ�м�Ŀ���ID�ҵ���Ӧ��վ������
                    strSql = "select վ������ from ���վ��ֲ� where ִ�п���id=[1]"
                    Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡվ������", lngCurDeptID)
                    
                    If rsRecord.RecordCount > 0 Then
                        mLcdCommonParameter.strQueryQueueNames = mLcdCommonParameter.strQueryQueueNames & "," & _
                                                                 Nvl(rsRecord!վ������) & ":" & Split(Split(strQueryQueueNames, ",")(i), ":")(1)
                        mstrQueueListQueryNames = mstrQueueListQueryNames & "," & Split(Split(strQueryQueueNames, ",")(i), ":")(1)
                    End If
                'case
                '
                '
            End Select
            
            lngPreDeptID = lngCurDeptID
        Next

    Else            '''''''''�°�������Ƹ�ʽ
        For i = 0 To UBound(Split(mstrQueryQueueNames, ","))
            lngCurDeptID = Split(Split(Split(strQueryQueueNames, ",")(i), "|")(1), "_")(0)
            
            Select Case glngBusinessType
                Case TBusinessType.btClinical   '�������ƴ洢����"����ID1,����ID2,����ID3"�� ��"63,64,65"
                    mLcdCommonParameter.strQueryQueueNames = mLcdCommonParameter.strQueryQueueNames & "," & lngCurDeptID
                    
                    If InStr(Split(mstrQueryQueueNames, ",")(i), "���Ҷ���") > 0 Then
                        mstrClinicNames = mstrClinicNames & "," & lngCurDeptID & "-"
                        mstrQueueListQueryNames = mstrQueueListQueryNames & "," & Split(Split(strQueryQueueNames, ",")(i), "|")(0)
                    Else
                        mstrClinicNames = mstrClinicNames & "," & lngCurDeptID & "-" & Split(Split(strQueryQueueNames, ",")(i), ":")(1)
                        mstrQueueListQueryNames = mstrQueueListQueryNames & "," & Split(Split(strQueryQueueNames, ",")(i), ":")(1)
                    End If
                    
                Case TBusinessType.btPacs       '�������ƴ洢����"������-ִ�м�1,������-ִ�м�2,������-ִ�м�3"�� ��"�����-CTһ�����,�����-CT�������,�����-CT�������"
                    mLcdCommonParameter.strQueryQueueNames = mLcdCommonParameter.strQueryQueueNames & "," & _
                                                             Split(Split(strQueryQueueNames, ",")(i), "|")(0) & "-" & Split(Split(Split(strQueryQueueNames, ",")(i), "|")(1), ":")(1)
                    
                    If InStr(Split(mstrQueryQueueNames, ",")(i), "δ�������") > 0 Then
                        mstrQueueListQueryNames = mstrQueueListQueryNames & "," & Split(Split(strQueryQueueNames, ",")(i), "|")(0)
                    Else
                        mstrQueueListQueryNames = mstrQueueListQueryNames & "," & Split(Split(strQueryQueueNames, ",")(i), ":")(1)
                    End If
                Case TBusinessType.btPeis       '�������ƴ洢����"վ����:ִ�н�"����"վ��һ:ִ�м�1,վ���:ִ�м�1"
                    '����ִ�м�Ŀ���ID�ҵ���Ӧ��վ������
                    strSql = "select վ������ from ���վ��ֲ� where ִ�п���id=[1]"
                    Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡվ������", lngCurDeptID)
                    
                    If rsRecord.RecordCount > 0 Then
                        mLcdCommonParameter.strQueryQueueNames = mLcdCommonParameter.strQueryQueueNames & "," & _
                                                                 Nvl(rsRecord!վ������) & ":" & Split(Split(strQueryQueueNames, ",")(i), ":")(1)
                        mstrQueueListQueryNames = mstrQueueListQueryNames & "," & Split(Split(strQueryQueueNames, ",")(i), ":")(1)
                    End If
                'case
                '
                '
            End Select
            
            lngPreDeptID = lngCurDeptID
        Next
    End If
    mLcdCommonParameter.strQueryQueueNames = Mid(mLcdCommonParameter.strQueryQueueNames, 2)
    mstrClinicNames = Mid(mstrClinicNames, 2)
    mstrQueueListQueryNames = Mid(mstrQueueListQueryNames, 2)
    GetQueueNames = mLcdCommonParameter.strQueryQueueNames
End Function

Private Sub InitDataList()
'��ʼ�������б�
    Dim i As Integer
    
    With vsfCallingList
        .Rows = mLcdCommonParameter.lngCallingRows
        
        For i = 0 To .Rows - 1
            .RowHeight(i) = .Height / .Rows
        Next
        
        For i = 0 To .Cols - 1
            .ColWidth(i) = .Width / .Cols
        Next
        
        If .Rows > 0 And .Cols > 0 Then .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
    End With
    
    With vsfQueueList
        .Cols = 3
        .Rows = mLcdCommonParameter.lngQueueRows
        
        For i = 0 To .Rows - 1
            .RowHeight(i) = .Height / .Rows
        Next
        
        .ColWidth(0) = .Width * 2 / 10
        .ColWidth(1) = .Width * 6 / 10
        .ColWidth(2) = .Width * 2 / 10
        
        .TextMatrix(0, 0) = "  ��������"
        .TextMatrix(0, 1) = "׼�������б�"
        .TextMatrix(0, 2) = "�Ŷ�����"
        
        If .Rows > 0 And .Cols > 0 Then
            .Cell(flexcpAlignment, 0, 0) = flexAlignLeftCenter
            .Cell(flexcpAlignment, 0, 1) = flexAlignCenterCenter
        End If
        
        If .Rows > 1 And .Cols > 0 Then
            .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
            .Cell(flexcpAlignment, 0, .Cols - 1, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        End If
        
        If mstrQueryQueueNames = "" Then Exit Sub
        
        For i = 1 To .Rows - 1
            .TextMatrix(i, 0) = "  " & Split(mstrQueueListQueryNames, ",")(i - 1)
            .TextMatrix(i, 2) = "�� 0  ��"
        Next
    End With
End Sub

Private Sub LoadListData(Optional ByVal lngQueueId As Long)
'�����ŶӺ����б�����
    Dim i As Integer, j As Integer, k As Integer
    Dim lngRow As Long, lngCol As Long
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    Dim rsClone As ADODB.Recordset
    Dim strSortStyle As String      '����ʽ
    Dim strQueuePatients As String
    Dim dblHeightScale As Double, dblWidhtScale As Double
    Dim strTemp As String

On Error GoTo ErrorHand:

    Call InitDataList
    
    lblPatientName.Caption = ""
    lblClinicName.Caption = ""
    
    If mLcdCommonParameter.strQueryQueueNames = "" Or glngBusinessType < 0 Then Exit Sub
    
    strSql = "select Zl_�ŶӽкŶ���_��ȡ����ʽ([1]) as ����ʽ from dual"
    Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡ����ʽ", glngBusinessType)
    
    If rsRecord.RecordCount > 0 Then strSortStyle = Nvl(rsRecord!����ʽ)
    
    Select Case glngBusinessType
        Case TBusinessType.btClinical
            strSql = "select distinct A.ID,A.�ŶӺ���,A.��������,A.�Ŷ�״̬,A.����ʱ��,A.��ע,A.�Ŷ����,A.��������,A.����,C.���� " & _
                     "from �ŶӽкŶ��� A, " & _
                     "(select ��������,���� from " & _
                     "(select Column_Value as �������� from Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) m) C, " & _
                     "(select substr(Column_Value,1,instr(Column_Value,'-')-1) as ����,substr(Column_Value,instr(Column_Value,'-')+1) as ���� from Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) n) D " & _
                     "where C.��������=D.����) B,���ű� C " & _
                     "Where A.�������� =B.�������� and a.����id=c.id and (A.����=B.���� or A.���� is null  or B.���� is null) and ҵ������=[3] and �Ŷ�״̬ in (0,1,7,8,9) And �Ŷ�ʱ�� > trunc(sysdate)" & _
                     IIf(mLcdCommonParameter.strFilter <> "", " and " & mLcdCommonParameter.strFilter, "") & " " & _
                     IIf(strSortStyle <> "", " order by " & strSortStyle, "")
                     
            Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡ�Ŷ���Ϣ", mLcdCommonParameter.strQueryQueueNames, mstrClinicNames, glngBusinessType)
            
        Case TBusinessType.btPacs
            strSql = "select ID,�ŶӺ���,��������,�Ŷ�״̬,����ʱ��,��ע,�Ŷ����,��������,���� from �ŶӽкŶ��� A,Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) B " & _
                     "Where A.�������� =B.Column_Value and ҵ������=[2] and �Ŷ�״̬ in (0,1,5,6,7,8,9) And �Ŷ�ʱ�� > trunc(sysdate) " & _
                     IIf(mLcdCommonParameter.strFilter <> "", " and " & mLcdCommonParameter.strFilter, "") & " " & _
                     IIf(strSortStyle <> "", " order by " & strSortStyle, "")
                     
            Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡ�Ŷ���Ϣ", mLcdCommonParameter.strQueryQueueNames, glngBusinessType)
            
        Case TBusinessType.btPeis
            strSql = "select ID,�ŶӺ���,��������,�Ŷ�״̬,����ʱ��,��ע,�Ŷ����,��������,���� from �ŶӽкŶ��� A,Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) B " & _
                     "Where A.�������� =B.Column_Value and ҵ������=[2] and �Ŷ�״̬ in (0,1,7,8,9) And �Ŷ�ʱ�� > trunc(sysdate) " & _
                     IIf(mLcdCommonParameter.strFilter <> "", " and " & mLcdCommonParameter.strFilter, "") & " " & _
                     IIf(strSortStyle <> "", " order by " & strSortStyle, "")
                     
            Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡ�Ŷ���Ϣ", mLcdCommonParameter.strQueryQueueNames, glngBusinessType)
        'case
        '
        '
    End Select
    If rsRecord.RecordCount <= 0 Then Exit Sub
    Set rsClone = rsRecord.Clone

    '�������ں��еĿ��Ҽ�������Ϣ
    If gstrCompareVersion < "010.034.000" Then
        rsRecord.Filter = "id=" & lngQueueId
    Else
        rsRecord.Filter = "�Ŷ�״̬=9"
        If rsRecord.RecordCount <= 0 Then rsRecord.Filter = "�Ŷ�״̬=1"
    End If
    
    '��ǰ������Ϣ
    If rsRecord.RecordCount > 0 Then
        lblPatientName.Caption = Format(Nvl(rsRecord!�ŶӺ���), "000") & "�� " & Nvl(rsRecord!��������)
        If glngBusinessType = TBusinessType.btClinical Then
            lblClinicName.Caption = IIf(Nvl(rsRecord!����) = "", Nvl(rsRecord!����), Nvl(rsRecord!����))
        Else
            lblClinicName.Caption = Nvl(rsRecord!����)
        End If
        
        '�������õ�ǰ������Ϣ��ʾλ��
        dblHeightScale = imgBack.Height / mtpPageObj.tpBackImage.lngHeight
        dblWidhtScale = imgBack.Width / mtpPageObj.tpBackImage.lngWidth
        
        lblPatientName.Left = dblWidhtScale * mtpPageObj.tpCurCallingInf.lngLeft + (dblWidhtScale * mtpPageObj.tpCurCallingInf.lngWidth - lblPatientName.Width) / 2
        lblPatientName.Top = dblHeightScale * mtpPageObj.tpCurCallingInf.lngTop + 60
        
        lblClinicName.Left = dblWidhtScale * mtpPageObj.tpCurCallingInf.lngLeft + (dblWidhtScale * mtpPageObj.tpCurCallingInf.lngWidth - lblClinicName.Width) / 2
        lblClinicName.Top = lblPatientName.Top + 780
    End If
    
    For i = 1 To vsfQueueList.Rows - 1
        '����׼�������б�����
        rsRecord.Filter = ""
        strQueuePatients = ""
        
        If mstrClinicNames <> "" Then strTemp = Split(Split(mstrClinicNames, ",")(i - 1), "-")(1)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�Ŷ�״̬=5''''''''''''''''''''''''''''''''''''''''''''''''''''���غ�������
        If strTemp = "" Then
            rsRecord.Filter = "��������='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And �Ŷ�״̬=5"
        Else
            rsRecord.Filter = "��������='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And ����='" & strTemp & "' And �Ŷ�״̬=0"
        End If
        
        For j = 0 To rsRecord.RecordCount - 1
            strQueuePatients = strQueuePatients & ", " & Format(Nvl(rsRecord!�ŶӺ���), "000") & "-" & Nvl(rsRecord!��������) & IIf(Nvl(rsRecord!��ע) <> "", "(" & Nvl(rsRecord!��ע) & ")", "")
            If j = mtpPageObj.tplngQueueListShowNum - 1 Then Exit For 'ÿ��������ʾ׼����������
            rsRecord.MoveNext
        Next
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�Ŷ�״̬=6
        If strTemp = "" Then
            rsRecord.Filter = "��������='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And �Ŷ�״̬=6"
        Else
            rsRecord.Filter = "��������='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And ����='" & strTemp & "' And �Ŷ�״̬=0"
        End If
        
        For j = 0 To rsRecord.RecordCount - 1
            strQueuePatients = strQueuePatients & ", " & Format(Nvl(rsRecord!�ŶӺ���), "000") & "-" & Nvl(rsRecord!��������) & IIf(Nvl(rsRecord!��ע) <> "", "(" & Nvl(rsRecord!��ע) & ")", "")
            If j = mtpPageObj.tplngQueueListShowNum - 1 Then Exit For 'ÿ��������ʾ׼����������
            rsRecord.MoveNext
        Next
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�Ŷ�״̬=0
        If strTemp = "" Then
            rsRecord.Filter = "��������='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And �Ŷ�״̬=0 "
        End If

        For j = 0 To rsRecord.RecordCount - 1
            strQueuePatients = strQueuePatients & ", " & Format(Nvl(rsRecord!�ŶӺ���), "000") & "-" & Nvl(rsRecord!��������) & IIf(Nvl(rsRecord!��ע) <> "", "(" & Nvl(rsRecord!��ע) & ")", "")
            If j = mtpPageObj.tplngQueueListShowNum - 1 Then Exit For 'ÿ��������ʾ׼����������
            rsRecord.MoveNext
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''���غ����������
        vsfQueueList.TextMatrix(i, 1) = Mid(strQueuePatients, 2)
        vsfQueueList.TextMatrix(i, 2) = "�� " & strFormat(rsRecord.RecordCount) & "��"
        
        '���غ����б�����
        If strTemp = "" Then
            rsClone.Filter = "��������='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And �Ŷ�״̬=9"
            If rsClone.RecordCount <= 0 Then rsClone.Filter = "��������='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And �Ŷ�״̬=1"
            If rsClone.RecordCount <= 0 Then rsClone.Filter = "��������='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And �Ŷ�״̬=7"
            If rsClone.RecordCount <= 0 Then rsClone.Filter = "��������='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And �Ŷ�״̬=8"
        Else
            rsClone.Filter = "��������='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And ����='" & strTemp & "' And �Ŷ�״̬=9"
            If rsClone.RecordCount <= 0 Then rsClone.Filter = "��������='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And ����='" & strTemp & "' And �Ŷ�״̬=1"
            If rsClone.RecordCount <= 0 Then rsClone.Filter = "��������='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And ����='" & strTemp & "' And �Ŷ�״̬=7"
            If rsClone.RecordCount <= 0 Then rsClone.Filter = "��������='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And ����='" & strTemp & "' And �Ŷ�״̬=8"
        End If
        
        If rsClone.RecordCount > 0 Then
            rsClone.Sort = "����ʱ�� desc"
            k = k + 1
            
            If vsfCallingList.Cols = 1 Then
                If glngBusinessType = TBusinessType.btClinical Then
                    vsfCallingList.TextMatrix(k - 1, 0) = "���� " & Format(Nvl(rsClone!�ŶӺ���), "000") & " �� " & Nvl(rsClone!��������) & " �� " & IIf(Nvl(rsClone!����) = "", Nvl(rsClone!����), Nvl(rsClone!����)) & " ���� "
                Else
                    vsfCallingList.TextMatrix(k - 1, 0) = "���� " & Format(Nvl(rsClone!�ŶӺ���), "000") & " �� " & Nvl(rsClone!��������) & " �� " & Nvl(rsClone!����) & " ���� "
                End If
            Else
                If glngBusinessType = TBusinessType.btClinical Then
                    vsfCallingList.TextMatrix(lngRow, lngCol) = "���� " & Format(Nvl(rsClone!�ŶӺ���), "000") & " �� " & Nvl(rsClone!��������) & " �� " & IIf(Nvl(rsClone!����) = "", Nvl(rsClone!����), Nvl(rsClone!����)) & " ���� "
                Else
                    vsfCallingList.TextMatrix(lngRow, lngCol) = "���� " & Format(Nvl(rsClone!�ŶӺ���), "000") & " �� " & Nvl(rsClone!��������) & " �� " & Nvl(rsClone!����) & " ���� "
                End If
                lngCol = k Mod 2
                lngRow = k \ 2
            End If
        End If
    Next
    
    lblCallContext.Caption = ""
    
    '��ȡ���ڡ������С��͡��Ѻ��С�������
    If mLcdCommonParameter.blnScrollDisplay Then
        For i = 1 To vsfQueueList.Rows - 1
            If mstrClinicNames <> "" Then strTemp = Split(Split(mstrClinicNames, ",")(i - 1), "-")(1)
            
            If strTemp = "" Then
                rsRecord.Filter = "��������='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And �Ŷ�״̬=7"
                rsClone.Filter = "��������='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And �Ŷ�״̬=1"
            Else
                rsRecord.Filter = "��������='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And ����='" & strTemp & "' And �Ŷ�״̬=7"
                rsClone.Filter = "��������='" & Split(mLcdCommonParameter.strQueryQueueNames, ",")(i - 1) & "' And ����='" & strTemp & "' And �Ŷ�״̬=1"
            End If
            
            If rsRecord.RecordCount > 0 Then
                rsRecord.Sort = "����ʱ�� asc"
                rsRecord.MoveFirst
                
                For j = 0 To IIf(rsClone.RecordCount > 0, rsRecord.RecordCount - 1, rsRecord.RecordCount - 2)
                    If glngBusinessType = TBusinessType.btClinical Then
                        lblCallContext.Caption = lblCallContext.Caption & " ��" & Format(Nvl(rsRecord!�ŶӺ���), "000") & "�� " & Nvl(rsRecord!��������) & " �� " & IIf(Nvl(rsRecord!����) = "", Nvl(rsRecord!����), Nvl(rsRecord!����)) & " ����"
                    Else
                        lblCallContext.Caption = lblCallContext.Caption & " ��" & Format(Nvl(rsRecord!�ŶӺ���), "000") & "�� " & Nvl(rsRecord!��������) & " �� " & Nvl(rsRecord!����) & " ����"
                    End If
                    
                    rsRecord.MoveNext
                Next
            End If
        Next
        
        lblRemarkInfo.Caption = ""
    End If
    
    If lblCallContext.Caption = "" Then lblCallContext.Caption = "��δ�е��ŵĻ������ĵȴ���"
    
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Function strFormat(ByVal strVal As String) As String
    Dim lngLength As Long
    
    lngLength = Len(strVal)
    
    strFormat = strVal & String(3 - Len(strVal), " ")
End Function

Private Sub SetStyleFont()
'���ý�����ؼ���������
    Dim i As Integer
    Dim strFontPropertys As String           '��ʽ:"����:����|�ֺ�:20|����:FALSE|б��:FALSE|�»���:FALSE|ǰ��ɫ:16777215"
    Dim strFontPropertys1 As String
    Dim strFontPropertys2 As String
    Dim strFontPropertys3 As String
    Dim strFontProperty() As String
    Dim strFontProperty1() As String
    Dim strFontProperty2() As String
    Dim strFontProperty3() As String
    
    Dim strRegPath As String
On Error GoTo ErrorHand
    strRegPath = G_STR_REGPATH & "\�������ʽ"
    
    'ҽԺ����
    strFontPropertys = Trim(ReadValue("��������", "ҽԺ��������", "����:����|�ֺ�:26|����:TRUE|б��:FALSE|�»���:FALSE|ǰ��ɫ:16777215"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblHospitalName, strFontProperty)
    End If
    
    '����
    strFontPropertys = Trim(ReadValue("��������", "��������", "����:����|�ֺ�:20|����:TRUE|б��:FALSE|�»���:FALSE|ǰ��ɫ:16777215"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblWeek, strFontProperty)
    End If

    '����
    strFontPropertys = Trim(ReadValue("��������", "��������", "����:����|�ֺ�:15|����:FALSE|б��:FALSE|�»���:FALSE|ǰ��ɫ:16777215"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblDate, strFontProperty)
    End If
    
    '������Ϣ
    strFontPropertys = Trim(ReadValue("��������", "������Ϣ����", "����:����|�ֺ�:26|����:TRUE|б��:FALSE|�»���:FALSE|ǰ��ɫ:194300"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblPatientName, strFontProperty)
        Call SetControlFont(lblClinicName, strFontProperty)
    End If
    
    '��ע����
    strFontPropertys = Trim(ReadValue("��������", "��ע��������", "����:����|�ֺ�:26|����:TRUE|б��:FALSE|�»���:FALSE|ǰ��ɫ:16777215"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblRemarkInfo, strFontProperty)
        Call SetControlFont(lblCallContext, strFontProperty)
    End If
    
    '�б�����
    strFontPropertys1 = Trim(ReadValue("��������", "����״̬������", "����:����|�ֺ�:18|����:TRUE|ǰ��ɫ:55871"))
    strFontPropertys2 = Trim(ReadValue("��������", "�Ŷ��б��������", "����:����|�ֺ�:20|����:TRUE|ǰ��ɫ:14721613"))
    strFontPropertys3 = Trim(ReadValue("��������", "׼������״̬������", "����:����|�ֺ�:18|����:FALSE|ǰ��ɫ:16777215"))
    
    strFontProperty1 = Split(strFontPropertys1, "|")
    strFontProperty2 = Split(strFontPropertys2, "|")
    strFontProperty3 = Split(strFontPropertys3, "|")
    
    For i = 0 To vsfCallingList.Rows - 1
        SetVSFListFont vsfCallingList, i, strFontProperty1
    Next
    
    SetVSFListFont vsfQueueList, 0, strFontProperty2
    
    For i = 1 To vsfQueueList.Rows - 1
        SetVSFListFont vsfQueueList, i, strFontProperty3
    Next
    
    If mLcdCommonParameter.blnFontAutoSizeToList Then
        If vsfCallingList.Rows >= 1 And vsfCallingList.Cols >= 1 Then
            If vsfCallingList.Cols = 1 Then
                vsfCallingList.Cell(flexcpFontSize, 0, 0, vsfCallingList.Rows - 1, vsfCallingList.Cols - 1) = vsfCallingList.RowHeight(0) / 34
            Else
                vsfCallingList.Cell(flexcpFontSize, 0, 0, vsfCallingList.Rows - 1, vsfCallingList.Cols - 1) = vsfCallingList.RowHeight(0) / 46
            End If
        End If
        
        If vsfQueueList.Rows >= 1 And vsfQueueList.Cols > 1 Then
            vsfQueueList.Cell(flexcpFontSize, 0, 0, vsfQueueList.Rows - 1, vsfQueueList.Cols - 1) = vsfQueueList.RowHeight(0) / 42
        End If
    End If
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHand
    If KeyAscii = vbKeyEscape Then Call CloseStyleWindow
Exit Sub
ErrorHand:
    Debug.Print Err.Description
    Err.Clear
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHand
    mlngInterval = 0
    tmrRefreshInterval.Interval = 1000
    
    Call refreshWeekLab
    
    lblDate.Caption = Format(Now, "yyyy��mm��dd�� hh:mm")
Exit Sub
ErrorHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    Dim i As Integer
    Dim dblHeightScale As Double, dblWidhtScale As Double
    
    '���屳��
    imgBack.Left = 0
    imgBack.Top = 0
    imgBack.Height = Me.ScaleHeight
    imgBack.Width = Me.ScaleWidth
    
    dblHeightScale = imgBack.Height / mtpPageObj.tpBackImage.lngHeight
    dblWidhtScale = imgBack.Width / mtpPageObj.tpBackImage.lngWidth
    
    'ҽԺͼ��
    Call ResizeImg(imgLOGO, dblWidhtScale * mtpPageObj.tpHospitalLOGO.lngLeft, dblHeightScale * mtpPageObj.tpHospitalLOGO.lngTop, dblWidhtScale * mtpPageObj.tpHospitalLOGO.lngWidth, dblHeightScale * mtpPageObj.tpHospitalLOGO.lngHeight)

    'ҽԺ����
    lblHospitalName.Left = dblWidhtScale * mtpPageObj.tpHospitalName.lngLeft
    lblHospitalName.Top = dblHeightScale * mtpPageObj.tpHospitalName.lngTop + dblHeightScale * mtpPageObj.tpHospitalName.lngHeight / 2 - lblHospitalName.Height / 2
    
    '����
    lblDate.Left = dblWidhtScale * mtpPageObj.tpDate.lngLeft + dblWidhtScale * mtpPageObj.tpDate.lngWidth / 2 - lblDate.Width / 2
    lblDate.Top = dblHeightScale * mtpPageObj.tpDate.lngTop + dblHeightScale * mtpPageObj.tpDate.lngHeight / 2 - lblDate.Height / 2
    
    '����
    lblWeek.Left = dblWidhtScale * mtpPageObj.tpWeek.lngLeft + dblWidhtScale * mtpPageObj.tpWeek.lngWidth / 2 - lblWeek.Width / 2
    lblWeek.Top = dblHeightScale * mtpPageObj.tpWeek.lngTop + dblHeightScale * mtpPageObj.tpWeek.lngHeight / 2 - lblWeek.Height / 2
    
    '��ǰ������Ϣ
    lblPatientName.Left = dblWidhtScale * mtpPageObj.tpCurCallingInf.lngLeft + (dblWidhtScale * mtpPageObj.tpCurCallingInf.lngWidth - lblPatientName.Width) / 2
    lblPatientName.Top = dblHeightScale * mtpPageObj.tpCurCallingInf.lngTop + 60
    
    lblClinicName.Left = dblWidhtScale * mtpPageObj.tpCurCallingInf.lngLeft + (dblWidhtScale * mtpPageObj.tpCurCallingInf.lngWidth - lblClinicName.Width) / 2
    lblClinicName.Top = lblPatientName.Top + 780

    '�����б�
    vsfCallingList.Left = dblWidhtScale * mtpPageObj.tpCurCalledList.lngLeft
    vsfCallingList.Top = dblHeightScale * mtpPageObj.tpCurCalledList.lngTop
    vsfCallingList.Height = dblHeightScale * mtpPageObj.tpCurCalledList.lngHeight
    vsfCallingList.Width = dblWidhtScale * mtpPageObj.tpCurCalledList.lngWidth
    
    '׼�������б�
    vsfQueueList.Left = dblWidhtScale * mtpPageObj.tpCurQueuedList.lngLeft
    vsfQueueList.Top = dblHeightScale * mtpPageObj.tpCurQueuedList.lngTop
    vsfQueueList.Height = dblHeightScale * mtpPageObj.tpCurQueuedList.lngHeight
    vsfQueueList.Width = dblWidhtScale * mtpPageObj.tpCurQueuedList.lngWidth
    
    '������Ϣ
    lblCallContext.Left = imgBack.Width
    lblCallContext.Top = dblHeightScale * mtpPageObj.tpBottomArea.lngTop + dblHeightScale * mtpPageObj.tpBottomArea.lngHeight / 2 - lblRemarkInfo.Height / 2
    
    '��ע����
    lblRemarkInfo.Left = imgBack.Width
    lblRemarkInfo.Top = lblCallContext.Top
    
    For i = 0 To vsfCallingList.Rows - 1
        vsfCallingList.RowHeight(i) = vsfCallingList.Height / vsfCallingList.Rows
    Next
        
    For i = 0 To vsfQueueList.Rows - 1
        vsfQueueList.RowHeight(i) = vsfQueueList.Height / vsfQueueList.Rows
    Next
    
    If mLcdCommonParameter.blnFontAutoSizeToList Then
        If vsfCallingList.Rows >= 1 And vsfCallingList.Cols >= 1 Then
            If vsfCallingList.Cols = 1 Then
                vsfCallingList.Cell(flexcpFontSize, 0, 0, vsfCallingList.Rows - 1, vsfCallingList.Cols - 1) = vsfCallingList.RowHeight(0) / 34
            Else
                vsfCallingList.Cell(flexcpFontSize, 0, 0, vsfCallingList.Rows - 1, vsfCallingList.Cols - 1) = vsfCallingList.RowHeight(0) / 46
            End If
        End If
        
        If vsfQueueList.Rows >= 1 And vsfQueueList.Cols > 1 Then
            vsfQueueList.Cell(flexcpFontSize, 0, 0, vsfQueueList.Rows - 1, vsfQueueList.Cols - 1) = vsfQueueList.RowHeight(0) / 42
        End If
    End If
End Sub


Private Sub tmrRefreshInterval_Timer()
On Error GoTo ErrorHand
    mlngInterval = mlngInterval + 1

    '��Timer�ۼƵ�ʱ��С����ѯʱ��ʱ������ˢ���Ŷ�����
    If mlngInterval < mlngRefreshInterval Then Exit Sub
    '���ۼƵ�ʱ����0
    mlngInterval = 0
    
    Call LoadListData
    
    Call SetStyleFont
Exit Sub
ErrorHand:
    Debug.Print Err.Description
    Err.Clear
End Sub

Private Sub tmrRemarkInfo_Timer()
On Error GoTo ErrorHand
    lblRemarkInfo.Left = lblRemarkInfo.Left - 15
    
    If lblRemarkInfo.Left <= -lblRemarkInfo.Width Or lblRemarkInfo.Caption = "" Then
        If lblCallContext.Caption <> "" Then    '������ʾ���ڡ������С��͡��Ѻ��С�������
            lblCallContext.Left = lblCallContext.Left - 15
            
            If lblCallContext.Left <= -lblCallContext.Width Then
                lblCallContext.Left = imgBack.Width
                lblRemarkInfo.Left = imgBack.Width
            End If
        Else
            lblRemarkInfo.Left = imgBack.Width
        End If
    End If
Exit Sub
ErrorHand:
    Debug.Print Err.Description
    Err.Clear
End Sub

Private Sub tmrTime_Timer()
On Error GoTo ErrorHand
    Call refreshWeekLab
    lblDate.Caption = Format(Now, "yyyy��mm��dd�� hh:mm")
Exit Sub
ErrorHand:
    Debug.Print Err.Description
    Err.Clear
End Sub

Private Sub refreshWeekLab()
    Select Case Weekday(Date)
        Case 1
            lblWeek.Caption = "������"
        Case 2
            lblWeek.Caption = "����һ"
        Case 3
            lblWeek.Caption = "���ڶ�"
        Case 4
            lblWeek.Caption = "������"
        Case 5
            lblWeek.Caption = "������"
        Case 6
            lblWeek.Caption = "������"
        Case 7
            lblWeek.Caption = "������"
    End Select
End Sub
