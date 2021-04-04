VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmStyle_SingleQueue 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11715
   Icon            =   "frmStyle_SingleQueue.frx":0000
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
   Begin VSFlex8Ctl.VSFlexGrid vsfCallingData 
      Height          =   2175
      Left            =   1440
      TabIndex        =   7
      Top             =   2280
      Width           =   2895
      _cx             =   5106
      _cy             =   3836
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
      ForeColor       =   0
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
   Begin VB.Image imgDoctor 
      Height          =   1215
      Left            =   6360
      Picture         =   "frmStyle_SingleQueue.frx":000C
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1035
   End
   Begin VB.Label lblDeptInfo 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   6480
      TabIndex        =   11
      Top             =   4920
      Width           =   4215
   End
   Begin VB.Label lblDoctorIntro 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   6480
      TabIndex        =   10
      Top             =   4200
      Width           =   3780
   End
   Begin VB.Label lblDoctorJob 
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
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   8400
      TabIndex        =   9
      Top             =   3480
      Width           =   240
   End
   Begin VB.Label lblDoctorName 
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
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   9000
      TabIndex        =   8
      Top             =   2280
      Width           =   240
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
      TabIndex        =   1
      Top             =   2640
      Width           =   240
   End
   Begin VB.Image imgLOGO 
      Height          =   720
      Left            =   240
      Picture         =   "frmStyle_SingleQueue.frx":13BD
      Stretch         =   -1  'True
      Top             =   60
      Width           =   840
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
      TabIndex        =   2
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
Attribute VB_Name = "frmStyle_SingleQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements ISty
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                               ��ʽ1˵��
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'����ʽ������Ҫ�ǰ���ָ����һ���ŶӶ��н��н�����ʾ
'���������Ҿ���Ķ�Ӧ��������
'�����������Ϊa������������������˷ֱ�a���Һ�b���ң�����������ʾ��ʽ�£�ֻ��ʾ���䵽a���ҵļ��
'
'
'Ŀǰ�Ŷӽкŵ��Ŷ�״̬�ֱ�ȡֵΪ��
'-1ռλ��������ʽ��ʼ�Ŷӣ�,0-�Ŷ���,1-������,2-������,3-����ͣ,4-�����,7-�Ѻ���,8-������,9-������
'�Ŷ�״̬��ת����ϵ����
'��������ݽ�����к�Ĭ�ϵ��Ŷ�״̬Ϊ-1
'���Խ���Ķ�������ִ��startqueue�����󣬿�ʼ��ʽ�Ŷ�,�Ŷ�״̬���޸�Ϊ0
'���Ըö������ݽ��к���ʱ���Ŷ�״̬���������״̬����Ϊ9
'���������ŶζԸö��в�������ʱ���Ŷ�״̬���޸�Ϊ1
'���������Ž����󣬸ö������ݵ��Ŷ�״̬���޸�Ϊ7
'�������еĲ��˽�������Һ���ҽ��ѡ���˽����������£����Ŷ�״̬���޸�Ϊ8
'����ɸò��˵Ľ����������ҽ��ѡ������ɲ���������Ŷ�״̬���޸�Ϊ4
'ֻ���ڲ���������������Ҫ��ͣ������߲��ٽ��о���ʱ���Ŷ�״̬�ű��޸�Ϊ3��ͣ����2����
'
'�ڸ���ʽ�£���Ҫ��ʾ������Ϊָ�������£��Ŷ�״̬Ϊ�����У��Ѻ��кͺ����е����ݣ����Ｔ��ǰ�����Ŷӵ����ݣ���ʾ�����ʽ����
' 004  ����  ������
'
' 003  ����  �Ѻ���
' 002  ����  �Ѻ���
'
' 005  ����  �����
' 006  ����  �����
'
'�Ѻ����ŶӼ�����ʾ�����ͺ����ŶӼ�����ʾ����Ӧ����ͨ���������ý���ȷ����
'�Ѻ�������ֻ��ȡ��󼸸������е��ŶӼ������

'��ʾ��ʽ��Ч��ͼƬ�ɲο�ͼ���ļ�����ʽ1��



'��Ҫʵ�ֵĽӿڷ������£�
'
'
'��lcd��ʾ����
'public sub ISty_Show(byval lngWindowNo as long)
'lngWindowNo:���ڱ�ţ����ݴ��ڱ�Ŷ�ȡ������Ϣ����������ʾ
'
'end sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const M_LNG_׼��������ʾ�� = 2  '�Ŷ��б��У�׼������������ʾ����

Private mlngWindowNo As Long            '���ڱ��
Private mlngRefreshInterval As Long     '��ѯʱ����
Private mlngInterval As Long            '�ۼ�ʱ����
Private mstrStyleTylePath As String     '������ʽͼƬ·��
Private mblnShowCallTarget As Boolean   '�Ƿ���ʾ����Ŀ�ĵ�
Private mstrClinicNames As String       '�ٴ��Ŷ�ҵ���µ���������

Private mLcdCommonParameter As TLcdCommonParameter

Private Type TPageObj
    tpBackImage  As TRect        '����(Ƥ��)
    tpTopAera    As TRect        '����
    tpMiddleAera As TRect        '�в�
    tpBottomAera As TRect        '�ײ�
    
    tpHospitalLOGO   As TRect    'ҽԺͼ��
    tpHospitalName   As TRect    'ҽԺ����
    tpWeek           As TRect    '����
    tpDate           As TRect    '����
    
    tpClinicName        As TRect       '��������
    tpDoctorPhotoAera   As TRect       'ҽ����Ƭ����
    tpDoctorInfo        As TRect       'ҽ����Ϣ
    tpDoctorIntro       As TRect       'ҽ�����
    tpCurQueuedList     As TRect       '׼�������б�
    
    tpblnShowListHeader As Boolean      '��ʾ�б����
    tpstrListHeaderName As String       '�б������
    tplngQueueListMaxRows As Long       '�������б������ʾ���������������ʾ�б�����������ͷ��
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
        .tpTopAera.lngLeft = Val(ReadValue("��������", "��"))
        .tpTopAera.lngTop = Val(ReadValue("��������", "��"))
        .tpTopAera.lngWidth = Val(ReadValue("��������", "��"))
        .tpTopAera.lngHeight = Val(ReadValue("��������", "��"))
        
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
        .tpMiddleAera.lngLeft = Val(ReadValue("�в�����", "��"))
        .tpMiddleAera.lngTop = Val(ReadValue("�в�����", "��"))
        .tpMiddleAera.lngWidth = Val(ReadValue("�в�����", "��"))
        .tpMiddleAera.lngHeight = Val(ReadValue("�в�����", "��"))
        
        '�ײ�����
        .tpBottomAera.lngLeft = Val(ReadValue("�ײ�����", "��"))
        .tpBottomAera.lngTop = Val(ReadValue("�ײ�����", "��"))
        .tpBottomAera.lngWidth = Val(ReadValue("�ײ�����", "��"))
        .tpBottomAera.lngHeight = Val(ReadValue("�ײ�����", "��"))
        
        '�Ŷ��б�����
        .tpCurQueuedList.lngLeft = Val(ReadValue("�Ŷ��б�����", "��"))
        .tpCurQueuedList.lngTop = Val(ReadValue("�Ŷ��б�����", "��"))
        .tpCurQueuedList.lngWidth = Val(ReadValue("�Ŷ��б�����", "��"))
        .tpCurQueuedList.lngHeight = Val(ReadValue("�Ŷ��б�����", "��"))
        
        .tpblnShowListHeader = CBool(ReadValue("�Ŷ��б�����", "�Ƿ���ʾ�б����"))
        
        If .tpblnShowListHeader Then
            .tpstrListHeaderName = Trim(ReadValue("�Ŷ��б�����", "�б������"))
            .tplngQueueListMaxRows = Val(ReadValue("�Ŷ��б�����", "������")) - 1
        Else
            .tpstrListHeaderName = ""
            .tplngQueueListMaxRows = Val(ReadValue("�Ŷ��б�����", "������"))
        End If
        
        '������������
        .tpClinicName.lngLeft = Val(ReadValue("������������", "��"))
        .tpClinicName.lngTop = Val(ReadValue("������������", "��"))
        .tpClinicName.lngWidth = Val(ReadValue("������������", "��"))
        .tpClinicName.lngHeight = Val(ReadValue("������������", "��"))
        
        '��Ƭ����
        .tpDoctorPhotoAera.lngLeft = Val(ReadValue("��Ƭ����", "��"))
        .tpDoctorPhotoAera.lngTop = Val(ReadValue("��Ƭ����", "��"))
        .tpDoctorPhotoAera.lngWidth = Val(ReadValue("��Ƭ����", "��"))
        .tpDoctorPhotoAera.lngHeight = Val(ReadValue("��Ƭ����", "��"))
        
        'ҽ����Ϣ
        .tpDoctorInfo.lngLeft = Val(ReadValue("ҽ����Ϣ", "��"))
        .tpDoctorInfo.lngTop = Val(ReadValue("ҽ����Ϣ", "��"))
        .tpDoctorInfo.lngWidth = Val(ReadValue("ҽ����Ϣ", "��"))
        .tpDoctorInfo.lngHeight = Val(ReadValue("ҽ����Ϣ", "��"))
        
        'ҽ�����ܺͿ��Ҽ�����ʾλ��
        .tpDoctorIntro.lngLeft = Val(ReadValue("�������", "��"))
        .tpDoctorIntro.lngTop = Val(ReadValue("�������", "��"))
        .tpDoctorIntro.lngWidth = Val(ReadValue("�������", "��"))
        .tpDoctorIntro.lngHeight = Val(ReadValue("�������", "��"))
    End With
End Sub

Public Sub ISty_RefreshQueueData(Optional ByVal lngQueueId As Long)
'ˢ�½�����ʾ����
    Dim blnExist���� As Boolean
    
    Call LoadCallingData(blnExist����)
    Call SetStyleFont(blnExist����)
    
    '����ˢ�º󽫼�ʱ����0
    mlngInterval = 0
End Sub

'��lcd��ʾ����
Public Sub ISty_Show(ByVal lngWindowNo As Long)
'lngWindowNo:���ڱ�ţ����ݴ��ڱ�Ŷ�ȡ������Ϣ����������ʾ
    Dim blnExist���� As Boolean     '�Ŷ��б����Ƿ����״̬���ڡ����������
    
    mlngWindowNo = lngWindowNo
    
    Call InitMonitor    '��ʼ������������
    
    If Not InitLocalPars Then Exit Sub
    
    Call LoadCallingData(blnExist����)
    
    Call SetStyleFont(blnExist����)

    Call Show
End Sub

Public Function ISty_ShowCfg(ByVal lngWindowNo As Long, objOwner As Object) As Boolean
'�򿪶�Ӧ����ʽ���ô���
    Dim objConfig As frmStyle_CommonCfg
    
    Set objConfig = New frmStyle_CommonCfg
            
    ISty_ShowCfg = objConfig.OpenShowConfig(lngWindowNo, TShowStyle.ssSingleQueue, Me)
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
'��ʼ�����ز�������
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    Dim lngCurLCDNo As Long
    Dim strBusinessType As String
    Dim strLCDLocation As String
    Dim strQueryQueueNames As String

On Error GoTo ErrorHand
    If gobjFile.FolderExists(App.Path & "\Skin\��������ʽ") Then
        mstrStyleTylePath = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "Ƥ����ʽ", App.Path & "\Skin\��������ʽ\�����п�������") & ".jpg"
    Else
        mstrStyleTylePath = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "Ƥ����ʽ", App.Path & "\zlQueueShow\Skin\��������ʽ\�����п�������") & ".jpg"
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
    strQueryQueueNames = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "��ʾ����", "")
    
    mLcdCommonParameter.blnConvertQueueName = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "ת����������", 0)) = 1
    
    If strQueryQueueNames <> "" Then
        If mLcdCommonParameter.blnConvertQueueName Then    'ת�����ϰ汾��ʽ�Ķ�������
            Select Case glngBusinessType
                Case TBusinessType.btClinical
                    mLcdCommonParameter.strQueryQueueNames = Split(Split(strQueryQueueNames, "|")(1), "_")(0)
                    mstrClinicNames = Split(strQueryQueueNames, ":")(1)
                    
                Case TBusinessType.btPacs
                    If InStr(strQueryQueueNames, "���Ҷ���") > 0 Then       '�������Ŷ�
                        mLcdCommonParameter.strQueryQueueNames = Split(Split(Split(strQueryQueueNames, "|")(1), ":")(0), "_")(1) & "-" & Split(strQueryQueueNames, "|")(0)
                    Else
                        mLcdCommonParameter.strQueryQueueNames = Split(Split(strQueryQueueNames, "|")(1), "_")(0) & ":" & Split(strQueryQueueNames, ":")(1)
                    End If
                    
                Case TBusinessType.btPeis
                    strSql = "select վ������ from ���վ��ֲ� where ִ�п���id=[1]"
                    Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡվ������", Val(Split(Split(strQueryQueueNames, "|")(1), "_")(0)))
                    
                    If rsRecord.RecordCount > 0 Then
                        mLcdCommonParameter.strQueryQueueNames = Nvl(rsRecord!վ������) & ":" & Split(strQueryQueueNames, ":")(1)
                    End If
                'case
                '
                '
            End Select
        Else                                                                '''''''''�°�������Ƹ�ʽ
            Select Case glngBusinessType
                Case TBusinessType.btClinical
                    mLcdCommonParameter.strQueryQueueNames = Split(Split(strQueryQueueNames, "|")(1), "_")(0)
                    mstrClinicNames = Split(strQueryQueueNames, ":")(1)
                    
                Case TBusinessType.btPacs
                    mLcdCommonParameter.strQueryQueueNames = Split(strQueryQueueNames, "|")(0) & "-" & Split(strQueryQueueNames, ":")(1)
                    
                Case TBusinessType.btPeis
                    strSql = "select վ������ from ���վ��ֲ� where ִ�п���id=[1]"
                    Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡվ������", Val(Split(Split(strQueryQueueNames, "|")(1), "_")(0)))
                    
                    If rsRecord.RecordCount > 0 Then
                        mLcdCommonParameter.strQueryQueueNames = Nvl(rsRecord!վ������) & ":" & Split(strQueryQueueNames, ":")(1)
                    End If
                'case
                '
                '
            End Select
        End If
    End If
    
    '��ǰ��������
    mLcdCommonParameter.strCurDiagnoseRoom = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "����ִ�м�", "")
    If mLcdCommonParameter.strCurDiagnoseRoom = "" Then
        strSql = "select d.���� from �ϻ���Ա�� A,��Ա�� B,������Ա C,���ű� D " & _
                 "where A.��ԱID=B.ID And b.id=c.��Աid and c.����id=d.id and c.ȱʡ=1 and A.�û���=[1]"
        
        Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "", UCase(gstrUserName))
        
        If rsRecord.RecordCount > 0 Then lblClinicName.Caption = Nvl(rsRecord!����)
    Else
        If InStr(strQueryQueueNames, "���Ҷ���") > 0 Then
            lblClinicName.Caption = Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(0)
        Else
            If Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "���ұ����Ƿ���ʾ������", 0)) = 1 Then
                If Trim(Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(1)) <> "" Then
                    lblClinicName.Caption = Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(0) & Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(1)
                Else
                    lblClinicName.Caption = Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(0)
                End If
            Else
                If Trim(Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(1)) <> "" Then
                    lblClinicName.Caption = Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(1)
                Else
                    lblClinicName.Caption = Split(mLcdCommonParameter.strCurDiagnoseRoom, "-")(0)
                End If
            End If
        End If
    End If
    
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
    
    '������ʾ�����ñ�񱳾�ͼ
    If mtpPageObj.tpblnShowListHeader Then
        mLcdCommonParameter.lngQueueRows = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "�Ŷ��б���ʾ��", mtpPageObj.tplngQueueListMaxRows)) + 1
        vsfCallingData.WallPaper = CutPicture(mstrStyleTylePath, picTemp, mtpPageObj.tpCurQueuedList.lngLeft * Screen.TwipsPerPixelX, mtpPageObj.tpCurQueuedList.lngTop * Screen.TwipsPerPixelY, mtpPageObj.tpCurQueuedList.lngWidth * Screen.TwipsPerPixelX, mtpPageObj.tpCurQueuedList.lngHeight * Screen.TwipsPerPixelY * mLcdCommonParameter.lngQueueRows / (mtpPageObj.tplngQueueListMaxRows + 1))
    Else
        mLcdCommonParameter.lngQueueRows = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "�Ŷ��б���ʾ��", mtpPageObj.tplngQueueListMaxRows))
        vsfCallingData.WallPaper = CutPicture(mstrStyleTylePath, picTemp, mtpPageObj.tpCurQueuedList.lngLeft * Screen.TwipsPerPixelX, mtpPageObj.tpCurQueuedList.lngTop * Screen.TwipsPerPixelY, mtpPageObj.tpCurQueuedList.lngWidth * Screen.TwipsPerPixelX, mtpPageObj.tpCurQueuedList.lngHeight * Screen.TwipsPerPixelY * mLcdCommonParameter.lngQueueRows / mtpPageObj.tplngQueueListMaxRows)
    End If
    
    mLcdCommonParameter.blnFontAutoSizeToList = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "�б���������Ӧ", True)
    mblnShowCallTarget = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "��ʾ����Ŀ�ĵ�", 0)) = 1
    
    Call LoadDoctorInfo
    
    tmrRefreshInterval.Enabled = True
    tmrRemarkInfo.Enabled = True
    
    InitLocalPars = True
Exit Function
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Function

Private Sub LoadDoctorInfo()
'���ض�Ӧִ�м��ҽ���Ϳ��������Ϣ
    Dim i As Integer
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    Dim blnNotWorkingTime As Boolean

    Dim strDoctorInfo As String     '�����ʽ��"ҽ��1��������ְλ|ҽ��2��������ְλ|��������"
    Dim strDoctorPhoto As String    '�����ʽ��"ҽ��1����Ƭ|ҽ��2����Ƭ|��������"
    Dim strIntroduction As String   '�����ʽ��"ҽ��1�ļ��|ҽ��2�ļ��|��������"
    Dim strWorkingTime As String    '�����ʽ��"ҽ��1��ֵ��ʱ��|ҽ��2��ֵ��ʱ��|��������"
    
    blnNotWorkingTime = True
    
    strDoctorInfo = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "ҽ����Ϣ")    '
    strDoctorPhoto = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "ҽ����Ƭ")    '
    strWorkingTime = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "ֵ��ʱ��")   '
    strIntroduction = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "ҽ�����")   '
    
    '���ݵ�ǰ���ڶ�ȡ��Ӧ��ҽ��������Ϣ
    For i = 0 To UBound(Split(Mid(strWorkingTime, 2), "|"))
        If Split(Mid(strWorkingTime, 2), "|")(i) = lblWeek.Caption Then
            blnNotWorkingTime = False
            
            lblDoctorName.Caption = Split(Split(Mid(strDoctorInfo, 2), "|")(i), "-")(1)
            
            If lblDoctorName.Caption <> "" Then
                strSql = "select רҵ����ְ�� from ��Ա�� where id=[1]"
                Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "", Val(Split(Split(Mid(strDoctorInfo, 2), "|")(i), "-")(0)))
                
                If rsRecord.RecordCount > 0 Then
                    lblDoctorJob.Caption = Nvl(rsRecord!רҵ����ְ��)
                End If
            End If
            
            Call LoadPictureInfo(imgDoctor, Split(Mid(strDoctorPhoto, 2), "|")(i))
            
            lblDoctorIntro.Caption = Split(Mid(strIntroduction, 2), "|")(i)
            lblDeptInfo.Visible = False
            Exit Sub
        End If
    Next
    
    '���ֻ����ҽ����û��ָ��ҽ����ֵ����Ϣ�����ȡ��½��Ա��Ϣ
    If blnNotWorkingTime Then
        strSql = "select B.����,B.ID from �ϻ���Ա�� A,��Ա�� B where A.��ԱID=B.ID And A.�û���=[1]"
        Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "", UCase(gstrUserName))

        If rsRecord.RecordCount > 0 Then
            For i = 0 To UBound(Split(Mid(strDoctorInfo, 2), "|"))
                If Split(Split(Mid(strDoctorInfo, 2), "|")(i), "-")(0) = Nvl(rsRecord!����) Then
                    lblDoctorName.Caption = Split(Split(Mid(strDoctorInfo, 2), "|")(i), "-")(1)
                    
                    If lblDoctorName.Caption <> "" Then
                        strSql = "select רҵ����ְ�� from ��Ա�� where id=[1]"
                        Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "", Split(Split(Mid(strDoctorInfo, 2), "|")(i), "-")(0))
                        
                        If rsRecord.RecordCount > 0 Then
                            lblDoctorJob.Caption = Nvl(rsRecord!רҵ����ְ��)
                        End If
                    End If
                    
                    Call LoadPictureInfo(imgDoctor, Split(Mid(strDoctorPhoto, 2), "|")(i))
            
                    lblDoctorIntro.Caption = Split(Mid(strIntroduction, 2), "|")(i)
                    lblDeptInfo.Visible = False
                    Exit Sub
                End If
            Next
        End If
    End If
    
    imgDoctor.Visible = False
    lblDoctorName.Visible = False
    lblDoctorJob.Visible = False
    lblDoctorIntro.Visible = False
    lblDoctorIntro.Caption = ""
End Sub

Private Sub SetStyleFont(ByVal blnExist���� As Boolean)
'���ý�����ؼ���������
    Dim i As Integer
    Dim strFontPropertys As String           '��ʽ:"����:����|�ֺ�:20|����:FALSE|б��:FALSE|�»���:FALSE|ǰ��ɫ:16777215"
    Dim strFontPropertys1 As String
    Dim strFontPropertys2 As String
    Dim strFontPropertys3 As String
    Dim strFontPropertys4 As String
    Dim strFontProperty() As String
    Dim strFontProperty1() As String
    Dim strFontProperty2() As String
    Dim strFontProperty3() As String
    Dim strFontProperty4() As String
    
On Error GoTo ErrorHand
    
    'ҽ����Ϣ����
    strFontPropertys = Trim(ReadValue("��������", "ҽ����Ϣ����", "����:����|�ֺ�:24|����:TRUE|б��:FALSE|�»���:FALSE|ǰ��ɫ:16777215"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblDoctorName, strFontProperty)
        Call SetControlFont(lblDoctorJob, strFontProperty)
    End If
    
    '����/ҽ���������
    strFontPropertys = Trim(ReadValue("��������", "ҽ��\���Ҽ������", "����:����|�ֺ�:15|����:TRUE|б��:FALSE|�»���:FALSE|ǰ��ɫ:16777215"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblDoctorIntro, strFontProperty)
        Call SetControlFont(lblDeptInfo, strFontProperty)
    End If
    
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

    '��������
    strFontPropertys = Trim(ReadValue("��������", "������������", "����:����|�ֺ�:26|����:TRUE|б��:FALSE|�»���:FALSE|ǰ��ɫ:16777215"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblClinicName, strFontProperty)
    End If

   '��ע����
    strFontPropertys = Trim(ReadValue("��������", "��ע��������", "����:����|�ֺ�:26|����:TRUE|б��:FALSE|�»���:FALSE|ǰ��ɫ:16777215"))
    
    If strFontPropertys <> "" Then
        strFontProperty = Split(strFontPropertys, "|")
        Call SetControlFont(lblRemarkInfo, strFontProperty)
        Call SetControlFont(lblCallContext, strFontProperty)
    End If
    
    '�����б�����
    strFontPropertys1 = Trim(ReadValue("��������", "�Ŷ��б��������", "����:����|�ֺ�:22|����:TRUE|ǰ��ɫ:4471868"))
    strFontPropertys2 = Trim(ReadValue("��������", "����״̬������", "����:����|�ֺ�:20|����:TRUE|ǰ��ɫ:260872"))
    strFontPropertys3 = Trim(ReadValue("��������", "׼������״̬������", "����:����|�ֺ�:20|����:FALSE|ǰ��ɫ:1681613"))
    strFontPropertys4 = Trim(ReadValue("��������", "����״̬������", "����:����|�ֺ�:20|����:FALSE|ǰ��ɫ:16777215"))
    
    strFontProperty1 = Split(strFontPropertys1, "|")
    strFontProperty2 = Split(strFontPropertys2, "|")
    strFontProperty3 = Split(strFontPropertys3, "|")
    strFontProperty4 = Split(strFontPropertys4, "|")
    
    SetVSFListFont vsfCallingData, 0, strFontProperty1
    
    For i = IIf(mtpPageObj.tpblnShowListHeader, 1, 0) To vsfCallingData.Rows - 1
        If InStr(vsfCallingData.TextMatrix(i, 3), "׼������") > 0 Then
            SetVSFListFont vsfCallingData, i, strFontProperty3
        ElseIf InStr(vsfCallingData.TextMatrix(i, 3), "����") > 0 Then
            SetVSFListFont vsfCallingData, i, strFontProperty2
        ElseIf InStr(vsfCallingData.TextMatrix(i, 3), "����") > 0 Then
            SetVSFListFont vsfCallingData, i, strFontProperty4
        End If
    Next
    
    If mLcdCommonParameter.blnFontAutoSizeToList And vsfCallingData.Rows > 1 And vsfCallingData.Cols > 1 Then
        vsfCallingData.Cell(flexcpFontSize, 0, 0, vsfCallingData.Rows - 1, vsfCallingData.Cols - 1) = vsfCallingData.RowHeight(0) / 42
    End If
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub InitDataList(ByVal blnExist���� As Boolean)
    Dim i As Integer
    Dim lngRow As Long
    '��ʼ���Ŷ��б�
    With vsfCallingData
        .Rows = 0   '�������
        .BackColorSel = &H838C00
        .ForeColorSel = .ForeColor
        .Cols = IIf(blnExist����, 5, 4)
        .Rows = mLcdCommonParameter.lngQueueRows
        
        '�����п�
        If blnExist���� Then
            .ColWidth(0) = .Width * 1 / 9
            .ColWidth(1) = .Width * 2 / 9
            .ColWidth(2) = .Width * 2 / 9
            .ColWidth(3) = .Width * 2 / 9
            .ColWidth(4) = .Width * 2 / 9
        Else
            .ColWidth(0) = .Width * 1 / 7
            .ColWidth(1) = .Width * 2 / 7
            .ColWidth(2) = .Width * 2 / 7
            .ColWidth(3) = .Width * 2 / 7
        End If
        
        For i = 0 To .Rows - 1
            .RowHeight(i) = .Height / .Rows    '�����и�
        Next
        
        If mtpPageObj.tpblnShowListHeader Then
            .TextMatrix(0, 0) = Trim(ReadValue("�Ŷ��б�����", "�б������", "�Ŷ�״̬��Ϣ"))
            .TextMatrix(0, 1) = .TextMatrix(0, 0)
            .TextMatrix(0, 2) = .TextMatrix(0, 0)
            .TextMatrix(0, 3) = .TextMatrix(0, 0)
            
            If blnExist���� Then .TextMatrix(0, 4) = .TextMatrix(0, 0)
        
            '�����кϲ�
            .MergeRow(0) = True
            .MergeCells = flexMergeRestrictRows
            .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        End If
        
        '����������ʾ��ʽ
        For lngRow = IIf(mtpPageObj.tpblnShowListHeader, 1, 0) To .Rows - 1
            .Cell(flexcpAlignment, lngRow, 0) = flexAlignRightCenter
            .Cell(flexcpAlignment, lngRow, 1) = flexAlignCenterCenter
            .Cell(flexcpAlignment, lngRow, 2) = flexAlignLeftCenter
            .Cell(flexcpAlignment, lngRow, 3) = flexAlignLeftCenter
            If blnExist���� Then .Cell(flexcpAlignment, lngRow, 4) = flexAlignLeftCenter
        Next
    End With
End Sub

Private Sub LoadCallingData(ByRef blnExist���� As Boolean)
'�����Ŷ��б�����
'blnExist������Ŷ��б����С����ʱ����true
    Dim i As Integer
    Dim lngRow As Long
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    Dim rsClone As ADODB.Recordset
    Dim strSortStyle As String      '����ʽ

On Error GoTo ErrorHand:
    Call InitDataList(False) '�Ŷ��б�Ĭ��Ϊ4��,û�С�����ԭ����
    
    If mLcdCommonParameter.strQueryQueueNames = "" Or glngBusinessType < 0 Then Exit Sub

    strSql = "select Zl_�ŶӽкŶ���_��ȡ����ʽ([1]) as ����ʽ from dual"
    Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡ����ʽ", glngBusinessType)
    
    If rsRecord.RecordCount > 0 Then strSortStyle = Nvl(rsRecord!����ʽ)
    
    Select Case glngBusinessType
        Case TBusinessType.btClinical
            If mstrClinicNames = "���Ҷ���" Then
                strSql = "select a.�ŶӺ���,a.��������,a.�Ŷ�״̬,a.����ʱ��,a.��ע,a.�Ŷ����,a.����,b.���� from �ŶӽкŶ��� a,���ű� b where ��������=[1] and ҵ������=[2]  and " & _
                         "�Ŷ�״̬ in (0,1,7,8,9) And �Ŷ�ʱ�� > trunc(sysdate) and a.����id=b.id " & IIf(mLcdCommonParameter.strFilter <> "", " and " & mLcdCommonParameter.strFilter, "") & " " & _
                         IIf(strSortStyle <> "", " order by " & strSortStyle, "")
        
                Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡ�Ŷ���Ϣ", mLcdCommonParameter.strQueryQueueNames, glngBusinessType)
            Else
                strSql = "select a.�ŶӺ���,a.��������,a.�Ŷ�״̬,a.����ʱ��,a.��ע,a.�Ŷ����,a.����,b.���� from �ŶӽкŶ��� a,���ű� b where ��������=[1] and (����=[2] or ���� is null) and ҵ������=[3]  and " & _
                         "�Ŷ�״̬ in (0,1,7,8,9) And �Ŷ�ʱ�� > trunc(sysdate) and a.����id=b.id " & IIf(mLcdCommonParameter.strFilter <> "", " and " & mLcdCommonParameter.strFilter, "") & " " & _
                         IIf(strSortStyle <> "", " order by " & strSortStyle, "")
        
                Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡ�Ŷ���Ϣ", mLcdCommonParameter.strQueryQueueNames, mstrClinicNames, glngBusinessType)
            End If
            
        Case TBusinessType.btPacs
            strSql = "select �ŶӺ���,��������,�Ŷ�״̬,����ʱ��,��ע,�Ŷ����,���� from �ŶӽкŶ��� where ��������=[1] and ҵ������=[2] and " & _
                     "�Ŷ�״̬ in (0,1,7,8,9) And �Ŷ�ʱ�� > trunc(sysdate) " & IIf(mLcdCommonParameter.strFilter <> "", " and " & mLcdCommonParameter.strFilter, "") & " " & _
                     IIf(strSortStyle <> "", " order by " & strSortStyle, "")
    
            Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡ�Ŷ���Ϣ", mLcdCommonParameter.strQueryQueueNames, glngBusinessType)
        
        Case TBusinessType.btPeis
            strSql = "select �ŶӺ���,��������,�Ŷ�״̬,����ʱ��,��ע,�Ŷ����,���� from �ŶӽкŶ��� where ��������=[1] and ҵ������=[2] and " & _
                     "�Ŷ�״̬ in (0,1,7,8,9) And �Ŷ�ʱ�� > trunc(sysdate) " & IIf(mLcdCommonParameter.strFilter <> "", " and " & mLcdCommonParameter.strFilter, "") & " " & _
                     IIf(strSortStyle <> "", " order by " & strSortStyle, "")
    
            Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡ�Ŷ���Ϣ", mLcdCommonParameter.strQueryQueueNames, glngBusinessType)
        'case
        '
        '
    End Select

    If rsRecord.RecordCount < 1 Then Exit Sub
    Set rsClone = rsRecord.Clone

    rsRecord.Filter = "��ע<>"""
    If rsRecord.RecordCount > 0 Then Call InitDataList(True) '������ԭ��ʱ���Ŷ��б���5��
    
    If mLcdCommonParameter.blnScrollDisplay Then        '��ȡ���ڡ������С��͡��Ѻ��С�������
        rsRecord.Filter = "�Ŷ�״̬=7"
        rsClone.Filter = "�Ŷ�״̬=1"
        
        lblCallContext.Caption = ""
        
        If rsRecord.RecordCount > 0 Then
            rsRecord.Sort = "����ʱ�� asc"
            rsRecord.MoveFirst
            
            For i = 0 To IIf(rsClone.RecordCount > 0, rsRecord.RecordCount - 1, rsRecord.RecordCount - 2)
                If glngBusinessType = TBusinessType.btClinical Then
                    lblCallContext.Caption = lblCallContext.Caption & " ��" & Format(Nvl(rsRecord!�ŶӺ���), "000") & "�� " & Nvl(rsRecord!��������) & " �� " & IIf(Nvl(rsRecord!����) = "", Nvl(rsRecord!����), Nvl(rsRecord!����)) & " ����"
                Else
                    lblCallContext.Caption = lblCallContext.Caption & " ��" & Format(Nvl(rsRecord!�ŶӺ���), "000") & "�� " & Nvl(rsRecord!��������) & " �� " & Nvl(rsRecord!����) & " ����"
                End If
                rsRecord.MoveNext
            Next
        End If
        
        lblRemarkInfo.Caption = ""
        If lblCallContext.Caption = "" Then lblCallContext.Caption = "��δ�е��ŵĻ������ĵȴ���"
    End If
        
    With vsfCallingData
        rsRecord.Filter = "�Ŷ�״̬=9"  '��ȡ���ڡ������С�������
        If rsRecord.RecordCount <= 0 Then rsRecord.Filter = "�Ŷ�״̬=1" '��ȡ���ڡ������С�������
        
        '��״̬���ڡ������С�û������ʱ����ȡ״̬���ڡ��Ѻ��С��Ľ��к���,��״̬���ڡ��Ѻ��С���������ֻ1��ʱ����ȡ������е�һ��
        If rsRecord.RecordCount <= 0 Then rsRecord.Filter = "�Ŷ�״̬=7"
        
        '��״̬���ڡ������С�û������ʱ����ȡ״̬���ڡ������С�������
        If rsRecord.RecordCount <= 0 Then rsRecord.Filter = "�Ŷ�״̬=8"
        
        blnExist���� = False
        lngRow = IIf(mtpPageObj.tpblnShowListHeader, 1, 0)
        
        If rsRecord.RecordCount > 0 Then
            rsRecord.Sort = "����ʱ�� desc"

            .TextMatrix(lngRow, 0) = " ��"
            .TextMatrix(lngRow, 1) = Format(rsRecord!�ŶӺ���, "000") & "��"
            .TextMatrix(lngRow, 2) = Nvl(rsRecord!��������)
            
            If mblnShowCallTarget Then
                If glngBusinessType = TBusinessType.btClinical Then
                    .TextMatrix(lngRow, 3) = "����(" & IIf(Nvl(rsRecord!����) = "", Nvl(rsRecord!����), Nvl(rsRecord!����)) & ")"
                Else
                    .TextMatrix(lngRow, 3) = "����(" & Nvl(rsRecord!����) & ")"
                End If
            Else
                .TextMatrix(lngRow, 3) = "����"
            End If
            
            If .Cols = 5 Then .TextMatrix(lngRow, 4) = Nvl(rsRecord!��ע)
            
            lngRow = IIf(mtpPageObj.tpblnShowListHeader, 2, 1) '�ӵ�lngRow�п�ʼ��ʾ�����׼����������
            
            blnExist���� = True
        End If
        
        If .Rows < lngRow + 1 Then Exit Sub
        rsClone.Filter = "�Ŷ�״̬=0"
        
        Do While Not rsClone.EOF
            .TextMatrix(lngRow, 0) = " ��"
            .TextMatrix(lngRow, 1) = Format(rsClone!�ŶӺ���, "000") & "��"
            .TextMatrix(lngRow, 2) = Nvl(rsClone!��������)
            
            '�����Ƿ��С����ȷ�����������
            If mtpPageObj.tpblnShowListHeader Then
                If lngRow >= IIf(blnExist����, 2, 1) And lngRow <= IIf(blnExist����, 2 + M_LNG_׼��������ʾ�� - 1, M_LNG_׼��������ʾ��) Then
                    .TextMatrix(lngRow, 3) = "׼������"
                Else
                    .TextMatrix(lngRow, 3) = "����"
                End If
            Else
                If lngRow >= IIf(blnExist����, 1, 0) And lngRow <= IIf(blnExist����, M_LNG_׼��������ʾ��, M_LNG_׼��������ʾ�� - 1) Then
                    .TextMatrix(lngRow, 3) = "׼������"
                Else
                    .TextMatrix(lngRow, 3) = "����"
                End If
            End If
            
            If .Cols = 5 Then .TextMatrix(lngRow, 4) = Nvl(rsClone!��ע)
            
            lngRow = lngRow + 1
            
            If lngRow > mLcdCommonParameter.lngQueueRows - 1 Then Exit Do
            rsClone.MoveNext
        Loop
    End With
Exit Sub
ErrorHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
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
    
    '�Ŷ��б�
    vsfCallingData.Left = dblWidhtScale * mtpPageObj.tpCurQueuedList.lngLeft
    vsfCallingData.Top = dblHeightScale * mtpPageObj.tpCurQueuedList.lngTop
    vsfCallingData.Width = dblWidhtScale * mtpPageObj.tpCurQueuedList.lngWidth
    vsfCallingData.Height = dblHeightScale * mtpPageObj.tpCurQueuedList.lngHeight

    '��������
    lblClinicName.Left = dblWidhtScale * mtpPageObj.tpClinicName.lngLeft + dblWidhtScale * mtpPageObj.tpClinicName.lngWidth / 2 - lblClinicName.Width / 2
    lblClinicName.Top = dblHeightScale * mtpPageObj.tpClinicName.lngTop + dblHeightScale * mtpPageObj.tpClinicName.lngHeight / 2 - lblClinicName.Height / 2
    
    'ҽ����Ƭ
    Call ResizeImg(imgDoctor, dblWidhtScale * mtpPageObj.tpDoctorPhotoAera.lngLeft, dblHeightScale * mtpPageObj.tpDoctorPhotoAera.lngTop, dblWidhtScale * mtpPageObj.tpDoctorPhotoAera.lngWidth, dblHeightScale * mtpPageObj.tpDoctorPhotoAera.lngHeight)

    'ҽ������
    lblDoctorName.Left = dblWidhtScale * mtpPageObj.tpDoctorInfo.lngLeft + dblWidhtScale * mtpPageObj.tpDoctorInfo.lngWidth / 2 - lblDoctorName.Width / 2
    lblDoctorName.Top = dblHeightScale * mtpPageObj.tpDoctorInfo.lngTop + dblHeightScale * mtpPageObj.tpDoctorInfo.lngHeight / 5
    
    'ҽ��ְλ
    lblDoctorJob.Left = dblWidhtScale * mtpPageObj.tpDoctorInfo.lngLeft + dblWidhtScale * mtpPageObj.tpDoctorInfo.lngWidth / 2 - lblDoctorJob.Width / 2
    lblDoctorJob.Top = dblHeightScale * mtpPageObj.tpDoctorInfo.lngTop + dblHeightScale * mtpPageObj.tpDoctorInfo.lngHeight * 3 / 5
    
    'ҽ�����
    lblDoctorIntro.Left = dblWidhtScale * mtpPageObj.tpDoctorIntro.lngLeft
    lblDoctorIntro.Top = dblHeightScale * mtpPageObj.tpDoctorIntro.lngTop + 60
    lblDoctorIntro.Width = dblWidhtScale * mtpPageObj.tpDoctorIntro.lngWidth
    lblDoctorIntro.Height = dblHeightScale * mtpPageObj.tpDoctorIntro.lngHeight
    
    '���Ҽ��
    If lblDoctorIntro.Caption = "" Then
        lblDeptInfo.Left = dblWidhtScale * mtpPageObj.tpDoctorIntro.lngLeft
        lblDeptInfo.Top = dblHeightScale * mtpPageObj.tpDoctorPhotoAera.lngTop + 60
        lblDeptInfo.Width = dblWidhtScale * mtpPageObj.tpDoctorIntro.lngWidth
        lblDeptInfo.Height = imgDoctor.Height + lblDoctorIntro.Height
        
        lblDeptInfo.Caption = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "���Ҽ��")
    End If
    
    '������Ϣ
    lblCallContext.Left = imgBack.Width
    lblCallContext.Top = dblHeightScale * mtpPageObj.tpBottomAera.lngTop + dblHeightScale * mtpPageObj.tpBottomAera.lngHeight / 2 - lblRemarkInfo.Height / 2
    
    '��ע����
    lblRemarkInfo.Left = imgBack.Width
    lblRemarkInfo.Top = lblCallContext.Top
    
    With vsfCallingData
        '�����Ŷ��б��п�
        If .Cols = 5 Then
            .ColWidth(0) = .Width * 1 / 9
            .ColWidth(1) = .Width * 2 / 9
            .ColWidth(2) = .Width * 2 / 9
            .ColWidth(3) = .Width * 2 / 9
            .ColWidth(4) = .Width * 2 / 9
        Else
            .ColWidth(0) = .Width * 1 / 7
            .ColWidth(1) = .Width * 2 / 7
            .ColWidth(2) = .Width * 2 / 7
            .ColWidth(3) = .Width * 2 / 7
        End If
        '�����Ŷ��б��и�
        For i = 0 To .Rows - 1
            .RowHeight(i) = .Height / .Rows
        Next
    End With
    
    If mLcdCommonParameter.blnFontAutoSizeToList And vsfCallingData.Rows > 1 And vsfCallingData.Cols > 1 Then
        vsfCallingData.Cell(flexcpFontSize, 0, 0, vsfCallingData.Rows - 1, vsfCallingData.Cols - 1) = vsfCallingData.RowHeight(0) / 42
    End If
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

Private Sub tmrRefreshInterval_Timer()
    Dim blnExist���� As Boolean
    
On Error GoTo ErrorHand
    mlngInterval = mlngInterval + 1

    '��Timer�ۼƵ�ʱ��С����ѯʱ��ʱ������ˢ���Ŷ�����
    If mlngInterval < mlngRefreshInterval Then Exit Sub
    '���ۼƵ�ʱ����0
    mlngInterval = 0
    
    Call LoadCallingData(blnExist����)
    
    Call SetStyleFont(blnExist����)
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHand
    If KeyAscii = vbKeyEscape Then Call CloseStyleWindow
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
