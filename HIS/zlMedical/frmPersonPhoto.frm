VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPersonPhoto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ƭ�ɼ�"
   ClientHeight    =   5355
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8625
   Icon            =   "frmPersonPhoto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "��Ƭ"
      Height          =   2355
      Left            =   5025
      TabIndex        =   2
      Top             =   45
      Width           =   3555
      Begin VB.CommandButton cmdLoad 
         Caption         =   "�������(&L)"
         Height          =   350
         Left            =   2025
         TabIndex        =   11
         Top             =   1830
         Width           =   1365
      End
      Begin VB.CommandButton cmdFile 
         Caption         =   "�ļ�����(&F)"
         Height          =   350
         Left            =   2025
         TabIndex        =   10
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "�����Ƭ(&C)"
         Height          =   350
         Left            =   2025
         TabIndex        =   9
         Top             =   720
         Width           =   1365
      End
      Begin VB.PictureBox picPhoto 
         AutoRedraw      =   -1  'True
         Height          =   1984
         Left            =   165
         ScaleHeight     =   1920
         ScaleWidth      =   1350
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   1417
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "����"
      Height          =   2760
      Left            =   5010
      TabIndex        =   1
      Top             =   2550
      Width           =   3570
      Begin VB.CommandButton cmdSource 
         Caption         =   "��Դ����(&S)"
         Height          =   350
         Left            =   2055
         TabIndex        =   8
         Top             =   1065
         Width           =   1365
      End
      Begin VB.CommandButton cmdStyle 
         Caption         =   "��ʽ����(&A)"
         Height          =   350
         Left            =   2055
         TabIndex        =   7
         Top             =   645
         Width           =   1365
      End
      Begin VB.ComboBox cboDev 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   195
         Width           =   2220
      End
      Begin VB.PictureBox picFilm 
         Height          =   1984
         Left            =   210
         ScaleHeight     =   1920
         ScaleWidth      =   1350
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   645
         Width           =   1417
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "�ɼ��豸(&D)"
         Height          =   180
         Left            =   195
         TabIndex        =   6
         Top             =   255
         Width           =   990
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   5265
      Left            =   15
      TabIndex        =   0
      Top             =   60
      Width           =   4935
      _cx             =   8705
      _cy             =   9287
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
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   12698049
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
      Begin VB.Line lnX 
         Index           =   0
         Visible         =   0   'False
         X1              =   -555
         X2              =   1230
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Line lnY 
         Index           =   0
         Visible         =   0   'False
         X1              =   270
         X2              =   270
         Y1              =   435
         Y2              =   1650
      End
   End
   Begin MSComctlLib.ImageList ils13 
      Left            =   5325
      Top             =   5775
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonPhoto.frx":000C
            Key             =   "����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonPhoto.frx":03A6
            Key             =   "����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonPhoto.frx":063C
            Key             =   "״̬"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonPhoto.frx":09D6
            Key             =   "סԺ"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonPhoto.frx":0D70
            Key             =   "����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonPhoto.frx":110A
            Key             =   "����"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonPhoto.frx":14A4
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonPhoto.frx":173A
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonPhoto.frx":19D0
            Key             =   "GChecked"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonPhoto.frx":1C66
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonPhoto.frx":1EFC
            Key             =   "Checked"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   6300
      Top             =   5715
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPersonPhoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                          '����������־
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long
Private mblnChanged As Boolean
Private mblnStarted As Boolean
Private mlng����id As Long

Private mbytPopMenu As Byte
'--------------------------------------------------------
'��  �ܣ�������Ƶ�����豸��
'�����ˣ�����
'�������ڣ�2005.11.8
'���̺����嵥��
'       mConnCapDevice() �����豸����
'       mGetCapSureDevice()
'       mParentWindowResize
'�޸ļ�¼��
'
'-------------------------------------------------------
Private Const WM_USER As Long = &H400
Private Const WM_CAP_START As Long = WM_USER

Private Const WM_CAP_GET_CAPSTREAMPTR As Long = WM_CAP_START + 1

Private Const WM_CAP_SET_CALLBACK_ERROR As Long = WM_CAP_START + 2
Private Const WM_CAP_SET_CALLBACK_STATUS As Long = WM_CAP_START + 3
Private Const WM_CAP_SET_CALLBACK_YIELD As Long = WM_CAP_START + 4
Private Const WM_CAP_SET_CALLBACK_FRAME As Long = WM_CAP_START + 5
Private Const WM_CAP_SET_CALLBACK_VIDEOSTREAM As Long = WM_CAP_START + 6
Private Const WM_CAP_SET_CALLBACK_WAVESTREAM As Long = WM_CAP_START + 7
Private Const WM_CAP_GET_USER_DATA As Long = WM_CAP_START + 8
Private Const WM_CAP_SET_USER_DATA As Long = WM_CAP_START + 9
    
Private Const WM_CAP_DRIVER_CONNECT As Long = WM_CAP_START + 10
Private Const WM_CAP_DRIVER_DISCONNECT As Long = WM_CAP_START + 11
Private Const WM_CAP_DRIVER_GET_NAME As Long = WM_CAP_START + 12
Private Const WM_CAP_DRIVER_GET_VERSION As Long = WM_CAP_START + 13
Private Const WM_CAP_DRIVER_GET_CAPS As Long = WM_CAP_START + 14

Private Const WM_CAP_FILE_SET_CAPTURE_FILE As Long = WM_CAP_START + 20
Private Const WM_CAP_FILE_GET_CAPTURE_FILE As Long = WM_CAP_START + 21
Private Const WM_CAP_FILE_ALLOCATE As Long = WM_CAP_START + 22
Private Const WM_CAP_FILE_SAVEAS As Long = WM_CAP_START + 23
Private Const WM_CAP_FILE_SET_INFOCHUNK As Long = WM_CAP_START + 24
Private Const WM_CAP_FILE_SAVEDIB As Long = WM_CAP_START + 25

Private Const WM_CAP_EDIT_COPY As Long = WM_CAP_START + 30

Private Const WM_CAP_SET_AUDIOFORMAT As Long = WM_CAP_START + 35
Private Const WM_CAP_GET_AUDIOFORMAT As Long = WM_CAP_START + 36

Private Const WM_CAP_DLG_VIDEOFORMAT As Long = WM_CAP_START + 41
Private Const WM_CAP_DLG_VIDEOSOURCE As Long = WM_CAP_START + 42
Private Const WM_CAP_DLG_VIDEODISPLAY As Long = WM_CAP_START + 43
Private Const WM_CAP_GET_VIDEOFORMAT As Long = WM_CAP_START + 44
Private Const WM_CAP_SET_VIDEOFORMAT As Long = WM_CAP_START + 45
Private Const WM_CAP_DLG_VIDEOCOMPRESSION As Long = WM_CAP_START + 46

Private Const WM_CAP_SET_PREVIEW As Long = WM_CAP_START + 50
Private Const WM_CAP_SET_OVERLAY As Long = WM_CAP_START + 51
Private Const WM_CAP_SET_PREVIEWRATE As Long = WM_CAP_START + 52
Private Const WM_CAP_SET_SCALE As Long = WM_CAP_START + 53
Private Const WM_CAP_GET_STATUS As Long = WM_CAP_START + 54
Private Const WM_CAP_SET_SCROLL As Long = WM_CAP_START + 55

Private Const WM_CAP_GRAB_FRAME As Long = WM_CAP_START + 60
Private Const WM_CAP_GRAB_FRAME_NOSTOP As Long = WM_CAP_START + 61

Private Const WM_CAP_SEQUENCE As Long = WM_CAP_START + 62
Private Const WM_CAP_SEQUENCE_NOFILE As Long = WM_CAP_START + 63
Private Const WM_CAP_SET_SEQUENCE_SETUP As Long = WM_CAP_START + 64
Private Const WM_CAP_GET_SEQUENCE_SETUP As Long = WM_CAP_START + 65
Private Const WM_CAP_SET_MCI_DEVICE As Long = WM_CAP_START + 66
Private Const WM_CAP_GET_MCI_DEVICE As Long = WM_CAP_START + 67
Private Const WM_CAP_STOP As Long = WM_CAP_START + 68
Private Const WM_CAP_ABORT As Long = WM_CAP_START + 69

Private Const WM_CAP_SINGLE_FRAME_OPEN As Long = WM_CAP_START + 70
Private Const WM_CAP_SINGLE_FRAME_CLOSE As Long = WM_CAP_START + 71
Private Const WM_CAP_SINGLE_FRAME As Long = WM_CAP_START + 72

Private Const WM_CAP_PAL_OPEN As Long = WM_CAP_START + 80
Private Const WM_CAP_PAL_SAVE As Long = WM_CAP_START + 81
Private Const WM_CAP_PAL_PASTE As Long = WM_CAP_START + 82
Private Const WM_CAP_PAL_AUTOCREATE As Long = WM_CAP_START + 83
Private Const WM_CAP_PAL_MANUALCREATE As Long = WM_CAP_START + 84

Private Const WM_CAP_SET_CALLBACK_CAPCONTROL As Long = WM_CAP_START + 85

Private Const WS_CHILD As Long = &H40000000
Private Const WS_VISIBLE As Long = &H10000000
Private Const SWP_NOSIZE As Long = &H1&
Private Const SWP_NOMOVE As Long = &H2&
Private Const SWP_NOZORDER As Long = &H4&
Private Const SWP_NOSENDCHANGING As Long = &H400&
Private Const HWND_BOTTOM As Long = 1&

Private hCapWnd As Long                          '�ɼ�������

Private Type VFWPOINT
        x As Long
        y As Long
End Type

Private Type CAPSTATUS
    uiImageWidth As Long
    uiImageHeight As Long
    fLiveWindow As Long
    fOverlayWindow As Long
    fScale As Long
    ptScroll As VFWPOINT
    fUsingDefaultPalette As Long
    fAudioHardware As Long
    fCapFileExists As Long
    dwCurrentVideoFrame As Long
    dwCurrentVideoFramesDropped As Long
    dwCurrentWaveSamples As Long
    dwCurrentTimeElapsedMS As Long
    hPalCurrent As Long
    fCapturingNow As Long
    dwReturn As Long
    wNumVideoAllocated As Long
    wNumAudioAllocated As Long
End Type



'�õ��ɼ������б�
Private Declare Function capGetDriverDescription Lib "avicap32.dll" Alias "capGetDriverDescriptionA" _
                                        (ByVal dwDriverIndex As Long, _
                                        ByVal lpszName As String, _
                                        ByVal cbName As Long, _
                                        ByVal lpszVer As String, _
                                        ByVal cbVer As Long) As Long
'�����ɼ�����
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" _
                                        (ByVal lpszWindowName As String, _
                                        ByVal dwStyle As Long, _
                                        ByVal x As Long, _
                                        ByVal y As Long, _
                                        ByVal nWidth As Long, _
                                        ByVal nHeight As Long, _
                                        ByVal hwndParent As Long, _
                                        ByVal nID As Long) As Long
'��Ϣ����
Private Declare Function SendMessageAsLong Lib "user32" Alias "SendMessageA" _
                                            (ByVal hWnd As Long, _
                                            ByVal wMsg As Long, _
                                            ByVal wParam As Long, _
                                            ByVal lParam As Long) As Long
Private Declare Function SendMessageAsAny Lib "user32" Alias "SendMessageA" _
                                            (ByVal hWnd As Long, _
                                            ByVal wMsg As Long, _
                                            ByVal wParam As Long, _
                                            ByRef lParam As Any) As Long
Private Declare Function SendMessageAsString Lib "user32" Alias "SendMessageA" _
                                            (ByVal hWnd As Long, _
                                            ByVal wMsg As Long, _
                                            ByVal wParam As Long, _
                                            ByVal lParam As String) As Long
                                            
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long 'C BOOL


Private Function GetCapSureDevice() As String
    '---------------------------------------------------------------------
    '���ܣ���ȡ��Ƶ�豸�嵥
    '������
    '���أ��豸�嵥��";"�ֿ�
    '�ϼ���������̣�
    '�¼���������̣�capGetDriverDescription
    '���õ��ⲿ������
    '�����ˣ�����
    '�޸��ˣ�
    '---------------------------------------------------------------------
    '��ȡ�����б�
    Const MAXVIDDRIVERS As Long = 9
    Const CAP_STRING_MAX As Long = 128
    
    Dim Index As Long
    Dim Device As String
    Dim Version As String
    Dim strTmp As String
    
    Device = String$(CAP_STRING_MAX, 0)
    Version = String$(CAP_STRING_MAX, 0)
    For Index = 0 To 8
        If 0 <> capGetDriverDescription(Index, Device, CAP_STRING_MAX, Version, CAP_STRING_MAX) Then
             strTmp = Left(Device, InStr(Device, vbNullChar) - 1) & Left$(Version, InStr(Version, vbNullChar) - 1)
             If Len(Trim(GetCapSureDevice)) > 0 Then
                GetCapSureDevice = GetCapSureDevice & ";"
             End If
             GetCapSureDevice = GetCapSureDevice & strTmp
        End If
    Next
End Function


Private Function ConnCapDevice(ParentWindowWnd As Long, CapDeviceIndex As Integer) As Boolean
    '-----------------------------------------------------------------------------------------
    '���ܣ����ӵ��豸
    '������ParentWindowWnd �������� ; CapDeviceIndex �豸������
    '���أ�True = �ɹ� False = ʧ��
    '�ϼ���������̣�
    '�¼���������̣�capCreateCaptureWindow;SendMessageAsLong;SendMessageAsAny;SetWindowPos
    '���õ��ⲿ������hCapWnd
    '�����ˣ�����
    '�޸��ˣ�
    '-----------------------------------------------------------------------------------------
    Dim retVal As Boolean
    Dim capStat As CAPSTATUS
    Dim strTmp() As String
    Dim i  As Integer
    
    hCapWnd = capCreateCaptureWindow("ZLSOFT_CAPTURE", WS_CHILD Or WS_VISIBLE, 0, 0, 5, 5, ParentWindowWnd, 0)
    
    If hCapWnd = 0 Then
        MsgBox "�����ɼ�����ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If

    retVal = SendMessageAsLong(hCapWnd, WM_CAP_DRIVER_CONNECT, CapDeviceIndex, 0&)
    
    If retVal = False Then
        MsgBox "�����豸ʧ�ܣ�", vbInformation, gstrSysName
        DestroyWindow hCapWnd
        Exit Function
    End If
    
    Call SendMessageAsLong(hCapWnd, WM_CAP_SET_PREVIEWRATE, 66, 0&)
    Call SendMessageAsLong(hCapWnd, WM_CAP_SET_PREVIEW, -(True), 0&)
    
    retVal = SendMessageAsAny(hCapWnd, WM_CAP_GET_STATUS, Len(capStat), capStat)

    Call SetWindowPos(hCapWnd, _
                0&, _
                0&, _
                0&, _
                capStat.uiImageWidth, _
                capStat.uiImageHeight, _
                SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSENDCHANGING)
                
    ConnCapDevice = True

End Function

Private Function SelectCapDevice(CapDeviceIndex As Integer) As Boolean
    '---------------------------------------------------------------------
    '���ܣ����ӵ�ָ���豸
    '������CapDeviceIndex �豸����(0--8)
    '���أ�True = �ɹ� False = ʧ��
    '�ϼ���������̣�
    '�¼���������̣�SendMessageAsLong;SetWindowPos
    '���õ��ⲿ������hCapWnd
    '�����ˣ�����
    '�޸��ˣ�
    '---------------------------------------------------------------------
    SelectCapDevice = SendMessageAsLong(hCapWnd, WM_CAP_DRIVER_CONNECT, CapDeviceIndex, 0&)
    
End Function

Private Function ParentWindowResize(ParentWindowWidth As Long, ParentWindowHeight As Long) As Boolean
    '---------------------------------------------------------------------
    '���ܣ�������ʾ���ڵ�λ���ڸ���������
    '������ParentWindowWidth �������� ParentWindowHeight ������߶�
    '���أ�True = �ɹ� False = ʧ��
    '�ϼ���������̣�
    '�¼���������̣�SendMessageAsAny
    '���õ��ⲿ������hCapWnd
    '�����ˣ�����
    '�޸��ˣ�
    '---------------------------------------------------------------------
    Dim retVal As Boolean
    Dim capStat As CAPSTATUS
    Dim lngWidth As Long
    Dim lngHeight As Long
    
    retVal = SendMessageAsAny(hCapWnd, WM_CAP_GET_STATUS, Len(capStat), capStat)
    
    If retVal Then
        If ParentWindowWidth - capStat.uiImageWidth <= 0 Then
            lngWidth = ParentWindowWidth
        Else
            lngWidth = (ParentWindowWidth - capStat.uiImageWidth) / 2
        End If
        If ParentWindowHeight - capStat.uiImageHeight <= 0 Then
            lngHeight = ParentWindowHeight
        Else
            lngHeight = (ParentWindowHeight - capStat.uiImageHeight) / 2
        End If
        Call SetWindowPos(hCapWnd, 0&, lngWidth, lngHeight, 0&, 0&, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOSENDCHANGING)
    End If
    
    ParentWindowResize = True
    
End Function

Private Function SaveImageFile(SavePath As String) As Boolean
    '---------------------------------------------------------------------
    '���ܣ����浱ǰ��ʾ��ͼ��
    '������SavePath=����·��
    '���أ�True = �ɹ� False = ʧ��
    '�ϼ���������̣�
    '�¼���������̣�SendMessageAsString
    '���õ��ⲿ������hCapWnd
    '�����ˣ�����
    '�޸��ˣ�
    '---------------------------------------------------------------------
    SaveImageFile = SendMessageAsString(hCapWnd, WM_CAP_FILE_SAVEDIB, 0&, SavePath)
    
End Function

Private Function ViewerFormat() As Boolean
    '---------------------------------------------------------------------
    '���ܣ���ʾͼ���ʽ
    '������
    '���أ�True = �ɹ� False = ʧ��
    '�ϼ���������̣�
    '�¼���������̣�SendMessageAsLong
    '���õ��ⲿ������hCapWnd
    '�����ˣ�����
    '�޸��ˣ�
    '---------------------------------------------------------------------
    ViewerFormat = SendMessageAsLong(hCapWnd, WM_CAP_DLG_VIDEOFORMAT, 0&, 0&)
    
End Function

Private Function ViewerSource() As Boolean
    '---------------------------------------------------------------------
    '���ܣ���ʾͼ����Դ
    '������
    '���أ�True = �ɹ� False = ʧ��
    '�ϼ���������̣�
    '�¼���������̣�SendMessageAsLong
    '���õ��ⲿ������hCapWnd
    '�����ˣ�����
    '�޸��ˣ�
    '---------------------------------------------------------------------
    ViewerSource = SendMessageAsLong(hCapWnd, WM_CAP_DLG_VIDEOSOURCE, 0&, 0&)
    
End Function

Public Function ShowEdit(ByVal frmMain As Object, ByRef lngKey As Long, Optional lng����id As Long = 0) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʾ�༭���壬������ô���Ľӿں���
    '����:  frmMain         ���ô������
    '       lngKey          ԤԼ�Ǽ�id
    '����:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
    mlng����id = lng����id
    mlngKey = lngKey
    Set mfrmMain = frmMain
        
    If InitData = False Then Exit Function
    If ReadData(mlngKey, lng����id) = False Then Exit Function

    Me.Show 1, frmMain
    
    ShowEdit = mblnOK

End Function


Private Function ReadData(ByVal lngKey As Long, ByVal lng����id As Long) As Boolean
     '------------------------------------------------------------------------------------------------------------------
    '����:  ��ȡ����
    '����:  lngKey      ����������
    '����:  True        ��ȡ�ɹ�
    '       False       ��ȡʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset

    On Error GoTo errHand
    
    gstrSQL = "SELECT Decode(C.����id,Null,0,0,0,1) AS ��Ƭ,A.����id AS ID,A.����,B.�����,B.�Ա�,TO_CHAR(B.��������,'yyyy-mm-dd') AS �������� " & _
                "FROM �����Ա���� A,������Ϣ B,������Ƭ C " & _
                "WHERE C.����id(+)=A.����id AND A.���״̬ IN (4,5) AND A.����id=B.����id and A.�Ǽ�id=[1]"
    If lng����id > 0 Then gstrSQL = gstrSQL & " AND B.����id=[2]"
    
    gstrSQL = gstrSQL & " Order By B.�����"
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, lng����id)
    If rs.BOF = False Then
        Call FillGrid(vsf, rs)
        Call AppendRows(vsf, lnX, lnY)
    End If
        
    
    
    ReadData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʼ������
    '����:  True        ��ʼ���ɹ�
    '       False       ��ʼ��ʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    Dim intLoop  As Integer
    Dim strTmp() As String
    
    On Error GoTo errHand
    
    strVsf = "��Ƭ,450,1,1,1,;����,1080,1,1,1,;�����,810,7,1,1,;�Ա�,600,1,1,1,;��������,990,1,1,1,"
    
    Call CreateVsf(vsf, strVsf)
    vsf.Cols = vsf.Cols + 1
    vsf.ColWidth(vsf.Cols - 1) = 15
    vsf.ColDataType(0) = flexDTBoolean
        
    Call AppendRows(vsf, lnX, lnY)
    
    '��ʼ���豸
    
    strTmp = Split(GetCapSureDevice(), ";")
    For intLoop = 0 To UBound(strTmp)
        cboDev.AddItem strTmp(intLoop)
    Next
    
    If cboDev.ListCount > 0 Then
        
        cboDev.ListIndex = 0
        Call ConnCapDevice(picFilm.hWnd, cboDev.ListIndex)
        
    End If
    
    InitData = True
    
    Exit Function

errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function SavePhoto(ByVal lng����id As Long, ByVal strFile As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ���没�˵���Ƭ����
    '����:  lng����id
    '       strFile        ������Ƭ�ļ�
    '����:  ����ɹ�����True;���򷵻�False
    '------------------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    blnTran = True
    gcnOracle.BeginTrans
    gcnOracle.Execute "Delete From ������Ƭ Where ����id=" & lng����id
    
    gstrSQL = "Select ����id,��Ƭ From ������Ƭ where ����id=" & lng����id
    
    rs.Open gstrSQL, gcnOracle, adOpenStatic, adLockOptimistic
    
    If rs.BOF Then
        
        If rs.EOF Then rs.AddNew
        
        rs("����id").Value = lng����id
        rs("��Ƭ").Value = Null
        rs.Update
        
        If zlDatabase.SavePicture(strFile, rs, "��Ƭ") = False Then
        
            ShowSimpleMsg "������Ƭ����,��ȷ���ļ��Ƿ�ɾ��!"
            
            gcnOracle.RollbackTrans
            blnTran = False
            
            Exit Function
            
        End If
        
        rs.Close
    End If
    
    gcnOracle.CommitTrans
    blnTran = False
    
    SavePhoto = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
    If blnTran Then gcnOracle.RollbackTrans
    
End Function

Private Sub cboDev_Click()
    If cboDev.ListIndex <> -1 Then Call SelectCapDevice(cboDev.ListIndex)
End Sub

Private Sub cmdClear_Click()
    Dim blnTran As Boolean
    
    On Error GoTo errHand
    picPhoto.Tag = ""
    picPhoto.Cls
    
    blnTran = True
    gcnOracle.BeginTrans
    gcnOracle.Execute "Delete From ������Ƭ Where ����id=" & Val(vsf.RowData(vsf.Row))
    gcnOracle.CommitTrans
    blnTran = False
    
    Exit Sub
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTran Then gcnOracle.RollbackTrans
End Sub

Private Sub cmdFile_Click()
    
    Dim objFile As New FileSystemObject
    Dim strTmp As String
    
    If Val(vsf.RowData(vsf.Row)) = 0 Then Exit Sub
    
    dlg.DialogTitle = "��ѡ��Ҫ��ӵ���Ƭ�ļ�"
    dlg.Filter = "ͼƬ(*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif"
    
    On Error Resume Next
    
    dlg.Flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
    dlg.FileName = ""
    dlg.MaxFileSize = 32767
    dlg.CancelError = True
    dlg.ShowOpen
    
    If Err.Number = 0 And dlg.FileName <> "" Then
        
        picPhoto.Tag = dlg.FileName
        
        On Error GoTo errHand
        
        Call DrawPicture(picPhoto, VB.LoadPicture(picPhoto.Tag), picPhoto.Width, picPhoto.Height)
        Call SavePhoto(Val(vsf.RowData(vsf.Row)), picPhoto.Tag)
        
        vsf.TextMatrix(vsf.Row, 0) = 1
    Else
        Err.Clear
    End If
    
    Exit Sub
    
errHand:
    ShowSimpleMsg "���ܴ��ļ�(" & picPhoto.Tag & "),���ļ���������ʹ�û��ļ�������!"
End Sub

Private Sub cmdLoad_Click()
    Dim strTmpFile As String
    
    On Error GoTo errHand
    
    If Val(vsf.RowData(vsf.Row)) = 0 Then Exit Sub
    
    strTmpFile = CreateTmpFile("bmp")
    
    Call SaveImageFile(strTmpFile)
    
    picPhoto.Tag = strTmpFile
    
    Call DrawPicture(picPhoto, VB.LoadPicture(picPhoto.Tag), picPhoto.Width, picPhoto.Height)
    Call SavePhoto(Val(vsf.RowData(vsf.Row)), picPhoto.Tag)
    
    vsf.TextMatrix(vsf.Row, 0) = 1
    Exit Sub
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Sub

Private Sub cmdSource_Click()
    Call ViewerSource
End Sub

Private Sub cmdStyle_Click()
    Call ViewerFormat
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
    Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
    
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim rs As New ADODB.Recordset
    Dim strTmpFile As String
    Dim objStd As IPictureDisp
    '������Ƭ
    If NewRow = OldRow Then Exit Sub
    
    picPhoto.Cls
    
    gstrSQL = "Select B.* From ������Ƭ B Where B.����id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(NewRow)))
    
    If rs.BOF = False Then
        strTmpFile = ""
        strTmpFile = ReadPicture(rs, "��Ƭ", strTmpFile)
        
        If strTmpFile <> "" Then
            Set objStd = VB.LoadPicture(strTmpFile)
            Call DrawPicture(picPhoto, objStd, objStd.Width, objStd.Height)
        End If
    End If
End Sub

