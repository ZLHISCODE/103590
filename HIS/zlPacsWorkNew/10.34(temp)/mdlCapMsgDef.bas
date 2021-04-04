Attribute VB_Name = "mdlCapMsgDef"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'基于VFW采集所要用到的消息定义
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Const WM_USER As Long = &H400
Public Const WM_CAP_START As Long = WM_USER

Public Const WM_CAP_GET_CAPSTREAMPTR As Long = WM_CAP_START + 1

Public Const WM_CAP_SET_CALLBACK_ERROR As Long = WM_CAP_START + 2
Public Const WM_CAP_SET_CALLBACK_STATUS As Long = WM_CAP_START + 3
Public Const WM_CAP_SET_CALLBACK_YIELD As Long = WM_CAP_START + 4
Public Const WM_CAP_SET_CALLBACK_FRAME As Long = WM_CAP_START + 5
Public Const WM_CAP_SET_CALLBACK_VIDEOSTREAM As Long = WM_CAP_START + 6
Public Const WM_CAP_SET_CALLBACK_WAVESTREAM As Long = WM_CAP_START + 7
Public Const WM_CAP_GET_USER_DATA As Long = WM_CAP_START + 8
Public Const WM_CAP_SET_USER_DATA As Long = WM_CAP_START + 9
    
Public Const WM_CAP_DRIVER_CONNECT As Long = WM_CAP_START + 10
Public Const WM_CAP_DRIVER_DISCONNECT As Long = WM_CAP_START + 11
Public Const WM_CAP_DRIVER_GET_NAME As Long = WM_CAP_START + 12
Public Const WM_CAP_DRIVER_GET_VERSION As Long = WM_CAP_START + 13
Public Const WM_CAP_DRIVER_GET_CAPS As Long = WM_CAP_START + 14

Public Const WM_CAP_FILE_SET_CAPTURE_FILE As Long = WM_CAP_START + 20
Public Const WM_CAP_FILE_GET_CAPTURE_FILE As Long = WM_CAP_START + 21
Public Const WM_CAP_FILE_ALLOCATE As Long = WM_CAP_START + 22
Public Const WM_CAP_FILE_SAVEAS As Long = WM_CAP_START + 23
Public Const WM_CAP_FILE_SET_INFOCHUNK As Long = WM_CAP_START + 24
Public Const WM_CAP_FILE_SAVEDIB As Long = WM_CAP_START + 25

Public Const WM_CAP_EDIT_COPY As Long = WM_CAP_START + 30

Public Const WM_CAP_SET_AUDIOFORMAT As Long = WM_CAP_START + 35
Public Const WM_CAP_GET_AUDIOFORMAT As Long = WM_CAP_START + 36

Public Const WM_CAP_DLG_VIDEOFORMAT As Long = WM_CAP_START + 41
Public Const WM_CAP_DLG_VIDEOSOURCE As Long = WM_CAP_START + 42
Public Const WM_CAP_DLG_VIDEODISPLAY As Long = WM_CAP_START + 43
Public Const WM_CAP_GET_VIDEOFORMAT As Long = WM_CAP_START + 44
Public Const WM_CAP_SET_VIDEOFORMAT As Long = WM_CAP_START + 45
Public Const WM_CAP_DLG_VIDEOCOMPRESSION As Long = WM_CAP_START + 46

Public Const WM_CAP_SET_PREVIEW As Long = WM_CAP_START + 50
Public Const WM_CAP_SET_OVERLAY As Long = WM_CAP_START + 51
Public Const WM_CAP_SET_PREVIEWRATE As Long = WM_CAP_START + 52
Public Const WM_CAP_SET_SCALE As Long = WM_CAP_START + 53
Public Const WM_CAP_GET_STATUS As Long = WM_CAP_START + 54
Public Const WM_CAP_SET_SCROLL As Long = WM_CAP_START + 55

Public Const WM_CAP_GRAB_FRAME As Long = WM_CAP_START + 60
Public Const WM_CAP_GRAB_FRAME_NOSTOP As Long = WM_CAP_START + 61

Public Const WM_CAP_SEQUENCE As Long = WM_CAP_START + 62
Public Const WM_CAP_SEQUENCE_NOFILE As Long = WM_CAP_START + 63
Public Const WM_CAP_SET_SEQUENCE_SETUP As Long = WM_CAP_START + 64
Public Const WM_CAP_GET_SEQUENCE_SETUP As Long = WM_CAP_START + 65
Public Const WM_CAP_SET_MCI_DEVICE As Long = WM_CAP_START + 66
Public Const WM_CAP_GET_MCI_DEVICE As Long = WM_CAP_START + 67
Public Const WM_CAP_STOP As Long = WM_CAP_START + 68
Public Const WM_CAP_ABORT As Long = WM_CAP_START + 69

Public Const WM_CAP_SINGLE_FRAME_OPEN As Long = WM_CAP_START + 70
Public Const WM_CAP_SINGLE_FRAME_CLOSE As Long = WM_CAP_START + 71
Public Const WM_CAP_SINGLE_FRAME As Long = WM_CAP_START + 72

Public Const WM_CAP_PAL_OPEN As Long = WM_CAP_START + 80
Public Const WM_CAP_PAL_SAVE As Long = WM_CAP_START + 81
Public Const WM_CAP_PAL_PASTE As Long = WM_CAP_START + 82
Public Const WM_CAP_PAL_AUTOCREATE As Long = WM_CAP_START + 83
Public Const WM_CAP_PAL_MANUALCREATE As Long = WM_CAP_START + 84

Public Const WM_CAP_SET_CALLBACK_CAPCONTROL As Long = WM_CAP_START + 85

Public Const WS_VISIBLE As Long = &H10000000
Public Const WS_CHILD = &H40000000

Public Const AVSTREAMMASTER_AUDIO As Long = 0  '/* Audio master (VFW 1.0, 1.1) */
Public Const AVSTREAMMASTER_NONE  As Long = 1  '/* No master */

Public Const INDEX_15_MINUTES As Long = 27000 '(30fps * 60sec * 15min)
Public Const INDEX_3_HOURS As Long = 324000 ' (30fps * 60sec * 60min * 3hr)

Public Const SWP_NOSIZE As Long = &H1&
Public Const SWP_NOMOVE As Long = &H2&
Public Const SWP_NOZORDER As Long = &H4&
Public Const SWP_NOSENDCHANGING As Long = &H400&   ' /* Don't send WM_WINDOWPOSCHANGING */
